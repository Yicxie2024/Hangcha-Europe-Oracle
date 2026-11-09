"""Separate Teams invoice-upload workspace for the Hangcha desktop GUI.

No automatic cloud write. Login -> verify target -> preview -> explicit approval
-> upload with no-overwrite policy. Independent of Excel writes and PDF archive.
"""
from __future__ import annotations
import json
import os
import queue
import threading
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

from teams_graph import GraphClient, TeamsError, invoice_candidates
from teams_store import TeamsLedger, sha_file
from teams_drag import DragSelectionError, prepare_files, validate_batch, expected_teams_folder

ROOT = Path(os.environ.get('LOCALAPPDATA', str(Path(__file__).resolve().parent))) / 'HangchaSalesAssistant'
SCOPES = ['Files.ReadWrite', 'User.Read']


class TeamsUploadWindow(tk.Toplevel):
    def __init__(self, main):
        super().__init__(main)
        self.main = main
        self.title('HCE Oracle | Teams 发票上传 · 拖拽模式')
        self.geometry('1250x790'); self.minsize(900, 560)
        self.config(bg='#F3F6FB')
        self.protocol('WM_DELETE_WINDOW', self._close)
        self.store = TeamsLedger(ROOT / 'teams_upload_history.sqlite3')
        self.cfgfile = ROOT / 'teams_upload_settings.json'
        self.profile = getattr(main, '_current_archive_profile', 'default')
        self.config_data = self._read_config()
        self.graph = None   # delegated access token held in RAM only
        self.auth_app = None
        self.account = ''
        self.target = None
        self.preview_day = None
        self.preview_hash = {}
        self.remote_names = {}
        self.busy = False
        self.events = queue.Queue()
        self.rows = []
        self.selected = set()
        self.drag_batch = None          # explicitly approved batch, not yet transferred
        self.last_drag = None           # Windows reports COPY, not cloud upload success
        self.drag_source_ready = False
        self.drag_available = False
        self.drag_error = ''
        self.stage_dirs = []
        self.transfer_mode = tk.StringVar(value='drag')
        self.by_iid = {}
        self.sort_key = 'invoice'
        self.descending = True
        self.status = tk.StringVar(value='未连接 · 不会修改 Teams 或 Excel')
        self.tenant = tk.StringVar(value=self.config_data.get('tenant', ''))
        self.client_id = tk.StringVar(value=self.config_data.get('client_id', ''))
        self.share_link = tk.StringVar(value=self.config_data.get('share_link', ''))
        self.override_link = tk.StringVar(value='')
        self.search = tk.StringVar(value='')
        self.filter = tk.StringVar(value='全部')
        self.only_new = tk.BooleanVar(value=True)
        self._build()
        self._setup_native_drag()
        self._switch_transfer_mode()
        self._load_candidates()
        self.after(180, self._poll)

    def _read_config(self):
        try:
            p=json.loads(self.cfgfile.read_text('utf-8'))
            return p if isinstance(p,dict) else {}
        except (FileNotFoundError,ValueError,OSError):
            return {}

    def _save_config(self):
        self.cfgfile.parent.mkdir(parents=True,exist_ok=True)
        cfg={'tenant':self.tenant.get().strip(),'client_id':self.client_id.get().strip(),
             'share_link':self.share_link.get().strip()}
        # Keep verified root ONLY if the authentication and link context matches.
        if all(cfg[k]==self.config_data.get(k) for k in ('tenant','client_id','share_link')):
            if 'base' in self.config_data:
                cfg['base']=self.config_data['base']
        else:
            self.target=None
        # Credential-free! No tokens/secrets are serialized to disk.
        self.cfgfile.write_text(json.dumps(cfg,ensure_ascii=False,indent=2),'utf-8')
        self.config_data=cfg

    def _build(self):
        top=tk.Frame(self,bg='#15375F');top.pack(fill='x')
        tk.Label(top,text='Teams 发票上传',font=('Microsoft YaHei UI',18,'bold'),bg='#15375F',fg='white').pack(side='left',padx=18,pady=12)
        tk.Label(top,text='与 Excel 业务写入独立 · 先预览，再确认上传',font=('Microsoft YaHei UI',10),bg='#15375F',fg='#D8E5F2').pack(side='left',padx=10)
        pane=tk.Frame(self,bg='#F3F6FB');pane.pack(fill='both',expand=True,padx=15,pady=10)
        mode=ttk.Frame(pane);mode.pack(fill='x',pady=(0,5))
        ttk.Label(mode,text='传送方式：').pack(side='left')
        ttk.Radiobutton(mode,text='拖拽到 Teams（无需 Entra）',variable=self.transfer_mode,value='drag',command=self._switch_transfer_mode).pack(side='left',padx=(4,16))
        ttk.Radiobutton(mode,text='Microsoft Graph API（企业授权）',variable=self.transfer_mode,value='graph',command=self._switch_transfer_mode).pack(side='left')
        ctrl=ttk.LabelFrame(pane,text='公司 SharePoint 连接 / 目标目录（仅 Graph API 模式）',padding=9)
        self.graph_ctrl=ctrl
        ctrl.pack(fill='x')
        r=ttk.Frame(ctrl);r.pack(fill='x',pady=2)
        ttk.Label(r,text='Tenant ID / 域：').pack(side='left')
        ttk.Entry(r,textvariable=self.tenant,width=35).pack(side='left',padx=4)
        ttk.Label(r,text='应用 Client ID：').pack(side='left',padx=(12,0))
        ttk.Entry(r,textvariable=self.client_id,width=37).pack(side='left',padx=4)
        self.login_button=ttk.Button(r,text='连接公司 Microsoft 365',command=self._login)
        self.login_button.pack(side='left',padx=10)
        r=ttk.Frame(ctrl);r.pack(fill='x',pady=4)
        ttk.Label(r,text='频道日期文件夹共享链接：').pack(side='left')
        ttk.Entry(r,textvariable=self.share_link).pack(side='left',fill='x',expand=True,padx=5)
        self.resolve_button=ttk.Button(r,text='验证目标根目录',command=self._resolve_base)
        self.resolve_button.pack(side='left')
        r=ttk.Frame(ctrl);r.pack(fill='x',pady=2)
        ttk.Label(r,text='本次手动目标（可选）：').pack(side='left')
        ttk.Entry(r,textvariable=self.override_link).pack(side='left',fill='x',expand=True,padx=5)
        ttk.Label(r,text='默认按电脑今天日期定位；手动路径只影响本次，不会永久覆盖设置。').pack(side='left')
        self.target_label=ttk.Label(ctrl,text='待验证：月份 → 当天日期 → 直接存 PDF；不自动创建或覆盖文件。')
        self.target_label.pack(anchor='w',pady=(5,0))

        self.drag_panel=ttk.LabelFrame(pane,text='把勾选的发票直接拖到 Teams 已打开的日期文件夹',padding=9)
        self.drag_panel.pack(fill='x',pady=(4,6))
        ttk.Label(self.drag_panel,text='① 在 Teams → 对应频道 → 当月 → 今天日期文件夹中打开文件列表。').pack(anchor='w',pady=2)
        ttk.Label(self.drag_panel,text='② 在下方勾选 PDF → 确认清单 → 按住鼠标左键拖动蓝色区域至 Teams。').pack(anchor='w',pady=2)
        self.drag_day_text=tk.StringVar(value='建议目标：'+expected_teams_folder(datetime.now().date()))
        ttk.Label(self.drag_panel,textvariable=self.drag_day_text,foreground='#225B98').pack(anchor='w',pady=(2,6))
        drow=ttk.Frame(self.drag_panel);drow.pack(fill='x')
        self.drag_prepare_button=ttk.Button(drow,text='准备勾选发票（需确认）',command=self._prepare_drag)
        self.drag_prepare_button.pack(side='left',padx=(0,8))
        self.drag_handle=tk.Label(drow,text='⇥  先勾选并准备发票，随后从这里拖入 Teams  ',
                                  bg='#B8C8D8',fg='#193A61',padx=15,pady=10,cursor='hand2',font=('Microsoft YaHei UI',10,'bold'))
        self.drag_handle.pack(side='left',fill='x',expand=True,padx=(0,8))
        ttk.Button(drow,text='备用：在资源管理器中拖拽',command=self._open_drag_staging).pack(side='left')
        drow2=ttk.Frame(self.drag_panel);drow2.pack(fill='x',pady=(6,0))
        self.drag_hint=tk.StringVar(value='尚未准备。Teams 中的实际目标由你手动决定，本软件无法核验拖入位置。')
        ttk.Label(drow2,textvariable=self.drag_hint,foreground='#6B7380').pack(side='left',fill='x',expand=True)
        self.drag_mark_button=ttk.Button(drow2,text='已在 Teams 核实 → 标记完成',command=self._mark_drag_verified,state='disabled')
        self.drag_mark_button.pack(side='right')

        bar=ttk.Frame(pane);bar.pack(fill='x',pady=(11,6))
        self.preview_button=ttk.Button(bar,text='检查云端 / 预览待上传',command=self._preview)
        self.preview_button.pack(side='left')
        self.refresh_button=ttk.Button(bar,text='刷新本地发票',command=self._load_candidates)
        self.refresh_button.pack(side='left',padx=6)
        self.upload_button=ttk.Button(bar,text='上传勾选发票（需确认）',command=self._upload_selected,state='disabled')
        self.upload_button.pack(side='left',padx=9)
        self.history_button=ttk.Button(bar,text='上传历史',command=self._show_history)
        self.history_button.pack(side='left')
        self.remote_folder_button=ttk.Button(bar,text='查看目标文件夹',command=self._open_remote_folder)
        self.remote_folder_button.pack(side='left',padx=7)
        ttk.Label(bar,text='搜索：').pack(side='left',padx=(12,0))
        ttk.Entry(bar,textvariable=self.search,width=21).pack(side='left',padx=4)
        ttk.Combobox(bar,textvariable=self.filter,state='readonly',values=('全部','待上传','需复核','远端同名','已上传'),width=10).pack(side='left')
        self.search.trace_add('write',lambda *_:self._refresh_table())
        self.filter.trace_add('write',lambda *_:self._refresh_table())

        summary=ttk.Frame(pane);summary.pack(fill='x',pady=(0,5))
        self.stats=ttk.Label(summary,text='发票 0 · 待上传 0 · 已上传 0 · 需复核 0')
        self.stats.pack(side='left')
        ttk.Button(summary,text='全选当前可上传',command=self._select_all).pack(side='right',padx=5)
        ttk.Button(summary,text='取消勾选',command=self._deselect).pack(side='right')
        cols=('check','customer','contract','invoice','day','gross','state','file')
        self.table=ttk.Treeview(pane,columns=cols,show='headings',selectmode='browse')
        cols_labels=[('check','☐ 全选',72),('customer','客户',140),('contract','合同号',110),('invoice','发票号',135),('day','Invoice date',112),('gross','金额',95),('state','上传状态',170),('file','文件名',330)]
        for k,label,width in cols_labels:
            self.table.heading(k,text=label,command=self._toggle_select_all_visible if k=='check' else lambda x=k:self._sort(x))
            self.table.column(k,width=width,minwidth=35,stretch=k in ('customer','file'))
        self.table.pack(fill='both',expand=True,side='left')
        scroll=ttk.Scrollbar(pane,orient='vertical',command=self.table.yview)
        scroll.pack(side='right',fill='y');self.table.configure(yscrollcommand=scroll.set)
        self.table.tag_configure('warn',background='#FFF3D9')
        self.table.tag_configure('uploaded',background='#E4F2E9')
        self.table.bind('<Button-1>',self._on_table_click)
        self.table.bind('<Button-3>',self._on_right_click)
        self.table.bind('<Double-1>',self._on_double)
        self.menu=tk.Menu(self,tearoff=False)
        self.menu.add_command(label='上传此发票（需确认）',command=self._upload_current)
        self.menu.add_command(label='准备拖拽此发票到 Teams',command=self._prepare_current_drag)
        self.menu.add_separator()
        self.menu.add_command(label='打开 PDF',command=self._open_current)
        self.menu.add_command(label='在文件夹中显示',command=self._reveal_current)
        self.menu.add_command(label='查看 Teams 已上传文件',command=self._open_uploaded)
        foot=ttk.Frame(self);foot.pack(fill='x',padx=15,pady=(0,10))
        ttk.Label(foot,textvariable=self.status).pack(side='left',fill='x',expand=True)
        self.progress=ttk.Progressbar(foot,mode='indeterminate',length=175)
        self.progress.pack(side='right')

    def _switch_transfer_mode(self):
        graph=self.transfer_mode.get()=='graph'
        if graph:
            self.drag_panel.pack_forget()
            self.graph_ctrl.pack(fill='x',before=self.preview_button.master)
            self.preview_button.pack(side='left',before=self.refresh_button)
            self.upload_button.pack(side='left',after=self.refresh_button,padx=9)
            self.remote_folder_button.pack(side='left',after=self.history_button,padx=7)
        else:
            self.graph_ctrl.pack_forget()
            self.drag_panel.pack(fill='x',before=self.preview_button.master,pady=(4,6))
            self.preview_button.pack_forget()
            self.upload_button.pack_forget()
            self.remote_folder_button.pack_forget()
        for btn in (self.login_button,self.resolve_button,self.preview_button):
            btn.configure(state='normal' if graph else 'disabled')
        self.upload_button.configure(state='normal' if graph and self.graph and self.target else 'disabled')
        self._invalidate_drag()
        if not graph:
            self.status.set('拖拽模式：在 Teams 中手动打开目标日期文件夹；完成拖拽后请核实云端文件。')

    def _setup_native_drag(self):
        if os.name!='nt':
            self.drag_error='跨应用文件拖拽需要 Windows；当前仅支持备用的资源管理器方式。'
            self.drag_hint.set(self.drag_error)
            return
        try:
            from tkinterdnd2 import TkinterDnD
            # Supports both the current public require() and older _require().
            require=getattr(TkinterDnD,'require',None) or getattr(TkinterDnD,'_require',None)
            if require is None:
                raise RuntimeError('tkinterdnd2 不提供 TkDND 初始化接口')
            require(self.main)
            self.drag_available=True
        except Exception as exc:
            self.drag_error=str(exc)
            self.drag_hint.set('原生拖拽暂不可用：'+str(exc)+'。请运行 01_INSTALL.bat，或使用备用资源管理器方式。')

    def _invalidate_drag(self):
        self.drag_batch=None
        self.last_drag=None
        self.drag_source_ready=False
        if hasattr(self,'drag_handle'):
            try:self.drag_handle.drag_source_unregister()
            except (AttributeError,tk.TclError):pass
            self.drag_handle.configure(bg='#B8C8D8',text='⇥  先勾选并准备发票，随后从这里拖入 Teams')
        if hasattr(self,'drag_mark_button'):
            self.drag_mark_button.configure(state='disabled')
        if hasattr(self,'drag_hint') and self.transfer_mode.get()=='drag':
            self.drag_hint.set('请选择发票，并确认本次拖拽清单。不会自动修改 Excel 或上传状态。')

    def _prepare_current_drag(self):
        row=self._current()
        if row is None:return
        if self._state(row)!='待上传':
            messagebox.showwarning('不可拖拽','此发票需先复核或已有上传记录。',parent=self);return
        self.selected={row.source}
        self._invalidate_drag()
        self._refresh_table()
        self._prepare_drag()

    def _prepare_drag(self):
        if self.transfer_mode.get()!='drag':return
        records=[(r,self._state(r)) for r in self.rows if r.source in self.selected]
        try:
            batch=prepare_files(records,datetime.now().date())
        except DragSelectionError as exc:
            messagebox.showwarning('无法准备拖拽',str(exc),parent=self);return
        names='\n'.join('• '+Path(x).name for x in batch.paths[:12])
        if len(batch.paths)>12:names+=f'\n…… 另有 {len(batch.paths)-12} 份'
        explanation=(f'本次将提供 {batch.size} 个原始 PDF 供你拖入 Teams。\n\n'
                    f'应在 Teams 手动打开目标：Rechnung-Ersatzteile Garantie / {batch.label}\n'
                    '实际上传位置完全由你在 Teams 打开的文件夹决定，软件无法自动检查。\n\n'
                    f'文件清单：\n{names}\n\n'
                    '拖拽仅提供文件副本；不会删除原始 PDF。\n'
                    '拖拽结束不等于云端上传成功，请在 Teams 核实后再手动标记。\n\n'
                    '确认准备这批文件吗？')
        if not messagebox.askyesno('确认准备发票拖拽',explanation,parent=self,icon='warning'):
            return
        special=[r for r,_ in records if '有误' in Path(r.source).name or r.status=='待复核']
        if special and not messagebox.askyesno('特殊文件二次确认',
                f'所选发票中有 {len(special)} 张带「有误」或曾处于待复核状态。\n'
                '请先打开原始 PDF 核查内容，确认仍要拖给 Teams 吗？',
                parent=self,icon='warning'):
            return
        self._invalidate_drag()
        self.drag_batch=batch
        self.drag_day_text.set('本次应拖入：'+batch.label+'  （电脑当天日期；请在 Teams 亲自确认）')
        if self.drag_available:
            try:
                from tkinterdnd2 import DND_FILES
                self.drag_handle.drag_source_register(1,DND_FILES)
                self.drag_handle.dnd_bind('<<DragInitCmd>>',self._drag_init)
                self.drag_handle.dnd_bind('<<DragEndCmd>>',self._drag_end)
                self.drag_source_ready=True
                self.drag_handle.configure(bg='#BFE9CF',text=f'⇥  按住这里，将 {batch.size} 个 PDF 拖到 Teams')
                self.drag_hint.set('已准备：直接按住绿色区域拖到 Teams 文件列表；暂不会标记为已上传。')
            except Exception as exc:
                self.drag_error=str(exc)
                self.drag_hint.set('TkDND 初始化失败，可用「备用：在资源管理器中拖拽」：'+str(exc))
        else:
            self.drag_hint.set('拖拽扩展未准备好。请使用右侧备用资源管理器方式。'+self.drag_error)
        self.status.set(f'已准备 {batch.size} 张发票，目标由用户在 Teams 中选定：{batch.label}')

    def _drag_init(self,event):
        from tkinterdnd2 import COPY,DND_FILES
        try:
            if not self.drag_batch or not self.drag_source_ready:
                raise DragSelectionError('尚未确认拖拽清单')
            validate_batch(self.drag_batch,datetime.now().date())
        except DragSelectionError as exc:
            self.status.set('取消拖拽：'+str(exc))
            return (COPY,DND_FILES,())
        return (COPY,DND_FILES,tuple(self.drag_batch.paths))

    def _drag_end(self,event):
        # OLE COPY only means the target accepted the drop, NOT Teams cloud success.
        action=str(getattr(event,'action','')).lower()
        if action=='copy' and self.drag_batch:
            self.last_drag=self.drag_batch
            self.drag_hint.set('Teams 已接收拖拽动作；请等待 Teams 上传完成并核实文件，再点「已核实」。')
            self.drag_mark_button.configure(state='normal')
            self.status.set('已拖出 PDF，但还没有验证云端。请在 Teams 确认文件存在。')
        else:
            self.drag_hint.set('拖拽被取消或目标未接收文件；可再次拖动，未标记为已上传。')
        # Do not unregister until after Tcl's callback finishes.

    def _mark_drag_verified(self):
        batch=self.last_drag
        if batch is None:return
        try:validate_batch(batch,datetime.now().date())
        except DragSelectionError as exc:
            messagebox.showwarning('无法标记',str(exc),parent=self);return
        if not messagebox.askyesno('人工确认 Teams 已上传',
                f'请确认你已经在公司 Teams 正确的日期文件夹中看到这 {batch.size} 个 PDF，'
                '且 Teams 显示上传已经完成。\n\n'
                '软件未能直接核验 SharePoint，这将被记录为「人工确认已上传」。\n'
                '如果尚未核实，请点击「否」，待核实后再操作。',parent=self,icon='warning'):
            return
        try:
            rows={str(Path(r.source).resolve()):r for r in self.rows}
            selected_rows=[]
            for p in batch.paths:
                row=rows.get(p)
                if row is None:raise DragSelectionError('发票列表已改变，请刷新并重新准备')
                if self._state(row)!='待上传':raise DragSelectionError(f'记录状态已改变：{Path(p).name}')
                selected_rows.append(row)
            self.store.record_manual_batch(self.profile, selected_rows,
                 expected_teams_folder(batch.prepared_on), batch.fingerprints)
        except Exception as exc:
            messagebox.showerror('记录上传状态失败',str(exc),parent=self);return
        count=batch.size
        self.selected.clear();self._invalidate_drag();self._refresh_table()
        self.status.set(f'已人工确认 {count} 张发票在 Teams 上传完成（非 API 自动验证）。')
        messagebox.showinfo('标记完成',f'已人工确认并记录 {count} 张发票；下次扫描不会重复进入待上传列表。',parent=self)

    def _open_drag_staging(self):
        batch=self.drag_batch
        if batch is None:
            messagebox.showinfo('请先准备','先勾选发票，点击「准备勾选发票（需确认）」。',parent=self);return
        try:validate_batch(batch,datetime.now().date())
        except DragSelectionError as exc:
            messagebox.showwarning('文件已变化',str(exc),parent=self);return
        import shutil, tempfile
        base=ROOT/'drag_staging'
        base.mkdir(parents=True,exist_ok=True)
        folder=Path(tempfile.mkdtemp(prefix='drag_',dir=base))
        self.stage_dirs.append(folder)
        try:
            for path in batch.paths:
                shutil.copy2(path,folder/Path(path).name)
        except Exception:
            shutil.rmtree(folder,ignore_errors=True)
            raise
        self.last_drag=batch   # User must still verify that Teams actually uploaded it.
        self.drag_mark_button.configure(state='normal')
        if os.name=='nt':os.startfile(str(folder))
        else:webbrowser.open(folder.as_uri())
        self.status.set(f'已打开临时文件夹（{batch.size} 个副本）；Ctrl+A 选择全部，拖入 Teams 日期文件夹。')
        self.drag_hint.set('备用方式已打开资源管理器；这些是临时副本，不会更改原始 PDF。请手动在 Teams 核查。')

    def _load_candidates(self):
        if getattr(self.main,'worker',None) and self.main.worker.is_alive():
            messagebox.showwarning('请稍后','订单管理正在扫描 PDF，请等待扫描完成后刷新。',parent=self)
            return
        self.profile=getattr(self.main,'_current_archive_profile','default')
        self.rows=invoice_candidates(self.main.rows,self.store,self.profile)
        from collections import Counter
        counts=Counter((r.contract,r.invoice) for r in self.rows)
        self._duplicate_invoice_keys={k for k,n in counts.items() if n>1}
        self.selected.clear()
        self._invalidate_drag()
        self.preview_day=None;self.target=None;self.preview_hash={};self.remote_names={}
        self._refresh_table()
        uploaded=sum(self._state(r)=='已上传' for r in self.rows)
        pending=sum(self._state(r)=='待上传' for r in self.rows)
        review=len(self.rows)-uploaded-pending
        self.status.set(f'本次扫描发票 {len(self.rows)} 张：待上传 {pending} · 已上传 {uploaded} · 需核对 {review}。' +
                        ('在 Teams 拖拽完成并核实后，请点击「标记完成」。' if self.transfer_mode.get()=='drag'
                         else 'Graph 模式上传前请先连接并预览目标文件夹。'))
        self.upload_button.configure(state='disabled')

    def _state(self,row):
        # Confirmed uploads remain visible even when another local copy shares
        # the same invoice ID. Unconfirmed duplicates still require review.
        if self.store.has_uploaded(self.profile,row.source,row.sha):
            return '已上传'
        if (row.contract,row.invoice) in getattr(self,'_duplicate_invoice_keys',set()):
            return '需复核'
        if row.status == '待复核':
            archive=getattr(self.main,'archive_store',None)
            try: reviewed=bool(archive and archive.is_reviewed(row))
            except Exception: reviewed=False
            if not reviewed:
                return '需复核'
        if Path(row.source).name.casefold() in self.remote_names:
            return '远端同名'
        return '待上传'

    def _refresh_table(self):
        self.table.delete(*self.table.get_children())
        self.by_iid={}
        q=self.search.get().casefold().strip();f=self.filter.get()
        items=[]
        for row in self.rows:
            state=self._state(row)
            if f!='全部' and f!=state: continue
            if q and q not in (' '.join([row.customer,row.contract,row.invoice,row.source])).casefold():continue
            items.append((row,state))
        def sorter(elem):
            row,state=elem
            if self.sort_key=='gross':
                try:return float(row.gross)
                except:return -1
            if self.sort_key=='day':
                try:return datetime.strptime(row.day,'%d.%m.%Y')
                except:return datetime.min
            if self.sort_key=='state':return state
            return str(getattr(row,self.sort_key,'') if self.sort_key!='file' else Path(row.source).name).casefold()
        items.sort(key=sorter,reverse=self.descending)
        for i,(r,state) in enumerate(items):
            iid=f'file_{i}'; self.by_iid[iid]=r
            self.table.insert('', 'end', iid=iid,
                    values=('☑' if r.source in self.selected else '☐',r.display_customer,r.contract,r.invoice,r.day,r.gross,state,Path(r.source).name),
                    tags=('warn' if state in ('需复核','远端同名') else 'uploaded' if state=='已上传' else '',))
        ready=sum(self._state(r)=='待上传' for r in self.rows)
        review=sum(self._state(r) in ('需复核','远端同名') for r in self.rows)
        uploaded=sum(self._state(r)=='已上传' for r in self.rows)
        # Reflect selection state of the currently visible, uploadable rows.
        visible_ready={r.source for r in self.by_iid.values() if self._state(r)=='待上传'}
        selected_ready=len(visible_ready & self.selected)
        if visible_ready and selected_ready==len(visible_ready):
            header='☑ 全选'
        elif selected_ready:
            header='▣ 全选'
        else:
            header='☐ 全选'
        self.table.heading('check',text=header)
        self.stats.configure(text=f'本地发票 {len(self.rows)} · 待上传 {ready} · 已上传 {uploaded} · 需核对 {review} · 已勾选 {len(self.selected)}')

    def _sort(self,k):
        if k=='check': return
        self.descending=not self.descending if self.sort_key==k else True
        self.sort_key=k;self._refresh_table()

    def _on_table_click(self,event):
        if self.table.identify('region',event.x,event.y)!='cell':return
        column=self.table.identify_column(event.x)
        iid=self.table.identify_row(event.y)
        r=self.by_iid.get(iid)
        if not r:return
        if column=='#1':
            if r.source in self.selected:self.selected.remove(r.source)
            else:self.selected.add(r.source)
            self._invalidate_drag()
            self._refresh_table()
            return 'break'
        if column in ('#3','#4'):
            # This window has no Excel remark column; copy the complete
            # contract/invoice number directly from the scan row.
            value=(r.contract if column=='#3' else r.invoice).strip()
            if not value:return
            try:
                self.clipboard_clear()
                self.clipboard_append(value)
            except tk.TclError:
                return
            self.table.selection_set(iid)
            self.table.focus(iid)
            self.status.set(f'已复制{"合同号" if column=="#3" else "发票号"}：{value}')
            return 'break'

    def _toggle_select_all_visible(self):
        """Checkbox heading toggles only visible invoices safe to upload."""
        paths={r.source for r in self.by_iid.values() if self._state(r)=='待上传'}
        if not paths:
            self.status.set('当前筛选结果中没有可勾选的待上传发票。')
            return
        if paths.issubset(self.selected):
            self.selected.difference_update(paths)
        else:
            self.selected.update(paths)
        self._invalidate_drag()
        self._refresh_table()

    def _select_all(self):
        self.selected.update(r.source for r in self.by_iid.values() if self._state(r)=='待上传')
        self._invalidate_drag()
        self._refresh_table()

    def _deselect(self):
        self.selected.clear();self._invalidate_drag();self._refresh_table()

    def _current(self):
        s=self.table.selection()
        return self.by_iid.get(s[0]) if s else None

    def _on_right_click(self,event):
        if self.table.identify('region',event.x,event.y)!='cell':return
        iid=self.table.identify_row(event.y)
        if iid not in self.by_iid:return
        self.table.selection_set(iid)
        row=self.by_iid[iid]
        self.menu.entryconfigure('上传此发票（需确认）',state='normal' if self._state(row)=='待上传' and self.target and self.graph and self.transfer_mode.get()=='graph' else 'disabled')
        self.menu.entryconfigure('准备拖拽此发票到 Teams',state='normal' if self._state(row)=='待上传' and self.transfer_mode.get()=='drag' else 'disabled')
        self.menu.entryconfigure('查看 Teams 已上传文件',state='normal' if self._uploaded_url(row) else 'disabled')
        try:self.menu.tk_popup(event.x_root,event.y_root)
        finally:self.menu.grab_release()

    def _on_double(self,event):
        if self.table.identify('region',event.x,event.y)=='cell':self._open_current()

    def _open_current(self):
        r=self._current()
        if not r:return
        try:
            if os.name=='nt':os.startfile(r.source)
            else:webbrowser.open(Path(r.source).as_uri())
        except Exception as exc:messagebox.showerror('无法打开 PDF',str(exc),parent=self)

    def _reveal_current(self):
        r=self._current()
        if r:self._reveal(r.source)

    def _reveal(self,source):
        try:
            from gui import _reveal_windows_pdf
            if os.name=='nt':_reveal_windows_pdf(Path(source))
            else:webbrowser.open(Path(source).parent.as_uri())
        except Exception as exc:
            messagebox.showerror('文件定位失败',str(exc),parent=self)

    def _uploaded_url(self,row):
        try:
            source=str(Path(row.source).resolve())
            return next((x['url'] for x in self.store.list_history(self.profile,limit=1000)
                         if x['invoice']==row.invoice and x['contract']==row.contract and x['url']), '')
        except Exception:return ''

    def _open_uploaded(self):
        r=self._current();url=self._uploaded_url(r) if r else ''
        if url:webbrowser.open(url)

    def _open_remote_folder(self):
        if self.target and self.target.get('folder_url'):webbrowser.open(self.target['folder_url'])
        else:messagebox.showinfo('尚未验证','请先登录并预览确认今日目标文件夹。',parent=self)

    def _submit(self,label,work):
        if self.busy:return
        self.busy=True;self.status.set(label);self.progress.start(15)
        for btn in (self.login_button,self.resolve_button,self.preview_button,self.upload_button):
            btn.configure(state='disabled')
        def run():
            try:self.events.put(('ok',work()))
            except Exception as e:self.events.put(('error',str(e)))
        threading.Thread(target=run,daemon=True).start()

    def _login(self):
        tenant=self.tenant.get().strip()
        clientid=self.client_id.get().strip()
        if not tenant or not clientid:
            messagebox.showwarning('需要公司授权','请输入公司 IT 批准的 Microsoft Entra Tenant ID 与应用 Client ID。\n不需要也不应填写客户端密钥或账户密码。',parent=self);return
        self._save_config()
        def task():
            try:import msal
            except ImportError:raise TeamsError('缺少 msal 库，请先运行 01_INSTALL.bat 更新依赖。')
            auth=msal.PublicClientApplication(clientid,authority='https://login.microsoftonline.com/'+tenant)
            result=auth.acquire_token_interactive(scopes=SCOPES,port=0)
            if not result.get('access_token'):
                raise TeamsError('登录或授权失败：'+str(result.get('error_description',result.get('error','未知错误'))))
            graph=GraphClient(result['access_token'])
            user=graph.call('GET','/me')
            return ('login',graph,auth,user.get('userPrincipalName',user.get('displayName','Microsoft 365 用户')))
        self._submit('请在浏览器中完成公司 Microsoft 365 登录…',task)

    def _resolve_base(self):
        if not self.graph:
            messagebox.showwarning('尚未登录','先连接公司 Microsoft 365。',parent=self);return
        link=self.share_link.get().strip()
        if not link:messagebox.showwarning('缺少链接','请粘贴一个日期文件夹或月份文件夹的 SharePoint :f: 链接。',parent=self);return
        def task():
            result=self.graph.resolve_base(link)
            result['configured_year']=datetime.now().year
            return ('root',result)
        self._submit('正在核验共享链接对应的真实文档库…',task)

    def _preview(self):
        if not self.graph:
            messagebox.showwarning('未连接','先通过 Microsoft 365 登录并获得授权。',parent=self);return
        self._save_config()
        if self.config_data.get('base') is None:
            messagebox.showwarning('尚未验证根目录','请先点击「验证目标根目录」。',parent=self);return
        day=datetime.now().date()
        manual=self.override_link.get().strip()
        base=dict(self.config_data['base'])
        def task():
            target=self.graph.explicit_target(manual,day) if manual else self.graph.target_today(base,day)
            # Scan the complete destination once; case-insensitive filename conflicts
            remote={x.get('name','').casefold():x for x in self.graph.children(target['drive_id'],target['folder_id']) if x.get('file') is not None}
            hashes={r.source:sha_file(r.source) for r in self.rows if Path(r.source).is_file()}
            return ('preview',target,remote,hashes)
        self._submit('正在检查 SharePoint 目标文件夹和远程重名文件…',task)

    def _upload_current(self):
        r=self._current()
        if r:self._upload([r])

    def _upload_selected(self):
        picked=[r for r in self.rows if r.source in self.selected]
        if not picked:
            messagebox.showinfo('未勾选','请先在发票列表中勾选需要上传的文件。',parent=self);return
        self._upload(picked)

    def _approval(self,files,target):
        d=tk.Toplevel(self);d.title('确认上传至公司 Teams');d.geometry('800x510');d.transient(self);d.grab_set()
        ok={'value':False}
        ttk.Label(d,text=f'将上传 {len(files)} 份 PDF。请核对目标路径和完整文件清单：').pack(anchor='w',padx=12,pady=9)
        box=scrolledtext.ScrolledText(d,wrap='word',font=('Consolas',10))
        box.pack(fill='both',expand=True,padx=12)
        text=f'目标：{target["label"]}\n上传日期（电脑当地日期）：{target["day"]}\n'
        text+=('注意：已手动改选目标文件夹，本次仅对该目录有效。\n' if target.get('manual') else '')
        text+='\n==== 上传清单 ====\n'
        text+='\n'.join(f'{i+1}. {r.invoice} / {r.contract} / {Path(r.source).name}\n   {r.source}' for i,r in enumerate(files))
        text+='\n\n不覆盖云端同名文件；不移动或删除本地 PDF；与 Excel 写入无关。'
        box.insert('1.0',text);box.configure(state='disabled')
        bottom=ttk.Frame(d);bottom.pack(fill='x',padx=12,pady=10)
        ttk.Button(bottom,text='取消',command=d.destroy).pack(side='right',padx=4)
        def yes():ok['value']=True;d.destroy()
        ttk.Button(bottom,text='确认上传',command=yes).pack(side='right')
        self.wait_window(d)
        return ok['value']

    def _upload(self,files):
        if self.transfer_mode.get()!='graph':
            messagebox.showinfo('当前为拖拽模式','请勾选发票并点击「准备勾选发票」，随后拖进 Teams。',parent=self);return
        if not self.graph or not self.target or not self.preview_day:
            messagebox.showwarning('请先预览','请先登录、验证目标文件夹并执行云端预览。',parent=self);return
        if self.preview_day != datetime.now().date():
            self.target=None;self.upload_button.configure(state='disabled')
            messagebox.showwarning('日期已变化','电脑日期已变化。请重新预览今日目标文件夹。',parent=self);return
        if any(self.config_data.get(k,'')!=v.get().strip() for k,v in [('tenant',self.tenant),('client_id',self.client_id),('share_link',self.share_link)]):
            messagebox.showwarning('配置已变化','公司授权参数或基础目录链接已变化，请重新验证目标。',parent=self);return
        if self.override_link.get().strip() != getattr(self, '_preview_manual_link', ''):
            messagebox.showwarning('目标已变化','手动目标链接在预览后被修改，请重新预览。',parent=self);return
        if any(self._state(r)!='待上传' for r in files):
            messagebox.showwarning('不能上传','含有待复核、云端同名或已上传的文件，请先处理这些项目。',parent=self);return
        try:
            for r in files:
                if sha_file(r.source)!=self.preview_hash.get(r.source):
                    raise TeamsError(f'PDF 在云端预览后被修改：{r.source}')
        except (OSError,TeamsError) as exc:
            messagebox.showwarning('源文件已变化',str(exc),parent=self);return
        if not self._approval(files,self.target):return
        if any('有误' in Path(r.source).name or r.status=='待复核' for r in files):
            if not messagebox.askyesno('特殊发票二次确认',
                  '所选文件包含「有误」或已人工标记复核的发票。\n这些发票可能仍有金额或数据问题。\n请确认你已经打开并检查过 PDF，确定仍需上传 Teams。',parent=self,icon='warning'):
                return
        target=dict(self.target);graph=self.graph;profile=self.profile;hashes=dict(self.preview_hash)
        def task():
            if datetime.now().date().isoformat()!=target['day']:
                raise TeamsError('执行前电脑日期已变化，未上传。请重新选择当天目标目录。')
            # Verify target folder identity right before uploading, even if it
            # was manually overridden or the day folder was renamed.
            folder=graph.item(target['drive_id'],target['folder_id'])
            if 'folder' not in folder:
                raise TeamsError('目标文件夹不再有效。')
            results=[]
            for r in files:
                if datetime.now().date().isoformat()!=target['day']:
                    results.append((r,'未上传：跨过午夜，须重新预览当天目标文件夹'))
                    break
                if sha_file(r.source)!=hashes[r.source]:
                    results.append((r,'失败：PDF 内容已改变'));continue
                try:
                    # A fresh cloud duplicate check AND conflictBehavior=fail
                    # handle concurrent uploads by coworkers without overwrite.
                    if graph.remote_file(target['drive_id'],target['folder_id'],Path(r.source).name):
                        outcome='冲突：远端已有同名 PDF，未覆盖'
                        self.store.log_attempt(profile,r,target,outcome)
                        results.append((r,outcome));continue
                    item=graph.upload_no_overwrite(Path(r.source),target)
                    self.store.record(profile,r,item,target)
                    self.store.log_attempt(profile,r,target,'上传成功')
                    results.append((r,'上传成功'))
                except Exception as e:
                    outcome='失败：'+str(e)
                    try:self.store.log_attempt(profile,r,target,outcome)
                    except Exception:pass
                    results.append((r,outcome))
            return ('upload',results)
        self._submit('正在上传已确认的发票…',task)

    def _show_history(self):
        items=self.store.list_history(self.profile)
        attempts=self.store.list_attempts(self.profile)
        d=tk.Toplevel(self);d.title('Teams 上传历史 / 失败记录');d.geometry('870x490')
        box=scrolledtext.ScrolledText(d,wrap='word');box.pack(fill='both',expand=True,padx=10,pady=10)
        lines=['=== 最近尝试 / 失败记录 ===']
        lines += [f'{x["attempted_at"]} | {x["invoice"]} | {x["outcome"]}\n{x["target"]} | {x["source"]}' for x in attempts]
        lines += ['\n=== 成功上传文件 ===']
        lines += [f'{x["uploaded_at"]}  {x["invoice"]}  {x["name"]} | '+ ('人工确认（非 API 核验）' if x['drive_id']=='user-confirmed' else 'Graph 上传确认')+f'\n{x["url"]}' for x in items]
        box.insert('1.0','\n'.join(lines) if (items or attempts) else '暂无上传历史')
        box.configure(state='disabled')

    def _poll(self):
        try:
            while True:
                kind,payload=self.events.get_nowait()
                self.busy=False;self.progress.stop()
                for btn in (self.login_button,self.resolve_button,self.preview_button):btn.configure(state='normal')
                self.upload_button.configure(state='normal' if self.transfer_mode.get()=='graph' and self.graph and self.target else 'disabled')
                if kind=='error':
                    self.status.set('操作未完成：'+payload[:220]);messagebox.showerror('Teams 操作失败',payload,parent=self);continue
                action,*args=payload
                if action=='login':
                    self.graph,self.auth_app,self.account=args
                    self.target=None;self.config_data.pop('base',None)
                    self.status.set('已登录：'+self.account+'。请验证目标 SharePoint 根目录。')
                elif action=='root':
                    base=args[0];self.config_data['base']=base
                    self.config_data.update({'tenant':self.tenant.get().strip(),'client_id':self.client_id.get().strip(),
                                             'share_link':self.share_link.get().strip()})
                    self.cfgfile.parent.mkdir(parents=True,exist_ok=True)
                    self.cfgfile.write_text(json.dumps(self.config_data,ensure_ascii=False,indent=2),'utf-8')
                    self.target_label.configure(text='已定位根目录：'+base['root_name']+'  |  文档库 '+base['drive_id'][:18]+'…')
                    self.status.set('目标根目录已验证。请点击「检查云端 / 预览待上传」。')
                elif action=='preview':
                    self.target,self.remote_names,self.preview_hash=args
                    self._preview_manual_link=self.override_link.get().strip()
                    self.preview_day=__import__('datetime').date.fromisoformat(self.target['day'])
                    self.status.set(f'云端预览完成 · {self.target["label"]} · 与 Excel 相互独立')
                    self.target_label.configure(text='本次目标：'+self.target['label']+('（手动选择）' if self.target.get('manual') else '（自动）'))
                    self.upload_button.configure(state='normal' if self.transfer_mode.get()=='graph' else 'disabled')
                    self._refresh_table()
                elif action=='upload':
                    results=args[0];self.selected.clear();self._refresh_table()
                    success=sum(msg=='上传成功' for _,msg in results)
                    warnings=[Path(r.source).name+'：'+msg for r,msg in results if msg!='上传成功']
                    self.status.set(f'上传完成 {success}/{len(results)}；失败/冲突 {len(warnings)} · 不影响 Excel')
                    messagebox.showinfo('上传执行结果',f'上传成功：{success}/{len(results)}\n'+('\n'.join(warnings[:12]) if warnings else '全部成功'),parent=self)
        except queue.Empty:pass
        if self.winfo_exists():self.after(160,self._poll)

    def _close(self):
        # Worker has its own immutable snapshot; shutting window while active
        # cannot abort in-flight uploads. Keep the window until they finish.
        if self.busy:
            messagebox.showwarning('正在执行','请在当前操作完成后关闭窗口。',parent=self);return
        # Only remove disposable temporary copies created for Explorer fallback.
        import shutil
        for folder in self.stage_dirs:
            try:shutil.rmtree(folder,ignore_errors=True)
            except OSError:pass
        self.destroy()
