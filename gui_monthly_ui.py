"""Monthly internal shipping statistics: scan -> review -> export -> AI merge.

Source workbook and PDFs are read-only. Manual decisions are stored under
LOCALAPPDATA and bound to the exact weekly-row/PDF fingerprint.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import hashlib
import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk

from gui_monthly_summary import scan_month_for_review
from gui_monthly_review import (ReviewEntry, apply_decisions, outcome_rows,
                                selected_invoices, store_decision)
from gui_monthly_xlsx import (generate_report, build_prompt, import_ai,
    update_approved_library, export_clean_report)

BASE=Path(__file__).resolve().parent
SETTINGS=Path(os.environ.get('LOCALAPPDATA',str(BASE))) / 'HangchaSalesAssistant' / 'monthly_report_config.json'
LIBRARY=SETTINGS.parent/'monthly_approved_dimensions.json'
DECISIONS=SETTINGS.parent/'monthly_invoice_review.json'


def _open_path(path: Path) -> None:
    if os.name=='nt':
        os.startfile(str(path))
    elif sys.platform=='darwin':
        subprocess.Popen(['open',str(path)])
    else:
        subprocess.Popen(['xdg-open',str(path)])


def _backup_and_install(files: list[tuple[Path,Path]], directory: Path) -> None:
    """Back up all existing outputs; roll back if a staged file fails to install."""
    existing=[dst for _,dst in files if dst.exists()]
    backup=None
    if existing:
        backup=directory/'_历史版本'/datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        backup.mkdir(parents=True,exist_ok=True)
        for dst in existing:shutil.copy2(dst,backup/dst.name)
    installed=[]
    try:
        for src,dst in files:
            os.replace(src,dst)
            installed.append(dst)
    except OSError:
        for dst in installed:
            old=backup/dst.name if backup else None
            if old is not None and old.exists():shutil.copy2(old,dst)
            elif dst.exists():dst.unlink()
        raise


class MonthlyReportWindow(tk.Toplevel):
    def __init__(self,master):
        super().__init__(master)
        self.title('HCE Oracle | 月度发运统计 / 发票复核 / ChatGPT 补全')
        self.geometry('1160x780'); self.minsize(840,570)
        self.configure(background='#F4F7FB')
        self.busy=False
        self._worker_queue=None
        self.entries: list[ReviewEntry]=[]
        self.by_invoice={}
        self.scan_signature=None
        self.last_report='';self.last_prompt=''
        previous=(date.today().replace(day=1)-timedelta(days=1))
        self.month=tk.StringVar(value=previous.strftime('%Y-%m'))
        self.folder=tk.StringVar(value=(master.sources[0]['path'] if getattr(master,'sources',[]) else ''))
        self.summary=tk.StringVar(value='')
        self.reference=tk.StringVar(value='')
        self.inventory=tk.StringVar(value='')
        self.output=tk.StringVar(value=str(Path.home()/'Documents'/'Hangcha Monthly Reports'))
        self.status=tk.StringVar(value='请选择月份和来源，先扫描并审核发票，再导出月报。原始 Excel 和 PDF 均只读。')
        self.counts=tk.StringVar(value='尚未扫描')
        self._load_options();self._build()
        for variable in (self.month,self.summary,self.folder):
            variable.trace_add('write',lambda *_args:self._invalidate_scan())
        self.protocol('WM_DELETE_WINDOW',self._close)

    def _build(self):
        root=ttk.Frame(self,padding=9);root.pack(fill='both',expand=True)
        ttk.Label(root,text='月度发运统计  |  先复核，再导出',font=('Microsoft YaHei UI',15,'bold')).pack(anchor='w',pady=(0,4))
        ttk.Label(root,text='仅显示所选月份发票。右键可查看 PDF / 明细、确认部分识别、移除及恢复。',
                  foreground='#506783').pack(anchor='w',pady=(0,6))
        conf=ttk.LabelFrame(root,text='① 统计月份及数据来源',padding=5);conf.pack(fill='x')
        form=ttk.Frame(conf);form.pack(fill='x');form.columnconfigure(1,weight=1)
        fields=[('月份 (YYYY-MM)',self.month,None),('业务周汇总 Excel',self.summary,'summary'),
                ('发票 PDF 根目录',self.folder,'folder'),('历史月报参考（可选）',self.reference,'file'),
                ('库存重量尺寸参考（可选）',self.inventory,'file'),('导出根目录',self.output,'out')]
        for i,(label,var,mode) in enumerate(fields):
            ttk.Label(form,text=label).grid(row=i,column=0,sticky='w',padx=(0,8),pady=2)
            ttk.Entry(form,textvariable=var).grid(row=i,column=1,sticky='ew',pady=2)
            if mode:
                ttk.Button(form,text='浏览…',width=8,command=lambda m=mode,v=var:self._browse(m,v)).grid(row=i,column=2,padx=(6,0),pady=2)
        actions=ttk.Frame(root);actions.pack(fill='x',pady=(7,3))
        self.scan_btn=ttk.Button(actions,text='① 扫描当月发票',command=self._scan)
        self.scan_btn.pack(side='left',padx=(0,8))
        self.export_btn=ttk.Button(actions,text='② 导出月报 + AI 文件',command=self._export,state='disabled')
        self.export_btn.pack(side='left',padx=(0,8))
        ttk.Button(actions,text='查看选中明细',command=self._show_details).pack(side='left',padx=(0,8))
        ttk.Button(actions,text='打开年月文件夹',command=self._open_dir).pack(side='right')
        actions2=ttk.Frame(root);actions2.pack(fill='x',pady=(0,6))
        ttk.Button(actions2,text='导入 ChatGPT 补全结果',command=self._import_ai).pack(side='left',padx=(0,8))
        ttk.Button(actions2,text='复制 Prompt',command=self._copy_prompt).pack(side='left',padx=(0,8))
        ttk.Button(actions2,text='保存已确认尺寸',command=self._save_confirmed).pack(side='left',padx=(0,8))
        ttk.Label(actions2,textvariable=self.counts,foreground='#20578B').pack(side='right')
        ttk.Label(root,textvariable=self.status,foreground='#20578B',wraplength=1100).pack(fill='x',pady=(0,5))
        panel=ttk.LabelFrame(root,text='② 当前月份发票复核列表（右键操作）',padding=5)
        panel.pack(fill='both',expand=True,pady=(0,8))
        columns=('invoice','state','count','reason','source')
        self.tree=ttk.Treeview(panel,columns=columns,show='headings',height=11,selectmode='browse')
        for col,name,width in [('invoice','发票号',132),('state','状态',96),('count','识别货物',76),
                               ('reason','说明 / 异常',300),('source','PDF 文件',170)]:
            self.tree.heading(col,text=name)
            self.tree.column(col,width=width,minwidth=60,stretch=col in ('reason','source'),anchor='w' if col!='count' else 'center')
        for state,bg,fg in [('正常','#E4F5EB','#15432A'),('已确认','#D8F1E4','#12452F'),
                            ('待复核','#FFF1D6','#7A4A12'),('来源已变化','#FCE0E0','#8A2020'),
                            ('已移除','#EEF0F3','#59667A'),('不符合条件','#F0F2F5','#657285')]:
            self.tree.tag_configure(state,background=bg,foreground=fg)
        self.tree.pack(side='left',fill='both',expand=True)
        scroll=ttk.Scrollbar(panel,orient='vertical',command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set);scroll.pack(side='right',fill='y')
        self.menu=tk.Menu(self,tearoff=False)
        self.menu.add_command(label='打开原始 PDF',command=self._open_pdf)
        self.menu.add_command(label='查看识别明细',command=self._show_details)
        self.menu.add_separator()
        self.menu.add_command(label='复核确认（仅纳入已识别货物）',command=lambda:self._decision('confirmed'))
        self.menu.add_command(label='移除发票（本月不纳入）',command=lambda:self._decision('removed'))
        self.menu.add_command(label='恢复自动筛选状态',command=lambda:self._decision('reset'))
        self.tree.bind('<Button-3>',self._popup)
        self.tree.bind('<Double-1>',lambda _ev:self._show_details())
        # File output guidance is shown in the status line / README. Keep the
        # table fully visible even on a 840x570 unmaximized window.

    def _invalidate_scan(self):
        """Never display another month's results after month/source selection changes."""
        if self.busy:return
        if self.scan_signature is not None:
            self.entries=[];self.by_invoice={};self.scan_signature=None
            self.tree.delete(*self.tree.get_children())
            self.export_btn.configure(state='disabled')
            self.counts.set('月份或数据来源已变更')
            self.status.set('月份或来源发生变化，请重新点击「① 扫描当月发票」。')

    def _browse(self,mode,var):
        if mode=='summary':
            p=filedialog.askopenfilename(parent=self,title='选择业务周汇总 Excel（含「发货」子表）',filetypes=[('Excel','*.xlsx'),('全部文件','*.*')])
        elif mode=='file':
            p=filedialog.askopenfilename(parent=self,title='选择历史重量尺寸参考',filetypes=[('Excel','*.xlsx'),('全部文件','*.*')])
        else:p=filedialog.askdirectory(parent=self,title='选择文件夹')
        if p:var.set(p)

    def _save_options(self):
        SETTINGS.parent.mkdir(parents=True,exist_ok=True)
        SETTINGS.write_text(json.dumps({'summary':self.summary.get(),'folder':self.folder.get(),
            'reference':self.reference.get(),'inventory':self.inventory.get(),
            'output':self.output.get()},ensure_ascii=False,indent=2),encoding='utf-8')

    def _load_options(self):
        try:
            opts=json.loads(SETTINGS.read_text(encoding='utf-8'))
            for key,var in [('summary',self.summary),('folder',self.folder),('reference',self.reference),
                            ('inventory',self.inventory),('output',self.output)]:
                if opts.get(key):var.set(opts[key])
        except (ValueError,OSError,TypeError):pass

    def _parse_month(self):
        import re
        value=self.month.get().strip()
        if not re.fullmatch(r'20\d\d-(0[1-9]|1[0-2])',value):
            raise ValueError('月份格式应为 YYYY-MM，例如 2026-10')
        y,m=map(int,value.split('-'))
        return y,m

    def _month_dir(self) -> Path:
        y,m=self._parse_month()
        root=self.output.get().strip()
        if not root:raise ValueError('请设置导出根目录')
        return Path(root)/f'{y}-{m:02d}'

    def _signature(self):
        y,m=self._parse_month()
        summary=Path(self.summary.get().strip())
        stat=summary.stat() if summary.is_file() else None
        return (y,m,os.path.normcase(os.path.abspath(str(summary))),
                (stat.st_size,stat.st_mtime_ns) if stat else None,
                os.path.normcase(os.path.abspath(self.folder.get().strip())))

    def _check_scan(self):
        if self.scan_signature is None or not self.entries:
            raise ValueError('请先扫描本月发票并复核，再生成月报')
        if self.scan_signature!=self._signature():
            raise ValueError('月份或来源已经改变，请重新扫描再导出')

    def _set_busy(self,value):
        self.busy=value
        self.scan_btn.configure(state='disabled' if value else 'normal')
        self.export_btn.configure(state='disabled' if value or self.scan_signature is None else 'normal')

    def _listen_worker(self, worker_queue):
        """Drain worker results on Tk's main thread; worker never calls Tk."""
        try:
            result=worker_queue.get_nowait()
        except queue.Empty:
            if self.winfo_exists():self.after(80,lambda:self._listen_worker(worker_queue))
            return
        if not self.winfo_exists():return
        kind=result[0]
        if kind=='scan':self._scan_done(result[1],result[2])
        elif kind=='export':self._export_done(result[1],result[2],result[3])
        elif kind=='error':self._failed(result[1])

    def _scan(self):
        if self.busy:return
        try:
            y,m=self._parse_month()
            workbook=Path(self.summary.get().strip());folder=Path(self.folder.get().strip())
            if not workbook.is_file() or workbook.suffix.lower()!='.xlsx':
                raise ValueError('请先选择正确的业务周汇总 Excel（.xlsx）')
            if not folder.is_dir():raise ValueError('发票 PDF 根目录不存在')
            signature=self._signature()
            self._save_options()
        except (ValueError,OSError) as exc:
            messagebox.showerror('配置错误',str(exc),parent=self);return
        self.entries=[];self.by_invoice={};self.scan_signature=None
        self.tree.delete(*self.tree.get_children())
        self._set_busy(True);self.status.set('正在扫描所选月份的发货汇总与对应 PDF…')
        worker_queue=queue.Queue()
        self.after(80,lambda:self._listen_worker(worker_queue))
        def worker():
            try:
                entries=scan_month_for_review(workbook,folder,y,m)
                apply_decisions(entries,DECISIONS,f'{y}-{m:02d}')
                worker_queue.put(('scan',entries,signature))
            except Exception as exc:
                worker_queue.put(('error',str(exc)))
        threading.Thread(target=worker,daemon=True,name='HCEOracleMonthlyScan').start()

    def _scan_done(self,entries,signature):
        if signature!=self._signature():
            self.entries=[];self.by_invoice={};self.scan_signature=None
            self._set_busy(False)
            self.status.set('扫描过程中月份或来源发生变化，请重新扫描。')
            return
        self.entries=entries
        self.by_invoice={x.invoice:x for x in entries}
        self.scan_signature=signature
        self._set_busy(False);self._refresh()
        self.status.set('扫描完成：可右键复核或移除发票，确认后再点击「导出月报 + AI 文件」。已保存的决定自动恢复。')

    def _refresh(self):
        selected=self.tree.selection()
        old=self.tree.item(selected[0],'values')[0] if selected else None
        self.tree.delete(*self.tree.get_children())
        for e in self.entries:
            state=e.display_state
            item=self.tree.insert('','end',iid=e.invoice,values=(e.invoice,state,
                len(e.record.items) if e.record else e.count,e.reason,Path(e.file).name),tags=(state,))
        if old and old in self.by_invoice:
            self.tree.selection_set(old);self.tree.see(old)
        num_in=sum(e.included for e in self.entries)
        confirmed=sum(e.display_state=='已确认' for e in self.entries)
        pending=sum(e.display_state in ('待复核','来源已变化') for e in self.entries)
        removed=sum(e.display_state=='已移除' for e in self.entries)
        self.counts.set(f'共 {len(self.entries)} 票 | 纳入 {num_in} | 手动确认 {confirmed} | 待复核 {pending} | 已移除 {removed}')

    def _get_selected(self):
        sel=self.tree.selection()
        if not sel:return None
        return self.by_invoice.get(sel[0])

    def _popup(self,event):
        row=self.tree.identify_row(event.y)
        if not row:return
        self.tree.selection_set(row);self.tree.focus(row)
        entry=self._get_selected()
        self.menu.entryconfigure(0,state='normal' if entry and Path(entry.file).suffix.lower()=='.pdf' and Path(entry.file).is_file() else 'disabled')
        self.menu.entryconfigure(3,state='normal' if entry and entry.can_confirm and not self.busy else 'disabled')
        self.menu.entryconfigure(4,state='normal' if not self.busy else 'disabled')
        self.menu.entryconfigure(5,state='normal' if not self.busy else 'disabled')
        try:self.menu.tk_popup(event.x_root,event.y_root)
        finally:self.menu.grab_release()

    def _open_pdf(self):
        e=self._get_selected()
        if not e or not Path(e.file).is_file() or Path(e.file).suffix.lower()!='.pdf':
            messagebox.showwarning('没有 PDF','此发票没有可打开的原始 PDF。',parent=self);return
        try:_open_path(Path(e.file))
        except Exception as exc:messagebox.showerror('打开失败',str(exc),parent=self)

    def _show_details(self):
        e=self._get_selected()
        if not e:return
        dialog=tk.Toplevel(self);dialog.title(f'发票 {e.invoice} · 识别明细')
        dialog.geometry('950x500');dialog.minsize(640,340);dialog.transient(self)
        main=ttk.Frame(dialog,padding=10);main.pack(fill='both',expand=True)
        ttk.Label(main,text=f'{e.invoice}  |  {e.display_state}',font=('Microsoft YaHei UI',13,'bold')).pack(anchor='w')
        ttk.Label(main,text=e.reason,wraplength=880,foreground='#9B5215').pack(fill='x',pady=6)
        if not e.record:
            ttk.Label(main,text='当前没有通过发票号/日期/金额校验的明细。此单不可手动强制纳入。').pack(anchor='w')
        else:
            r=e.record
            ttk.Label(main,text=f'日期 {r.day}  ·  客户 {r.customer or "未识别"}  ·  国家 {r.country or "未识别"}  ·  '
                         f'Total {r.gross}  /  Netto {r.net}  ·  已识别 {len(r.items)} 项',wraplength=870).pack(anchor='w',pady=4)
            cols=('erp','part','description','qty','price')
            tbl=ttk.Treeview(main,columns=cols,show='headings',height=12)
            for col,name,w in [('erp','ERP No.',180),('part','Part No.',200),('description','货物描述',300),('qty','数量',75),('price','行金额 €',100)]:
                tbl.heading(col,text=name);tbl.column(col,width=w,stretch=(col=='description'))
            tbl.pack(fill='both',expand=True,pady=(6,0))
            for item in r.items:
                tbl.insert('','end',values=(item.erp,item.part,item.description,str(item.quantity),str(item.line_total)))
            if r.notes:
                ttk.Label(main,text='未识别或其他异常：'+'；'.join(r.notes),wraplength=870,foreground='#9B5215').pack(fill='x',pady=6)
            if e.can_confirm and e.state=='待复核':
                ttk.Label(main,text='复核确认后，只导出上表已识别到的货物；未识别行不进入月报。',foreground='#245D41').pack(anchor='w')
        ttk.Button(main,text='关闭',command=dialog.destroy).pack(anchor='e',pady=(7,0))

    def _decision(self,action):
        if self.busy:return
        e=self._get_selected()
        if not e:return
        if action=='confirmed':
            if not e.can_confirm or not e.record:
                messagebox.showwarning('不可确认','编号、日期、金额、目的国必须先核对；至少需要一行有效配件。',parent=self);return
            errors='\n'.join(e.record.notes[:5]) if e.record.notes else '没有识别异常。'
            if not messagebox.askyesno('确认仅纳入已识别货物',
                f'发票 {e.invoice}\n已识别 {len(e.record.items)} 行配件。\n\n{errors}\n\n'
                '确认后，只有这几行进入月报，未识别的行将被忽略。继续吗？',parent=self):return
        elif action=='removed':
            if not messagebox.askyesno('确认移除',f'本次及以后重新生成本月月报时，将排除发票 {e.invoice}。\n可随时右键恢复。确定吗？',parent=self):return
        try:
            y,m=self._parse_month()
            self._check_scan()
            store_decision(DECISIONS,f'{y}-{m:02d}',e,action)
            self._refresh()
            self.status.set(f'发票 {e.invoice}：{e.display_state}。决定已自动保存。')
        except (ValueError,OSError) as exc:messagebox.showerror('保存复核决定失败',str(exc),parent=self)

    def _export(self):
        if self.busy:return
        try:
            self._check_scan()
            y,m=self._parse_month();directory=self._month_dir()
            ref=self.reference.get().strip();inventory=self.inventory.get().strip()
            if ref and not Path(ref).is_file():raise ValueError('历史参考 Excel 不存在')
            if inventory and not Path(inventory).is_file():raise ValueError('库存参考 Excel 不存在')
            chosen=selected_invoices(self.entries)
            if not chosen:raise ValueError('当前没有纳入的发票，请先检查筛选列表')
            # An edited/replaced PDF must NEVER be exported using a previous
            # review approval. Detect changes between scan and export too.
            for entry in self.entries:
                if not entry.record:continue
                src=Path(entry.record.path)
                if not src.is_file() or hashlib.sha256(src.read_bytes()).hexdigest()!=entry.record.digest:
                    raise ValueError(f'发票 {entry.invoice} 的 PDF 自上次扫描后发生变化，请重新扫描复核')
            pending=[e for e in self.entries if e.display_state in ('待复核','来源已变化')]
            if pending and not messagebox.askyesno('待复核发票',
                f'仍有 {len(pending)} 份发票待复核或来源已变化。它们不会纳入报表。\n确定继续导出吗？',parent=self):return
            directory.mkdir(parents=True,exist_ok=True)
            ym=f'{y}-{m:02d}'
            filenames=[f'Versendungsstatistikmeldung_{ym}_半成品.xlsx',
                       f'Versendungsstatistikmeldung_{ym}.xlsx',
                       f'ChatGPT_补全Prompt_{ym}.txt']
            if any((directory/f).exists() for f in filenames):
                if not messagebox.askyesno('保留旧版本',
                    '对应月份的报表已存在。软件会先复制旧文件到「_历史版本」文件夹，再生成新版本。继续吗？',parent=self):return
            logs=outcome_rows(self.entries)
            self._save_options()
        except (ValueError,OSError) as exc:
            messagebox.showerror('无法导出',str(exc),parent=self);return
        self._set_busy(True);self.status.set('正在生成月报、独立 Excel 和 AI 提示词…')
        worker_queue=queue.Queue()
        self.after(80,lambda:self._listen_worker(worker_queue))
        def worker():
            try:
                with tempfile.TemporaryDirectory(prefix='monthly_export_',dir=directory) as tmp:
                    temp=Path(tmp)
                    half=temp/filenames[0];clean=temp/filenames[1];prompt=temp/filenames[2]
                    result=generate_report(chosen,logs,y,m,half,ref or None,LIBRARY,inventory or None)
                    export_clean_report(half,clean)
                    prompt.write_text(build_prompt(y,m),encoding='utf-8')
                    _backup_and_install([(half,directory/filenames[0]),(clean,directory/filenames[1]),
                                         (prompt,directory/filenames[2])],directory)
                worker_queue.put(('export',result,directory,filenames))
            except Exception as ex:
                worker_queue.put(('error',str(ex)))
        threading.Thread(target=worker,daemon=True,name='HCEOracleMonthlyExport').start()

    def _export_done(self,result,directory,filenames):
        self._set_busy(False)
        self.last_report=str(directory/filenames[0]);self.last_prompt=str(directory/filenames[2])
        self.status.set(f"已输出 {result['invoices']} 份发票、{result['lines']} 行货物；待补字段 {result['tasks']} 项。目录：{directory}")
        messagebox.showinfo('月报导出成功',
            f'已生成独立月报、AI 半成品和 Prompt。\n\n独立月报：{filenames[1]}\nAI 半成品：{filenames[0]}\n提示词：{filenames[2]}\n\n目录：{directory}',parent=self)

    def _failed(self,err):
        self._set_busy(False);self.status.set('月度统计操作失败：'+err)
        messagebox.showerror('操作失败',err,parent=self)

    def _copy_prompt(self):
        try:
            y,m=self._parse_month()
            self.clipboard_clear();self.clipboard_append(build_prompt(y,m))
            self.status.set('Prompt 已复制。请连同当月「半成品.xlsx」上传 ChatGPT。')
        except (ValueError,tk.TclError) as exc:messagebox.showerror('操作失败',str(exc),parent=self)

    def _import_ai(self):
        if self.busy:return
        try:
            y,m=self._parse_month();directory=self._month_dir();ym=f'{y}-{m:02d}'
            original=directory/f'Versendungsstatistikmeldung_{ym}_半成品.xlsx'
            if not original.is_file():
                answer=messagebox.askyesno('未找到原始半成品',
                    '对应年月文件夹没有原始半成品。是否手动选择它？',parent=self)
                if not answer:return
                p=filedialog.askopenfilename(parent=self,title='选择原始半成品 Excel',initialdir=str(directory),filetypes=[('Excel','*.xlsx')])
                if not p:return
                original=Path(p)
            ai=filedialog.askopenfilename(parent=self,title='选择 ChatGPT 返回的补全 Excel',
                initialdir=str(directory),filetypes=[('Excel','*.xlsx')])
            if not ai:return
            merged=directory/f'Versendungsstatistikmeldung_{ym}_AI补全待复核.xlsx'
            clean=directory/f'Versendungsstatistikmeldung_{ym}.xlsx'
            expected=[merged,clean,merged.with_name(merged.stem+'_导入核验.txt')]
            if any(x.exists() for x in expected) and not messagebox.askyesno('备份旧版本',
                '会先将已有 AI 补全结果、独立月报及核验说明备份到「_历史版本」，再覆盖。继续吗？',parent=self):return
            directory.mkdir(parents=True,exist_ok=True)
        except (ValueError,OSError) as exc:messagebox.showerror('路径错误',str(exc),parent=self);return
        try:
            with tempfile.TemporaryDirectory(prefix='monthly_ai_',dir=directory) as tmp:
                work=Path(tmp)/merged.name
                result=import_ai(original,ai,work)
                clean_tmp=Path(tmp)/clean.name
                export_clean_report(work,clean_tmp)
                audit=Path(result['review_txt'])
                _backup_and_install([(work,merged),(clean_tmp,clean),
                                     (audit,expected[2])],directory)
            self.status.set(f"已导入 {result['filled']} 项重量/体积；正式独立月报已更新，待复核 {len(result['errors'])} 项。")
            messagebox.showinfo('AI安全导入完成',
                f"导入 {result['filled']} / {result['tasks']} 项。\n未通过 {len(result['errors'])} 项。"
                f'\n独立月报已同步更新：{clean.name}\n目录：{directory}',parent=self)
        except Exception as exc:messagebox.showerror('AI导入失败（原文件未更改）',str(exc),parent=self)

    def _save_confirmed(self):
        try:directory=self._month_dir()
        except ValueError as exc:messagebox.showerror('月份格式',str(exc),parent=self);return
        path=filedialog.askopenfilename(parent=self,title='选择经人工确认的 AI 补全报告',
                    filetypes=[('Excel','*.xlsx')],initialdir=str(directory))
        if not path:return
        if not messagebox.askyesno('保存已确认尺寸',
            '仅保存标记为“实测 / 供应商参数 / 历史已确认”且有依据的单件货物重量和体积。\n'
            'AI估算及包裹数据不进入资料库。确定吗？',parent=self):return
        try:
            result=update_approved_library(path,LIBRARY)
            self.status.set(f"本次确认 {result['updated_fields']} 项尺寸，资料库共有 {result['parts']} 个配件。")
            messagebox.showinfo('已保存',self.status.get(),parent=self)
        except Exception as exc:messagebox.showerror('保存失败',str(exc),parent=self)

    def _open_dir(self):
        try:
            directory=self._month_dir()
            directory.mkdir(parents=True,exist_ok=True)
            _open_path(directory)
        except Exception as exc:messagebox.showerror('无法打开文件夹',str(exc),parent=self)

    def _close(self):
        if self.busy and not messagebox.askyesno('任务尚未完成',
            '当前操作仍在进行，关闭窗口不会立即停止文件生成。确定关闭吗？',parent=self):return
        self.destroy()
