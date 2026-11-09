"""Tkinter batch search and print/picking UI extension for InventoryWindow.

All confirmations default to the safe action on Enter; Esc cancels. This UI
never changes the source inventory, PDF files, sales Excel or archives.
"""
from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path
import os
import re
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog, font as tkfont

from gui_inventory_core import readable_number, read_order_parts
from gui_picking_core import (PickRequest, PickResult, parse_paste, import_requests_xlsx,
    requests_from_order, lookup_requests, allocate_and_route, parse_rack, MAX_ITEMS)
from gui_picking_pdf import make_pick_pdf, find_sumatra, print_pdf
from gui_picking_history import count_submitted,record_submitted
from gui_picking_location_stock import LocationStockStore, validate_group
from gui_picking_core import split_rack_locations
from gui_inventory_core import number_or_none


class DownwardQueryButton(tk.Canvas):
    """Clickable, keyboard-accessible query action with a *drawn* down arrow.

    The arrow is a Canvas shape instead of a Unicode glyph, so it also works
    on Windows machines whose Tk font cannot render U+2B07 (⬇).
    """

    WIDTH = 174
    HEIGHT = 58
    NORMAL = '#156879'
    HOVER = '#0D5267'
    ACTIVE = '#093D51'

    def __init__(self, parent, *, command):
        background = ttk.Style(parent).lookup('TFrame', 'background') or '#F0F0F0'
        super().__init__(parent, width=self.WIDTH, height=self.HEIGHT,
                         background=background, borderwidth=0,
                         highlightthickness=0, takefocus=1, cursor='hand2')
        self._command = command
        self._hover = False
        self._pressed = False
        self._focused = False
        self._label_font = tkfont.nametofont('TkDefaultFont').copy()
        self._label_font.configure(weight='bold', size=11)
        self.bind('<Enter>', self._enter)
        self.bind('<Leave>', self._leave)
        self.bind('<ButtonPress-1>', self._press)
        self.bind('<ButtonRelease-1>', self._release)
        self.bind('<Key-Return>', self._keyboard_click)
        self.bind('<Key-space>', self._keyboard_click)
        self.bind('<FocusIn>', self._focus_in)
        self.bind('<FocusOut>', self._focus_out)
        self.bind('<Configure>', lambda event: self._redraw())
        self._redraw()

    def _enter(self, _event):
        self._hover = True
        self._redraw()

    def _leave(self, _event):
        self._hover = False
        self._pressed = False
        self._redraw()

    def _press(self, _event):
        self.focus_set()
        self._pressed = True
        self._redraw()

    def _release(self, event):
        was_pressed = self._pressed
        self._pressed = False
        self._redraw()
        if was_pressed and 0 <= event.x < self.winfo_width() and 0 <= event.y < self.winfo_height():
            self._command()

    def _keyboard_click(self, _event):
        self._command()
        return 'break'

    def _focus_in(self, _event):
        self._focused = True
        self._redraw()

    def _focus_out(self, _event):
        self._focused = False
        self._redraw()

    def _redraw(self):
        self.delete('all')
        w = max(self.winfo_width(), self.WIDTH) if self.winfo_width() <= 1 else self.winfo_width()
        h = max(self.winfo_height(), self.HEIGHT) if self.winfo_height() <= 1 else self.winfo_height()
        fill = self.ACTIVE if self._pressed else self.HOVER if self._hover else self.NORMAL
        # Rounded rectangle; the label uses ordinary CJK text, and the only
        # decorative icon is drawn below it using Tk primitives.
        inset = 3
        radius = 12
        x0,y0,x1,y1 = inset,inset,w-inset,h-inset
        outline = '#E0A948' if self._focused else '#0B4656'
        coords = [x0+radius,y0, x1-radius,y0, x1,y0, x1,y0+radius,
                  x1,y1-radius, x1,y1, x1-radius,y1, x0+radius,y1,
                  x0,y1, x0,y1-radius, x0,y0+radius, x0,y0]
        self.create_polygon(coords, smooth=True, splinesteps=16,
                            fill=fill, outline=outline,
                            width=2 if self._focused else 1)
        cx=w/2
        self.create_text(cx, h*0.34, text='查询库存', font=self._label_font,
                         fill='white', anchor='center')
        # Broad shaft and solid arrowhead: no Unicode arrows or emoji.
        self.create_line(cx, h*0.55, cx, h*0.73,
                         fill='white', width=4, capstyle=tk.ROUND)
        self.create_polygon(cx-11,h*0.70, cx+11,h*0.70, cx,h*0.88,
                            fill='white', outline='white')


class PickingMixin:
    def _picking_state_init(self):
        self.batch_input=tk.StringVar()
        self.batch_default_qty=tk.StringVar(value='1')
        self.batch_status=tk.StringVar(value='可手工添加、导入 Excel，或分次添加多份订单 PDF')
        self.batch_requests:list[PickRequest]=[]
        self.batch_sources:list[str]=[]  # parallel to requests: 手工, Excel or stable PDF key
        self.batch_order_files:dict[str,dict]={}  # key -> path/contract/customer/signature
        self._batch_order_inflight:set[str]=set()
        self._batch_order_autoprint:set[str]=set()
        self._batch_epoch=0
        self.batch_results:list[PickResult]=[]
        self._batch_iid_to_result={}
        self._auto_print_order=False
        self._order_print_identity=''
        self._order_file_version=None

    def _picking_build(self):
        self.batch_tab=ttk.Frame(self.tabs,padding=8)
        self.tabs.add(self.batch_tab,text='统一库存搜索与拣货单')
        inputs=ttk.LabelFrame(self.batch_tab,text='添加查询内容 · 手工 / Excel / 订单 PDF',padding=7)
        inputs.pack(fill='x')
        line=ttk.Frame(inputs);line.pack(fill='x')
        ttk.Label(line,text='ERP No. / Part No.：').pack(side='left')
        self.batch_input_entry=ttk.Entry(line,textvariable=self.batch_input)
        self.batch_input_entry.pack(side='left',fill='x',expand=True,padx=(4,6))
        self.batch_input_entry.bind('<Return>',lambda e:self._batch_add())
        ttk.Label(line,text='默认数量').pack(side='left')
        qty=ttk.Entry(line,textvariable=self.batch_default_qty,width=6)
        qty.pack(side='left',padx=4)
        qty.bind('<Return>',lambda e:self._batch_add())
        ttk.Button(line,text='添加编号',command=self._batch_add).pack(side='left',padx=3)
        line2=ttk.Frame(inputs);line2.pack(fill='x',pady=(6,0))
        ttk.Button(line2,text='＋ 添加订单 PDF…',command=self._batch_choose_order_pdfs).pack(side='left',padx=3)
        ttk.Button(line2,text='导入 Excel…',command=self._batch_import).pack(side='left',padx=3)
        ttk.Label(line2,text='支持分次添加 PDF；双击上方清单可修改数量或补录编号。',
                  foreground='#516679').pack(side='left',padx=(12,0))

        # Deterministic three-row layout: list → permanent action bar → result.
        # Unlike a vertical Panedwindow, the arrow and delete buttons cannot
        # be squeezed out by sash position or a non-maximized window.
        workspace=ttk.Frame(self.batch_tab)
        workspace.pack(fill='both',expand=True,pady=(5,3))
        workspace.columnconfigure(0,weight=1)
        workspace.rowconfigure(0,weight=2,minsize=108)
        workspace.rowconfigure(1,weight=0)
        workspace.rowconfigure(2,weight=3,minsize=155)
        box1=ttk.LabelFrame(workspace,text='① 待搜索清单 · 双击修改数量 / 补录编号',padding=5)
        box1.grid(row=0,column=0,sticky='nsew')
        self.batch_pending_tree=self._create_table(box1,[
            ('code','ERP / Part No.',240),('qty','需求量',115),('type','来源（订单编号 / 文件名）',285)])
        self.batch_pending_tree.tag_configure('missing',foreground='#B04A18')
        self.batch_pending_tree.bind('<Double-1>',lambda e:self._batch_edit_qty())
        self.batch_pending_tree.bind('<Button-3>',self._batch_pending_menu)

        between=ttk.Frame(workspace,padding=(4,5))
        between.grid(row=1,column=0,sticky='ew')
        between.columnconfigure(0,weight=1)
        between.columnconfigure(1,weight=1)
        between.columnconfigure(2,weight=1)
        left=ttk.Frame(between);left.grid(row=0,column=0,sticky='w')
        self.batch_delete_btn=ttk.Button(left,text='删除选中',command=self._batch_delete)
        self.batch_delete_btn.pack(side='left',padx=(0,5))
        self.batch_clear_btn=ttk.Button(left,text='清空清单',command=self._batch_clear)
        self.batch_clear_btn.pack(side='left')
        # Center CTA explicitly points *from* the pending list *to* results.
        # Canvas arrow avoids missing-glyph square boxes on Windows Tk.
        self.batch_search_btn=DownwardQueryButton(between,command=self._batch_search)
        self.batch_search_btn.grid(row=0,column=1)
        ttk.Label(between,text='上方：待查配件\n下方：库存结果',
                  foreground='#516679',justify='right').grid(
            row=0,column=2,sticky='e',padx=(3,2))

        box2=ttk.LabelFrame(workspace,text='② 库存查询结果 · 勾选要打印的配件',padding=5)
        box2.grid(row=2,column=0,sticky='nsew')
        # Pack fixed-height actions FIRST and anchor to the bottom so they
        # remain visible even when the tree has only ~100px of height.
        action2=ttk.Frame(box2);action2.pack(side='bottom',fill='x',pady=(5,0))
        ttk.Button(action2,text='剔除 / 恢复选中',command=self._batch_toggle_selected).pack(side='left',padx=3)
        ttk.Button(action2,text='查看库位详情',command=lambda:self._batch_result_detail(None)).pack(side='left',padx=3)
        ttk.Button(action2,text='核对多库位库存…',command=self._batch_edit_location_stock).pack(side='left',padx=3)
        ttk.Button(action2,text='生成 / 预览 A4 拣货单',command=lambda:self._batch_print(preview=True)).pack(side='right',padx=3)
        ttk.Button(action2,text='打印拣货单…',command=self._batch_print).pack(side='right',padx=3)
        self.batch_result_tree=self._create_table(box2,[
            ('include','打印',46),('code','输入编号',174),('qty','需求',62),
            ('status','状态',115),('erp','ERP No.',158),('part','Part No.',155),
            ('rack','Rack No.',182),('stock','库存合计',85),('desc','描述',220)])
        self.batch_result_tree.bind('<Button-1>',self._batch_toggle_click,add='+')
        self.batch_result_tree.bind('<Double-1>',self._batch_result_detail,add='+')
        ttk.Label(self.batch_tab,textvariable=self.batch_status,foreground='#516679',
                  wraplength=900).pack(fill='x',pady=(2,0))

    def _batch_pending_menu(self,event):
        row=self.batch_pending_tree.identify_row(event.y)
        if not row:return
        if row not in self.batch_pending_tree.selection():
            self.batch_pending_tree.selection_set(row)
        menu=tk.Menu(self,tearoff=0)
        menu.add_command(label='修改数量 / 补录编号（双击）',command=self._batch_edit_qty)
        index=int(row[1:])
        source=self.batch_sources[index]
        if source in self.batch_order_files:
            pdf=self.batch_order_files[source]['path']
            menu.add_command(label='打开来源订单 PDF',command=lambda:self._open_preview(Path(pdf)))
        menu.add_separator()
        menu.add_command(label='删除选中',command=self._batch_delete)
        try:menu.tk_popup(event.x_root,event.y_root)
        finally:menu.grab_release()

    @staticmethod
    def _batch_pdf_key(path):
        # Windows paths case-insensitive. Path.resolve also removes ./ and ..
        return str(Path(path).resolve()).casefold()

    @staticmethod
    def _batch_pdf_signature(path):
        st=Path(path).stat()
        return (st.st_size,st.st_mtime_ns)

    def _batch_choose_order_pdfs(self):
        paths=filedialog.askopenfilenames(parent=self,title='选择一份或多份订单 PDF（可再次添加）',
                                          filetypes=[('PDF 订单','*.pdf')])
        if paths:self._batch_add_order_paths(paths)

    def _batch_add_order_paths(self, paths, auto_print=False):
        """Asynchronous per-PDF import: additive, safely deduped and nonblocking."""
        started=0;duplicates=0;errors=[]
        for value in paths:
            path=Path(value)
            try:
                if not path.is_file() or path.suffix.lower()!='.pdf':
                    raise ValueError('文件不是可读取的 PDF')
                key=self._batch_pdf_key(path)
                stamp=self._batch_pdf_signature(path)
            except (OSError,ValueError) as exc:
                errors.append(f'{path.name}: {exc}');continue
            existing=self.batch_order_files.get(key)
            if existing and existing['signature']==stamp and not auto_print:
                duplicates+=1;continue
            if key in self._batch_order_inflight:
                if auto_print:self._batch_order_autoprint.add(key)
                else:duplicates+=1
                continue
            self._batch_order_inflight.add(key)
            if auto_print:self._batch_order_autoprint.add(key)
            epoch=self._batch_epoch
            def worker(p=path,k=key,signature=stamp,request_epoch=epoch):
                try:
                    parts=read_order_parts(p,allow_incomplete=True)
                    requests=requests_from_order(parts,allow_incomplete=True)
                    if not requests:
                        raise ValueError('订单没有需要库存查询的配件（运费等项目不计）')
                    if self._batch_pdf_signature(p)!=signature:
                        raise ValueError('读取期间 PDF 已修改，请重新选择')
                    from automation import pdf_order_identity
                    from gui_customer import extract_order_customer
                    ident=pdf_order_identity(p)
                    contract=ident[0] if ident else ''
                    customer=extract_order_customer(p) or '未确认客户'
                    self._events.put(('batch_order',0,None,dict(
                        key=k,path=str(p),signature=signature,parts=parts,
                        requests=requests,contract=contract,customer=customer,epoch=request_epoch)))
                except Exception as exc:
                    self._events.put(('batch_order_error',0,None,dict(
                        key=k,path=str(p),name=p.name,error=str(exc),epoch=request_epoch)))
            threading.Thread(target=worker,daemon=True).start()
            started+=1
        if errors:
            messagebox.showwarning('部分文件无法添加','\n'.join(errors[:12]),parent=self)
        if started:
            self.batch_status.set(f'正在读取 {started} 份订单 PDF；已有清单会保留。'
                                  +(f' 已跳过 {duplicates} 份重复订单。' if duplicates else ''))
        elif duplicates:
            self.batch_status.set(f'已跳过 {duplicates} 份已添加的订单；清单保持不变。')

    def _batch_order_error(self,data):
        self._batch_order_inflight.discard(data['key'])
        self._batch_order_autoprint.discard(data['key'])
        if data['epoch']!=self._batch_epoch:return
        messagebox.showwarning('订单 PDF 无法自动读取',
            f"{data['name']}\n{data['error']}\n\n将打开原始 PDF。请复制 ERP No. / Part No.，"
            '通过上方「添加编号」手动添加；其他清单项目不会丢失。',parent=self)
        self._open_preview(Path(data['path']))
        self.batch_input_entry.focus_set()
        self.batch_status.set(f"读取 {data['name']} 失败，已尝试打开 PDF 供人工复制；清单未清空。")

    def _batch_order_loaded(self,data):
        key=data['key'];self._batch_order_inflight.discard(key)
        auto_print=key in self._batch_order_autoprint
        self._batch_order_autoprint.discard(key)
        if data['epoch']!=self._batch_epoch:return
        path=Path(data['path'])
        try:
            if self._batch_pdf_signature(path)!=data['signature']:
                raise ValueError('订单文件已变化，请重新添加')
        except (OSError,ValueError) as exc:
            messagebox.showerror('订单 PDF 已变化',str(exc),parent=self);return
        current=self.batch_order_files.get(key)
        replacements=[]
        if current:
            if current['signature']==data['signature']:
                if not auto_print:
                    self.batch_status.set(f'{path.name} 已在待查询清单中，没有重复添加。')
            else:
                replacements=[key]
                if not self._confirm_enter('订单内容更新',
                    f'{path.name} 已经添加过，但 PDF 内容发生了变化。是否用新订单明细替换旧内容？'):
                    return
        elif data['contract']:
            for oldkey,info in self.batch_order_files.items():
                if info['contract'] and info['contract']==data['contract']:
                    replacements.append(oldkey)
            if replacements and not self._confirm_enter('检测到相同合同号',
                f'合同 {data["contract"]} 已有订单 PDF。是否替换原订单明细？\n'
                '选择取消可避免同一订单被重复统计。'):
                return
        if not current or replacements:
            after_count=len(self.batch_requests)-sum(x in replacements for x in self.batch_sources)
            if after_count+len(data['requests'])>MAX_ITEMS:
                messagebox.showerror('条目过多',f'合计最多 {MAX_ITEMS} 项，当前订单未添加。',parent=self)
                return
            if replacements:
                keep=[(r,src) for r,src in zip(self.batch_requests,self.batch_sources)
                      if src not in replacements]
                self.batch_requests=[r for r,src in keep]
                self.batch_sources=[src for r,src in keep]
                for oldkey in replacements:self.batch_order_files.pop(oldkey,None)
            self.batch_order_files[key]={k:data[k] for k in
                ('path','signature','contract','customer','requests')}
            self._batch_append(data['requests'],source=key)
            missing_ids=sum(not r.code for r in data['requests'])
            missing_qty=sum(r.quantity is None for r in data['requests'])
            if missing_ids:
                messagebox.showwarning('订单存在未识别编号',
                    f'{path.name}\n有 {missing_ids} 行未能识别 ERP No. / Part No.。\n'
                    '将打开来源 PDF。请复制编号，并双击待搜索清单中「待补编号」的条目手动粘贴；'
                    '补全前不允许查询或打印。',parent=self)
                self._open_preview(path)
            if missing_ids or missing_qty:
                self.batch_status.set(
                    f'已读取 {path.name}；待补编号 {missing_ids} 项，待补数量 {missing_qty} 项。'
                    '双击橙色待搜索条目完成补录，再点击中间的向下箭头。')
            else:
                self.batch_status.set(f'已添加 {path.name}（合同 {data["contract"] or "未识别"}，'
                    f'{len(data["requests"])} 项）；可继续添加 PDF / Excel / 编号。')
        # Even for a duplicate source, the right-click order-print shortcut
        # must keep printing THAT PDF (not a previously imported one).
        if auto_print:
            self._order_path=str(path)
            self._order_file_version=data['signature']
            self._order_parts=data['parts']
            self._auto_print_order=True
            self._picking_order_loaded()

    def _batch_validate_pdf_sources(self):
        for info in self.batch_order_files.values():
            path=Path(info['path'])
            try:
                if self._batch_pdf_signature(path)!=info['signature']:
                    raise ValueError('文件内容已改变')
            except (OSError,ValueError) as exc:
                messagebox.showwarning('订单来源发生变化',
                    f'订单 PDF {path.name} 不再与待查询清单一致：{exc}。\n'
                    '请重新添加此订单并确认替换后再搜索或打印。',parent=self)
                return False
        return True

    def _batch_add(self):
        try:
            new=parse_paste(self.batch_input.get(),self.batch_default_qty.get())
            if not new:raise ValueError('请先输入至少一个编号')
            self._batch_append(new,source='手工')
            self.batch_input.set('')
        except ValueError as exc:messagebox.showwarning('添加失败',str(exc),parent=self)

    def _batch_append(self,new,source='手工'):
        if len(self.batch_requests)+len(new)>MAX_ITEMS:
            messagebox.showerror('条目过多',f'最多 {MAX_ITEMS} 项',parent=self);return False
        self.batch_requests.extend(new)
        self.batch_sources.extend([source]*len(new))
        self._batch_invalidate_results()
        self._batch_render_pending()
        return True

    def _batch_invalidate_results(self):
        self.batch_results=[]
        self.batch_result_tree.delete(*self.batch_result_tree.get_children())
        self._batch_iid_to_result.clear()

    def _batch_import(self):
        path=filedialog.askopenfilename(parent=self,title='导入 ERP/Part + 需求量 Excel',
                             filetypes=[('Excel','*.xlsx *.xlsm')])
        if not path:return
        try:
            requests=import_requests_xlsx(path)
        except Exception as exc:
            messagebox.showerror('导入失败（没有添加任何条目）',str(exc),parent=self);return
        if not self._batch_append(requests,source='Excel'):return
        self.batch_status.set(f'已从 {Path(path).name} 导入 {len(requests)} 项；可编辑/删除后统一搜索')

    def _batch_render_pending(self):
        tree=self.batch_pending_tree;tree.delete(*tree.get_children())
        for i,(r,src) in enumerate(zip(self.batch_requests,self.batch_sources)):
            unresolved=not r.code or r.quantity is None or r.quantity<=0
            if src in self.batch_order_files:
                info=self.batch_order_files[src]
                source_label=f'订单 {info["contract"] or Path(info["path"]).name}'
            else:
                source_label=src
            tree.insert('','end',iid=f'r{i}',values=(
                r.code or '⚠ 待补 ERP/Part No.',
                ('⚠ 待补数量' if r.quantity is None or r.quantity<=0
                 else readable_number(r.quantity)),
                source_label),tags=('missing',) if unresolved else ())
        unknown=sum(not r.code or r.quantity is None or r.quantity<=0
                    for r in self.batch_requests)
        self.batch_status.set(f'待搜索 {len(self.batch_requests)} 项'
             + (f' · {unknown} 项需双击补录' if unknown else '')
             + '；修改清单后需要重新查询')

    def _batch_selected_indices(self):
        return sorted((int(i[1:]) for i in self.batch_pending_tree.selection() if i.startswith('r')))

    def _batch_edit_qty(self):
        """Double-click repairs either an unreadable ID or a quantity."""
        indices=self._batch_selected_indices()
        if len(indices)!=1:
            messagebox.showinfo('修改条目','请先选中一条待搜索记录',parent=self);return
        i=indices[0];old=self.batch_requests[i]
        code=old.code
        erp,part=old.erp,old.part
        if not code:
            code=simpledialog.askstring('手动补录 ERP / Part No.',
                'PDF 中未识别配件编号。请从打开的订单 PDF 复制 ERP No. 或 Part No.，'
                f'粘贴到下方：\n\n描述：{old.description or "未识别"}',
                parent=self)
            if code is None:return
            code=code.strip()
            # A manually pasted code is an untyped identifier, so ERP and
            # Part columns remain blank: lookup searches ERP first, Part next.
            if not code or any(c in code for c in ('\\', ':', ' ', '\t', '\n')):
                messagebox.showwarning('编号无效','请输入单个不含空格的 ERP / Part No.',parent=self);return
            erp=part=''
        # For recovered IDs with reliable quantities, do NOT force another
        # confirmation; ordinary rows retain double-click quantity editing.
        if old.quantity is None or old.quantity<=0:
            value=simpledialog.askstring('手动补录数量',
                f'编号：{code}\nPDF 数量未识别；请查看订单后填写实际数量（不默认填 1）：',
                initialvalue='',parent=self)
        elif not old.code:
            value=readable_number(old.quantity)
        else:
            value=simpledialog.askstring('修改需求量',
                f'编号：{code}\n请输入新的需求量：',
                initialvalue=readable_number(old.quantity),parent=self)
        if value is None:return
        from gui_inventory_core import number_or_none
        quantity=number_or_none(value)
        if quantity is None or quantity<=0:
            messagebox.showwarning('无效需求量','需求量必须大于 0',parent=self);return
        self.batch_requests[i]=PickRequest(code,quantity,erp,part,old.description)
        self._batch_invalidate_results()
        self._batch_render_pending()
        self.batch_pending_tree.selection_set(f'r{i}')
        self.batch_pending_tree.focus(f'r{i}')
        self.batch_status.set(f'条目 {code} 已手动确认数量 {readable_number(quantity)}；'
                              '请点击中间向下箭头重新查询库存。')

    def _batch_delete(self):
        indices=set(self._batch_selected_indices())
        if not indices:return
        self.batch_requests=[v for i,v in enumerate(self.batch_requests) if i not in indices]
        self.batch_sources=[v for i,v in enumerate(self.batch_sources) if i not in indices]
        self._batch_prune_order_files()
        self._batch_invalidate_results()
        self._batch_render_pending()

    def _batch_prune_order_files(self):
        active=set(self.batch_sources)
        for key in tuple(self.batch_order_files):
            if key not in active:self.batch_order_files.pop(key,None)

    def _batch_clear(self):
        if self.batch_requests and not self._confirm_enter('清空待搜索列表','确定清空全部待搜索条目？此操作不会更改库存 Excel。'):
            return
        self._batch_epoch+=1
        self.batch_requests=[];self.batch_sources=[];self.batch_order_files.clear()
        self._batch_order_inflight.clear();self._batch_order_autoprint.clear()
        self._batch_invalidate_results()
        self._batch_render_pending()

    def _batch_search(self):
        if not self.batch_requests:
            self.batch_status.set('待搜索清单为空');return
        pending=[i for i,r in enumerate(self.batch_requests)
                 if not r.code or r.quantity is None or r.quantity<=0]
        if pending:
            first=pending[0]
            iid=f'r{first}'
            self.batch_pending_tree.selection_set(iid)
            self.batch_pending_tree.focus(iid)
            self.batch_pending_tree.see(iid)
            messagebox.showwarning('清单尚未补录完整',
                f'有 {len(pending)} 个条目缺少 ERP/Part No. 或有效数量。\n'
                '已选中第一条橙色记录，请双击补录后再查询。\n'
                '系统不会默认猜测数量，也不会忽略未识别的配件。',parent=self)
            self.batch_status.set(f'尚有 {len(pending)} 条待补录；查询未执行。')
            return
        if self._batch_order_inflight:
            self.batch_status.set('仍有订单 PDF 正在读取，请在读取完成后查询库存。');return
        if not self._batch_validate_pdf_sources():return
        if not self._ensure_current_index():
            self.batch_status.set('库存 Excel 尚未加载；加载完成后再次点击「查询库存」');return
        self.batch_results=lookup_requests(self.index,self.batch_requests)
        self._render_batch_results()

    def _render_batch_results(self):
        tree=self.batch_result_tree;tree.delete(*tree.get_children())
        self._batch_iid_to_result={}
        for i,r in enumerate(self.batch_results):
            iid=f'b{i}'
            self._batch_iid_to_result[iid]=r
            racks=', '.join(s.rack or '未知' for s in r.records) or '—'
            tree.insert('','end',iid=iid,values=('☐' if r.excluded else '☑',r.request.code,
                readable_number(r.request.quantity),r.status,r.erp,r.part,racks,
                readable_number(r.stock_total),r.description),tags=(
                'dim' if r.excluded else 'ok' if r.status=='充足' else 'warn',))
        self.batch_status.set(f'已查询 {len(self.batch_results)} 个唯一配件项目；左侧 ☑ 为参与打印（可排除）；'
                              '多库位、缺货、编号冲突均保留详细提示')

    def _batch_toggle_click(self,event):
        if self.batch_result_tree.identify_region(event.x,event.y)!='cell':return
        if self.batch_result_tree.identify_column(event.x)=='#1':
            iid=self.batch_result_tree.identify_row(event.y)
            if iid:
                self._batch_toggle_iids([iid]);return 'break'

    def _batch_toggle_iids(self,iids):
        for iid in iids:
            r=self._batch_iid_to_result.get(iid)
            if r:r.excluded=not r.excluded
        self._render_batch_results()

    def _batch_toggle_selected(self):
        self._batch_toggle_iids(list(self.batch_result_tree.selection()))

    def _batch_result_detail(self,event):
        selection=self.batch_result_tree.selection()
        if event is not None:
            iid=self.batch_result_tree.identify_row(event.y)
            if iid:selection=(iid,)
        if len(selection)!=1:return
        r=self._batch_iid_to_result.get(selection[0])
        if not r:return
        lines=[f'{r.request.code} | 需求 {readable_number(r.request.quantity)} | {r.status}',
               f'匹配依据：{r.match_method or "无"}',f'编号：ERP={r.erp} | Part={r.part}',
               f'库存：{readable_number(r.stock_total)}',f'说明：{r.warning or "无"}', '']
        store=LocationStockStore(self.stock_path.get()) if self.index is not None else None
        for hit in r.records:
            locs=split_rack_locations(hit.rack)
            if len(locs)>1:
                for loc in locs:
                    value=store.get(hit,loc) if store else None
                    lines.append(f'Rack {loc} | 该库位库存 {readable_number(value) if value is not None else "待核"}'
                                 f' | Excel 行 {hit.row} | 货物 {hit.description or "—"}')
            else:
                lines.append(f'Rack {hit.rack or "未提供"} | 库存 {readable_number(hit.stock)} | Excel 行 {hit.row}'
                             f' | 重量 {hit.weight or "—"} | 尺寸 {hit.dimension or "—"}'
                             f' | 中文描述 {hit.description_cn or "—"}')
        messagebox.showinfo('库存详细记录','\n'.join(lines),parent=self)

    def _batch_edit_location_stock(self):
        """Manually confirm rack-level balances absent from the aggregate Excel.

        Read-only for source inventory. All inputs are validated before a
        snapshot-bound supplemental JSON is atomically written.
        """
        if not self._ensure_current_index():
            return
        if not self.batch_results:
            messagebox.showinfo('先搜索库存','请先查询库存，再核对本次需要拣取的多库位配件。',parent=self)
            return
        records={r.row:r for result in self.batch_results if not result.excluded
                 for r in result.records if len(split_rack_locations(r.rack))>1}
        if not records:
            messagebox.showinfo('没有多库位配件','本次查询没有需要拆分库位库存的配件。',parent=self)
            return
        try:
            store=LocationStockStore(self.stock_path.get())
        except OSError as exc:
            messagebox.showerror('库存文件无法访问',str(exc),parent=self)
            return
        pending={row:store.group(rec) for row,rec in records.items()}
        pop=tk.Toplevel(self)
        pop.title('核对每个库位的实际库存 · 仅针对当前 Excel 版本')
        pop.geometry('920x530')
        pop.minsize(730,400)
        pop.transient(self)
        root=ttk.Frame(pop,padding=12)
        root.pack(fill='both',expand=True)
        ttk.Label(root,wraplength=870,justify='left',text=(
            '双击“库位库存”填写该位置真实数量；空白表示尚未核实。Excel 的 MA 列只有整行总库存，'
            '软件不会自动平分或复制到多个位置。所有填写数据都不会修改原始 Excel；'
            '库存 Excel 变更后需要重新确认。')).pack(fill='x',pady=(0,8))
        columns=('row','part','loc','value','state')
        area=ttk.Frame(root)
        area.pack(fill='both',expand=True)
        area.rowconfigure(0,weight=1);area.columnconfigure(0,weight=1)
        tree=ttk.Treeview(area,columns=columns,show='headings',selectmode='browse')
        headers={'row':('Excel行',75),'part':('ERP / Part No.',240),
                 'loc':('库位',145),'value':('库位库存',95),'state':('校验说明',300)}
        for col in columns:
            title,width=headers[col]
            tree.heading(col,text=title)
            tree.column(col,width=width,minwidth=70,stretch=(col in ('part','state')))
        y=ttk.Scrollbar(area,orient='vertical',command=tree.yview)
        tree.configure(yscrollcommand=y.set)
        tree.grid(row=0,column=0,sticky='nsew')
        y.grid(row=0,column=1,sticky='ns')
        row_by_item={}
        for rec in sorted(records.values(),key=lambda v:v.row):
            for loc in split_rack_locations(rec.rack):
                iid=f'{rec.row}:{len(row_by_item)}'
                row_by_item[iid]=(rec.row,loc)
                tree.insert('','end',iid=iid)

        def refresh():
            for iid,(row,loc) in row_by_item.items():
                rec=records[row]
                values=pending[row]
                filled=sum(v is not None for v in values.values())
                group_sum=sum((v for v in values.values() if v is not None),Decimal('0'))
                err=validate_group(rec,values)
                if err:
                    status='! '+err
                elif filled==len(values):
                    status='已核对：数量合计符合库存表'
                else:
                    status=f'已核 {filled}/{len(values)} 处 · 尚有位置待核'
                val=values.get(loc)
                tree.item(iid,values=(rec.row,rec.erp or rec.part,loc,
                    readable_number(val) if val is not None else '待核',status))

        def edit(event=None):
            iid=tree.identify_row(event.y) if event is not None else ''
            if not iid:
                sel=tree.selection()
                iid=sel[0] if sel else ''
            if iid not in row_by_item:
                return
            row,loc=row_by_item[iid]
            rec=records[row]
            old=pending[row][loc]
            val=simpledialog.askstring('填写真实库位库存',
                f'配件：{rec.erp or rec.part}\n库位：{loc}\n'
                f'库存表该行总库存：{readable_number(rec.stock)}\n'
                '请输入现场核实的这个库位库存数量，留空表示未知：',
                initialvalue=readable_number(old) if old is not None else '',parent=pop)
            if val is None:
                return
            v=number_or_none(val) if val.strip() else None
            if v is not None and (not v.is_finite() or v<0):
                messagebox.showwarning('数量无效','库位库存必须是非负数字，或留空表示待核。',parent=pop)
                return
            test=dict(pending[row]);test[loc]=v
            # Enforce <= aggregate immediately, while allowing an incomplete
            # group to be finished across several double-clicks.
            if rec.stock is not None and sum((a for a in test.values() if a is not None),Decimal('0'))>rec.stock:
                messagebox.showwarning('超出总库存','填写的库位库存合计不能超过 Excel 的总库存。',parent=pop)
                return
            pending[row]=test
            refresh()

        tree.bind('<Double-1>',edit)
        ttk.Label(root,text='注意：同一配件的多个库位始终共用一个拣货序号；补录仅作为库位库存参考，不会扣库存。',
                  foreground='#53657A').pack(fill='x',pady=(8,4))
        controls=ttk.Frame(root)
        controls.pack(fill='x')
        ttk.Button(controls,text='编辑选中库位…',command=edit).pack(side='left')
        def save():
            try:
                store.update_groups([(rec,pending[row]) for row,rec in records.items()])
            except (ValueError,OSError) as exc:
                messagebox.showwarning('无法保存库位数量',str(exc),parent=pop)
                return
            self.summary_var.set(f'已核对本次 {len(records)} 条多库位库存记录；生成拣货单将显示各库位真实数量或“待核”。')
            pop.destroy()
        ttk.Button(controls,text='保存核对结果',command=save).pack(side='right',padx=(6,0))
        ttk.Button(controls,text='取消',command=pop.destroy).pack(side='right')
        pop.bind('<Escape>',lambda e:pop.destroy())
        pop.grab_set()
        refresh()
        tree.focus_set()

    def _picking_index_loaded(self):
        if self._auto_print_order and self._order_parts and self.index:
            self._maybe_open_autoprint()

    def _picking_order_loaded(self):
        if self._auto_print_order and self._order_parts and self.index:
            self._maybe_open_autoprint()

    def request_order_print(self,pdf):
        self.tabs.select(self.batch_tab)
        self._batch_add_order_paths([pdf],auto_print=True)

    def _maybe_open_autoprint(self):
        if not self._auto_print_order:return
        self._auto_print_order=False
        self.after(120,self._order_print)

    def _order_file_signature(self):
        return self._batch_pdf_signature(self._order_path)

    def _order_identity(self) -> tuple[str,str]:
        if not self._order_path:return '', ''
        from automation import pdf_order_identity
        from gui_customer import extract_order_customer
        p=Path(self._order_path)
        identity=pdf_order_identity(p)
        return (identity[0] if identity else '',extract_order_customer(p) or '未确认客户')

    def _linked_deliveries(self,contract:str):
        # Rows are trusted only from a COMPLETE PDF-only scan. Their contract
        # field was derived from delivery PDF Ordernumber, NOT Delivery No.
        master=self.master
        if not contract or not getattr(master,'scanned_complete',False):return []
        found=[];seen=set()
        from gui_archive import file_sha256
        for r in getattr(master,'rows',[]):
            if (r.kind=='发货单' and r.contract==contract and r.status not in
                ('忽略作废','待复核','读取失败') and Path(r.source).is_file()):
                if r.source not in seen:
                    if r.sha and file_sha256(r.source)!=r.sha:
                        raise ValueError('关联发货单 PDF 内容已变化，请先重新完整扫描，避免打印错误文件')
                    found.append((r.delivery_no or '未知编号',Path(r.source)));seen.add(r.source)
        return sorted(found,key=lambda x:x[0])

    def _order_print(self,preview:bool=False):
        if not self._order_path or not self._order_parts:
            messagebox.showwarning('先加载订单','请选择订单 PDF，并等待明细识别完成',parent=self);return
        if not self._ensure_current_index():
            messagebox.showwarning('库存尚未加载','请先加载库存 Excel 并重新执行打印',parent=self);return
        try:
            if self._order_file_version and self._order_file_version != self._order_file_signature():
                raise ValueError('订单 PDF 已在读取后修改，请重新查询该订单，避免使用过期需求量')
            contract,customer=self._order_identity()
            if not contract:raise ValueError('订单 PDF 合同号不能可靠提取，禁止生成关联发货单打印任务')
            key=self._batch_pdf_key(self._order_path)
            if key in self.batch_order_files:
                requests=[req for req,src in zip(self.batch_requests,self.batch_sources)
                          if src==key]
                if not requests or any(not r.code or r.quantity is None or r.quantity<=0 for r in requests):
                    raise ValueError('订单仍存在未补录完整的 ERP/Part 或数量，请先双击清单补录')
            else:
                requests=requests_from_order(self._order_parts)
            results=lookup_requests(self.index,requests)
            self._prepare_print(results,contract,customer,preview=preview)
        except Exception as exc:messagebox.showerror('生成拣货单失败',str(exc),parent=self)

    def _batch_print(self,preview:bool=False):
        if any(not r.code or r.quantity is None or r.quantity<=0 for r in self.batch_requests):
            messagebox.showwarning('未完成手动补录',
                '待搜索清单中仍有缺失编号或数量的配件，无法生成可能遗漏物料的拣货单。',parent=self)
            return
        if not self.batch_results:
            messagebox.showwarning('请先搜索','先把待搜索清单中的编号全部搜索，再选择打印',parent=self);return
        if not self._batch_validate_pdf_sources():return
        if not self._ensure_current_index():return
        results=[r for r in self.batch_results if not r.excluded]
        if not results:
            messagebox.showwarning('没有打印项目','所有结果都已排除，无法生成拣货单',parent=self);return
        # Only a pure single-PDF request may print a linked delivery note.
        sources=set(self.batch_sources)
        if len(sources)==1:
            info=self.batch_order_files.get(next(iter(sources)))
            current=[req for req,src in zip(self.batch_requests,self.batch_sources) if src in self.batch_order_files]
            # All original order rows must still be present. Manual repairs
            # of an unreadable qty/code are allowed, but deleting a row turns
            # this into an internal/partial list (no linked delivery printing).
            if info and info['contract'] and len(current)==len(info['requests']):
                self._order_path=info['path']
                self._order_file_version=info['signature']
                self._prepare_print(results,info['contract'],info['customer'],preview=preview)
                return
        orders=[d['contract'] or Path(d['path']).stem for d in self.batch_order_files.values()]
        desc=('合并订单：'+', '.join(orders[:8])+('…' if len(orders)>8 else '')) if orders else '内部备货'
        self._prepare_print(results,'',desc,preview=preview)

    def _prepare_print(self,results,contract,customer,preview=False):
        # Never print stale Excel snapshot; already checked in caller.
        from copy import deepcopy
        safe=deepcopy(results)
        steps,notes=allocate_and_route(safe)
        ambiguous=sum(1 for r in safe if r.shortage>0 or r.status not in {'充足'})
        output=Path.home()/'Documents'/'Hangcha Picking Lists'
        ts=datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        name=f'Picking_{contract or "Internal"}_{ts}.pdf'
        try:
            confirmed=LocationStockStore(self.stock_path.get()).pdf_values(safe)
        except (OSError,ValueError) as exc:
            messagebox.showerror('库位库存核对数据不可用',str(exc),parent=self)
            return
        pdf=make_pick_pdf(output/name,contract,customer,safe,steps,notes,
                          self.stock_path.get(),rack_stock=confirmed)
        from gui_picking_core import split_rack_locations
        display_count=sum(1 for result in safe if not result.excluded for _ in result.records)
        display_count+=sum(1 for result in safe if not result.excluded and not result.records)
        self.summary_var.set(f'已生成 A4 拣货单：{pdf.name} | 配件种类 {len(safe)} | 库存记录行 {display_count} | 待人工核对 {ambiguous}')
        deliveries=self._linked_deliveries(contract) if contract else []
        if preview:
            self._open_preview(pdf);return
        config=self._print_dialog(contract,display_count,ambiguous,deliveries)
        if config is None:return
        p_copies,selected_deliveries,delivery_copies=config
        # Catch changes while the modal confirmation was displayed.
        if not self._ensure_current_index():
            messagebox.showwarning('库存表已更改','库存文件发生变化。请重新搜索/重新生成拣货单后再打印。',parent=self)
            return
        if contract:
            if self._order_file_version != self._order_file_signature():
                messagebox.showwarning('订单已更改','订单 PDF 已变化，请重新加载订单后再打印。',parent=self)
                return
            try:
                current=set(str(path) for _,path in self._linked_deliveries(contract))
            except ValueError as exc:
                messagebox.showerror('发货单已更改',str(exc),parent=self);return
            if any(str(p) not in current for _,p in selected_deliveries):
                messagebox.showerror('发货单发生变化','请重新扫描，再重新选择需要打印的发货单。',parent=self)
                return
        # User confirmed: Enter presses the primary "提交打印" button.
        jobs=[(pdf,p_copies)]+[(p,delivery_copies) for _,p in selected_deliveries]
        printer=find_sumatra()
        if printer is None:
            self._open_preview(pdf)
            for _,p in selected_deliveries:self._open_preview(p)
            messagebox.showwarning('无法自动打印',
                '未发现 SumatraPDF。已打开拣货单和选中的发货单，请在 PDF 阅读器中手动设置打印份数。\n'
                '如需一键按份数直接打印，请安装 SumatraPDF 并设置系统默认打印机。',parent=self)
            return
        def runner():
            issues=[]
            for file,copies in jobs:
                try:
                    proc=print_pdf(file,copies,printer)
                    status=proc.wait(timeout=120)
                    if status!=0:issues.append(f'{file.name}：打印程序返回 {status}')
                except Exception as exc:issues.append(f'{file.name}：{exc}')
            if not issues:
                try:
                    record_submitted(contract,str(pdf),p_copies,
                        [(no,delivery_copies) for no,_ in selected_deliveries],
                        self.stock_path.get())
                except Exception as exc:
                    issues.append('打印程序已提交任务，但写入本地打印历史失败：'+str(exc))
            self.after(0,lambda:self._print_done(issues,len(jobs),pdf))
        threading.Thread(target=runner,daemon=True).start()
        self.summary_var.set(f'已提交 {len(jobs)} 个打印任务，正在等待打印程序反馈…')

    def _print_done(self,issues,count,pdf):
        if issues:
            messagebox.showerror('打印未全部完成','\n'.join(issues),parent=self)
        else:
            self.summary_var.set(f'已向默认打印机提交 {count} 个文档；请在打印机检查实际出纸（文件保存在 {pdf.parent}）')

    def _open_preview(self,pdf):
        try:
            if os.name=='nt':os.startfile(str(pdf))
            else:
                import shutil
                exe=shutil.which('xdg-open')
                if exe:subprocess.Popen([exe,str(pdf)],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        except Exception as exc:
            messagebox.showinfo('PDF 已生成',f'{pdf}\n\n无法自动打开：{exc}',parent=self)

    def _confirm_enter(self,title,text):
        return self._dialog_enter(title,text)

    def _dialog_enter(self,title,text):
        pop=tk.Toplevel(self);pop.title(title);pop.transient(self);pop.resizable(False,False)
        value=tk.BooleanVar(value=False)
        ttk.Label(pop,text=text,wraplength=460,padding=(17,16)).pack(fill='x')
        foot=ttk.Frame(pop,padding=(12,8));foot.pack(fill='x')
        def done(confirm):value.set(confirm);pop.destroy()
        ok=ttk.Button(foot,text='确认（Enter）',command=lambda:done(True));ok.pack(side='right',padx=5)
        ttk.Button(foot,text='取消（Esc）',command=lambda:done(False)).pack(side='right')
        pop.bind('<Return>',lambda e:done(True));pop.bind('<Escape>',lambda e:done(False))
        pop.protocol('WM_DELETE_WINDOW',lambda:done(False));pop.grab_set();ok.focus_set();self.wait_window(pop)
        return value.get()

    def _print_dialog(self,contract,count,unresolved,deliveries):
        pop=tk.Toplevel(self);pop.title('打印确认 - Enter 提交 / Esc 取消')
        pop.transient(self);pop.resizable(False,False)
        output={'result':None}
        body=ttk.Frame(pop,padding=15);body.pack(fill='both',expand=True)
        ttk.Label(body,text=f'拣货单：{contract or "内部备货"} | 库存记录行 {count} | 待核项 {unresolved}',
                  font=('',11,'bold')).pack(anchor='w',pady=(0,8))
        ttk.Label(body,text='A4 横向；库存以已保存的 Excel 快照为准。打印不会扣减库存。',
                  foreground='#666666').pack(anchor='w',pady=(0,9))
        if contract:
            try:
                old=count_submitted(contract)
            except Exception:old=0
            if old:
                ttk.Label(body,text=f'⚠ 这个订单此前有 {old} 次打印提交记录；确认后仍可重打。',
                          foreground='#A65C1D').pack(anchor='w',pady=(0,9))
        row=ttk.Frame(body);row.pack(fill='x')
        ttk.Label(row,text='拣货单份数：').pack(side='left')
        pick_var=tk.StringVar(value='1')
        pick_spin=ttk.Spinbox(row,from_=1,to=20,textvariable=pick_var,width=5)
        pick_spin.pack(side='left',padx=(2,12))
        ttk.Label(row,text='发货单每份份数：').pack(side='left')
        delivery_var=tk.StringVar(value='1')
        delivery_spin=ttk.Spinbox(row,from_=1,to=20,textvariable=delivery_var,width=5)
        delivery_spin.pack(side='left')
        ttk.Label(body,text=f'通过 PDF Ordernumber 关联的发货单：{len(deliveries)} 份',
                  foreground='#3B5979').pack(anchor='w',pady=(12,4))
        choices=[]
        for no,p in deliveries:
            include=tk.BooleanVar(value=True)
            ttk.Checkbutton(body,text=f'同时打印 Delivery No. {no}  ({p.name})',
                            variable=include).pack(anchor='w',pady=2)
            choices.append((no,p,include))
        if not deliveries:
            ttk.Label(body,text='本次扫描未找到关联发货单，或尚未完成扫描；只打印拣货单。',
                      foreground='#936023',wraplength=530).pack(anchor='w',pady=(4,10))
        if unresolved:
            ttk.Label(body,text='⚠ 存在缺货/数量未分配的条目；可定位库位直接排入对应楼层，需现场核对数量。',
                      foreground='#A34A2A').pack(anchor='w',pady=(9,4))
        btns=ttk.Frame(body);btns.pack(fill='x',pady=(14,0))
        def close(ok):
            if ok:
                try:
                    a,b=int(pick_var.get()),int(delivery_var.get())
                    if not 1<=a<=20 or not 1<=b<=20:raise ValueError()
                except ValueError:
                    messagebox.showwarning('打印份数无效','请填写 1–20 之间的整数',parent=pop);return
                output['result']=(a,[(no,p) for no,p,var in choices if var.get()],b)
            pop.destroy()
        confirm=ttk.Button(btns,text='提交打印（Enter）',command=lambda:close(True))
        confirm.pack(side='right',padx=5)
        ttk.Button(btns,text='取消（Esc）',command=lambda:close(False)).pack(side='right')
        pop.bind('<Return>',lambda e:close(True))
        pop.bind('<Escape>',lambda e:close(False))
        pick_spin.bind('<Return>',lambda e:close(True))
        delivery_spin.bind('<Return>',lambda e:close(True))
        pop.protocol('WM_DELETE_WINDOW',lambda:close(False))
        pop.grab_set();pick_spin.focus_set();self.wait_window(pop)
        return output['result']
