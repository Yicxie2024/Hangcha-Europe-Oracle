"""Read-only inventory search workspace for Hangcha Sales Assistant.

Independent of the sales Excel, archive SQLite, and Teams upload workflow.
Open for a single SKU before making an order, or inspect all parts of a
selected sales-order PDF. The user chooses the source inventory workbook.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from gui_picking_ui import PickingMixin

from gui_inventory_core import (
    OrderPart, InventoryIndex, SearchResult, assess_stock, load_inventory,
    read_order_parts, readable_number,
)

PREFS = Path(os.environ.get('LOCALAPPDATA',str(Path.home()))) / 'HangchaSalesAssistant' / 'inventory_settings.json'
# Cache an immutable snapshot for repeated window opens within this process.
_SESSION_CACHE: dict[tuple[str,int,int], InventoryIndex] = {}


class InventoryWindow(PickingMixin, tk.Toplevel):
    def __init__(self, parent: tk.Misc, order_pdf: str | None = None):
        super().__init__(parent)
        self.title('HCE Oracle | 统一库存查询与拣货单')
        self.geometry('1350x780')
        self.minsize(930, 570)
        self.index: InventoryIndex | None = None
        self._index_signature = None
        self._loading = False
        self._generation = 0
        self._order_parts: list[OrderPart] = []
        self._order_path = ''
        self._order_loading = False
        self._events: queue.Queue = queue.Queue()
        self._picking_state_init()
        self._prefs = self._load_prefs()
        self.stock_path = tk.StringVar(value=self._prefs.get('inventory_path',''))
        self.lookup_code = tk.StringVar()
        self.order_file = tk.StringVar()
        self.summary_var = tk.StringVar(value='请选择一份库存 Excel；所有操作均为只读')
        self.single_status = tk.StringVar(value='支持 ERP No. / Part No. 精确搜索（ERP 优先）')
        self.order_status = tk.StringVar(value='选择已扫描订单，或手动打开一份订单 PDF')
        self._build()
        self.after(120, self._poll_events)
        if self.stock_path.get():
            self.after(160,self._load_index)
        if order_pdf:
            self.after(200,lambda: self.open_order(order_pdf))

    def _load_prefs(self) -> dict:
        try:
            data=json.loads(PREFS.read_text(encoding='utf-8'))
            return data if isinstance(data,dict) else {}
        except (OSError, ValueError, TypeError):
            return {}

    def _save_prefs(self):
        try:
            PREFS.parent.mkdir(parents=True,exist_ok=True)
            prefs=dict(self._prefs)
            prefs['inventory_path']=self.stock_path.get().strip()
            # Atomic replacement, no changes to any business file.
            temp=PREFS.with_suffix('.tmp')
            temp.write_text(json.dumps(prefs,ensure_ascii=False,indent=2),encoding='utf-8')
            temp.replace(PREFS)
        except OSError as exc:
            messagebox.showwarning('未能保存库存配置', f'本次查询仍可使用，但下次可能需要重新选择文件：\n{exc}',parent=self)

    def _build(self):
        top=ttk.Frame(self,padding=12)
        top.pack(fill='x')
        ttk.Label(top,text='库存来源 Excel：').pack(side='left')
        ttk.Entry(top,textvariable=self.stock_path,state='readonly',width=69).pack(side='left',fill='x',expand=True,padx=(4,6))
        ttk.Button(top,text='选择库存 Excel…',command=self._choose_inventory).pack(side='left',padx=3)
        ttk.Button(top,text='重新加载',command=lambda:self._load_index(force=True)).pack(side='left',padx=3)
        ttk.Label(self,textvariable=self.summary_var,foreground='#5B6472',padding=(12,0,12,4)).pack(fill='x')

        # One searchable workspace: manual code, Excel and any number of
        # sales-order PDFs all append to the SAME pending list.
        self.tabs=ttk.Notebook(self)
        self.tabs.pack(fill='both',expand=True,padx=12,pady=(5,12))
        self._picking_build()

    def _create_table(self,parent, cols):
        area=ttk.Frame(parent)
        area.pack(fill='both',expand=True)
        keys=[c[0] for c in cols]
        tree=ttk.Treeview(area,columns=keys,show='headings',selectmode='extended')
        vs=ttk.Scrollbar(area,orient='vertical',command=tree.yview)
        hs=ttk.Scrollbar(area,orient='horizontal',command=tree.xview)
        tree.configure(yscrollcommand=vs.set,xscrollcommand=hs.set)
        for key,label,width in cols:
            tree.heading(key,text=label)
            tree.column(key,width=width,minwidth=65,stretch=False,anchor='w')
        area.rowconfigure(0,weight=1)
        area.columnconfigure(0,weight=1)
        tree.grid(row=0,column=0,sticky='nsew')
        vs.grid(row=0,column=1,sticky='ns')
        hs.grid(row=1,column=0,sticky='ew')
        tree.bind('<Control-c>',lambda e:self._copy_selected(tree))
        return tree

    def _set_note_tree_colors(self,tree):
        tree.tag_configure('ok',foreground='#006B45')
        tree.tag_configure('bad',foreground='#AE3A3A')
        tree.tag_configure('warn',foreground='#966200')
        tree.tag_configure('dim',foreground='#677586')

    def _copy_selected(self,tree):
        selection=tree.selection()
        if not selection:
            return
        lines=['\t'.join(str(v) for v in tree.item(iid,'values')) for iid in selection]
        self.clipboard_clear()
        self.clipboard_append('\n'.join(lines))
        self.summary_var.set(f'已复制 {len(lines)} 行到剪贴板')

    def _choose_inventory(self):
        p=filedialog.askopenfilename(parent=self,title='选择仓库库存 Excel（仅读取）',
               filetypes=[('Excel 工作簿','*.xlsx *.xlsm'),('所有文件','*.*')])
        if p:
            self.stock_path.set(p)
            self._save_prefs()
            self._load_index(force=True)

    def _signature(self,p:Path):
        st=p.stat()
        return str(p.resolve()),st.st_mtime_ns,st.st_size

    def _load_index(self,force:bool=False):
        raw=self.stock_path.get().strip()
        if not raw:
            self.summary_var.set('请先选择库存 Excel；不需要先扫描 PDF')
            return
        p=Path(raw)
        try:
            stamp=self._signature(p)
        except OSError as exc:
            self.index=None
            self.summary_var.set('库存文件不可读取，请重新选择 Excel')
            messagebox.showerror('库存文件不可读取',str(exc),parent=self)
            return
        if not force and self.index is not None and self._index_signature==stamp:
            return
        if not force and stamp in _SESSION_CACHE:
            self.index = _SESSION_CACHE[stamp]
            self._index_signature=stamp
            self._loading=False
            self.batch_results=[]
            self.batch_result_tree.delete(*self.batch_result_tree.get_children())
            dt=datetime.fromtimestamp(stamp[1]/1e9).strftime('%d.%m.%Y %H:%M')
            self.summary_var.set(f'已加载 {len(self.index.records)} 条库存记录 · 文件保存时间 {dt} · 本次会话缓存')
            self._picking_index_loaded()
            return
        self._generation+=1
        generation=self._generation
        self._loading=True
        self.index=None
        self.summary_var.set('正在后台读取库存 Excel…较大的库存文件首次加载可能需要一些时间')
        def work():
            try:
                idx=load_inventory(p)
                self._events.put(('index',generation,stamp,idx))
            except Exception as exc:
                self._events.put(('index_error',generation,stamp,str(exc)))
        threading.Thread(target=work,daemon=True).start()

    def _ensure_current_index(self) -> bool:
        if not self.stock_path.get().strip():
            messagebox.showwarning('请选择库存 Excel','请先点击上方「选择库存 Excel…」。',parent=self)
            return False
        if self._loading:
            return False
        if self.index is None:
            self._load_index()
            return False
        try:
            if self._index_signature!=self._signature(Path(self.stock_path.get())):
                self._load_index(force=True)
                return False
        except OSError:
            self._load_index(force=True)
            return False
        return True

    def _poll_events(self):
        if not self.winfo_exists():
            return
        for _ in range(12):
            try: kind,generation,stamp,payload=self._events.get_nowait()
            except queue.Empty: break
            if kind.startswith('index'):
                if generation!=self._generation: continue
                self._loading=False
                if kind=='index_error':
                    self.summary_var.set('库存文件读取失败：'+str(payload))
                    continue
                self.index=payload
                self._index_signature=stamp
                self.batch_results=[]
                self.batch_result_tree.delete(*self.batch_result_tree.get_children())
                _SESSION_CACHE.clear()
                _SESSION_CACHE[stamp]=payload
                dt=datetime.fromtimestamp(stamp[1]/1e9).strftime('%d.%m.%Y %H:%M')
                self.summary_var.set(f'已加载 {len(payload.records)} 条库存记录 · 文件保存时间 {dt} · 只读快照')
                self._picking_index_loaded()
            elif kind == 'batch_order':
                self._batch_order_loaded(payload)
            elif kind == 'batch_order_error':
                self._batch_order_error(payload)
        self.after(130,self._poll_events)

    def search(self):
        """Compatibility with older callers; use the unified request list."""
        from gui_picking_core import PickRequest
        from decimal import Decimal
        code=self.lookup_code.get().strip()
        if code and self._batch_append([PickRequest(code,Decimal('1'))],source='手工'):
            self._batch_search()

    def open_order(self, pdf: str):
        """Main GUI's selected-order action now adds to the shared batch list."""
        self.tabs.select(self.batch_tab)
        self._batch_add_order_paths([pdf])

    def _choose_order(self):
        self._batch_choose_order_pdfs()

    def _reload_order(self):
        if self._order_path:
            self._batch_add_order_paths([self._order_path])
        else:
            self._batch_choose_order_pdfs()

    def _render_order(self):
        """Legacy hook: the shared results table replaces the old order tab."""
        return
