"""Explicitly approved business preview / Excel write for the universal GUI.

No filesystem mutation while planning (except private temporary copies). Writing is
allowed ONLY by GUI confirmation; a scheduled run uses preview exclusively.
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import re
import tempfile
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import automation as legacy
from gui_scanner import ScanRow
from gui_notes import NoteDraft, NoteStore


@dataclass(frozen=True)
class NoteChange:
    draft: NoteDraft
    sheet: str
    old: str
    final: str
    location: str  # '新增记录' or '已有记录'


@dataclass
class BusinessPlan:
    target: Path
    existing: bool
    workbook_fingerprint: str | None
    source_fingerprints: dict[str, str]
    orders: list
    invoices: list
    warnings: list[str]
    notices: list[str]
    skipped: list[str]
    files_seen: int
    scanned_delivery: int
    cancelled: int
    generated_at: datetime
    source_specs: tuple[tuple[str,str], ...] = ()
    exclude_names: tuple[str, ...] = ()
    scanned_paths: frozenset[str] = frozenset()
    invoiced_contracts: frozenset[str] = frozenset()
    note_changes: tuple[NoteChange, ...] = ()
    note_completed: tuple[NoteDraft, ...] = ()
    note_snapshot: tuple[NoteDraft, ...] = ()
    note_store_path: Path | None = None
    note_profile: str = ''

    @property
    def added(self):
        return len(self.orders) + len(self.invoices)

    @property
    def has_changes(self):
        return bool(self.added or self.note_changes)


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _input_records(rows: list[ScanRow], salesperson: str, archived_paths=None):
    """Map a GUI scan snapshot to legacy business Record objects.

    Invalid PDFs are NOT admitted to the write candidates. Still keep all valid
    invoices for month-end presence checks, including PDFs missing their totals.
    """
    records = []
    warnings = []
    fallback_orders = []
    delivery = []
    invoices_seen = []
    hashes = {}
    cancelled = 0
    archived_paths = set(archived_paths or ())
    for row in rows:
        pdf = Path(row.source)
        if row.kind == '目录异常':
            warnings.append('【目录不可访问】' + str(pdf) + '；扫描结果不完整，禁止写入')
            continue
        if not pdf.is_file():
            warnings.append('【PDF 已移动】' + str(pdf) + '；必须重新扫描')
            continue
        if os.path.normcase(os.path.abspath(row.source)) in archived_paths:
            # Hide from new Excel inserts, but preserve *valid* invoice presence
            # evidence so archiving cannot create a false missing-invoice alarm.
            if row.kind == '发票' and row.status != '忽略作废':
                invoices_seen.append((pdf, row.customer))
            hashes[str(pdf)] = file_hash(pdf)
            continue
        if row.status == '忽略作废':
            cancelled += 1
            continue
        # Duplicate PDF copies are already correctly removed in GUI scanner.
        if row.status == '重复文件':
            continue
        if row.kind == '发货单':
            delivery.append((pdf, row.customer))
        if row.kind == '发票':
            invoices_seen.append((pdf, row.customer))
        hashes[str(pdf)] = file_hash(pdf)
        if row.kind not in ('订单', '发票'):
            if row.status == '待复核' and row.kind != '发货单':
                warnings.append(f'【未分类 PDF】{pdf}：{row.note}')
            elif row.kind == '发货单' and row.status == '待复核':
                warnings.append(f'【发货单待核对】{pdf}：{row.note}')
            continue
        kind = 'order' if row.kind == '订单' else 'invoice'
        if row.status == '待复核' or row.customer == '未确定':
            warnings.append(f'【PDF 字段待复核】{pdf}：{row.note}')
            if kind == 'order':
                try:
                    ident = legacy.pdf_order_identity(pdf)
                    if ident:
                        fallback_orders.append((ident[0], ident[1], row.customer, str(pdf)))
                except Exception:
                    pass
            continue
        try:
            record = legacy.extract_pdf(pdf, kind, row.customer, salesperson,
                                        customer_selected_by_scanner=True)
            if row.sha and record.sha != row.sha:
                raise RuntimeError('PDF 内容自上次扫描以来发生变化')
            records.append(record)
        except Exception as exc:
            warnings.append(f'【无法安全提取】{pdf}：{exc}')
    records, dup_warnings = legacy.normalize_duplicates(records)
    warnings += dup_warnings
    return records, warnings, fallback_orders, delivery, invoices_seen, hashes, cancelled


def _open_workbook(excel_app, workbook: Path, preview: bool):
    """Open a PRIVATE copy for both previews and confirmed edits.

    Excel should never open the live OneDrive/desktop workbook during an automated
    write. All edits occur in a private temp copy; only the final verified file
    replaces the live workbook after the user has explicitly approved the plan.
    """
    if not workbook.exists() or workbook.stat().st_size == 0:
        return excel_app.Workbooks.Add(), None
    fd, filename = tempfile.mkstemp(prefix='hangcha_excel_work_', suffix='.xlsx')
    os.close(fd)
    copied = Path(filename)
    try:
        shutil.copy2(workbook, copied)
        return (excel_app.Workbooks.Open(str(copied), UpdateLinks=0, ReadOnly=preview,
                                        IgnoreReadOnlyRecommended=True), copied)
    except BaseException:
        copied.unlink(missing_ok=True)
        raise


def _with_excel(target: Path, preview: bool, callback):
    if os.name != 'nt':
        raise RuntimeError('业务汇总核对需要 Windows + 桌面版 Microsoft Excel。')
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = None
    wb = None
    copied = None
    try:
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        wb, copied = _open_workbook(app, target, preview)
        if not preview and bool(wb.ReadOnly):
            raise RuntimeError('Excel 正在被占用或只读；请先保存并关闭工作簿。')
        return callback(wb)
    finally:
        if wb is not None:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if app is not None:
            try: app.Quit()
            except Exception: pass
        if copied is not None:
            try: copied.unlink(missing_ok=True)
            except Exception: pass
        pythoncom.CoUninitialize()


def _plan_with_book(wb, records, warnings, fallback, delivery, invoice_files):
    legacy.prepare_workbook(wb)
    old_orders, _, _ = legacy.sheet_snapshot(wb.Worksheets('订单'), 'order')
    old_invoices, _, _ = legacy.sheet_snapshot(wb.Worksheets('发货'), 'invoice')
    new_o, new_i, issues, skipped, notices = legacy.calculate_plan(records, old_orders, old_invoices)
    month_errs, month_notices, _ = legacy.missing_month_end_invoices(
        records, old_orders, old_invoices, delivery, invoice_files,
        today=datetime.now().date(), warning_days=5, fallback_orders=fallback)
    # Excel may already contain an invoice even when its PDF is not in the scan.
    # An exact contract number is required; do not infer from delivery-note IDs.
    excel_invoiced_contracts = frozenset(str(contract).strip() for contract, _ in old_invoices)
    return new_o, new_i, warnings + issues + month_errs, notices + month_notices, skipped, excel_invoiced_contracts


def _note_destination_text(old: str, text: str) -> str:
    """Append; never replace existing human edits or duplicate a prior note."""
    old = old.strip()
    if not old:
        return text
    if text in [s.strip() for s in old.split('；')]:
        return old
    return old + '；' + text


def _plan_notes(wb, drafts: list[NoteDraft], new_orders, new_invoices):
    """Determine the exact Excel remark updates, without modifying any cells."""
    if not drafts:
        return (), (), []
    ws_o, ws_i = wb.Worksheets('订单'), wb.Worksheets('发货')
    old_o, _, _ = legacy.sheet_snapshot(ws_o, 'order')
    old_i, _, _ = legacy.sheet_snapshot(ws_i, 'invoice')
    new_order_ids = {r.contract for r in new_orders}
    new_invoice_ids = {(r.contract, r.invoice) for r in new_invoices}
    changes, completed, warnings = [], [], []
    for draft in drafts:
        key = draft.contract if draft.kind == 'order' else (draft.contract, draft.invoice)
        indexed = old_o if draft.kind == 'order' else old_i
        is_new = key in (new_order_ids if draft.kind == 'order' else new_invoice_ids)
        label = f'{"订单" if draft.kind == "order" else "发货"} {draft.contract}' + (f'/{draft.invoice}' if draft.invoice else '')
        sheet = '订单' if draft.kind == 'order' else '发货'
        if is_new:
            changes.append(NoteChange(draft, sheet, '', draft.text, '新增记录'))
            continue
        candidates = indexed.get(key, [])
        if len(candidates) != 1:
            warnings.append(f'【待写备注暂存】{label}：Excel 中对应记录数为 {len(candidates)}，'
                            '暂不写入备注，请检查 PDF / Excel 或手动填写。')
            continue
        rr = candidates[0]['row']
        cell = wb.Worksheets(sheet).Cells(rr, 7 if draft.kind == 'order' else 12)
        if getattr(cell, 'HasFormula', False):
            warnings.append(f'【待写备注暂存】{label}：备注单元格包含公式，禁止覆盖。')
            continue
        old = str(cell.Value2 if cell.Value2 is not None else '').strip()
        final = _note_destination_text(old, draft.text)
        if final == old:
            completed.append(draft)
        elif len(final) > 32767:
            warnings.append(f'【待写备注暂存】{label}：备注超出 Excel 单元格限制。')
        else:
            changes.append(NoteChange(draft, sheet, old, final, f'已有记录第 {rr} 行'))
    return tuple(changes), tuple(completed), warnings


def _apply_notes_to_workbook(wb, changes: tuple[NoteChange, ...]):
    """Apply note changes AFTER all row insertions, to locate shifted rows safely."""
    if not changes:
        return
    order_index, _, _ = legacy.sheet_snapshot(wb.Worksheets('订单'), 'order')
    invoice_index, _, _ = legacy.sheet_snapshot(wb.Worksheets('发货'), 'invoice')
    for change in changes:
        draft = change.draft
        if draft.kind == 'order':
            rows = order_index.get(draft.contract, [])
        else:
            rows = invoice_index.get((draft.contract, draft.invoice), [])
        if len(rows) != 1:
            raise RuntimeError(f'无法唯一定位备注写入行：{change.sheet} {draft.contract}/{draft.invoice}')
        cell = wb.Worksheets(change.sheet).Cells(rows[0]['row'], 7 if draft.kind == 'order' else 12)
        if getattr(cell, 'HasFormula', False):
            raise RuntimeError('备注单元格包含公式，停止写入')
        current = str(cell.Value2 if cell.Value2 is not None else '').strip()
        if current != change.old:
            raise RuntimeError(f'备注内容发生变化：{change.sheet} {draft.contract}/{draft.invoice}；请重新预览')
        cell.NumberFormat = '@'
        cell.Value2 = change.final


def make_plan(rows: list[ScanRow], target_path: str, salesperson: str,
              source_specs=None, exclude_names=None, archived_paths=None, pending_notes=None, note_store_path=None, note_profile="") -> BusinessPlan:
    if not rows:
        raise ValueError('请先扫描 PDF 文件。')
    target = Path(target_path.strip()).expanduser().resolve()
    if target.suffix.lower() != '.xlsx' or target.name.startswith('~$'):
        raise ValueError('请选择本员工的 .xlsx 汇总文件（允许指定尚不存在的新文件）。')
    if not target.parent.is_dir():
        raise ValueError('目标 Excel 的父文件夹不存在。')
    records, warnings, fallback, delivery, invoices_seen, hashes, cancelled = _input_records(rows, salesperson, archived_paths)
    if not records and not fallback and not archived_paths:
        raise ValueError('没有能识别的订单或发票；请先排除识别问题。')
    originals = [r for r in rows if r.kind == '目录异常']
    if originals:
        raise ValueError('至少有一个来源目录无法访问，禁止在扫描不完整时核对或写入。')
    fingerprint = file_hash(target) if target.is_file() else None
    pending_notes = list(pending_notes or [])
    def preview_in_book(wb):
        results = _plan_with_book(wb, records, warnings, fallback, delivery, invoices_seen)
        changes, completed, note_warnings = _plan_notes(wb, pending_notes, results[0], results[1])
        return results, changes, completed, note_warnings
    (new_o, new_i, issues, notices, skipped, excel_invoiced_contracts), note_changes, note_completed, note_warnings = _with_excel(
        target, True, preview_in_book)
    issues += note_warnings
    specs = tuple((str(d['path']), str(d.get('hint', '自动识别'))) for d in (source_specs or []))
    inventory = frozenset(str(Path(r.source).resolve()) for r in rows if r.kind != '目录异常')
    return BusinessPlan(target, target.is_file(), fingerprint, hashes, new_o, new_i, issues,
                        notices, skipped, len(rows), len(delivery), cancelled, datetime.now(),
                        source_specs=specs, exclude_names=tuple(exclude_names or ()), scanned_paths=inventory,
                        invoiced_contracts=excel_invoiced_contracts, note_changes=note_changes, note_completed=note_completed,
                        note_snapshot=tuple(pending_notes), note_store_path=Path(note_store_path) if note_store_path else None,
                        note_profile=note_profile)


def format_preview(plan: BusinessPlan) -> str:
    lines = [f'目标 Excel：{plan.target}',
             f'待新增订单：{len(plan.orders)}  |  待新增发票：{len(plan.invoices)}',
             f'异常/需复核：{len(plan.warnings)}  |  已存在：{len(plan.skipped)}',
             f'发货单：{plan.scanned_delivery}  |  忽略作废：{plan.cancelled}', '',
             '==== 订单（待新增） ====']
    for r in plan.orders:
        lines.append(f'{r.document_date:%d.%m.%Y} | {r.contract} | {r.customer} | {r.gross} / Netto {r.net} | {r.file}')
    lines.append('\n==== 发票（待新增） ====')
    for r in plan.invoices:
        lines.append(f'{r.document_date:%d.%m.%Y} | {r.contract} | {r.invoice} | {r.gross} / Netto {r.net} | {r.file}')
    lines.append('\n==== 待写 Excel 备注（需审核） ====')
    if plan.note_changes:
        for change in plan.note_changes:
            draft = change.draft
            ident = draft.contract + ('/' + draft.invoice if draft.invoice else '')
            lines.append(f'{change.sheet} {ident} | {change.location} | 原备注：{change.old or "（空）"} | 写入后：{change.final}')
    else:
        lines.append('没有可安全自动写入的备注')
    if plan.note_completed:
        lines.append('已包含相同备注、不需重复写入：' + ', '.join(d.contract + (('/' + d.invoice) if d.invoice else '') for d in plan.note_completed))
    lines.append('\n==== 异常 ====')
    lines += plan.warnings or ['无']
    lines.append('\n==== 提示 ====')
    lines += plan.notices or ['无']
    return '\n'.join(lines)


def _check_input_unchanged(plan: BusinessPlan):
    # Any notes edited/added/deleted in another instance invalidate an old preview.
    if plan.note_store_path is not None:
        current = tuple(NoteStore(plan.note_store_path, plan.note_profile).list_drafts())
        if current != plan.note_snapshot:
            raise RuntimeError('本地备注草稿在上次预览后发生了变化，请重新核对，再确认写入。')
    if plan.source_specs:
        from gui_scanner import iter_files, norm, DEFAULT_EXCLUDED
        excluded = set(DEFAULT_EXCLUDED) | {norm(x) for x in plan.exclude_names}
        current = set()
        for pdf, root, hint in iter_files([{'path': path, 'hint': hint} for path,hint in plan.source_specs], excluded):
            if pdf is None:
                raise RuntimeError('某个 PDF 来源目录无法访问，必须重新扫描。')
            current.add(str(pdf.resolve()))
        if current != plan.scanned_paths:
            raise RuntimeError('来源文件夹在预览后新增或移除了 PDF，请重新扫描以避免遗漏。')
    if (file_hash(plan.target) if plan.target.is_file() else None) != plan.workbook_fingerprint:
        raise RuntimeError('目标 Excel 在预览后发生了变化，请重新预览并确认。')
    for name, value in plan.source_fingerprints.items():
        p = Path(name)
        if not p.is_file() or file_hash(p) != value:
            raise RuntimeError(f'来源 PDF 在预览后发生变化，禁止写入：{p}')


def _current_plan_keys(new_o, new_i):
    return ({(r.contract, r.document_date, r.gross, r.net, r.sha) for r in new_o},
            {(r.contract, r.invoice, r.document_date, r.gross, r.net, r.sha) for r in new_i})


def dated_output_target(path: Path, on: datetime | None = None) -> Path:
    """Only rewrite an explicit trailing yyyy-mm-dd-like update date in a name.

    No assumptions about employee name; no renames outside Excel file names.
    """
    match = re.search(r'更新至\d{1,2}\.\d{1,2}\.\d{4}(?=\.xlsx$)', path.name, re.I)
    if not match:
        return path
    today = (on or datetime.now()).strftime('%d.%m.%Y')
    return path.with_name(path.name[:match.start()] + '更新至' + today + path.name[match.end():])


def apply_confirmed(plan: BusinessPlan, rename_date: bool = False) -> tuple[Path, Path | None]:
    """Only call after an explicit positive GUI confirmation. Never schedule."""
    _check_input_unchanged(plan)
    target = plan.target
    output = dated_output_target(target) if rename_date else target
    if output != target and output.exists():
        raise RuntimeError(f'今日日期的 Excel 已存在：{output}，停止避免覆盖。')
    stage = None
    backup = None
    temp_ref = {}
    lock_fd = None
    lock_acquired = False
    try:
        # Excel typically creates a sibling ~$.xlsx owner lock while someone
        # has the workbook open. Do not overwrite an in-use human workbook.
        if target.is_file():
            excel_owner_lock = target.with_name('~$' + target.name[2:])
            if excel_owner_lock.exists():
                raise RuntimeError('检测到 Excel 工作簿仍被打开。请先在 Excel 中保存并关闭该文件，再重新核对和确认写入。')
        # Independent per-target lock. O_EXCL prevents simultaneous GUI writers.
        lock_path = target.with_name('.' + target.name + '.hangcha.lock')
        lock_fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        lock_acquired = True
        os.write(lock_fd, str(os.getpid()).encode('ascii'))
        os.close(lock_fd)
        lock_fd = None
        _check_input_unchanged(plan)
        def do_write(wb):
            nonlocal stage, backup
            legacy.prepare_workbook(wb)
            # Abort if someone has added a row to the workbook since preview.
            order_index, _, _ = legacy.sheet_snapshot(wb.Worksheets('订单'), 'order')
            invoice_index, _, _ = legacy.sheet_snapshot(wb.Worksheets('发货'), 'invoice')
            for r in plan.orders:
                if r.contract in order_index:
                    raise RuntimeError(f'订单 {r.contract} 已存在，请重新预览。')
            for r in plan.invoices:
                if (r.contract, r.invoice) in invoice_index:
                    raise RuntimeError(f'发票 {r.invoice} 已存在，请重新预览。')
            if target.exists():
                backups = target.parent / 'Hangcha_Backups'
                backups.mkdir(parents=True, exist_ok=True)
                backup = backups / f'{datetime.now():%Y%m%d_%H%M%S_%f}_{target.name}'
                shutil.copy2(target, backup)
                # Ensure the backup is exactly the version approved in preview.
                if file_hash(backup) != plan.workbook_fingerprint:
                    raise RuntimeError('Excel 在制作备份时发生变化；已停止写入，请重新核对。')
            for r in sorted(plan.orders, key=lambda x: (x.document_date, x.contract)):
                legacy.append_rec(wb.Worksheets('订单'), r)
            for r in sorted(plan.invoices, key=lambda x: (x.document_date, x.contract, x.invoice)):
                legacy.append_rec(wb.Worksheets('发货'), r)
            _apply_notes_to_workbook(wb, plan.note_changes)
            temp_ref['tmp'] = output.with_name(f'~hangcha_stage_{os.getpid()}_{datetime.now():%H%M%S%f}.xlsx')
            stage = temp_ref['tmp']
            wb.SaveAs(str(stage), FileFormat=51)
            wb.Close(SaveChanges=False)  # Release Windows file handles before atomic replacement
            if not legacy.safe_zip(stage):
                raise RuntimeError('生成的新 Excel 无法通过 XLSX 结构检查；原文件未更改。')
            _check_input_unchanged(plan)
            os.replace(stage, output)
            stage = None
            if output != target and target.exists():
                target.unlink()  # Old version already backed up above; only after successful replacement.
            return output
        saved = _with_excel(target, False, do_write)
        return saved, backup
    finally:
        if stage is not None:
            try: stage.unlink(missing_ok=True)
            except Exception: pass
        if lock_fd is not None:
            try: os.close(lock_fd)
            except Exception: pass
        if lock_acquired:
            try: lock_path.unlink(missing_ok=True)
            except Exception: pass


def save_preview_report(plan: BusinessPlan) -> Path:
    """Write a local audit report without touching any business PDF or Excel."""
    base = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'HangchaSalesAssistant' / 'reports'
    base.mkdir(parents=True, exist_ok=True)
    path = base / f'{datetime.now():%Y%m%d_%H%M%S_%f}_business_preview.txt'
    path.write_text(format_preview(plan), encoding='utf-8-sig')
    return path


def save_write_report(plan: BusinessPlan, output: Path, backup: Path | None) -> Path:
    base = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'HangchaSalesAssistant' / 'reports'
    base.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now()
    path = base / f'{stamp:%Y%m%d_%H%M%S_%f}_business_applied.txt'
    path.write_text(f'写入时间：{stamp.isoformat(timespec="seconds")}\n'
                    f'确认写入输出文件：{output}\n备份路径：{backup}\n\n'
                    + format_preview(plan), encoding='utf-8-sig')
    return path
