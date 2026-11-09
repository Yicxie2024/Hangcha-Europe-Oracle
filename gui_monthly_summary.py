"""Weekly 发货 workbook -> invoice-PDF identity matching for monthly *internal* reports.

The weekly spreadsheet is the sole eligibility index: selected month and equal
invoice gross/Netto. PDF identity, monetary amounts, destination, and line items
are checked only for indexed invoice numbers. This module never writes sources.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import unicodedata

from automation import EXCEL_HEADERS, normalize_id, money_to_decimal, extract_invoice_header
from gui_monthly_core import Invoice, Outcome, parse_invoice, EURO
from gui_monthly_xlsx import read_sheet
from gui_monthly_review import ReviewEntry, fingerprint

INVOICE_ID_RE = re.compile(r'(?<![A-Z0-9])H\d{8}(?![A-Z0-9])', re.I)
# A fixed-id spreadsheet column is the primary source; use direct matching
# and NOT fuzzy filenames containing a contract/order number.
SUMMARY_COLS = tuple('ABCDEFGHIJKL')
DATE_COLUMNS = ('I', 'G', 'F')  # invoice sent, shipped, expected-delivery (all invoice date in app)


@dataclass(frozen=True)
class SummaryInvoice:
    number: str
    gross: Decimal
    net: Decimal
    day: date
    excel_row: int
    contract: str


def _summary_date(value: str) -> date | None:
    s = str(value or '').strip()
    if not s:
        return None
    # All dates that the weekly automation writes are plain text dd.mm.yyyy.
    for fmt in ('%d.%m.%Y', '%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d'):
        try:
            return datetime.strptime(s.split(' ')[0], fmt).date()
        except ValueError:
            pass
    # Date cells in Excel may be real serials. Guard against non-date numbers.
    try:
        serial = float(s)
        if 36526 <= serial <= 73050 and serial.is_integer():
            return date(1899,12,30) + timedelta(days=int(serial))
    except ValueError:
        pass
    return None


def _summary_money(value: str) -> Decimal | None:
    s = str(value or '').strip()
    if not s:
        return None
    if s.startswith('#'):
        return None
    # OOXML numeric cached values are most commonly dot-decimal, with no group
    # separators; keep German-formatted user-entered strings supported too.
    try:
        if re.fullmatch(r'-?\d+(?:\.\d+)?', s):
            return Decimal(s).quantize(Decimal('0.01'))
        return money_to_decimal(s)
    except (ValueError, InvalidOperation, ArithmeticError):
        return None


def read_eligible_invoices(workbook: str | Path, year: int, month: int) -> tuple[list[SummaryInvoice], list[Outcome]]:
    """Only user-approved weekly 发货 rows may become candidate PDF lookups."""
    if not (2000 <= year <= 2100 and 1 <= month <= 12):
        raise ValueError('统计月份无效')
    path=Path(workbook)
    if not path.is_file():
        raise ValueError('业务周汇总 Excel 不存在')
    rows=read_sheet(path,'发货',SUMMARY_COLS)
    if not rows:
        raise ValueError('业务周汇总 Excel 的「发货」工作表为空')
    expected=EXCEL_HEADERS['发货']
    actual=[str(rows[0].get(chr(65+i),'')).strip() for i in range(len(expected))]
    if actual != expected:
        raise ValueError('发货子表表头与当前业务周汇总模板不符；未继续筛选。\n预期：'
                         +'、'.join(expected)+'\n实际：'+'、'.join(actual))
    candidates=[];logs=[];grouped=defaultdict(list)
    all_numbers=[normalize_id(row.get('C','')) for row in rows[1:]]
    repetitions=Counter(n for n in all_numbers if INVOICE_ID_RE.fullmatch(n))
    warned_duplicates=set()
    for line,row in enumerate(rows[1:],2):
        inv=normalize_id(row.get('C',''))
        if not inv:continue
        dates=[(column,_summary_date(row.get(column,''))) for column in DATE_COLUMNS]
        dates=[(column,value) for column,value in dates if value is not None]
        # Month filtering BEFORE creating display logs. Rows with no usable
        # date cannot be confidently assigned to the selected month.
        if not dates or not any((d.year,d.month)==(year,month) for _,d in dates):
            continue
        if not INVOICE_ID_RE.fullmatch(inv):
            logs.append(Outcome(str(path),inv,'待复核',f'发货表第 {line} 行发票号格式不正确',0));continue
        if repetitions[inv]>1:
            if inv not in warned_duplicates:
                warned_duplicates.add(inv)
                logs.append(Outcome(str(path),inv,'待复核',
                            f'发货表同号发票出现 {repetitions[inv]} 行，需复核重复记录',0))
            continue
        day=dates[0][1]
        if any(d!=day for _,d in dates):
            logs.append(Outcome(str(path),inv,'待复核',f'发货表第 {line} 行 F/G/I 日期不一致',0));continue
        if (day.year,day.month)!=(year,month):continue
        gross=_summary_money(row.get('D',''))
        net=_summary_money(row.get('E',''))
        if gross is None or net is None:
            logs.append(Outcome(str(path),inv,'待复核',f'发货表第 {line} 行金额或 Netto 无效',0));continue
        if gross<=0 or net<=0:
            logs.append(Outcome(str(path),inv,'待复核',f'发货表第 {line} 行金额为零或负值',0));continue
        if gross!=net:
            logs.append(Outcome(str(path),inv,'排除',f'发货表第 {line} 行金额 {gross} != Netto {net}',0));continue
        grouped[inv].append(SummaryInvoice(inv,gross,net,day,line,normalize_id(row.get('B',''))))
    for inv,items in sorted(grouped.items()):
        if len(items)>1:
            for item in items:
                logs.append(Outcome(str(path),inv,'待复核',f'发货表第 {item.excel_row} 行：同一发票号多行，需人工核对是否重复',0))
        else:
            candidates.append(items[0])
    return candidates,logs


def _invoice_header_quick(path: Path) -> str:
    """Fallback for a PDF filename with no invoice number: only its header."""
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        text='\n'.join((pg.extract_text() or '') for pg in pdf.pages[:2])
    try:
        number,_=extract_invoice_header(text)
        return normalize_id(number)
    except (ValueError,IndexError):
        return ''


def _index_pdfs(folder: Path, wanted: set[str]) -> tuple[dict[str,list[Path]], list[Path]]:
    names=defaultdict(list); unknown=[]
    for p in sorted(folder.rglob('*')):
        if not p.is_file() or p.suffix.casefold()!='.pdf':
            continue
        if '作废' in unicodedata.normalize('NFKC',p.name):
            continue
        matches={m.upper() for m in INVOICE_ID_RE.findall(p.name)}
        for number in matches & wanted:
            names[number].append(p)
        if not (matches & wanted):unknown.append(p)  # fall back to PDF header if filename is missing or misleading
    return names,unknown


def scan_month_for_review(workbook: str|Path, pdf_folder: str|Path, year:int, month:int) -> list[ReviewEntry]:
    """Build ONE review entry per invoice in the selected month only.

    Only weekly summary invoice numbers are considered. For suspect PDFs,
    known-good parsed items remain available for explicit manual confirmation,
    but only when identity, amount, date and non-Swiss destination are verified.
    """
    root=Path(pdf_folder)
    if not root.is_dir():raise ValueError('发票 PDF 根目录不存在')
    candidates, summary_logs = read_eligible_invoices(workbook,year,month)
    entries={}
    for log in summary_logs:
        if not log.invoice:continue
        entries[log.invoice]=ReviewEntry(log.invoice,log.state,log.reason,log.file,log.count,
                                         fingerprint=fingerprint(log.invoice,log.state,log.reason))
    if not candidates:return sorted(entries.values(),key=lambda e:e.invoice)
    wanted={c.number for c in candidates}
    names,unknown=_index_pdfs(root,wanted)
    missing=wanted-set(names)
    if missing:
        for p in unknown:
            try:number=_invoice_header_quick(p)
            except Exception:continue
            if number in missing:names[number].append(p)
    for c in sorted(candidates,key=lambda q:(q.day,q.number)):
        number=c.number
        context=f'{c.excel_row}|{c.day.isoformat()}|{c.gross}|{c.net}|{c.contract}'
        paths=names.get(number,[])
        def add(state,reason,file,record=None,can_confirm=False):
            entries[number]=ReviewEntry(number,state,reason,str(file),
                    len(record.items) if record else 0,record,can_confirm,
                    fingerprint(number,state,reason,context,record))
        if not paths:
            add('待复核',f'发货表第 {c.excel_row} 行：未找到发票号对应 PDF',workbook)
            continue
        valid=[];errors=[]
        for p in paths:
            if '有误' in unicodedata.normalize('NFKC',p.name):
                errors.append(f'{p.name}: 文件名标记有误');continue
            try:r=parse_invoice(p)
            except Exception as ex:
                errors.append(f'{p.name}: PDF解析失败：{str(ex)[:130]}');continue
            if r.invoice.upper()!=number:
                errors.append(f'{p.name}: PDF内 Invoice No. 与汇总表不一致');continue
            if r.day!=c.day:
                errors.append(f'{p.name}: PDF日期 {r.day} 与汇总表 {c.day} 不一致');continue
            if EURO(r.gross)!=c.gross or EURO(r.net)!=c.net:
                errors.append(f'{p.name}: PDF金额 {r.gross}/{r.net} 与汇总表 {c.gross}/{c.net} 不一致');continue
            valid.append(r)
        # If any candidate file has a conflict, do not select a different PDF
        # as a silent workaround. This must be resolved at the file level.
        if errors:
            add('待复核','；'.join(errors)[:520],paths[0]);continue
        if not valid:
            add('待复核','没有可用的发票 PDF',paths[0]);continue
        if len({r.digest for r in valid})>1:
            add('待复核','同一发票号有多个内容不同的 PDF，必须先解决版本冲突',paths[0]);continue
        r=valid[0]
        if r.country=='SCHWEIZ':
            add('排除','瑞士目的地',r.path,r,False);continue
        if not r.country:
            add('待复核','PDF目的地不明确：'+'；'.join(r.notes),r.path,r,False);continue
        if not r.items:
            add('待复核','未识别到可纳入的货物：'+'；'.join(r.notes),r.path,r,False);continue
        if r.notes:
            add('待复核',f'识别 {len(r.items)} 行；'+'；'.join(r.notes),r.path,r,True)
        else:
            add('纳入','汇总表与 PDF 编号、日期、金额、国家核对通过',r.path,r,True)
    return sorted(entries.values(),key=lambda e:(e.record.day.isoformat() if e.record else '',e.invoice))


def scan_month_from_summary(workbook: str|Path, pdf_folder: str|Path,
                            year:int, month:int) -> tuple[list[Invoice],list[Outcome]]:
    """Backward-compatible API for code using the previous auto-only flow."""
    from gui_monthly_review import outcome_rows, selected_invoices
    entries=scan_month_for_review(workbook,pdf_folder,year,month)
    return selected_invoices(entries),outcome_rows(entries)
