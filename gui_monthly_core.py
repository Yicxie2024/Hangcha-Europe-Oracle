"""Read-only monthly *internal* dispatch report extraction (v1.5.19).

Policy: invoice date within selected month, explicitly labelled invoice NET equals TOTAL,
non-Swiss destination (known country), non-cancelled invoice. Missing/ambiguous facts
are REVIEW, never filled with guessed money, item codes or quantities.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable
import hashlib
import re
import unicodedata

from automation import extract_totals, extract_invoice_header, money_to_decimal

MONEY = Decimal('0.01')
SKIP_CODES = {'999-830','999-831','X10019'}
COUNTRIES = {
 'Germany':'Deutschland', 'Deutschland':'Deutschland', 'Denmark':'Denmark', 'Danmark':'Denmark',
 'Austria':'Österreich', 'Österreich':'Österreich', 'Poland':'Polen', 'Polen':'Polen',
 'Netherlands':'Niederlande','Nederland':'Niederlande','Belgium':'Belgien',
 'Belgien':'Belgien','France':'Frankreich','Frankreich':'Frankreich',
 'Italy':'Italien','Italia':'Italien','Spain':'Spanien','España':'Spanien',
 'Sweden':'Schweden','Sverige':'Schweden','Norway':'Norwegen','Norge':'Norwegen',
 'Finland':'Finnland','Suomi':'Finnland','Czech Republic':'Tschechien',
 'Czechia':'Tschechien','Slovakia':'Slowakei','Slovenia':'Slowenien',
 'Hungary':'Ungarn','Romania':'Rumänien','Portugal':'Portugal',
 'Ireland':'Irland','Croatia':'Kroatien','Greece':'Griechenland',
 'Luxembourg':'Luxemburg','United Kingdom':'Vereinigtes Königreich',
 'Great Britain':'Vereinigtes Königreich',
 'Switzerland':'SCHWEIZ','Schweiz':'SCHWEIZ','Suisse':'SCHWEIZ','Svizzera':'SCHWEIZ',
}
COUNTRY_RE = re.compile(r'(?i)(?<![A-Za-zÀ-ÿ])('+'|'.join(re.escape(k) for k in sorted(COUNTRIES,key=len,reverse=True))+r')(?![A-Za-zÀ-ÿ])')
AMOUNT_TOKEN = r'-?\d[\d.,]*[.,]\d{2}'
# First two numeric columns in each item row are qty and ERP; final columns are
# fixed money amounts. A separate percentage may appear before the line total.
ITEM_RE = re.compile(rf'^\s*(\d+(?:[.,]\d+)?)\s+(\S+)\s+(\S+)\s+(.*?)\s+({AMOUNT_TOKEN})(?:\s+(\d+(?:[.,]\d+)?))?\s+({AMOUNT_TOKEN})(?:\s+netto)?\s*$',re.I)
FREIGHT_RE = re.compile(rf'^\s*(\d+(?:[.,]\d+)?)\s+(999-830|999-831|X10019)\s+(.+?)\s+({AMOUNT_TOKEN})(?:\s+netto)?\s+({AMOUNT_TOKEN})\s*$',re.I)
ITEM_START = re.compile(r'^\s*\d+(?:[.,]\d+)?\s+(?:\d{2}\.\d{2}\.\d{2}\.\d{5,}|[A-Z]\d{5,}|999-\d{3})\s+',re.I)
HEAD_BEGIN = re.compile(r'^\s*Qty\.?\s+ERP\s+No\.',re.I)
END_TABLE = re.compile(r'^\s*(?:Total\s+quantity|Your\s+V\.?A\.?T|Payment\s*\(|Excluding VAT|Banktransfer\.|Hangcha\s+Europe\s+GmbH\.|IBAN)',re.I)
EURO = lambda value: value.quantize(MONEY,rounding=ROUND_HALF_UP)

@dataclass
class InvoiceItem:
    erp: str
    part: str
    description: str
    quantity: Decimal
    unit_price: Decimal
    line_total: Decimal
    line_discount: Decimal

@dataclass
class Invoice:
    invoice: str
    day: date
    customer: str
    country: str
    vat_id: str
    net: Decimal
    gross: Decimal
    freight: Decimal | None
    discount: Decimal | None
    items: list[InvoiceItem]
    path: str
    digest: str
    notes: list[str] = field(default_factory=list)

@dataclass
class Outcome:
    file: str
    invoice: str
    state: str
    reason: str
    count: int = 0


def _decimal(t: str) -> Decimal:
    return money_to_decimal(t)


def _determine_destination(text: str) -> tuple[str, str]:
    """Only the header address zone, never an arbitrary footer or 'sent to' prose.

    Mixed bill-to and ship-to countries are ambiguous and queued for review.
    """
    prefix=text.split('Invoice\n',1)[0]
    if not prefix or len(prefix)>3500: prefix=text[:2500]
    hits={COUNTRIES[m.group(0).casefold().title()] if m.group(0).casefold().title() in COUNTRIES else
          next((v for k,v in COUNTRIES.items() if k.casefold()==m.group(0).casefold()),'')
          for m in COUNTRY_RE.finditer(prefix)}
    hits.discard('')
    if len(hits)==1:return next(iter(hits)),''
    if not hits:return '','目的地国家未能从发票地址区可靠识别'
    return '','地址区存在多个不同国家（客户地址与送达地址可能不同），请人工核对'


def _customer_name(path: Path) -> str:
    try:
        from gui_customer import extract_invoice_customer
        return extract_invoice_customer(path) or ''
    except (OSError, ValueError, ImportError):
        return ''


def _parse_items(text: str) -> tuple[list[InvoiceItem], Decimal | None, list[str]]:
    items=[]; freight=[]; errors=[]; in_table=False; unparsed=[]
    for original in text.splitlines():
        line=' '.join(original.split())
        if HEAD_BEGIN.search(line): in_table=True;continue
        if not in_table:continue
        if END_TABLE.search(line):break
        match=FREIGHT_RE.match(line)
        if match:
            try:freight.append(_decimal(match.group(5)))
            except ValueError:errors.append('运费金额无法识别')
            continue
        m=ITEM_RE.match(line)
        if m:
            qty_s,erp,part,desc,unit_s,percent_s,total_s=m.groups()
            if erp.upper() in SKIP_CODES:continue
            try:
                qty=Decimal(qty_s.replace(',','.'))
                unit=_decimal(unit_s); total=_decimal(total_s)
                discount=EURO(qty*unit-total)
                if qty<=0 or unit<0 or total<0 or discount<Decimal('-0.03'):
                    raise ValueError('金额/数量无效')
                if percent_s is not None:
                    pct=Decimal(percent_s.replace(',','.'))
                    # Percent column can be informative only; authoritative money
                    # remains the explicitly printed line total.
                    if pct<0 or pct>100:errors.append(f'配件 {part} 折扣百分比异常')
                items.append(InvoiceItem(erp,part,desc,qty,unit,total,max(Decimal('0'),discount)))
            except (ValueError,ArithmeticError):errors.append('配件行数字无法解析：'+line[:85])
            continue
        if ITEM_START.search(line):unparsed.append(line[:115])
        # Some Hangcha descriptions wrap to next line. Append text only if it
        # contains no quantity / price and looks like the trailing part ID.
        elif items and re.fullmatch(r'[A-Z0-9][A-Z0-9./_-]{4,}',line,re.I):
            items[-1].description+=' '+line
    if unparsed:errors.extend('疑似配件行未识别：'+x for x in unparsed[:10])
    if not items:errors.append('未识别到有效配件明细')
    return items,sum(freight,Decimal('0')) if freight else None,errors


def parse_invoice(path: str | Path) -> Invoice:
    import pdfplumber
    p=Path(path)
    if '作废' in unicodedata.normalize('NFKC',p.name):raise ValueError('已作废的发票')
    if '有误' in unicodedata.normalize('NFKC',p.name):raise ValueError('文件名标记有误，需人工核查后重开正确发票')
    stat=(p.stat().st_size,p.stat().st_mtime_ns)
    with pdfplumber.open(p) as pdf:
        text='\n'.join(page.dedupe_chars(tolerance=1.5).extract_text() or '' for page in pdf.pages)
    if (p.stat().st_size,p.stat().st_mtime_ns)!=stat:raise ValueError('PDF正在更新，已取消')
    if re.search(r'(?i)\bcredit\s+note\b|\bstorno\b|\bgutschrift\b',text):raise ValueError('贷项/冲销单，不自动统计')
    if not re.search(r'\bInvoice\s+No\.',text,re.I):raise ValueError('不是已支持的发票格式')
    inv,day=extract_invoice_header(text)
    net,vat,gross=extract_totals(text)
    country,issue=_determine_destination(text)
    items,freight,errors=_parse_items(text)
    vat_match=re.search(r'Your\s+V\.?A\.?T\.?\s+number\s*:\s*([A-Z]{2}[A-Z0-9 ]{6,20})',text,re.I)
    if not vat_match:vat_match=re.search(r'(?m)^VAT\s*:\s*([A-Z]{2}[A-Z0-9 ]{6,20})',text,re.I)
    vat_id=re.sub(r'\s+','',vat_match.group(1)) if vat_match else ''
    if not vat_id:errors.append('发票缺少明确客户 VAT No.，需补充并核查')
    customer=_customer_name(p)
    if not customer:errors.append('客户名称未能从发票地址块可靠提取')
    if issue:errors.append(issue)
    parts_total=sum((r.line_total for r in items),Decimal('0'))
    adjustment=EURO(net-parts_total-(freight or Decimal('0')))
    if not errors and adjustment>Decimal('0.03'):
        errors.append(f'Netto 与配件金额+运费差 {adjustment} EUR，可能存在额外收费，请复核')
    if adjustment<Decimal('-0.03'):
        errors.append(f'Netto 与配件金额+运费差 {adjustment} EUR，可能存在未识别折扣，请复核')
    return Invoice(inv,day,customer,country,vat_id,net,gross,freight,
                   sum((r.line_discount for r in items),Decimal('0')),
                   items,str(p),hashlib.sha256(p.read_bytes()).hexdigest(),errors)


def scan_month(folder: str | Path, year: int, month: int) -> tuple[list[Invoice], list[Outcome]]:
    if year<2000 or year>2100 or month<1 or month>12:raise ValueError('月份无效')
    root=Path(folder)
    if not root.is_dir():raise ValueError('发票来源文件夹不存在')
    selected=[]; outcomes=[]; grouped={}
    for p in sorted(root.rglob('*.pdf')):
        try:
            rec=parse_invoice(p)
            if (rec.day.year,rec.day.month)!=(year,month):
                outcomes.append(Outcome(str(p),rec.invoice,'排除','非选择月份'))
                continue
            if rec.country=='SCHWEIZ':
                outcomes.append(Outcome(str(p),rec.invoice,'排除','瑞士订单'))
                continue
            if EURO(rec.net)!=EURO(rec.gross):
                outcomes.append(Outcome(str(p),rec.invoice,'排除','发票Total与Netto不相等'))
                continue
            if rec.net<=0:
                outcomes.append(Outcome(str(p),rec.invoice,'待复核','发票金额为0'))
                continue
            if not rec.country:
                outcomes.append(Outcome(str(p),rec.invoice,'待复核','目的地不明确：'+'；'.join(rec.notes)))
                continue
            grouped.setdefault(rec.invoice,[]).append(rec)
        except Exception as ex:
            # Other file kinds / unusable source are logged, not silently added.
            reason=str(ex).replace('\n',' ')[:230]
            outcomes.append(Outcome(str(p),'','跳过' if reason in ('不是已支持的发票格式','已作废的发票') else '待复核',reason))
    for inv,copies in sorted(grouped.items()):
        hashes={x.digest for x in copies}
        if len(hashes)!=1:
            for x in copies:outcomes.append(Outcome(x.path,inv,'待复核','同一发票号存在不同PDF版本，避免重复统计'))
            continue
        chosen=copies[0]
        if chosen.notes:
            outcomes.append(Outcome(chosen.path,inv,'待复核','；'.join(chosen.notes),len(chosen.items)))
        else:
            selected.append(chosen)
            outcomes.append(Outcome(chosen.path,inv,'纳入','筛选通过',len(chosen.items)))
        for duplicate in copies[1:]:outcomes.append(Outcome(duplicate.path,inv,'排除','相同PDF重复文件，已去重'))
    selected.sort(key=lambda rec:(rec.day,rec.invoice))
    return selected,outcomes
