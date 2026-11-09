"""Read-only inventory lookup and Hangcha sales-order line extraction.

Inventory fields come exclusively from the chosen XLSX workbook. Sales-order
identifiers and quantities come exclusively from the PDF text/table geometry.
No Excel/PDF modification, stock reservations, or remote connections.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable
import posixpath
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
PKG_REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
INVENTORY_SHEET = '库存已有备件'


def normalize_number(value: str) -> str:
    """Exact, case-insensitive identifiers; only whitespace is normalized."""
    return ''.join(unicodedata.normalize('NFKC', str(value or '')).upper().split())


def readable_number(value: Decimal | None) -> str:
    if value is None:
        return '未提供'
    return str(int(value)) if value == value.to_integral_value() else format(value.normalize(), 'f')


def number_or_none(value: object) -> Decimal | None:
    s = str(value or '').strip()
    if not s:
        return None
    try:
        return Decimal(s.replace(',', '.'))
    except InvalidOperation:
        return None


@dataclass(frozen=True)
class StockRecord:
    row: int
    erp: str
    part: str
    rack: str
    stock: Decimal | None
    weight: str
    dimension: str
    description: str
    description_cn: str


@dataclass(frozen=True)
class SearchResult:
    method: str
    hits: tuple[StockRecord, ...]
    warning: str = ''

    @property
    def total(self) -> Decimal | None:
        if self.warning and ('不一致' in self.warning or '不同 ERP/Part' in self.warning):
            return None
        return sum((r.stock for r in self.hits), Decimal('0')) if self.hits and all(r.stock is not None for r in self.hits) else None


class InventoryIndex:
    def __init__(self, records: Iterable[StockRecord], path: Path):
        self.records = tuple(records)
        self.path = Path(path)
        self.by_erp: dict[str, list[StockRecord]] = defaultdict(list)
        self.by_part: dict[str, list[StockRecord]] = defaultdict(list)
        for r in self.records:
            if normalize_number(r.erp):
                self.by_erp[normalize_number(r.erp)].append(r)
            if normalize_number(r.part):
                self.by_part[normalize_number(r.part)].append(r)

    def search(self, code: str) -> SearchResult:
        key = normalize_number(code)
        if not key:
            return SearchResult('', ())
        erp = tuple(self.by_erp.get(key, ()))
        if erp:
            return SearchResult('ERP No.', erp, _duplicate_warning(erp))
        part = tuple(self.by_part.get(key, ()))
        return SearchResult('Part No.' if part else '', part, _duplicate_warning(part))

    def find_item(self, erp: str, part: str) -> SearchResult:
        """ERP is authoritative, but exact Part matches with BLANK ERP add stock rows.

        The workbook sometimes has a second stock line for the same part with
        an empty ERP. Ignoring it under-counts stock (NP 267073 / AMB155).
        Never combine a different nonempty ERP merely because Part matches.
        """
        key = normalize_number(erp)
        part_key = normalize_number(part)
        matches = tuple(self.by_erp.get(key, ())) if key else ()
        if matches:
            if part_key:
                agreeing = tuple(r for r in matches if normalize_number(r.part) == part_key)
                if agreeing:
                    # Attach only rows with an exactly matching Part No. AND
                    # no ERP. These rows can be distinguished by Excel row ID,
                    # and their MA values are separate stock records.
                    supplements = tuple(r for r in self.by_part.get(part_key, ())
                                        if not normalize_number(r.erp) and r not in agreeing)
                    hits = tuple(dict.fromkeys((*agreeing, *supplements)))
                    warn = _duplicate_warning(hits)
                    if supplements:
                        extra = f'同一 Part No. 还有 {len(supplements)} 条 ERP 为空的库存行，已单独列出其库位与库存'
                        warn = (warn + '；' if warn else '') + extra
                    return SearchResult('ERP + Part No.' if supplements else
                                        'ERP + Part 核对' if len(matches)>1 else 'ERP No.',
                                        hits, warn)
                return SearchResult('ERP No.', matches,
                    'ERP 已命中，但库存表 Part No. 与订单 PDF 不一致，请人工核对')
            return SearchResult('ERP No.', matches, _duplicate_warning(matches))
        part_matches = tuple(self.by_part.get(part_key, ())) if part_key else ()
        warn = _duplicate_warning(part_matches)
        if part_matches and key:
            warn = ('ERP 未命中，已改用 Part No. 查询。' + (' ' + warn if warn else ''))
        return SearchResult('Part No.' if part_matches else '', part_matches, warn)


def _duplicate_warning(hits: tuple[StockRecord, ...]) -> str:
    if not hits:
        return ''
    erps = {normalize_number(r.erp) for r in hits if normalize_number(r.erp)}
    parts = {normalize_number(r.part) for r in hits if normalize_number(r.part)}
    if len(erps)>1 or len(parts)>1:
        return '该编号对应不同 ERP/Part 组合，需人工确认匹配范围，勿直接合计为可用库存'
    return ''


def _sheet_xml_name(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = wb.find(NS+'sheets')
    node = next((s for s in sheets if s.get('name')==sheet_name), None) if sheets is not None else None
    if node is None:
        raise ValueError(f'库存 Excel 找不到工作表「{sheet_name}」')
    rid = node.get(RID)
    rel = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    target = next((r.get('Target') for r in rel if r.get('Id') == rid), None)
    if not target:
        raise ValueError('库存 Excel 工作表位置无法定位')
    filename = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('xl',target))
    if filename not in z.namelist():
        raise ValueError('库存 Excel 工作表 XML 不存在')
    return filename


def _shared_strings(z: zipfile.ZipFile) -> list[str]:
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    result = []
    with z.open('xl/sharedStrings.xml') as stream:
        for _, element in ET.iterparse(stream, events=('end',)):
            if element.tag == NS+'si':
                result.append(''.join(t.text or '' for t in element.iter(NS+'t')))
                element.clear()
    return result


def _cell_text(c: ET.Element, strings: list[str]) -> str:
    raw = c.find(NS+'v')
    if raw is not None:
        val = raw.text or ''
        if c.get('t') == 's':
            try: return strings[int(val)]
            except (IndexError, ValueError): return ''
        if c.get('t') == 'e':
            return ''
        return val
    inline = c.find(NS+'is')
    return ''.join(t.text or '' for t in inline.iter(NS+'t')) if inline is not None else ''


def load_inventory(path: str | Path) -> InventoryIndex:
    """Stream the 70+ MB sheet XML; never open/save/recalculate the workbook.

    MA stock often contains formulas. We read Excel's stored cached result as of
    its latest save; if missing, stock is UNKNOWN, not zero.
    """
    p = Path(path).expanduser()
    if not p.is_file() or p.suffix.lower() not in ('.xlsx', '.xlsm'):
        raise ValueError('请选择有效的 .xlsx/.xlsm 库存文件')
    records = []
    cols = {'A':'weight','B':'dimension','C':'rack','D':'erp','E':'part',
            'F':'description','G':'description_cn','MA':'stock'}
    with zipfile.ZipFile(p) as z:
        strings = _shared_strings(z)
        sheet = _sheet_xml_name(z, INVENTORY_SHEET)
        header = {}
        with z.open(sheet) as stream:
            for _, elem in ET.iterparse(stream, events=('end',)):
                if elem.tag != NS+'row':
                    continue
                idx = int(elem.get('r','0'))
                vals = {}
                for c in elem.findall(NS+'c'):
                    letter = re.sub(r'\d+', '', c.get('r',''))
                    if letter in cols:
                        vals[letter] = _cell_text(c, strings).strip()
                elem.clear()
                if idx == 1:
                    header = vals
                    if not ('ERP' in header.get('D','').upper() and 'PART' in header.get('E','').upper()
                            and header.get('C','').casefold().startswith('rack')
                            and header.get('MA','').strip() == '存'):
                        raise ValueError('库存表列位置与预期不同（需要 C=Rack No., D=ERP, E=Part, MA=存），已停止查询')
                    continue
                if not vals.get('D') and not vals.get('E'):
                    continue
                records.append(StockRecord(idx, vals.get('D',''), vals.get('E',''),
                    vals.get('C',''), number_or_none(vals.get('MA')),
                    vals.get('A',''), vals.get('B',''), vals.get('F',''), vals.get('G','')))
    if not records:
        raise ValueError('库存表中没有有效 ERP/Part 编号')
    return InventoryIndex(records,p)


@dataclass(frozen=True)
class DeliveryPart:
    erp: str
    part: str
    quantity: Decimal | None
    description: str
    freight: bool = False


def _group_line_words(words: list[dict], tolerance: float = 2.0) -> list[list[dict]]:
    groups: list[list[dict]] = []
    for word in sorted(words, key=lambda x:(float(x['top']),float(x['x0']))):
        if not groups or abs(float(word['top'])-float(groups[-1][0]['top'])) > tolerance:
            groups.append([word])
        else:
            groups[-1].append(word)
    return groups


def _joined(words: list[dict]) -> str:
    return ' '.join(w['text'].strip() for w in sorted(words,key=lambda x:x['x0']) if w['text'].strip()).strip()


def read_delivery_parts(path: str | Path) -> list[DeliveryPart]:
    """Extract table columns from PDF geometry; never parse part IDs from filename.

    Returns unique ERP/Part entries with combined demand for repeated lines.
    Requires a recognizable delivery table to avoid accidental other-document
    matching. Non-stock 'Standard Freight' is retained as an excluded item.
    """
    import pdfplumber
    path = Path(path)
    raw: list[DeliveryPart] = []
    is_delivery = False
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            pg = page.dedupe_chars(tolerance=1.5)
            words = pg.extract_words(x_tolerance=2, y_tolerance=2)
            lines = _group_line_words(words)
            body = pg.extract_text() or ''
            if re.search(r'\bDelivery\s*(?:order|note)\b', body, flags=re.I):
                is_delivery = True
            for li, line in enumerate(lines):
                texts=[w['text'].strip().lower().rstrip('.') for w in line]
                if not ('erp' in texts and 'part' in texts and 'qty' in texts):
                    continue
                erp_head = next((w for w in line if w['text'].strip().lower()=='erp'),None)
                part_head = next((w for w in line if w['text'].strip().lower()=='part'),None)
                desc_head = next((w for w in line if w['text'].strip().lower().startswith('description')),None)
                qty_head = next((w for w in line if w['text'].strip().lower().startswith('qty')),None)
                if not all((erp_head, part_head, desc_head, qty_head)):
                    continue
                erp_x, part_x, desc_x, qty_x = [float(w['x0']) for w in (erp_head,part_head,desc_head,qty_head)]
                if not (qty_x < erp_x < part_x < desc_x):
                    continue
                for next_line in lines[li+1:]:
                    top=float(next_line[0]['top'])
                    if top > pg.height - 110:
                        break
                    if any(w['text'].strip().lower().startswith(('hangcha','total')) for w in next_line if w['x0']<erp_x-15):
                        break
                    # Each real line starts with numeric Ordered/Qty; narrative
                    # order references are all in Description and ignored.
                    quantity_words=[w for w in next_line if qty_x-3 <= float(w['x0']) < erp_x-3]
                    erp_words=[w for w in next_line if erp_x-3 <= float(w['x0']) < part_x-3]
                    part_words=[w for w in next_line if part_x-3 <= float(w['x0']) < desc_x-3]
                    desc_words=[w for w in next_line if float(w['x0']) >= desc_x-3]
                    qty_s = _joined(quantity_words)
                    erp_s = _joined(erp_words)
                    part_s = _joined(part_words)
                    desc = _joined(desc_words)
                    if not erp_s or not re.fullmatch(r'\d+(?:[,.]\d+)?',qty_s):
                        continue
                    qty = number_or_none(qty_s)
                    is_freight = erp_s.strip() in {'999-830','999-831'} or bool(re.search(r'\b(?:freight|shipping|transport charge)\b',desc,re.I))
                    raw.append(DeliveryPart(erp_s, part_s, qty, desc, is_freight))
    if not is_delivery:
        raise ValueError('所选 PDF 未找到明确的 Delivery order/Delivery note 标题，请确认文件类型')
    if not raw:
        raise ValueError('找到发货单标题，但无法识别明细表中的 ERP No./Part No./Qty.，请人工查看 PDF')
    # Multiple lines for the same part must be checked against TOTAL demand.
    grouped: dict[tuple[str,str,bool],DeliveryPart] = {}
    for item in raw:
        key = (normalize_number(item.erp),normalize_number(item.part),item.freight)
        prev = grouped.get(key)
        if prev is None:
            grouped[key]=item
        else:
            qty = prev.quantity+item.quantity if prev.quantity is not None and item.quantity is not None else None
            grouped[key]=DeliveryPart(prev.erp,prev.part,qty,prev.description or item.description, prev.freight)
    return list(grouped.values())


@dataclass(frozen=True)
class OrderPart:
    """One stock-check demand from a PDF sales order (not a delivery note)."""
    erp: str
    part: str
    quantity: Decimal | None
    description: str
    freight: bool = False


def read_order_parts(path: str | Path, allow_incomplete: bool = False) -> list[OrderPart]:
    """Read actual order-detail table rows without mistaking footer labels for SKUs.

    Strict callers reject incomplete rows; the unified GUI may explicitly opt
    into placeholders for manual verification. The parser never guesses IDs or
    fills an unreadable quantity with 1. Blank identifiers become placeholders
    ONLY if the line is visibly a priced order line in the PDF table.
    """
    import pdfplumber

    p = Path(path)
    raw: list[OrderPart] = []
    saw_order_title = False
    saw_order_number = False
    saw_header = False
    with pdfplumber.open(p) as pdf:
        if not pdf.pages:
            raise ValueError('PDF 没有页面')
        for page in pdf.pages:
            pg = page.dedupe_chars(tolerance=1.5)
            body = pg.extract_text() or ''
            if re.search(r'(?mi)^\s*Sales\s*order\s*$', body):
                saw_order_title = True
            if re.search(r'\bOrder\s+No\.?\b', body, re.I) and re.search(r'\bOrderdatum\b', body, re.I):
                saw_order_number = True
            lines = _group_line_words(pg.extract_words(x_tolerance=2, y_tolerance=2))
            for li, line in enumerate(lines):
                keys = [re.sub(r'[^a-z]', '', w['text'].lower()) for w in line]
                if not all(x in keys for x in ('qty', 'erp', 'part')):
                    continue
                if not any(x.startswith('description') for x in keys):
                    continue
                if 'unit' not in keys or 'total' not in keys:
                    continue
                def field(name):
                    return next((w for w, k in zip(line, keys) if k == name), None)
                qty_h, erp_h, part_h, unit_h, total_h = map(
                    field, ('qty', 'erp', 'part', 'unit', 'total'))
                desc_h = next((w for w,k in zip(line,keys)
                               if k.startswith('description')), None)
                if not all((qty_h, erp_h, part_h, desc_h, unit_h, total_h)):
                    continue
                qty_x, erp_x, part_x, desc_x, unit_x, total_x = (
                    float(w['x0']) for w in (qty_h, erp_h, part_h, desc_h, unit_h, total_h))
                if not (qty_x < erp_x < part_x < desc_x < unit_x < total_x):
                    continue
                saw_header = True
                previous_detail_y = -1000.0
                for next_line in lines[li+1:]:
                    top = float(next_line[0]['top'])
                    if top > pg.height - 95:
                        break
                    line_text = _joined(next_line)
                    # HFT 266247: 'Customer Ref. B2615716' appears below
                    # the real order lines, but inside the ERP column. It is
                    # NOT a SKU with missing quantity.
                    if re.search(r'\b(?:Customer\s+Ref\.?|Customer\s+Reference|'
                                 r'Payment|Excluding\s+VAT|Date\s+of\s+delivery|'
                                 r'Grand\s+Total|Terms\s+of\s+payment|Uw\s+B\.?T\.?W\.?\s+nummer)\b',
                                 line_text, re.I):
                        break
                    def in_col(start, end):
                        return _joined([w for w in next_line
                                        if start <= float(w['x0']) < end])
                    qty_s = in_col(qty_x - 7, erp_x - 3)
                    erp_s = in_col(erp_x - 3, part_x - 3)
                    part_s = in_col(part_x - 3, desc_x - 3)
                    desc = in_col(desc_x - 3, unit_x - 3)
                    unit_s = in_col(unit_x - 3, total_x - 3)
                    total_s = _joined([w for w in next_line
                                       if float(w['x0']) >= total_x - 3])
                    # Prices are the strongest row-boundary check. A lone
                    # reference in the ERP column has no price and is not
                    # treated as an unrecognized order line.
                    priced = (number_or_none(unit_s) is not None or
                              number_or_none(total_s) is not None or
                              # German thousands separators (1.234,56) and
                              # currency-symbol variants still look like a
                              # priced order row, even if not decimal-parseable.
                              bool(re.search(r'\d', unit_s) or re.search(r'\d', total_s)))
                    qty = number_or_none(qty_s)
                    if not priced:
                        if raw and desc and not erp_s and not part_s and not qty_s and \
                                0 < top - previous_detail_y <= 18:
                            prev = raw[-1]
                            raw[-1] = OrderPart(prev.erp, prev.part, prev.quantity,
                                               (prev.description+' '+desc).strip(),prev.freight)
                            previous_detail_y = top
                        continue
                    # A priced line with a missing/unreadable ID is retained
                    # as a GUI placeholder rather than silently omitted.
                    freight = (normalize_number(erp_s) in {'999-830','999-831','X10019'} or
                               bool(re.search(r'\b(?:freight|shipping|transport\s+charge)\b',
                                              desc, re.I)))
                    if qty is not None and qty <= 0:
                        qty = None
                    raw.append(OrderPart(erp_s, part_s, qty, desc, freight))
                    previous_detail_y = top
    if not saw_order_title or not saw_order_number:
        raise ValueError('所选 PDF 缺少 Salesorder / Order No. / Orderdatum 标识，请选择订单 PDF')
    if not saw_header or not raw:
        raise ValueError('订单 PDF 中未找到可识别的 Qty./ERP No./Part No. 明细表，请人工核查')
    # Do not aggregate unresolved quantities or identifiers, because that
    # would hide missing rows or make a manual correction ambiguous.
    grouped: dict[tuple[str,str,bool],OrderPart] = {}
    for item in raw:
        if item.quantity is None or not (item.erp or item.part):
            grouped[('UNRESOLVED', str(len(grouped)), False)] = item
            continue
        key = (normalize_number(item.erp), normalize_number(item.part), item.freight)
        prev = grouped.get(key)
        if prev is None:
            grouped[key] = item
        else:
            grouped[key] = OrderPart(prev.erp, prev.part,
                                     prev.quantity+item.quantity,
                                     prev.description or item.description,prev.freight)
    parts = list(grouped.values())
    if not allow_incomplete:
        problematic = [x for x in parts if not x.freight and
                       (x.quantity is None or not (x.erp or x.part))]
        if problematic:
            raise ValueError('订单有数量或 ERP/Part No. 未识别的配件行；请使用统一库存搜索中的手动补录功能')
    return parts


def assess_stock(search: SearchResult, quantity: Decimal | None) -> str:
    if not search.hits:
        return '未找到库存记录'
    if search.warning and ('不一致' in search.warning or '不同 ERP/Part' in search.warning):
        return '编号冲突·待核'
    if search.total is None:
        return '库存未提供'
    if search.total <= 0:
        return '无库存'
    if quantity is None:
        return '有库存·数量待核'
    return '充足' if search.total >= quantity else '不足'
