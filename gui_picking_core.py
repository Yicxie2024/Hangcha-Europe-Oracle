"""Read-only warehouse picking planning and batch inventory queries (v1.5.30).

Ground all stock in the chosen saved workbook, all ordered quantities and
identifiers in the PDF body. Route estimates use a configurable aisle graph,
NOT measured travel distances; no stock movements are made.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from collections import defaultdict
from typing import Iterable
import heapq
import itertools
import re
import zipfile
import xml.etree.ElementTree as ET

from gui_inventory_core import (InventoryIndex, StockRecord, OrderPart, normalize_number,
    number_or_none, readable_number, _sheet_xml_name, _shared_strings, _cell_text, NS)

# Warehouse V1.5.30 - prefix-based location recognition and physical layout.
# A..P are plan columns. A-E / I-P use plan rows 01..12; H uses
# 01..13. F/G are special 8-position aisles 01..08 OUTWARD FROM THE
# STAIR LANDING (F01/G01
# nearest, F08/G08 farthest), NOT the same coordinates as ordinary plan rows.
# Shelf height codes 01..04 / 05..08 / 09..12 identify floors 1/2/3.
# The cross-aisle is BETWEEN ordinary plan rows 05 and 06, at the F/G stairs.
FLOOR_LEVELS = {1: range(1, 5), 2: range(5, 9), 3: range(9, 13)}
# The first FIVE characters (e.g. A0103) are sufficient to identify
# physical floor, plan position and vertical shelf tier. Optional cell/suffix
# information is retained but must never force a valid base into a candidate
# section. A01 / A010 are partial and use the user's explicit 1F fallback.
RACK_PATTERN = re.compile(r'^([A-P])((?:0[1-9]|1[0-3]))'
                          r'((?:0[1-9]|1[0-2]))(.*)$', re.I)
PARTIAL_PATTERN = re.compile(r'^([A-P])((?:0[1-9]|1[0-3]))(?:\d)?$', re.I)
LOCATION_SEPARATOR = re.compile(r'\s*[/;；,，、|\n]+\s*')
FREIGHT_CODES = {'999-830', '999-831', 'X10019'}
MAX_ITEMS = 2000

@dataclass(frozen=True)
class Rack:
    code: str
    col: int
    row: int
    level: int
    cell: int
    floor: int
    suffix: str = ''

    @property
    def shelf_height(self) -> int:
        """Relative vertical tier within the physical floor, 1..4."""
        return (self.level - 1) % 4 + 1


def split_rack_locations(value: str) -> tuple[str, ...]:
    """Keep all locations. Never silently select the first location."""
    original = str(value or '').strip()
    if not original:
        return ()
    return tuple(dict.fromkeys(x.strip().upper() for x in
               LOCATION_SEPARATOR.split(original) if x.strip()))


def parse_rack(value: str) -> Rack | None:
    """Recognize A-P + row + level prefix; F/G 01..08, H 01..13.

    Cell and suffix are optional and never affect physical floor/routing.
    F/G rack row numbers are measured away from the 05/06 landing, not
    regular A-F/I-P physical matrix rows. F/G 01..08 are real rack positions.
    """
    original = str(value or '').strip().upper()
    match = RACK_PATTERN.fullmatch(original)
    if not match:
        return None
    letter, row, level, tail = match.groups()
    row, level = int(row), int(level)
    if (letter in ('F', 'G') and row > 8) or (letter != 'H' and letter not in ('F', 'G') and row > 12):
        return None
    # A 2-digit cell may follow the prefix; all later characters are retained
    # for printing. Unusual tails are NOT used to invent a route segment.
    cell_match = re.match(r'^(\d{2})(.*)$', tail)
    cell = int(cell_match.group(1)) if cell_match else 0
    suffix = cell_match.group(2) if cell_match else tail
    return Rack(original, ord(letter)-ord('A'), row, level, cell,
                (level-1)//4+1, suffix)



def _valid_row(letter: str, row: int) -> bool:
    if letter in ('F','G'):
        return 1 <= row <= 8
    if letter == 'H':
        return 1 <= row <= 13
    return 1 <= row <= 12


def _row_error(letter: str) -> str:
    if letter in ('F','G'):
        return 'F/G列编号仅支持01–08'
    if letter == 'H':
        return 'H列编号仅支持01–13'
    return '普通列编号仅支持01–12'


def floor_from_uncertain_rack(value: str) -> int | None:
    """Recover a known physical floor even if the plan row is invalid.

    A01/A010 are explicitly placed at the END of floor 1, not routed.
    Other malformed positions with a valid tier (e.g. P1405) are displayed
    at the END of the corresponding floor, NOT in an independent table.
    Never infer physical floor from plan row.
    """
    word = str(value or '').strip().upper()
    full = RACK_PATTERN.fullmatch(word)
    if full:
        return (int(full.group(3)) - 1)//4+1
    partial = PARTIAL_PATTERN.fullmatch(word)
    if partial:
        if not _valid_row(partial.group(1), int(partial.group(2))):
            return None
        return 1
    m = re.match(r'^[A-P]\d{2}(0[1-9]|1[0-2])',word)
    return (int(m.group(1))-1)//4+1 if m else None


def rack_note(value: str) -> str:
    """Human-readable location caveat without discarding valid prefixes."""
    word = str(value or '').strip().upper()
    partial = PARTIAL_PATTERN.fullmatch(word)
    if partial:
        letter, row = partial.group(1), int(partial.group(2))
        if not _valid_row(letter, row):
            return _row_error(letter) + '；位置无法确定'
        return '层号未提供；按约定暂列1F末尾，现场确认'
    m = RACK_PATTERN.fullmatch(word)
    if m:
        if not _valid_row(m.group(1), int(m.group(2))):
            return _row_error(m.group(1)) + '；可识别楼层，位置待核'
        tail=m.group(4)
        if tail and not re.fullmatch(r'(?:\d{2})?(?:[-\s]?[IVX]+|[-\s]?[ⅠⅡⅢⅣ一二三四])?',tail,re.I):
            return '可定位前5位；尾码请现场核对'
        return ''
    if floor_from_uncertain_rack(word) is not None:
        return '货架行号/编码不标准；楼层已识别，现场定位'
    return '位置无法确认'


@dataclass(frozen=True)
class PickRequest:
    code: str
    quantity: Decimal | None
    erp: str = ''
    part: str = ''
    description: str = ''


def parse_paste(text: str, default_qty: str | int = '1') -> list[PickRequest]:
    """Comma/CN-comma/newline-separated IDs; CODE*QTY optional. No guessing CSV pairs."""
    dq = number_or_none(default_qty)
    if dq is None or dq <= 0:
        raise ValueError('默认需求量必须大于零')
    result = []
    for chunk in re.split(r'[,，;；\n\r\t]+', text):
        word = chunk.strip()
        if not word:
            continue
        code, sep, q = word.rpartition('*')
        if sep:
            quantity = number_or_none(q)
            if quantity is None or quantity <= 0:
                raise ValueError(f'无效数量：{word}（格式应为 ERP*2）')
            word = code.strip()
        else:
            quantity = dq
        if not word or any(c in word for c in ('\\', ':', ' ')):
            raise ValueError(f'编号不符合要求：{word!r}')
        result.append(PickRequest(word, quantity))
        if len(result)>MAX_ITEMS:
            raise ValueError(f'单次不能超过 {MAX_ITEMS} 项')
    return result


def import_requests_xlsx(path: str | Path) -> list[PickRequest]:
    """Import first sheet, read-only ZIP/XML; columns: identifier + demand.

    Supports [ERP No., Part No., Qty.] or [Code, Qty.] headers and headerless
    [code, qty]. Blank/invalid quantities are reported, never assumed as 1.
    """
    p=Path(path)
    if p.suffix.lower() not in {'.xlsx','.xlsm'} or not p.is_file():
        raise ValueError('请选择 .xlsx 或 .xlsm 表格')
    result=[]; problems=[]
    with zipfile.ZipFile(p) as z:
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        sheets=wb.find(NS+'sheets')
        if sheets is None or len(sheets)==0:
            raise ValueError('Excel 中找不到工作表')
        sheet_name=sheets[0].get('name','')
        sheet=_sheet_xml_name(z,sheet_name)
        strings=_shared_strings(z)
        layout=None
        with z.open(sheet) as stream:
            for _,element in ET.iterparse(stream,events=('end',)):
                if element.tag!=NS+'row': continue
                rownum=int(element.get('r','0'))
                values={}
                for c in element.findall(NS+'c'):
                    letter=re.sub(r'\d','',c.get('r',''))
                    if letter in ('A','B','C','D','E','F','G','H','I','J'):
                        values[letter]=_cell_text(c,strings).strip()
                element.clear()
                if not any(values.values()): continue
                if layout is None:
                    normalized={k:re.sub(r'[\s._\-/]+','',v).casefold() for k,v in values.items()}
                    def names(test):return [k for k,v in normalized.items() if test(v)]
                    qty_cols=names(lambda v: v in {'qty','quantity','bedarf','need','需求量','数量','订购数量','所需数量'} or v.startswith('qty'))
                    erp_cols=names(lambda v: 'erp' in v or v=='商品明细码')
                    part_cols=names(lambda v: 'partno' in v or v=='配件号')
                    code_cols=names(lambda v: v in {'code','number','编号','partnumber','erp/part','erp或partno'})
                    if qty_cols and (erp_cols or part_cols or code_cols):
                        layout=(erp_cols[0] if erp_cols else '',part_cols[0] if part_cols else '',
                                code_cols[0] if code_cols else '',qty_cols[0])
                        continue
                    # headerless, 2 columns A=code B=qty
                    layout=('','','A','B')
                erp_col,part_col,code_col,qty_col=layout
                erp=values.get(erp_col,'') if erp_col else ''
                part=values.get(part_col,'') if part_col else ''
                code=values.get(code_col,'') if code_col else ''
                identity=erp or part or code
                if not identity: continue
                qty=number_or_none(values.get(qty_col,''))
                if qty is None or qty<=0:
                    problems.append(f'第 {rownum} 行数量无效');continue
                if not normalize_number(identity) or any(c in identity for c in ('\\', ':', ' ')):
                    problems.append(f'第 {rownum} 行编号无效');continue
                result.append(PickRequest(identity,qty,erp,part))
                if len(result)>MAX_ITEMS:raise ValueError(f'单次最多导入 {MAX_ITEMS} 项')
    if problems:
        raise ValueError('导入未执行：需要先修正以下单元格：\n'+'\n'.join(problems[:20])+
                         (f'\n另外 {len(problems)-20} 行异常' if len(problems)>20 else ''))
    if not result:raise ValueError('Excel 中没有可用的 ERP/Part 编号与需求数量')
    return result


def requests_from_order(items: Iterable[OrderPart], allow_incomplete: bool = False) -> list[PickRequest]:
    requests=[]
    for item in items:
        if item.freight or normalize_number(item.erp) in FREIGHT_CODES:continue
        if item.quantity is None or item.quantity<=0:
            if not allow_incomplete:
                raise ValueError(f'订单中的配件 {item.erp or item.part} 数量不可用，请人工核对')
        if not (item.erp or item.part) and not allow_incomplete:
            raise ValueError('订单有配件行缺少 ERP 和 Part No.，请人工核对')
        requests.append(PickRequest(item.erp or item.part,item.quantity,item.erp,item.part,item.description))
    return requests


@dataclass(frozen=True)
class Allocation:
    rack: str
    quantity: Decimal
    stock: Decimal
    record: StockRecord
    position: Rack
    travel_index: int = 0


@dataclass
class PickResult:
    request: PickRequest
    records: tuple[StockRecord, ...]
    match_method: str
    status: str
    stock_total: Decimal | None
    warning: str = ''
    allocations: list[Allocation] = field(default_factory=list)
    shortage: Decimal = Decimal('0')
    excluded: bool = False
    @property
    def erp(self):return self.request.erp or (self.records[0].erp if self.records else '')
    @property
    def part(self):return self.request.part or (self.records[0].part if self.records else '')
    @property
    def description(self):return self.request.description or (self.records[0].description if self.records else '')


def lookup_requests(index: InventoryIndex, requests: Iterable[PickRequest]) -> list[PickResult]:
    """Create lookup rows; aggregate duplicate IDENTICAL user requests.

    Never sum ambiguous cross-number hits or print a route for unclear racks.
    """
    grouped={}
    for req in requests:
        if req.quantity is None or req.quantity <= 0 or not req.code.strip():
            raise ValueError('存在尚未补录完整编号或数量的订单条目，不能查询或生成拣货单')
        key=(normalize_number(req.erp),normalize_number(req.part)) if req.erp or req.part else (normalize_number(req.code),'')
        if key in grouped:
            old=grouped[key]
            grouped[key]=PickRequest(old.code,old.quantity+req.quantity,old.erp,old.part,old.description or req.description)
        else:grouped[key]=req
    results=[]
    for req in grouped.values():
        res=index.find_item(req.erp,req.part) if req.erp or req.part else index.search(req.code)
        if not res.hits:
            status='未找到';shortage=req.quantity
        elif res.warning and ('不一致' in res.warning or '不同 ERP/Part' in res.warning):
            status='编号冲突·待核';shortage=req.quantity
        elif any(r.stock is None for r in res.hits):
            status='库存未知·待核';shortage=req.quantity
        else:
            total=res.total
            status='充足' if total is not None and total>=req.quantity else '不足'
            shortage=max(Decimal(0),req.quantity-(total or Decimal(0)))
        results.append(PickResult(req,res.hits,res.method,status,res.total,res.warning,[],shortage))
    # Same stock row hit through ERP and Part queries: do not count twice;
    # consolidate demand under one material identity only for unambiguous rows.
    seen={}
    for item in results:
        if item.status=='编号冲突·待核' or not item.records:continue
        identities=tuple(sorted(set((normalize_number(r.erp),normalize_number(r.part)) for r in item.records)))
        if len(identities)!=1:continue
        ident=identities[0]
        old=seen.get(ident)
        if old is None:seen[ident]=item;continue
        # Merge separate ERP and Part search inputs when both conclusively
        # map to same ERP + Part identity; preserve one list row.
        old.request=PickRequest(old.request.code,old.request.quantity+item.request.quantity,
                                old.request.erp or item.request.erp,old.request.part or item.request.part,
                                old.request.description or item.request.description)
        old.stock_total=old.stock_total
        old.shortage=max(Decimal(0),old.request.quantity-(old.stock_total or Decimal(0)))
        old.status=('库存未知·待核' if old.stock_total is None else
                    '充足' if old.shortage==0 else '不足')
        item.excluded=True
        item.warning='已与相同 ERP/Part 的另一条需求合并'
    return [item for item in results if not item.excluded]

# Aisle graph is a ROUTE HEURISTIC, not a measured plan. We represent the
# between-row 05/06 cross-aisle as a virtual node at 05.5 on each column.
# A-E / I-P use rows 01..12; H uses rows 01..13, and all these
# ordinary columns enter via the mid aisle at rows
# 05/06. F/G 01..08 are separate branches whose #01 racks are nearest the
# F/G stair landing. Never add regular F/G row links or route through the
# F/G 01..06 physical staircase void in the ordinary grid.
MID = 55   # virtual 05.5, not a rack row number
FRONT = 0 # virtual ordinary row-01-end approach (legacy approximation)
STAIRS = (16, MID)  # virtual landing between F (col=5) and G (col=6)


def _distance_graph():
    graph:dict[tuple[int,int],list[tuple[tuple[int,int],int]]]=defaultdict(list)
    def edge(a,b,weight):
        graph[a].append((b,weight));graph[b].append((a,weight))
    for c in range(16):
        # Cross aisle extends A..P; the stairs meet it between F/G.
        if c < 15:
            edge((c,MID),(c+1,MID),2)
        if c in (5,6):
            # Physical racks F/G have a separate 01..08 numbering from the
            # stairs outwards. No accidental connection to ordinary row 06.
            edge((c,MID),(c,1),1)
            for r in range(1,8):
                edge((c,r),(c,r+1),2)
            continue
        # Ordinary columns, lower positions 01..05 and upper 06..12.
        edge((c,MID),(c,5),1)
        for r in range(1,5):
            edge((c,r),(c,r+1),2)
        edge((c,MID),(c,6),1)
        for r in range(6, 13 if c == 7 else 12):
            edge((c,r),(c,r+1),2)
        # Keep original optional entry-side corridor as an approximation,
        # split around the physical F/G stair void.
        edge((c,FRONT),(c,1),1)
        if c < 15 and c not in (4,5,6):
            edge((c,FRONT),(c+1,FRONT),2)
    # Virtual stairs node connects to both sides of the F/G cross-aisle.
    edge(STAIRS,(5,MID),1)
    edge(STAIRS,(6,MID),1)
    return graph

_GRAPH=_distance_graph()
_CACHE={}


def route_distance(a:tuple[int,int],b:tuple[int,int]) -> int:
    if a==b:return 0
    k=(a,b)
    if k in _CACHE:return _CACHE[k]
    pq=[(0,a)]; seen={}
    while pq:
        d,n=heapq.heappop(pq)
        if n in seen:continue
        seen[n]=d
        if n==b:break
        for target,dist in _GRAPH.get(n,()):
            if target not in seen:heapq.heappush(pq,(d+dist,target))
    v=seen.get(b,100000)
    _CACHE[k]=v;_CACHE[(b,a)]=v
    return v


def route_path(a:tuple[int,int], b:tuple[int,int]) -> list[tuple[int,int]]:
    """Explicit graph path for checking that no blocked staircase is crossed."""
    if a==b:return [a]
    pq=[(0,a)]; dist={a:0}; parent={}
    while pq:
        d,n=heapq.heappop(pq)
        if d != dist.get(n):continue
        if n==b:break
        for target,w in _GRAPH.get(n,()):
            nd=d+w
            if nd < dist.get(target,10**9):
                dist[target]=nd;parent[target]=n
                heapq.heappush(pq,(nd,target))
    if b not in dist:return []
    path=[b]
    while path[-1]!=a:path.append(parent[path[-1]])
    return path[::-1]


def order_floor_positions(positions: Iterable[tuple[int,int]]) -> list[tuple[int,int]]:
    """Route order for all safely located entries, whether stock is assigned.

    Uses the existing graph and bounded 2-opt heuristic. Valid multi-location
    candidates are included as walking positions, not silently omitted.
    """
    remaining=set(positions)
    ordered=[];cur=STAIRS
    while remaining:
        chosen=min(remaining,key=lambda pos:(route_distance(cur,pos),pos[0],pos[1]))
        ordered.append(chosen);remaining.remove(chosen);cur=chosen
    def cost(sequence):
        cur=STAIRS;val=0
        for pos in sequence:val+=route_distance(cur,pos);cur=pos
        return val+route_distance(cur,STAIRS)
    if 3<=len(ordered)<=100:
        for _ in range(12):
            improved=False;best=cost(ordered)
            for i in range(0,len(ordered)-1):
                for j in range(i+1,min(len(ordered),i+16)):
                    trial=ordered[:i]+ordered[i:j+1][::-1]+ordered[j+1:]
                    new=cost(trial)
                    if new<best:
                        ordered=trial;best=new;improved=True;break
                if improved:break
            if not improved:break
    return ordered


def _order_floor(steps:list[Allocation]) -> list[Allocation]:
    """Route by physical rack position; pick upper shelves before lower at a stop."""
    by_position:dict[tuple[int,int],list[Allocation]]=defaultdict(list)
    for item in steps:by_position[(item.position.col,item.position.row)].append(item)
    for rack_steps in by_position.values():
        rack_steps.sort(key=lambda a:(-a.position.shelf_height,a.position.cell,a.rack))
    return [item for pos in order_floor_positions(by_position) for item in by_position[pos]]


def allocate_and_route(results:list[PickResult]) -> tuple[list[Allocation],list[str]]:
    """Allocate only stock with a single unambiguous location.

    All other possible positions, including routable multi-location records,
    are displayed on the same physical-floor tables by make_pick_pdf().
    Never allocate aggregate stock to every location of a multi-location row.
    """
    used:dict[int,Decimal]=defaultdict(Decimal)
    active_floors=set();notes=[]
    for result in results:
        result.allocations=[]
        if result.excluded:continue
        if result.status in ('未找到','编号冲突·待核','库存未知·待核'):continue
        need=result.request.quantity
        candidates=[]
        for record in result.records:
            locations=split_rack_locations(record.rack)
            if len(locations)!=1:continue
            rack=parse_rack(locations[0])
            if rack is None or record.stock is None or record.stock<=0:continue
            remaining=max(Decimal(0),record.stock-used[record.row])
            if remaining>0:candidates.append((record,rack,remaining))
        if not candidates:
            result.status='库位数量待核'
            result.warning=(result.warning+'；' if result.warning else '')+'需要现场确定具体位置或多库位取货数量'
            result.shortage=need
            continue
        # Prefer staying on already required floors; then prefer higher floors
        # because worker starts upstairs and walks DOWN. No extra floor is
        # added just to satisfy an arbitrary lexical rack sort.
        candidates.sort(key=lambda t:(t[1].floor not in active_floors,-t[1].floor,
                                      t[1].col,t[1].row,-t[1].shelf_height,t[1].cell))
        for record,rack,available in candidates:
            if need<=0:break
            take=min(need,available)
            result.allocations.append(Allocation(record.rack,take,record.stock,record,rack))
            used[record.row]+=take;need-=take;active_floors.add(rack.floor)
        result.shortage=need
        if need>0:
            result.status='不足/库位待核'
            result.warning=(result.warning+'；' if result.warning else '')+'其他库存可能在多个库位，待现场确定取货数量'
    steps=[s for r in results if not r.excluded for s in r.allocations]
    by_floor=defaultdict(list)
    for step in steps:by_floor[step.position.floor].append(step)
    sequence=[]
    for floor in (3,2,1):
        if not by_floor[floor]:continue
        subset=_order_floor(by_floor[floor])
        notes.append(f'{floor}F：由 F/G 楼梯口的 05–06 横向通道出发并返回；F/G 的 01–08 自楼梯向内编号；H 为 01–13；路线仅为通道模型建议')
        sequence.extend(subset)
    sequence=[Allocation(a.rack,a.quantity,a.stock,a.record,a.position,i)
              for i,a in enumerate(sequence,1)]
    for result in results:
        owned={(a.record.row,a.rack,a.quantity) for a in result.allocations}
        result.allocations=[b for b in sequence if (b.record.row,b.rack,b.quantity) in owned]
    return sequence,notes
