"""Landscape A4 pick-list generation and opt-in Windows print dispatch.

Never sends a print job without the caller's explicit approval. PDFs are kept
in the user's Documents/Hangcha Picking Lists folder for preview and audit.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape
import os
import shutil
import subprocess

from gui_inventory_core import readable_number
from gui_picking_core import (PickResult, Allocation, split_rack_locations,
    floor_from_uncertain_rack, parse_rack, rack_note,
    order_floor_positions)


def _register_font():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfbase.ttfonts import TTFont
    # Avoid redistributing font binaries. Prefer Windows-installed TrueType
    # fonts; CID fallback works with most viewers but relies on substitution.
    candidates=[
        Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'/'msyh.ttc',
        Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'/'simsun.ttc',
        Path('/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc'),
        Path('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'),
    ]
    for item in candidates:
        if item.is_file():
            try:
                pdfmetrics.registerFont(TTFont('HangchaCJK',str(item),subfontIndex=0))
                return 'HangchaCJK'
            except Exception:
                pass
    if 'STSong-Light' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    return 'STSong-Light'


def make_pick_pdf(destination: str | Path, order_no: str, customer: str,
                  results: list[PickResult], route: list[Allocation], notes: list[str],
                  inventory_source: str, picked_at: datetime | None = None,
                  rack_stock: dict[tuple[int,str], object] | None = None) -> Path:
    """One picking row per source Excel stock record (per floor if necessary).

    Preserve each original Rack No. string and MA stock figure. Repeated Excel
    records of the same SKU carry the SAME picking serial; demand and the rack
    review remark are vertically merged where such records appear together.
    For a stock record spanning physical floors we must split its locations by
    floor to preserve 3F->2F->1F. Its aggregate MA is shown ONLY on the first
    displayed fragment; later fragments point back to the same source row.
    """
    from collections import defaultdict
    from decimal import Decimal
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                    TableStyle, CondPageBreak)

    rack_stock = rack_stock or {}
    dest = Path(destination)
    dest.parent.mkdir(parents=True,exist_ok=True)
    font = _register_font()
    normal = ParagraphStyle('normal', fontName=font,fontSize=8,leading=11,wordWrap='CJK')
    small = ParagraphStyle('small', parent=normal,fontSize=7,leading=9)
    tiny = ParagraphStyle('tiny', parent=normal,fontSize=6.8,leading=8.8)
    heading = ParagraphStyle('heading',parent=normal,fontSize=14,leading=19,
                             textColor=colors.HexColor('#133451'))
    floorhead = ParagraphStyle('floorhead',parent=normal,fontSize=11,leading=15,
                               textColor=colors.HexColor('#174366'),spaceAfter=5)

    def p(value, style=normal):
        return Paragraph(escape(str(value if value is not None and value != '' else '—')),style)

    page = landscape(A4)
    # A 14pt gutter outside the LEFT table border is reserved for paper checkboxes.
    doc = SimpleDocTemplate(str(dest),pagesize=page,leftMargin=42,rightMargin=26,
                            topMargin=28,bottomMargin=34,
                            title=f'HCE Oracle Picking List {order_no or "Internal"}',
                            author='HCE Oracle')
    stamp = (picked_at or datetime.now()).strftime('%d.%m.%Y %H:%M')
    order_label = order_no or '内部备货（无订单）'
    # Keep the essential identity; omit the three crossed-out guidance lines.
    story = [p('HANGCHA EUROPE · HCE Oracle | PICKING LIST / 仓库拣货单',heading),
             p(f'订单 / Order: {order_label}    客户 / Customer: {customer or "未指定"}    生成: {stamp}'),
             Spacer(1,7)]

    # Build ONE visual row per source workbook row. For a source row with
    # locations spanning distinct physical floors, one visual fragment per
    # floor is unavoidable; preserve source-row identity to prevent MA being
    # counted/showed as separate inventory on different floors.
    raw=[]
    for res_idx,res in enumerate(results):
        if res.excluded:
            continue
        for rec in res.records:
            source_locations = split_rack_locations(rec.rack) or ('未提供',)
            by_floor={}
            for loc in source_locations:
                position = parse_rack(loc)
                floor = position.floor if position else floor_from_uncertain_rack(loc)
                by_floor.setdefault(floor,[]).append((loc,position))
            # A single-floor record must retain its original Rack No. verbatim.
            for floor,locs in by_floor.items():
                original = rec.rack.strip() if len(by_floor)==1 else ' / '.join(x[0] for x in locs)
                raw.append({'res':res,'res_idx':res_idx,'record':rec,
                            'floor':floor,'rack':original,
                            'source_id':(res_idx,rec.row),
                            'positions':[pos for _,pos in locs if pos is not None],
                            'multiple':len(source_locations)>1,
                            'cross_floor':len(by_floor)>1,
                            'notes':[rack_note(loc) for loc,_ in locs if rack_note(loc)]})
        if not res.records:
            raw.append({'res':res,'res_idx':res_idx,'record':None,
                        'floor':None,'rack':'未提供','source_id':(res_idx,None),
                        'positions':[],'multiple':False,'cross_floor':False,
                        'notes':['库存记录未找到']})

    rank_by_floor={}
    for floor in (3,2,1):
        coords={(pos.col,pos.row) for x in raw if x['floor']==floor
                for pos in x['positions']}
        rank_by_floor[floor]={coord:i for i,coord in enumerate(order_floor_positions(coords))}

    def route_key(item):
        floor=item['floor']
        positions=item['positions']
        if positions and floor in rank_by_floor:
            rank=min(rank_by_floor[floor][(q.col,q.row)] for q in positions)
            # Within one source record/stop higher shelf tiers first.
            return (0,rank,-max(q.shelf_height for q in positions),item['record'].row)
        partial=bool(__import__('re').fullmatch(r'[A-P](?:0[1-9]|1[0-2])\d?',item['rack'].strip().upper()))
        return (2 if partial else 1,9999,0,item['record'].row if item['record'] else 99999)

    # At each floor, keep all source rows of one SKU adjacent. Their first
    # reachable stop determines where the group enters the planned route.
    ordered=[]
    for floor in (3,2,1,None):
        candidates=[x for x in raw if x['floor']==floor]
        grouped=defaultdict(list)
        for x in candidates:grouped[x['res_idx']].append(x)
        groups=sorted(grouped.values(), key=lambda g:
                      (min(route_key(x) for x in g),min(x['res_idx'] for x in g)))
        for g in groups:ordered.extend(sorted(g,key=route_key))

    # A type receives one serial regardless of how many source lines/floors.
    serial_by_result={}
    for item in ordered:
        serial_by_result.setdefault(item['res_idx'],len(serial_by_result)+1)

    # Demand is shown once per SKU across the ENTIRE PDF, with a vertically
    # merged cell when its source rows are adjacent within the same floor.
    demand_displayed=set()
    stock_displayed=set()
    # Every item type has one No. even if it has multiple stock records or
    # appears on several physical floors. Do not duplicate the demand.
    from collections import Counter
    all_counts=Counter(x['res_idx'] for x in ordered)
    floors_by_result={k:{x['floor'] for x in ordered if x['res_idx']==k and x['floor'] is not None}
                      for k in all_counts}
    repeated_parts={k for k, count in all_counts.items() if count>1}
    pale_red=colors.HexColor('#FBE7E8')
    pale_orange=colors.HexColor('#FFF0D9')
    labels=['#','Rack No.','ERP No.','Part No.','Description / 描述','应拣','库存','实拣','备注']
    # The removed '确认' column is deliberately NOT replaced by checkbox column.
    widths=[24,120,100,98,147,51,46,48,117]
    floor_names={3:'第三层 / 3F  (货架层号 09–12)',
                 2:'第二层 / 2F  (货架层号 05–08)',
                 1:'第一层 / 1F  (货架层号 01–04)',
                 None:'无法确认区域 / FLOOR UNKNOWN'}

    def draw_outside_checkboxes(table,event,*args):
        """Draw a print-only checkbox to the LEFT of the outer table border.

        The callback is retained by ReportLab when a long Table is split over
        multiple pages, and row markers are kept on the cell Paragraph objects.
        One checkbox per displayed SKU group / physical floor, not per stock row.
        """
        if event!='endTable':return
        canvas=table.canv
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor('#3B5568'))
        canvas.setLineWidth(.8)
        for i,row in enumerate(table._cellvalues):
            if i==0:continue  # header, including repeated page headers
            no_cell=row[0]
            # During layout ReportLab wraps a Paragraph in an
            # _ExpandedCellTuple; look inside it when reading our marker.
            if isinstance(no_cell,(tuple,list)) and no_cell:
                no_cell=no_cell[0]
            if not getattr(no_cell,'picking_checkbox',False):continue
            midpoint=(table._rowpositions[i]+table._rowpositions[i+1])/2
            canvas.rect(-16.3,midpoint-4.4,8.8,8.8,stroke=1,fill=0)
        canvas.restoreState()

    def make_table(rows,spans,highlighted_rows,orange_rows,no_split):
        table=Table(rows,colWidths=widths,repeatRows=1,hAlign='LEFT',
                    renderCB=draw_outside_checkboxes)
        style=[('BACKGROUND',(0,0),(-1,0),colors.HexColor('#DCEAF2')),
               ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F4F8FB')]),
               ('GRID',(0,0),(-1,-1),.35,colors.HexColor('#C6D3DE')),
               ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
               ('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),3),
               ('TOPPADDING',(0,0),(-1,-1),2.5),('BOTTOMPADDING',(0,0),(-1,-1),2.5)]
        # Orange takes priority on a particular source row when it has proven
        # ZERO stock or when there is no inventory record. Unknown is not zero.
        for row_no in highlighted_rows:
            style.append(('BACKGROUND',(0,row_no),(-1,row_no),pale_red))
        for row_no in orange_rows:
            style.append(('BACKGROUND',(0,row_no),(-1,row_no),pale_orange))
        for first,last in no_split:
            if last>first:
                style.append(('NOSPLIT',(0,first),(-1,last)))
        for col,first,last in spans:
            style.append(('SPAN',(col,first),(col,last)))
            style.append(('VALIGN',(col,first),(col,last),'MIDDLE'))
        table.setStyle(TableStyle(style))
        return table

    floors=[f for f in (3,2,1,None) if any(x['floor']==f for x in ordered)]
    if not floors:floors=[1]
    for floor in floors:
        items=[x for x in ordered if x['floor']==floor]
        rows=[[p(x,small) for x in labels]]
        spans=[]
        highlighted_rows=[]
        orange_rows=[]
        no_split=[]
        index=0
        while index<len(items):
            group_end=index+1
            while group_end<len(items) and items[group_end]['res_idx']==items[index]['res_idx']:
                group_end+=1
            # All lines of one part in this floor share one demand and note.
            group=items[index:group_end]
            res=group[0]['res'];res_idx=group[0]['res_idx']
            first_for_part=res_idx not in demand_displayed
            if first_for_part:demand_displayed.add(res_idx)
            demand=readable_number(res.request.quantity) if first_for_part else '—'
            unknown_location=any(x['multiple'] or x['notes'] or x['floor'] is None for x in group)
            remark='库位待核' if unknown_location else ''
            multi_rows=[x for x in group if x['multiple'] and x['record'] is not None]
            if multi_rows and all(
                all((x['record'].row,loc) in rack_stock for loc in split_rack_locations(x['record'].rack))
                for x in multi_rows) and not any(x['notes'] for x in group):
                remark='库位已核'
            if any(x['floor'] is None for x in group):remark='库位无法确认'
            if any(x['cross_floor'] for x in group):remark=(remark+'；' if remark else '')+'跨楼层库位'
            distinct_floors=floors_by_result[res_idx]
            if len(distinct_floors)>1:
                floor_text='、'.join(f'{f}F' for f in sorted(distinct_floors,reverse=True))
                remark=(remark+'；' if remark else '')+'同一配件跨物理楼层重复（'+floor_text+'）'
            if not remark and res.status in ('未找到','编号冲突·待核','库存未知·待核'):
                remark=res.status
            first_row=len(rows)
            for item in group:
                rec=item['record']; sid=item['source_id']
                if rec is None:
                    stock='—'
                elif sid in stock_displayed:
                    stock='同一Excel行'
                else:
                    stock=readable_number(rec.stock) if rec.stock is not None else '待核'
                    stock_displayed.add(sid)
                # The ERP value on a blank-ERP source record comes from the
                # order request; its own stock/rack still come from Excel.
                erp=(res.erp or (rec.erp if rec else ''))
                part=(res.part or (rec.part if rec else ''))
                desc=res.description or (rec.description if rec else '')
                serial_cell=p(serial_by_result[res_idx] if len(rows)==first_row else '',small)
                serial_cell.picking_checkbox=(len(rows)==first_row)
                row_no=len(rows)
                rows.append([serial_cell,p(item['rack'],tiny),
                             p(erp,tiny),p(part,tiny),p(desc,tiny),
                             p(demand if row_no==first_row else '',tiny),p(stock,tiny),
                             p('_______',small),
                             p(remark if row_no==first_row else '',tiny)])
                if res_idx in repeated_parts:
                    highlighted_rows.append(row_no)
                if rec is None or (rec.stock is not None and rec.stock <= 0):
                    orange_rows.append(row_no)
            last_row=len(rows)-1
            if last_row>first_row:
                # Merge No., demand, and remark across identical SKU stock rows
                # WITHIN the same physical floor. Other floor tables stay distinct.
                spans.extend([(0,first_row,last_row),
                              (5,first_row,last_row),(8,first_row,last_row)])
                no_split.append((first_row,last_row))
            index=group_end
        # Keep manual per-location checks useful without replacing source-row
        # stock totals. Source Rack No. remains unchanged in the main table.
        confirmed_notes=[]
        seen_sources=set()
        for item in items:
            rec=item['record']
            if rec is None or rec.row in seen_sources:
                continue
            seen_sources.add(rec.row)
            confirmed=[f'{loc}={readable_number(rack_stock[(rec.row,loc)])}'
                       for loc in split_rack_locations(rec.rack)
                       if (rec.row,loc) in rack_stock]
            if confirmed:
                confirmed_notes.append(f'Excel 行 {rec.row} / {rec.part or rec.erp}: '+', '.join(confirmed))
        floor_story=[]
        probe=make_table(rows[:2],[],[],[],[])
        _, rowheight=probe.wrap(doc.width,doc.height)
        floor_story.append(CondPageBreak(min(doc.height,rowheight+40)))
        floor_story.extend([p(floor_names[floor],floorhead),
                            p(('起终点：F/G 楼梯—05/06 横向通道；F/G库位01–08自楼梯向内编号，H排01–13。'
                               if floor is not None else
                               '未识别位置与普通库位使用相同表格格式。'),small),Spacer(1,4),
                            make_table(rows,spans,highlighted_rows,orange_rows,no_split),Spacer(1,4)])
        if confirmed_notes:
            floor_story.append(p('人工核对分库位数量（不重复计入上表库存）：'+'；'.join(confirmed_notes),tiny))
        floor_story.append(Spacer(1,6))
        story.extend(floor_story)
    if notes:story.append(p('路径提示：'+'；'.join(notes),tiny))
    story.append(p('拣货人: ____________________   复核人: ____________________   日期: _______________',small))
    story.append(p(f'库存来源: {Path(inventory_source).name}',tiny))

    def footer(canvas,doc):
        canvas.saveState();canvas.setFont(font,7)
        canvas.setFillColor(colors.HexColor('#61788C'))
        canvas.drawString(26,16,f'HCE Oracle | Order: {order_label}')
        canvas.drawRightString(page[0]-26,16,f'Page {doc.page} | A4 Landscape | {stamp}')
        canvas.restoreState()
    doc.build(story,onFirstPage=footer,onLaterPages=footer)
    return dest

def find_sumatra() -> Path|None:
    names=['SumatraPDF.exe','SumatraPDF']
    for cmd in names:
        found=shutil.which(cmd)
        if found:return Path(found)
    roots=[os.environ.get('ProgramFiles',''),os.environ.get('ProgramFiles(x86)',''),
           os.environ.get('LOCALAPPDATA','')]
    for root in roots:
        if root:
            for suffix in ('SumatraPDF/SumatraPDF.exe','Programs/SumatraPDF/SumatraPDF.exe'):
                p=Path(root)/suffix
                if p.is_file():return p
    return None


def print_pdf(path: Path, copies:int, sumatra:Path|None=None) -> subprocess.Popen:
    """Print exactly N copies to default printer; errors surfaced by caller.

    SumatraPDF is an independently installed Windows viewer. The caller must
    explicitly approve print jobs before invoking this method. No silent
    shell-verb fallback that could ignore the copy count.
    """
    if os.name!='nt':raise RuntimeError('直接打印仅支持 Windows')
    if copies < 1 or copies > 20:raise ValueError('打印份数必须在 1 到 20 之间')
    exe=sumatra or find_sumatra()
    if exe is None:raise RuntimeError('未发现 SumatraPDF，无法可靠控制打印份数；请先预览 PDF 并手动打印')
    p=Path(path)
    if not p.is_file():raise FileNotFoundError(str(p))
    return subprocess.Popen([str(exe),'-print-to-default','-print-settings',
                             f'{copies}x,paper=A4,fit',str(p)],
                             stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
