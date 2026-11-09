"""Self-contained XLSX writer/reader for Hangcha monthly internal reports.

Uses only OOXML ZIP and ElementTree: no external Excel engine, stale formula
caches or live Excel automation. Original user template is read-only reference.
"""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from xml.etree import ElementTree as ET
import hashlib
import json
import os
import re
import tempfile
import zipfile

from gui_inventory_core import _cell_text, _sheet_xml_name, _shared_strings, NS, normalize_number

URL_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CONTENT_NS='http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('',URL_NS)
ET.register_namespace('r',REL_NS)
SHEETS=['10月','AI待补全','筛选日志','历史参考','说明']  # workbook sheet1 name is dynamic
HEADERS=['No.','Art','Datum','Rechnung Nr.','Artikel','Code','Firmennamen',
         'Ort der Lieferung','VAT No.','Menge','Einzelpreis €','Betrag €','Rabatt €',
         'Freight','Summe','Gewicht KG','Maßen (m³)','包裹净重 (kg)','包裹体积 (m³)']
TASK_HEADERS=['任务ID','发票号','Part No.','产品名称','数量','待补字段','填写数值','来源状态','估算依据 / 实测记录','目标单元格','历史参考（仅建议）']
MONTHS='一二三四五六七八九十十一十二'
FIELD_COL={'货物重量(kg)':'P','货物体积(m³)':'Q','包裹净重(kg)':'R','包裹体积(m³)':'S'}
CONFIRMED={'实测','供应商参数','历史已确认'}
ACCEPTED=CONFIRMED|{'AI估算','人工估算','历史参考待确认'}


def _a1(col:int,row:int)->str:
    s=''
    while col:s=chr(65+(col-1)%26)+s;col=(col-1)//26
    return f'{s}{row}'


def _text_cell(parent, col:int, row:int, value, style:int=0):
    if value is None or value=='':return
    c=ET.SubElement(parent,NS+'c',{'r':_a1(col,row),'s':str(style)})
    if isinstance(value,(int,float,Decimal)) and not isinstance(value,bool):
        c.set('t','n');ET.SubElement(c,NS+'v').text=format(value,'f') if isinstance(value,Decimal) else str(value)
    else:
        c.set('t','inlineStr');inline=ET.SubElement(c,NS+'is');ET.SubElement(inline,NS+'t').text=str(value)


def _make_sheet(rows:list[list], widths:list[float], kind:str='report', style_func=None) -> bytes:
    ws=ET.Element(NS+'worksheet')
    views=ET.SubElement(ws,NS+'sheetViews');v=ET.SubElement(views,NS+'sheetView',{'workbookViewId':'0'})
    ET.SubElement(v,NS+'pane',{'ySplit':'4' if kind=='report' else '1','topLeftCell':'A5' if kind=='report' else 'A2','activePane':'bottomLeft','state':'frozen'})
    cols=ET.SubElement(ws,NS+'cols')
    for i,w in enumerate(widths,1):ET.SubElement(cols,NS+'col',{'min':str(i),'max':str(i),'width':str(w),'customWidth':'1'})
    data=ET.SubElement(ws,NS+'sheetData')
    for r,row in enumerate(rows,1):
        height=('30' if r==1 else '24' if (r==3 and kind=='report') or (r==1 and kind!='report') else '21')
        el=ET.SubElement(data,NS+'row',{'r':str(r),'ht':height,'customHeight':'1'})
        for j,val in enumerate(row,1):
            style=style_func(r,j,val) if style_func else (3 if r==1 else 0)
            _text_cell(el,j,r,val,style)
    filter_to=3 if kind=='report' else 1
    # No automatic filtering for grouped report; filtering would hide continuation rows.
    if kind!='report' and rows:
        ET.SubElement(ws,NS+'autoFilter',{'ref':f'A{filter_to}:{_a1(max(1,len(widths)),len(rows))}'})
    ET.SubElement(ws,NS+'pageMargins',{'left':'.25','right':'.25','top':'.4','bottom':'.4','header':'.15','footer':'.15'})
    return ET.tostring(ws,encoding='utf-8',xml_declaration=True)


def _styles()->bytes:
    """Style indices: 0 normal,1 title,2 subtitle,3 heading,4 group,5 alt,6 missing,7 numeric,8 EUR,9 note,10 bad."""
    stylesheet=ET.Element(NS+'styleSheet')
    ET.SubElement(stylesheet,NS+'numFmts',{'count':'3'})
    fmt=stylesheet.find(NS+'numFmts')
    ET.SubElement(fmt,NS+'numFmt',{'numFmtId':'164','formatCode':'#,##0.00'})
    ET.SubElement(fmt,NS+'numFmt',{'numFmtId':'165','formatCode':'0.000000'})
    ET.SubElement(fmt,NS+'numFmt',{'numFmtId':'166','formatCode':'0.000'})
    fonts=ET.SubElement(stylesheet,NS+'fonts',{'count':'3'})
    for bold,color,size in ((False,'27374A',10),(True,'FFFFFF',11),(True,'17365D',11)):
        f=ET.SubElement(fonts,NS+'font')
        if bold:ET.SubElement(f,NS+'b')
        ET.SubElement(f,NS+'sz',{'val':str(size)})
        ET.SubElement(f,NS+'name',{'val':'Calibri'})
        ET.SubElement(f,NS+'color',{'rgb':'FF'+color})
    fills=ET.SubElement(stylesheet,NS+'fills',{'count':'7'})
    for rgb in (None,'FFFFFFFF','FF17365D','FFDCEAF6','FFF3F7FC','FFFFEDCC','FFFFE2E2'):
        e=ET.SubElement(fills,NS+'fill');p=ET.SubElement(e,NS+'patternFill',{'patternType':'none' if rgb is None else 'solid'})
        if rgb:ET.SubElement(p,NS+'fgColor',{'rgb':rgb});ET.SubElement(p,NS+'bgColor',{'indexed':'64'})
    borders=ET.SubElement(stylesheet,NS+'borders',{'count':'1'});b=ET.SubElement(borders,NS+'border')
    for edge in ('left','right','top','bottom','diagonal'):ET.SubElement(b,NS+edge)
    ET.SubElement(stylesheet,NS+'cellStyleXfs',{'count':'1'});ET.SubElement(stylesheet.find(NS+'cellStyleXfs'),NS+'xf',{'numFmtId':'0','fontId':'0','fillId':'0','borderId':'0'})
    xfs=ET.SubElement(stylesheet,NS+'cellXfs',{'count':'13'})
    entries=[(0,0,0),(1,2,0),(2,3,0),(1,2,0),(2,3,0),(0,4,0),(0,5,0),(0,0,164),(0,0,164),(0,4,0),(0,6,0),(0,0,165),(0,0,166)]
    for font,fill,num in entries:
        x=ET.SubElement(xfs,NS+'xf',{'fontId':str(font),'fillId':str(fill),'borderId':'0',
                 'numFmtId':str(num),'xfId':'0','applyAlignment':'1', 'applyNumberFormat':'1' if num else '0'})
        ET.SubElement(x,NS+'alignment',{'vertical':'center','wrapText':'1'})
    styles=ET.SubElement(stylesheet,NS+'cellStyles',{'count':'1'});ET.SubElement(styles,NS+'cellStyle',{'name':'Normal','xfId':'0','builtinId':'0'})
    return ET.tostring(stylesheet,encoding='utf-8',xml_declaration=True)


def _write_book(path:Path, sheets:list[tuple[str,list[list],list[float],str,object]]) -> None:
    workbook=ET.Element(NS+'workbook')
    sheetlist=ET.SubElement(workbook,NS+'sheets')
    wbrel=ET.Element('Relationships',{'xmlns':PKG_NS})
    overrides=[]
    for k,(name,rows,widths,kind,style) in enumerate(sheets,1):
        ET.SubElement(sheetlist,NS+'sheet',{'name':name,'sheetId':str(k),f'{{{REL_NS}}}id':f'rId{k}'})
        ET.SubElement(wbrel,'Relationship',{'Id':f'rId{k}','Type':REL_NS+'/worksheet','Target':f'worksheets/sheet{k}.xml'})
        overrides.append((f'/xl/worksheets/sheet{k}.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'))
    ET.SubElement(wbrel,'Relationship',{'Id':f'rId{len(sheets)+1}','Type':REL_NS+'/styles','Target':'styles.xml'})
    types=ET.Element('Types',{'xmlns':CONTENT_NS})
    ET.SubElement(types,'Default',{'Extension':'rels','ContentType':'application/vnd.openxmlformats-package.relationships+xml'})
    ET.SubElement(types,'Default',{'Extension':'xml','ContentType':'application/xml'})
    for name,ct in [('/xl/workbook.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),('/xl/styles.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),*overrides]:
        ET.SubElement(types,'Override',{'PartName':name,'ContentType':ct})
    rootrels=ET.Element('Relationships',{'xmlns':PKG_NS})
    ET.SubElement(rootrels,'Relationship',{'Id':'rId1','Type':REL_NS+'/officeDocument','Target':'xl/workbook.xml'})
    with zipfile.ZipFile(path,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as z:
        z.writestr('[Content_Types].xml',ET.tostring(types,encoding='utf-8',xml_declaration=True))
        z.writestr('_rels/.rels',ET.tostring(rootrels,encoding='utf-8',xml_declaration=True))
        z.writestr('xl/workbook.xml',ET.tostring(workbook,encoding='utf-8',xml_declaration=True))
        z.writestr('xl/_rels/workbook.xml.rels',ET.tostring(wbrel,encoding='utf-8',xml_declaration=True))
        z.writestr('xl/styles.xml',_styles())
        for i,(_name,rows,widths,kind,style) in enumerate(sheets,1):
            z.writestr(f'xl/worksheets/sheet{i}.xml',_make_sheet(rows,widths,kind,style))


def read_sheet(path:str|Path, sheet:str, columns:tuple[str,...]|None=None)->list[dict]:
    """Read XLSX values safely (no formulas, macros or external links executed)."""
    result=[]
    with zipfile.ZipFile(path) as z:
        sheetfile=_sheet_xml_name(z,sheet);strings=_shared_strings(z)
        with z.open(sheetfile) as fd:
            for _,el in ET.iterparse(fd,events=('end',)):
                if el.tag!=NS+'row':continue
                data={}
                for c in el.findall(NS+'c'):
                    letter=re.sub(r'\d','',c.get('r',''))
                    if columns is not None and letter not in columns:continue
                    data[letter]=_cell_text(c,strings)
                result.append(data);el.clear()
    return result


def load_history(reference_file:str|Path|None)->dict[str,dict]:
    """October/other report is a *suggestion*, not approved measurement.

    Report row P/Q store total line values, so divide by the invoiced quantity.
    Only repeated identical values across months are kept as suggestions.
    """
    if not reference_file or not Path(reference_file).is_file():return {}
    try:
        with zipfile.ZipFile(reference_file) as z:
            wb=ET.fromstring(z.read('xl/workbook.xml'))
            tabs=[x.get('name','') for x in wb.find(NS+'sheets')]
        sh=next((x for x in tabs if re.search(r'\d+月$',x)),None)
        if sh is None:return {}
        rows=read_sheet(reference_file,sh,('F','J','P','Q'))
    except (OSError,ValueError,KeyError,zipfile.BadZipFile):return {}
    found={}
    for row in rows:
        part=str(row.get('F','')).strip();qty=row.get('J','')
        try:q=Decimal(qty)
        except (ValueError,InvalidOperation):continue
        if not part or q<=0:continue
        ref={}
        for col,k in (('P','unit_weight_kg'),('Q','unit_volume_m3')):
            try:
                val=Decimal(row.get(col,''))/q
                if val>0:ref[k]=float(val.quantize(Decimal('0.0000001')))
            except (ValueError,InvalidOperation):continue
        if not ref:continue
        key=normalize_number(part)
        if key in found:
            if found[key]!=ref:found[key]={'ambiguous':True}
        else:found[key]=ref
    return {k:v for k,v in found.items() if not v.get('ambiguous')}


def load_approved_library(path:str|Path)->dict:
    try:
        obj=json.loads(Path(path).read_text(encoding='utf-8'))
        if not isinstance(obj,dict):return {}
        return {str(k):v for k,v in obj.items() if isinstance(v,dict)}
    except (OSError,ValueError):return {}


def _s_money(x):return float(x) if x is not None else None

def generate_report(invoices, outcomes, year:int, month:int, output:str|Path,
                    reference_file:str|Path|None=None, library_path:str|Path|None=None,
                    inventory_file:str|Path|None=None)->dict:
    """Write 19-column worksheet + AI tasks + audit log + history suggestions."""
    output=Path(output);output.parent.mkdir(parents=True,exist_ok=True)
    history=load_history(reference_file)
    approved=load_approved_library(library_path) if library_path else {}
    inventory_index=None
    if inventory_file:
        from gui_inventory_core import load_inventory
        inventory_index=load_inventory(inventory_file)
    title=f'{month}月'
    detail=[['Hangcha Europe GmbH'],[f'Versendungsstatistikmeldung {month:02d}.{year}'],
            HEADERS,[]]
    tasks=[TASK_HEADERS]; task_count=0; automatic=0
    serial=0
    for inv in invoices:
        for idx,item in enumerate(inv.items):
            serial+=1;line_no=len(detail)+1
            first=idx==0
            line=[serial,'', inv.day.strftime('%d.%m.%Y') if first else '', inv.invoice if first else '',
                  item.description,item.part,inv.customer,inv.country,inv.vat_id,_s_money(item.quantity),
                  _s_money(item.unit_price),_s_money(item.line_total),
                  _s_money(inv.discount) if first and inv.discount is not None else None,
                  _s_money(inv.freight) if first and inv.freight is not None else None,
                  _s_money(inv.net) if first else None,None,None,None,None]
            ref=history.get(normalize_number(item.part),{})
            inventory_hint=''
            if inventory_index:
                found=inventory_index.find_item(item.erp,item.part)
                if found.hits:
                    sample=found.hits[0]
                    inventory_hint='库存 Excel 原始字段(未确认): 重量 '+(sample.weight or '空')+' / 尺寸 '+(sample.dimension or '空')
                    if len(found.hits)>1 or found.warning:inventory_hint+='；存在多个库位或编号需核对'
            approved_ref=approved.get(normalize_number(item.part),{})
            for col,field,code in [('P','货物重量(kg)','unit_weight_kg'),('Q','货物体积(m³)','unit_volume_m3')]:
                num=approved_ref.get(code)
                if isinstance(num,(int,float)) and 0<num<100000:
                    line[ord(col)-65]=round(float(num)*float(item.quantity),6);automatic+=1
                else:
                    suggestion=ref.get(code)
                    tasks.append([f'{inv.invoice}:{line_no}:{col}',inv.invoice,item.part,item.description,
                                  float(item.quantity),field,None,None,None,f'{col}{line_no}',
                                  ('历史参考每件 '+str(suggestion)+' (未确认)；' if suggestion else '')+inventory_hint])
                    task_count+=1
            if first:
                for col,field in [('R','包裹净重(kg)'),('S','包裹体积(m³)')]:
                    tasks.append([f'{inv.invoice}:{line_no}:{col}',inv.invoice,item.part,
                                  '【整张发票/所有包裹】 '+inv.customer,float(item.quantity),field,
                                  None,None,None,f'{col}{line_no}',
                                  '包裹值不能直接等同于货物值；仅供参考'])
                    task_count+=1
            detail.append(line)
        detail.append([])
    logs=[['来源文件','发票号','结果','原因','配件行数']]
    for x in outcomes:logs.append([x.file,x.invoice,x.state,x.reason,x.count])
    refs=[['Part No.','历史每件重量kg','历史每件体积m³','来源','可信程度']]
    for part,info in sorted(history.items()):refs.append([part,info.get('unit_weight_kg'),info.get('unit_volume_m3'),str(reference_file or ''),'仅参考，需核实'])
    guides=[['使用说明 / 使用前必读'],
        ['筛选规则','先从业务周汇总 Excel 的「发货」子表读取指定月份（F/G/I 发票日期），D「发票金额」等于 E「Netto」且金额大于零的发票号；然后按发票号查找 PDF，核对内容并排除瑞士。'],
        ['金额准确性','汇总 Excel 的 D/E 是筛选索引；单价、行金额、Netto、运费来自发票正文，与 D/E 不符时待复核，不导入报表。'],
        ['有效性','Excel 同一发票号重复、对应 PDF 缺失或编号/金额不符，以及作废/有误、贷项单、不同版本、缺明细等均需核查。'],
        ['资料来源','历史报告只作为参考，AI估算不当作实测值。经用户确认的数据可保存到本地尺寸库。'],
        ['AI补全过程','上传此半成品及 Prompt.txt 到 ChatGPT。只填写 AI待补全 的 G(数值)、H(来源状态)、I(依据)。'],
        ['导入结果','回软件选择原半成品及AI补全文件；软件只采纳对应 G/H/I，原财务字段完全不变。'],
        ['人工复核','包裹重量体积应来自实际包装/物流信息，AI估算值不得当作测量值。'],
        ['报表性质','公司内部使用；不是德国 Intrastat 申报表。'],
        ['提示','某些项目若未能从 PDF 识别会列入筛选日志，需检查后重新生成。']]
    widths1=[8,14,16,20,50,28,33,24,20,12,16,16,16,14,16,18,19,19,19]
    def main_style(r,c,v):
        if r==1:return 1
        if r==2:return 2
        if r==3:return 3
        if r==4:return 0
        if c in (16,17,18,19) and v in (None,''):return 6
        if c in (17,19):return 11  # volumes need six decimal places
        if c in (16,18):return 12 # weights need at least 0.001 kg
        if c in (10,11,12,13,14,15):return 8
        return 4 if len(detail[r-1])>3 and detail[r-1][3] else (5 if r%2==0 else 0)
    def table_style(r,c,v):
        if r==1:return 3
        if c==7 and (v is None or v==''):return 6
        return 5 if r%2==0 else 0
    sheets=[(title,detail,widths1,'report',main_style),
            ('AI待补全',tasks,[27,21,28,50,11,24,16,19,52,16,46],'table',table_style),
            ('筛选日志',logs,[76,21,15,90,12],'table',table_style),
            ('历史参考',refs,[32,21,23,70,23],'table',table_style),
            ('说明',guides,[30,112],'table',table_style)]
    _write_book(output,sheets)
    return {'invoices':len(invoices),'lines':serial,'tasks':task_count,'historical_applied':automatic,
            'review':sum(o.state=='待复核' for o in outcomes),'output':str(output)}



def export_clean_report(source_xlsx: str|Path, output_xlsx: str|Path) -> dict:
    """Export JUST the A:S monthly report from a half or AI-merged workbook.

    Rebuild the main worksheet independently from the trusted source workbook,
    without carrying any AI tasks / logs / explanatory material across.
    """
    src=Path(source_xlsx);dest=Path(output_xlsx)
    with zipfile.ZipFile(src) as z:
        names=_sheet_names(z)
    title=next((n for n in names if re.fullmatch(r'\d{1,2}月',n)),None)
    if not title:raise ValueError('输入文件未包含月份统计工作表')
    entries=read_sheet(src,title,tuple('ABCDEFGHIJKLMNOPQRS'))
    if len(entries)<3:raise ValueError('月度统计工作表为空')
    actual=[entries[2].get(chr(65+i),'') for i in range(19)]
    if actual!=HEADERS:raise ValueError('月报 A-S 表头不匹配，拒绝生成正式统计文件')
    rows=[]
    for row in entries:
        values=[]
        for ci in range(19):
            col=chr(65+ci);val=row.get(col,'')
            if len(rows)>=4 and ci in range(9,19) and val!='':
                try:val=Decimal(val)
                except (InvalidOperation,ValueError):
                    raise ValueError(f'统计主表 {col}{len(rows)+1} 不是有效数字')
            values.append(val)
        rows.append(values)
    # Never write a policy sentence into the business sheet, even if the
    # input came from an older version that had it in row four.
    while len(rows)<4:rows.append(['']*19)
    rows[3]=['']*19
    widths=[8,14,16,20,50,28,33,24,20,12,16,16,16,14,16,18,19,19,19]
    def style(r,c,v):
        if r==1:return 1
        if r==2:return 2
        if r==3:return 3
        if c in (16,17,18,19) and v in (None,''):return 6
        if c in (17,19):return 11
        if c in (16,18):return 12
        if c in (10,11,12,13,14,15):return 8
        return 4 if len(rows[r-1])>3 and rows[r-1][3] else (5 if r%2==0 else 0)
    dest.parent.mkdir(parents=True,exist_ok=True)
    if src.resolve()==dest.resolve():raise ValueError('原工作簿与独立月报不得是同一文件')
    _write_book(dest,[(title,rows,widths,'report',style)])
    return {'output':str(dest),'sheets':[title],'rows':len(rows)}

def build_prompt(year:int,month:int)->str:
    return f'''# Hangcha Europe GmbH - {year}-{month:02d} 内部月度发运统计补全\n\n我上传了由 HCE Oracle 基于业务周汇总 Excel（发货表）发票号筛选、再校验对应 PDF 所生成的半成品 Excel。请认真读取。\n此表仅用于公司内部统计，不是海关或 Intrastat 正式申报。\n\n【唯一允许修改的区域】\n工作表「AI待补全」：只填写 G 列（待补数值）、H 列（来源状态）、I 列（具体依据）。\n不要修改任何其他列、行号、任务ID、单据号、商品编码或其他工作表。\n每行仅为一个缺失值，严格按 F 列字段来填；G 列必须是正数，使用 Excel 数字格式（不要附 kg/m³ 字符）。\n\n【填写原则】\n1. 先参考「历史参考」的相同 Part No.，但历史参考并非已核实的真实测量值。\n2. P/Q 对应整行全部货物重量(kg)/体积(m³)，务必考虑 E 列数量。\n3. R/S 对应整张发票的实际包裹净重/包裹体积，不能简单把货物体积直接当成包裹体积。\n4. 若有可信型号规格或准确来源，在 I 列说明依据，H 列填「供应商参数」或「历史已确认」。\n5. 只能合理推断时，H 列填「AI估算」，在 I 列说明假设、合理区间及不确定性。\n6. 如果无法合理估计，就让 G 列保持空白，H 列写「需要人工测量」并注明原因；切勿编造精确值。\n7. 不得修改财务数字、币种、客户、税号或行顺序。不要伪造对网络或供应商的核验。\n\n【最终输出】\n请返回一份保留原文件全部工作表和格式的 .xlsx 文件。\n只在「AI待补全」的 G/H/I 列填内容。\n我将把原始半成品和你返回的文件一起导入 HCE Oracle 软件：软件会校验任务ID，只合并这三个字段对应的重量、体积，不信任其他任何单元格。\n'''


def _rows_map(rows:list[dict])->dict[str,dict]:
    result={}
    for row in rows[1:]:
        key=row.get('A','').strip()
        if key:
            if key in result:raise ValueError('AI待补全中存在重复任务ID：'+key)
            result[key]=row
    return result


def _set_value(root:ET.Element,addr:str,value:object)->None:
    data=root.find(NS+'sheetData'); rowno=re.search(r'\d+',addr).group(0)
    row=next((r for r in data.findall(NS+'row') if r.get('r')==rowno),None)
    if row is None:raise ValueError('找不到目标行 '+addr)
    old=next((c for c in row.findall(NS+'c') if c.get('r')==addr),None)
    if old is None:
        old=ET.SubElement(row,NS+'c',{'r':addr,'s':'11' if addr[0] in 'QS' else '12'})
    if old.find(NS+'v') is not None or old.find(NS+'is') is not None:
        raise ValueError(f'目标单元格 {addr} 已有内容，不允许AI覆盖')
    old.set('t','n');ET.SubElement(old,NS+'v').text=str(value)


def _set_task(root:ET.Element,row_number:int,column:str,value:object):
    data=root.find(NS+'sheetData')
    row=next((r for r in data.findall(NS+'row') if r.get('r')==str(row_number)),None)
    if row is None:raise ValueError('任务行不存在')
    addr=f'{column}{row_number}'
    cell=next((c for c in row.findall(NS+'c') if c.get('r')==addr),None)
    if cell is None:cell=ET.SubElement(row,NS+'c',{'r':addr,'s':'11' if column=='G' else '0'})
    if column=='G':cell.set('s','11')
    for ch in list(cell):cell.remove(ch)
    if column=='G':cell.set('t','n');ET.SubElement(cell,NS+'v').text=str(value)
    else:
        cell.set('t','inlineStr');ET.SubElement(ET.SubElement(cell,NS+'is'),NS+'t').text=str(value)


def import_ai(original:str|Path, completed:str|Path, output:str|Path)->dict:
    """Merge *only* numeric G fields keyed by immutable task ID into ORIGINAL.

    The AI workbook is never used as the output baseline; original invoices and
    amounts therefore cannot be accidentally modified by a language model.
    """
    original,completed,output=map(Path,(original,completed,output))
    if original.resolve()==output.resolve() or completed.resolve()==output.resolve():
        raise ValueError('导出文件必须与输入文件不同')
    base=read_sheet(original,'AI待补全'); ai=read_sheet(completed,'AI待补全')
    if not base or not ai:raise ValueError('缺少 AI待补全 工作表')
    expected=_rows_map(base);provided=_rows_map(ai)
    errors=[];updating=[];seen=set()
    for rowno,row in enumerate(base[1:],2):
        key=row.get('A','').strip()
        if not key:continue
        given=provided.get(key)
        if given is None:errors.append(f'{key} AI文件缺少任务');continue
        if any(str(row.get(c,'')).strip()!=str(given.get(c,'')).strip() for c in ('A','B','C','D','E','F','J')):
            errors.append(f'{key} 任务信息被AI修改，已拒绝');continue
        val=given.get('G','').strip();status=given.get('H','').strip();basis=given.get('I','').strip()
        if not val:continue
        if status not in ACCEPTED or not basis:
            errors.append(f'{key} 缺少可接受的来源状态/依据，已跳过');continue
        try:qty=Decimal(val.replace(',','.'))
        except InvalidOperation:errors.append(f'{key} 数值不是数字');continue
        if not qty.is_finite() or qty<=0 or qty>100000:
            errors.append(f'{key} 数值超出合理检查范围');continue
        target=row.get('J','').strip().upper()
        if not re.fullmatch('[PQRS][1-9][0-9]*',target):errors.append(f'{key} 目标列异常');continue
        if target in seen:errors.append(f'{key} 目标单元格重复');continue
        seen.add(target)
        updating.append((target,qty,rowno,status,basis))
    if not updating:raise ValueError('没有可安全导入的数值。'+('\n'+'\n'.join(errors[:8]) if errors else ''))
    with zipfile.ZipFile(original) as z:
        names=z.namelist()
        base_sheet=_sheet_xml_name(z,next((x for x in _sheet_names(z) if re.search(r'^\d+月$',x)),''))
        task_sheet=_sheet_xml_name(z,'AI待补全')
        main=ET.fromstring(z.read(base_sheet));tasks=ET.fromstring(z.read(task_sheet))
        done=[]
        for target,qty,rowno,status,basis in updating:
            try:
                _set_value(main,target,qty)
                _set_task(tasks,rowno,'G',qty);_set_task(tasks,rowno,'H',status);_set_task(tasks,rowno,'I',basis)
                done.append((target,qty))
            except ValueError as exc:errors.append(str(exc))
        if not done:raise ValueError('全部补全值都未通过安全校验')
        output.parent.mkdir(parents=True,exist_ok=True)
        with zipfile.ZipFile(output,'w',compression=zipfile.ZIP_DEFLATED) as dst:
            for inf in z.infolist():
                if inf.filename==base_sheet:data=ET.tostring(main,encoding='utf-8',xml_declaration=True)
                elif inf.filename==task_sheet:data=ET.tostring(tasks,encoding='utf-8',xml_declaration=True)
                else:data=z.read(inf.filename)
                dst.writestr(inf,data)
    note=output.with_name(output.stem+'_导入核验.txt')
    note.write_text('HCE Oracle 月度报告 - AI补全安全导入核验\n'
                    f'成功写入字段：{len(done)} / {len(expected)}\n'
                    '仅保留原半成品中的财务数据，其它AI工作表改动一律忽略。\n'
                    'AI估算字段仍需管理员复核，不能当作实测。\n\n'
                    +'\n'.join(f'{a} = {v}' for a,v in done)
                    +'\n\n未导入/异常：\n'+'\n'.join(errors),encoding='utf-8')
    return {'filled':len(done),'tasks':len(expected),'errors':errors,'output':str(output),'review_txt':str(note)}


def _sheet_names(z)->list[str]:
    root=ET.fromstring(z.read('xl/workbook.xml'))
    return [el.get('name','') for el in root.find(NS+'sheets')]


def update_approved_library(checked_report:str|Path,library_file:str|Path)->dict:
    """Only explicitly marked human-confirmed item totals can enter library."""
    rows=read_sheet(checked_report,'AI待补全')
    original=load_approved_library(library_file)
    updates=0
    for row in rows[1:]:
        status=row.get('H','').strip();basis=row.get('I','').strip()
        if status not in CONFIRMED or not basis:continue
        try: total=Decimal(row.get('G',''));quantity=Decimal(row.get('E',''))
        except (InvalidOperation,ValueError):continue
        if quantity<=0 or total<=0:continue
        field=row.get('F','');name={'货物重量(kg)':'unit_weight_kg','货物体积(m³)':'unit_volume_m3'}.get(field)
        if not name:continue
        part=normalize_number(row.get('C',''))
        if not part:continue
        unit=float((total/quantity).quantize(Decimal('0.0000001'),rounding=ROUND_HALF_UP))
        original.setdefault(part,{})[name]=unit
        original[part][name+'_source']=f'{status}: {basis}';updates+=1
    p=Path(library_file);p.parent.mkdir(parents=True,exist_ok=True)
    if updates:
        fd,tmp=tempfile.mkstemp(prefix='dims_',suffix='.tmp',dir=p.parent)
        try:
            with os.fdopen(fd,'w',encoding='utf-8') as f:json.dump(original,f,ensure_ascii=False,indent=2)
            os.replace(tmp,p)
        finally:
            if os.path.exists(tmp):os.unlink(tmp)
    return {'updated_fields':updates,'parts':len(original)}
