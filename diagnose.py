"""Read-only local PDF parsing diagnostic. Never opens or changes any Excel workbook."""
from pathlib import Path
from datetime import datetime
import json
from collections import Counter
from automation import CONFIG_FILE, REPORTS, discover, extract_pdf, summary_excerpt


def main():
    if not CONFIG_FILE.exists():
        raise SystemExit('请先运行 02_CONFIGURE.bat，保存工作文件夹路径。')
    cfg = json.loads(CONFIG_FILE.read_text(encoding='utf-8-sig'))
    root = Path(cfg['root_path']).expanduser()
    if not root.is_dir():
        raise SystemExit(f'工作目录不存在：{root}')
    docs = discover(root)
    good = []
    failed = []
    for path,kind,customer,week in docs:
        try:
            rec = extract_pdf(path,kind,customer,cfg.get('salesperson','Xie'))
            good.append((rec.kind,rec.contract,rec.invoice))
        except Exception as e:
            msg = str(e)
            failed.append((path,kind,msg))
    stamp=datetime.now().strftime('%Y%m%d_%H%M%S')
    REPORTS.mkdir(parents=True,exist_ok=True)
    out = REPORTS / f'{stamp}_PDF诊断_只读.txt'
    lines = ['杭叉 PDF 提取诊断（只读，不操作 Excel）', f'扫描 {len(docs)} 份，成功 {len(good)}，需要复核 {len(failed)}',
             '说明：本报告只显示本地文件路径与局部金额汇总信息，不读取发货单、运单，也不自动更正合同号。',
             '', '==== 需要复核的文件及精确原因 ====']
    for path,kind,why in failed:
        lines.append(f'[{kind}] {path}\n  失败原因：{why}')
    lines.extend(['', '==== 已成功提取的记录（仅编号，不含金额） ===='])
    for kind,contract,invoice in good:
        lines.append(f'{kind} {contract} {invoice}')
    out.write_text('\n'.join(lines),encoding='utf-8-sig')
    print(f'只读诊断结束：PDF {len(docs)}，提取成功 {len(good)}，需要复核 {len(failed)}')
    print(f'诊断报告：{out}')
    print('未修改 Excel、PDF 或任何业务文件夹。')

if __name__ == '__main__':
    main()
