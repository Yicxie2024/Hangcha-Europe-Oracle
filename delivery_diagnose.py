"""Read-only delivery note contract reference diagnostics (v1.0.5)."""
from datetime import datetime
from pathlib import Path
import json
from automation import (BASE, CONFIG_FILE, REPORTS, discover_delivery_notes,
                        pdf_delivery_contract_refs)


def main():
    if not CONFIG_FILE.exists():
        print('请先运行 02_CONFIGURE.bat，设置业务根目录。')
        return 1
    cfg = json.loads(CONFIG_FILE.read_text(encoding='utf-8-sig'))
    root = Path(cfg['root_path']).expanduser()
    if not root.is_dir():
        print('目录不存在：', root)
        return 1
    documents = discover_delivery_notes(root)
    good = []
    unmatched = []
    failures = []
    for path, client in documents:
        try:
            refs = sorted(pdf_delivery_contract_refs(path))
        except Exception as exc:
            failures.append(f'{client} | {path.name} | {exc}')
            continue
        if refs:
            good.append(f'{client} | 合同号 {", ".join(refs)} | {path.name}')
        else:
            unmatched.append(f'{client} | {path.name} | 未找到明确 Ordernumber/合同号（不要猜测发货单号）')
    stamp = datetime.now()
    REPORTS.mkdir(parents=True, exist_ok=True)
    out = REPORTS / f'{stamp:%Y%m%d_%H%M%S}_delivery_diagnostic.txt'
    lines = [f'发货单识别诊断（只读） | {stamp:%Y-%m-%d %H:%M:%S}',
             f'扫描发货单 PDF：{len(documents)}', f'明确关联合同号：{len(good)}',
             f'未能关联：{len(unmatched)}', f'读取失败：{len(failures)}', '',
             '==== 明确关联合同 ===='] + good + [
             '', '==== 无法关联（需要发货单样本） ===='] + unmatched + [
             '', '==== PDF 读取异常 ===='] + failures
    out.write_text('\n'.join(lines), encoding='utf-8-sig')
    print('\n'.join(lines[:5]))
    print('详细诊断：', out)
    print('未修改 PDF、Excel 或文件夹。')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
