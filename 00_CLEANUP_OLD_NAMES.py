"""Safe one-time tidy-up for the damaged filenames of the v1.1.0 ZIP.
Moves *only known release files* into backup instead of deleting. Does not
read, move, rename, or modify config.json, .venv, state, or business records.
"""
from pathlib import Path
from datetime import datetime
import shutil

ROOT = Path(__file__).resolve().parent
OLD_FILES = {
    '01_安装依赖.bat': '01_INSTALL.bat',
    '02_配置目录.bat': '02_CONFIGURE.bat',
    '03_只读预览.bat': '03_PREVIEW.bat',
    '04_正式执行.bat': '04_APPLY.bat',
    '05_安装工作日定时任务.bat': '05_ENABLE_SCHEDULE.bat',
    '06_删除定时任务.bat': '06_DISABLE_SCHEDULE.bat',
    '07_PDF识别诊断_只读.bat': '07_PDF_DIAGNOSTIC.bat',
    '08_规则自检_只读.bat': '08_RULE_TESTS.bat',
    '09_发货单识别诊断_只读.bat': '09_DELIVERY_DIAGNOSTIC.bat',
    '10_启动目录扫描GUI.bat': '10_START_GUI.bat',
    '11_GUI调试启动.bat': '11_GUI_DEBUG.bat',
    '12_GUI目录扫描自检.bat': '12_GUI_TESTS.bat',
    '开始安装.txt': 'README_START_HERE.txt',
    'GUI使用说明_v1.1_请先看.txt': 'README_GUI.txt',
    '使用说明_从这里开始.md': 'README_FULL.md',
    '版本信息.txt': 'VERSION.txt',
}

# CP437 represents the exact bad filenames from the previous ZIP; Windows may
# also have extracted the real Unicode names. Both forms are recognized.
old_mojibake = {k.encode('utf-8').decode('cp437'):v for k,v in OLD_FILES.items()}
recognized = {**OLD_FILES, **old_mojibake}
files = [p for p in ROOT.iterdir() if p.is_file() and p.name in recognized]
if not files:
    print('No old Chinese/garbled launcher files found. Folder is already clean.')
    raise SystemExit(0)
backup_dir=ROOT/'legacy_filename_backup'
backup_dir.mkdir(exist_ok=True)
print('Moving', len(files), 'known legacy files to:', backup_dir)
for p in files:
    if p.name.isascii():
        continue
    new_basename=recognized[p.name]
    # Keep originals, even if a user edited one manually.
    safe_suffix=p.suffix.lower().lstrip('.') or 'file'
    safe_name = 'old_' + new_basename + '.' + safe_suffix + '.bak'
    dst = backup_dir / safe_name
    counter=2
    while dst.exists():
        dst=backup_dir/f'old_{new_basename}.{counter}.{safe_suffix}.bak'
        counter+=1
    try:
        shutil.move(str(p), str(dst))
        print('BACKUP:', new_basename, '->', dst.name)
    except Exception as exc:
        print('SKIPPED:', new_basename, repr(exc))
print('Completed. Launch GUI with: 10_START_GUI.bat')
