import json
from pathlib import Path

base=Path(__file__).resolve().parent
cfg=base/'config.json'
print('业务周汇总自动化：请选择包含「周文件夹」及「业务周汇总表」的根目录。')
folder=''
try:
    import tkinter as tk
    from tkinter import filedialog
    root=tk.Tk();root.withdraw();root.attributes('-topmost',True)
    folder=filedialog.askdirectory(title='请选择周文件夹和Excel所在的总目录')
    root.destroy()
except Exception:
    pass
if not folder:
    folder=input('输入完整的根目录路径（如 C:\\Company\\Weekly）：').strip().strip('"')
if not folder or not Path(folder).is_dir():
    raise SystemExit('目录不存在，配置未保存。')
path=Path(folder).resolve()
print('目录：',path)
if not list(path.glob('业务周汇总表_XYC_更新至*.xlsx')):
    print('提示：目录中尚未找到业务周汇总表，请确认选对了目录。')
config={'root_path':str(path),'salesperson':'Xie','rename_current_week_folder':True,'popup_on_errors':True}
cfg.write_text(json.dumps(config,ensure_ascii=False,indent=2),encoding='utf-8')
print('配置已保存：',cfg)
