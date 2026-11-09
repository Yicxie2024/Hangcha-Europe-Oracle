"""Legacy task launcher disabled for safety.
Enable only PREVIEW-ONLY schedule after in-GUI confirmation.
"""
from __future__ import annotations

def main():
    print('Safety: scheduled Excel writing is disabled. Open 10_START_GUI.bat,')
    print('then use 定时检查设置 and approve a PREVIEW-ONLY scheduled scan.')
    return 3

if __name__ == '__main__':
    raise SystemExit(main())
