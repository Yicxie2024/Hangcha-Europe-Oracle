HCE Oracle v1.5.31 — 软件更名
================================

新品牌：HCE Oracle
原名称：Hangcha Sales Assistant

本次修改：
- 主 GUI 窗口标题与左上角品牌：HCE Oracle。
- 库存查询、月度发运统计、Teams 上传子窗口标题：HCE Oracle。
- 拣货单标题/PDF 作者信息与页脚：HCE Oracle，保留正式公司名称 Hangcha Europe。
- 月度 AI 补全 Prompt 中的软件称呼：HCE Oracle。
- 安装程序创建 HCE Oracle.lnk，而不是旧名称快捷方式。
- 可单独运行 02_UPDATE_SHORTCUT.bat 更新现有桌面快捷方式，无需重装。
- 版本号从 1.5.30 改为 1.5.31。

兼容性：
- 原 LOCALAPPDATA/HangchaSalesAssistant 目录、SQLite 数据、归档、设置、个人配置、PDF 文件及 Excel 路径完全保留。
- 旧版 Python module/class 名称保留，防止任何内部引用失效。
- 路径规划、库存查找与 PDF 识别逻辑未变。
- 替换补丁前关闭软件，并备份原程序目录。

提示：HCE Oracle 是用户选择的内部软件名称；如准备对外商业发布，建议事先核查第三方商标和品牌合规。
