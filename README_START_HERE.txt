HCE Oracle v1.5.34 | Hangcha Europe 内部业务助手
================================================

本版重点：统一订单/发货/发票号显示；按 PDF 正文 Ordernumber 关联历史归档订单；
人工确认的批量补充归档；默认快速增量扫描及 SQLite 识别缓存。
详细说明：README_v1.5.34.txt。

【从 v1.5.33 升级】
1. 关闭软件，备份当前程序文件夹和本地 archive_state.sqlite3。
2. 将 HCE_Oracle_v1.5.34_Incremental_Archive_Patch.zip 解压到 v1.5.33 程序目录，覆盖同名文件。
3. 使用桌面快捷方式或 10_START_GUI.bat 启动，无需重装 Python。
4. 点击“开始扫描（只读）”：默认快速增量模式，首次创建缓存需要解析 PDF。
5. 发货单遗留：进入「其他功能」→「检查历史发货单归档关联（需确认）…」，
   检查每份候选文件和无法确认的原因，确认后才会实际补归档。

【首次安装】
1. 解压 HCE_Oracle_v1.5.34_Full_Windows.zip。
2. 运行 01_INSTALL.bat 安装依赖并创建 HCE Oracle 桌面快捷方式。
3. 后续双击快捷方式或 10_START_GUI.bat 启动。

【提醒】
- 历史归档和员工配置保存在 %LOCALAPPDATA%\HangchaSalesAssistant；新增缓存 scan_cache.sqlite3。
- 不删除或改写原始 PDF、Excel；未确认前不会自动补充归档。
- 归档仍须确认 Excel 业务记录已登记，免得遗漏后续新增写入。
- 怀疑归档源文件被替换时选「其他功能」→「扫描模式」→「完整归档核验（SHA-256）」。
- 首次建立缓存需扫描旧文件，之后未改变的文件可复用缓存。

历史版本说明见 README_v1.5.33.txt 及更早说明。
