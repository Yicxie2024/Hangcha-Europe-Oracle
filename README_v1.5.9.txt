Hangcha Sales Assistant v1.5.9 — 客户名称跨列识别修复
==================================================
基于 v1.5.8，仅修复订单 PDF 页眉客户地址区域的文本识别。

复现与原因
  真实订单：BUW Stapler-Service Customer order 266788 (28-10-26).pdf
  正确客户：BUW Stapler-Service GmbH
  原 v1.5.8：识别出 Delivery ad（来自右侧 Delivery address）
  原因：左侧客户地址裁剪区域向右延伸到页面 67%，与右侧收货地址重叠。

修复
  * 严格读取已验证模板的左侧客户地址列（截至页面宽度 55%）。
  * 排除 Delivery address/Shipping address 等收货地址标签作为客户名。
  * 若客户名称可能在裁剪边界被截断，宁可返回不确定，使用现有的
    文件夹/文件名兜底逻辑，不自动生成截断的正式客户名称。
  * 合同号关联发票和发货单、Excel 已有记录不覆盖、备注、归档、
    Teams 发票上传均沿用 v1.5.8 的既有逻辑。

真实 PDF 验证
  266788 -> BUW Stapler-Service GmbH
  266716 -> Passmann Gabelstapler
  266686 -> KLOZ Gabelstapler GmbH
  共 16 项客户识别单元/集成测试通过。
  Windows 上的真实运行仍需用户复核。

升级（v1.5.8 -> v1.5.9）
  1. 关闭 Hangcha Sales Assistant。
  2. 备份整个安装目录（尤其 state）与用户 LOCALAPPDATA 数据。
  3. 用补丁中的 gui_customer.py 覆盖安装目录中的同名文件；
     VERSION.txt 和 README_v1.5.9.txt 可一同复制。
  4. 不要删除 SQLite、.venv、个人配置、PDF 或 Excel。
  5. 通过 10_START_GUI.bat 启动并重新扫描；检查合同 266788 的客户名。

备注：本补丁不重写历史 Excel 中已经存储的客户名称；如旧版曾把错误客户
名称实际写进 Excel，需要人工核查、更正已有的 Excel 单元格。
