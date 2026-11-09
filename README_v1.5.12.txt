Hangcha Sales Assistant v1.5.12 — 发货单 PDF 正文识别修复
========================================================

本次针对 v1.5.10 起大量发货单显示“未分类 PDF”的回归问题进行修复。
安装：关闭 GUI 后备份旧目录；把升级包中的同名 .py 覆盖到原程序目录，保留 state/config.json/.venv；运行 10_START_GUI.bat，重新完整扫描。

规则：
1. PDF 正文中的 Delivery order / Delivery note / Deliverynote / Lieferschein 识别为发货单，文件名和目录不能决定类型、客户、合同号。
2. 发货单左上角收货地址的公司名作为发货单行“客户”列；即使它与订单/发票客户不同，也不被订单客户覆盖。
3. 只用发货单正文明确标注的 Ordernumber 关联内部合同，不用 Delivery No. 和 Customer order number。示例：Delivery No. 266954 关联 Ordernumber 266953。
4. “合同号 / 单据号”列：订单显示 Order No.、发票显示 Invoice No.、发货单显示 Delivery No.；实际关联仍通过内部合同号。
5. 发货单行“PDF 日期”显示明确标注的 Date of departure。该日期仅用于展示，不用于 Excel“发货”表；后者继续仅凭有效发票 PDF，默认使用 Invoice date。
6. 正文没有唯一的 Ordernumber、收货公司或发货单号/离港日期时，仍识别为“发货单”，但标记待复核；多个 Ordernumber 均记录到说明中，暂不自动挂到单个合同以免误关联。
7. 用户约定的文件名操作标记“作废”、“有误”仍保留原有含义；其余业务字段不再从文件名或目录猜测。
8. 月末缺票判断也使用发货单正文的 Ordernumber；不把 Delivery No. 当作合同号。

实际 PDF 验证：
BUW Delivery note 266954 (03-11-26).pdf
-> 收货公司 BUW Stapler-Service GmbH；Delivery No. 266954；Ordernumber 266953；Date of departure 03.11.2026。

安全：不删除、不移动 PDF，不自动改写 Excel。保留原 SQLite 归档、备注、上传记录。
建议：重新扫描后检查“发货单”计数和剩余未分类文件，并针对其他模板用原 PDF 验证。
