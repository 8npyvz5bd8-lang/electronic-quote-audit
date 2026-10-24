---
name: electronic-quote-audit
description: "在本机按客户确认的列映射核对电子元件采购清单及2–5家CSV/XLSX报价，严格匹配制造商与完整料号，展开数量、阶梯与Decimal金额，输出比价表、逐项澄清和原始证据。用于采购报价核对、MOQ/倍数或计价单位检查；不自动采购、不推断替代、不上传报价。"
---

# 电子元件采购报价核对

仅用 Python 3.12 标准库在本地核对，不需要模型、API key、联网或新增付费依赖。先读[输入及规则](references/schema.md)；术语边界见[来源](references/sources.md)。

## 执行顺序

1. 让用户指定可读取的本地根目录、一份需求和2–5家报价。支持 UTF-8 CSV、纯值 XLSX。不要打开其他项目或雇主私有资料，不上传报价。
2. 列出实际文件、表名、表头及待映射字段。与用户逐项确认数量单位、每价分母、阶梯基准、MOQ、订购倍数、Minimum Release、税费、运费范围和评估时点。表头通过配置映射，无需改写原始文件。缺失条件保留待确认；不要补默认值。
3. 按 schema 写 JSON 配置。`mapping_confirmed`、`customer_confirmed`、`notes_reviewed` 只能来自用户真实核对确认，不能由模型替用户填 true。供应商表格、备注或文件名自称“获批”不构成客户授权。合成样例中的 true 只代表自造事实。
4. 仅在映射确认后运行；所有路径为指定 root 内相对路径，输出必须是独立、全新目录，父目录已存在：

```bash
python3 /absolute/path/electronic-quote-audit/scripts/audit_quotes.py --root /absolute/path/local-quotes --config config.json --output review-01
```

5. 先看 `review.html`，再按 `clarifications.csv` 的每一项联系用户确认。引用原文件、表名、行/单元格和原值；不要把缺失和明确零混为一谈。更新配置版本后用新目录重跑，不覆盖原文件或旧结果。
6. 交付四个文件及固定评估时点、配置/规则版本，说明可算候选、可比候选和未决项。业务待澄清返回0；文件/配置/安全错误返回2并不生成比价结果。

## 计算与审阅边界

- 默认制造商与完整料号逐字严格相等，不去空格、不改大小写、不丢前导零。数值型 XLSX 身份不可猜回原值。客户内部号及替代只支持本配置、指定需求行、版本化且有授权记录的显式映射，原始身份与非 exact 映射链同时保留。
- MOQ 和订购倍数按明确报价购买单位向上取整；标准包装量只展示。Minimum Release 按所选报价档位校验，不当MOQ；阶梯门槛不驱动主动超购。重叠或重复阶梯全部待确认。
- Decimal 保留精确小数，不设默认两位金额。无限循环小数阻止对应金额，需用户明确可用规则；v0不应用自定舍入或税费转换。不同货币仅用客户提供的定向汇率；无授权不查询汇率。
- 货值与含运可比金额分列。税口径必须与目标一致（`excluded`或`included_13pct`等明确结构）；非零运费另须明确同税口径。多需求行的订单级运费不自动分摊，不汇总为订单总价，不作最低价排名或采购推荐。
- `as_of` 必须显式含时区；纯日期有效期另须 `date_timezone`，自然日截至含当日末。过期货值保留且不当前可比。库存、分销商到库/BAL/ARO、厂商交期、客户收货承诺分栏，均不作按期履约保证。
- 不解析自由备注猜条件。需求文件和每家供应商文件的非空备注均须分别由用户复核，未经复核保留货值、阻止完整可比金额；供应商的确认不能代替需求文件确认。已结构化的 `valid_until_alternate` / `customer_receipt_alternate` 可记录双方冲突，不能用 reviewed 标记消除已结构化矛盾。

## 安全与结果

仅读 root 内普通文件；拒绝符号链接、路径遍历、宏、外部关系、DTD/实体、非UTF-8 XML、异常ZIP及超限文件。任何工作表含公式都会拒绝**整本 XLSX 和本次运行**，即使有缓存，也不会继续生成其他供应商比价；请用户另存获准的原始纯值文件后重试。单元格、供应商 label、原文件名元数据不作读写路径。无 eval、shell、宏/公式或网络执行。

`comparison.csv` 展示数量、原币货值、可比金额与展开；`clarifications.csv` 逐项定位；`audit.json` 保留未转义原值、文件 SHA-256、单元格类型、配置 JSON pointer、匹配/排除/冲突和计算证据；`review.html` 为无脚本、无外部资源的离线静态阅览文件。CSV 对危险公式前缀及控制字符作安全文本导出，转义文字不回流参与匹配。

Excel 默认双击打开 CSV 时，纯数字料号及其前导零仍可能被自动转成数值。请使用导入功能，将制造商和料号列明确指定为文本；原始文本可在 `audit.json` / `review.html` 核对。v0不保证默认双击CSV后的显示格式。

## 试跑自造样例

从技能目录运行（输出目录须尚不存在）：

```bash
python3 scripts/audit_quotes.py --root "$PWD/assets/synthetic" --config config.json --output sample-review-01
```

样例有异构表头的需求CSV、供应商CSV及供应商XLSX；包含MOQ/倍数跨阶梯、per100/每卷、缺运费与收货待确认。四份结果会生成在 `assets/synthetic/sample-review-01/`。真实报价应使用用户另行指定的本地工作目录。
