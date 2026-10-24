# 输入配置与核对规则 v1

命令 `python3 scripts/audit_quotes.py --root ABSOLUTE_DIRECTORY --config config.json --output new-review`。只读 root 中用户指定的文件；config/output 均为 root 下相对路径。输出目录不得已存在，父目录须存在。文件名、单元格、original_name 均不作为其他路径。

本程序面向已确认的结构化字段，不用模型猜表头/备注，不自动拆解“100件/卷”等自由文本。若原表把价格/阶梯合在一个描述单元格，先由用户确认拆解方式，并保留原文件与客户确认依据，再使用单独字段或明确配置常量。未知业务字段保持空白，不能填写猜测的0/1。

## 顶层配置

参考可直接运行的 `assets/synthetic/config.json`，完整顶层形式如下：

```json
{
  "schema_version": "electronic-quote-audit-config-v1",
  "config_version": "customer-batch-v1",
  "mapping_confirmed": true,
  "as_of": "2026-10-24T00:00:00Z",
  "date_timezone": "UTC",
  "rounding": "exact",
  "demand": {"file": "procurement.csv", "header_row": 1,
    "columns": {"manufacturer": "制造商", "mpn": "完整料号", "quantity": "需求数", "quantity_unit": "需求单位"}},
  "suppliers": [
    {"label": "供应商甲", "file": "a.csv", "columns": {"manufacturer": "Maker", "mpn": "Part", "unit_price": "Price"},
     "constants": {"quantity_unit": "pcs", "price_unit": "pcs", "price_quantity": "1", "moq": "0", "order_multiple": "1", "minimum_release": "0", "tier_min": "0", "tier_max": "unbounded", "tier_basis": "order_quantity", "currency": "CNY", "tax_basis": "included_13pct", "freight": "0", "freight_currency": "CNY", "freight_scope": "line", "valid_until": "2026-11-30T23:59:59Z"}},
    {"label": "供应商乙", "file": "b.xlsx", "sheet": "Quote", "columns": {"manufacturer": "制造商", "mpn": "料号", "unit_price": "单价"}}
  ],
  "comparison": {"currency": "CNY", "tax_basis": "included_13pct", "freight_basis": "goods_plus_freight"}
}
```

上例乙的业务条件刻意未补齐，会输出澄清。甲的constants仅演示**客户已明确确认的事实**，不是导入默认值。

- `mapping_confirmed=true` 只在用户确认每份文件、列映射、单位语义及比较口径后设置；不是供应商表格的自我声明。
- `config_version` 非空，`as_of` 必须含时间和时区偏移。程序不读取运行时钟决定有效性。`date_timezone` 可省略，但源有效期为纯日期时将待确认；含时区时间戳无需它。
- `suppliers` 2–5份，label 唯一且1–80字符，独立配置，不由文件名推断。
- `file` 为UTF-8 CSV或纯值XLSX。XLSX须有确切 `sheet`。`header_row` 默认1，只是文件结构规则；CSV行号是逻辑记录行号，带换行的引号字段仍是一行记录。
- `columns` 是“标准字段名 → 原表精确表头”，不去空格；表头重复/非文本直接报错。`constants` 是明确字段常量，全部写字符串；同一字段不能同时映射列与常量。字段值为空或未映射保留缺失，明确`"0"`不视为空。
- 可选 `original_name` 仅保留逻辑原文件名；即使值为 `../../outside.csv` 也不读取或写入该路径。
- 可选每文件 `notes_reviewed=true` 仅在客户确认该文件全部备注已审阅且无未决条件后设置。需求配置`demand`与各供应商配置分别复核，供应商标记不能授权需求文件；字符串`"true"`或单元格自述均不算确认。源备注非空且未确认，会保留货值、阻止完整可比；程序不自动理解自由文本。结构化冲突始终保留。需求备注的澄清证据保留原文件单元格及`/demand/notes_reviewed`配置hash/pointer。

## 字段语义

| 字段 | 确切含义/允许值 |
|---|---|
| manufacturer / mpn | 原始制造商/完整料号，必须文本；逐字匹配，大小写、空格及前导零均有意义 |
| internal_id | 原始客户内部号，可选；不是外部料号，不能自行推断 |
| quantity / quantity_unit | 需求数量及需求单位；数量必须正数。需求行的quantity_unit不默认传播为报价单位 |
| required_receipt | 客户要求收货时间，仅展示及提示缺承诺，不自动判断履约 |
| quantity_unit（报价） | 实际MOQ、倍数、阶梯及Minimum Release所使用的购买单位 |
| moq / order_multiple | 报价购买单位的MOQ（非负）及订购倍数（正数）；明确无MOQ写0、任意整件可买写1须有来源 |
| package_quantity | 厂商标准包装量，原值展示；不参与订购量或单位换算 |
| minimum_release | 该档报价价格的最低release，非负；选中阶梯后校验订购量，不提高订购量。不同档可不同 |
| tier_min / tier_max | 闭区间数量下/上界，非负；上界无限必须明确`unbounded`，空白不是无限 |
| tier_basis | `order_quantity`（最终订购量）或`demand_quantity`（需求换算到报价单位）；不支持其他猜测 |
| unit_price / price_quantity / price_unit | 单价、每价分母数量（正数）、分母单位。价格允许明确0，不能负数 |
| currency | 原始货值货币标识，不默认CNY |
| tax_basis | 货值完整已确认税口径：仅`excluded`或`included_<税率>pct`，税率十进制0–100，如`included_13pct`。未知/待定不放行；口径标签精确相等仍需真实客户确认 |
| freight / freight_currency | 运费非负金额及币种，空白不视为免运，明确免运须0及币种 |
| freight_tax_basis | 非零运费必需，同上tax_basis格式且与目标一致；零运费不需税转换。货值税basis不自动覆盖运费 |
| freight_scope | `line`：明确归属于这一需求行；`included`：已包含且freight明确0；`order`：本报价订单一次费用，仅一需求行时可加，多行时不分摊 |
| valid_from / valid_until | 起效/截至时间；valid_until缺失待确认，valid_from可选。时间戳必须含时区，纯日期用date_timezone当日初/末，端点包含 |
| valid_until_alternate | 其他明确来源的截至时间；与主值矛盾则当前有效性未知。不可用notes_reviewed覆盖 |
| stock | 原库存陈述，仅保留，不替代交期承诺 |
| distributor_receipt | 分销商到库、BAL/ARO等原文，非客户收货时间 |
| manufacturer_lead_time | 厂商标准交期原文，非客户收货承诺 |
| customer_receipt / customer_receipt_alternate | 客户收货承诺及另一来源；不同时进入澄清，不用库存消解 |
| notes | 原备注；不执行、不当指令、不当客户授权，不解析猜条件 |

同一供应商、同一原始制造商+料号的多行视为阶梯。MOQ/倍数/购买单位/币种/税费/有效期等共有条件跨行不同或部分缺失时保守待确认。所有阶梯的区间必须无重叠（即使当前数量不落在重叠处），不猜修订优先级。没有适用档不主动加购。仅所选档的Minimum Release决定价格是否满足。

数字使用十进制字符串，允许普通正负号和小数；拒绝货币符号、千位分隔、科学记数、非有限数。整数及小数部分各最多24位。订购量向上取整以精确有理数算出，金额使用Decimal且捕获不精确运算；无限循环小数不截断、不产生金额。v0只支持`rounding=exact`，其他舍入策略明确报不支持。

## 单位与汇率授权

仅支持单条明确方向，不自动取倒数、连乘多跳或猜别名。单位标签严格相等时无需转换。不同标签即使人看来同义也需授权。

```json
{
  "unit_conversions": [
    {"from": "reel", "to": "pcs", "factor": "2000", "suppliers": ["供应商乙"],
     "customer_confirmed": true, "version": "unit-v1", "source": "客户确认记录：1卷2000件"}
  ],
  "fx": [
    {"from": "USD", "to": "CNY", "rate": "7.1234", "direction": "multiply",
     "scope": "this_config", "suppliers": ["供应商甲"], "customer_confirmed": true,
     "version": "fx-v1", "source": "客户授权记录：本批次1USD=7.1234CNY"}
  ]
}
```

单位`factor`表示 **1 from = factor to**，必须正数。仅用于该供应商；购买整卷的MOQ/倍数仍要独立明确。报价购买单位→需求单位、计价单位→报价购买单位分别寻找直接关系，不由包装量推断。

汇率`rate`表示目标币/原币，用乘法；必须正数、本配置范围、明确supplier/来源/版本/客户确认。换汇分别作用于货值和运费；不自动税费转换。比较仅在目标currency/tax_basis/freight_basis明确、源同税口径、运费及当前有效性完整时形成可比金额。

## 极小客户映射

仅支持本配置的确切需求逻辑行+确切internal_id，无全局别名、模糊匹配、自动认可厂商：

```json
{
  "customer_mappings": {
    "version": "customer-map-v2", "authorization": "AUTH-001",
    "customer_confirmed": true, "scope": "this_config",
    "entries": [{"demand_row": 2, "internal_id": "INT-7", "manufacturer": "ACME", "mpn": "000123",
      "alternatives": [{"manufacturer": "BETA", "mpn": "R7"}]}]
  }
}
```

无客户确认/版本/授权记录/范围时不应用并澄清；同一需求行多份entry直接报错，先解决授权冲突。原始严格匹配标`exact`，内部号映射标`customer_mapping`，替代标`authorized_alternative`。不同原始身份永不宣称严格相等。映射链保留原始需求、目标、版本/授权和配置hash/pointer。

## 结果与安全限制

- `audit.json`：未转义原始每单元格、来源文件SHA-256、表/行/cell、类型/存储值、配置hash/JSON pointer、全部匹配/排除/冲突/计算。共享搜索范围使用指向`inputs`的pointer，避免重复嵌入全表。实现文件hash与规则版本同时保留。
- `comparison.csv`：按输入需求/配置供应商顺序，列出数量、所选阶梯、原币货值、税/运费、当前有效性和可比金额；没有可比金额为空，不补零。原始字段仍在JSON，不用导出转义值反向匹配。
- CSV默认在Excel双击打开时，纯数字料号及前导零仍可能被宿主自动转为数值；请使用Excel导入功能，将制造商/料号列指定为文本，并用`audit.json` / `review.html`核对原始文本。v0不保证默认双击CSV后的显示格式。
- `clarifications.csv`：逐项编号、业务原因、字段原值和来源。`blocking`说明该问题可能阻止相关价格/比较；非阻断交期提示也不可作履约结论。
- `review.html`：安全转义、无脚本/链接/外部资源，包含汇总、澄清及计算来源。CSV前置危险`= + - @`、控制字符及空白后的危险前缀以文本处理；内部原值不改。
- 不排名、不给最低价、不把单行货值称为未知落地总价、不跨采购行合计订单金额。
- 仅CSV/XLSX；CSV UTF-8（可BOM）。XLSX只支持UTF-8 XML、纯值、inline/shared文本和数值；数值身份/Excel日期序列不恢复。任何工作表公式、宏或外部关系整本及本次运行拒绝；不读取缓存当值。
- 每文件10MiB，ZIP最多1000项、展开40MiB、每项12MiB且压缩比不超过200，XML不接受DTD/实体；最多20000行、256列、200000单元格，每单元格10000字符，同一供应商严格身份最多100条阶梯。未做上限规模吞吐保证。检测ZIP路径、重复项及symlink，不解压到磁盘。
- 结构/安全/配置错误退出2且不生成结果目录；操作系统写盘中断可能留下部分目录，勿将其当成功。正常完成退出0仍可能全部待澄清，返回0不表示可采购。
