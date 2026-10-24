# electronic-quote-audit

**v0.1.0-beta.1 · 免费公开测试版 · MIT License**

原创电子元件采购报价核对 skill 和 Python 离线脚本。对一份采购清单与 2–5 家供应商的 CSV / 纯值 XLSX 报价，按制造商和完整料号逐字匹配，展开数量、MOQ、订购倍数、阶梯价和 Decimal 金额，保留原始来源并逐项列出待确认条件。

本版本免费提供，采用 [MIT 许可证](LICENSE)，允许使用、复制、修改、分发及商用，须保留版权和许可声明。软件按原样提供，不提供保证。本发布不含收款、托管或付费服务承诺。

## 从哪里开始

- 技能入口：[electronic-quote-audit/SKILL.md](electronic-quote-audit/SKILL.md)
- 配置与业务口径：[references/schema.md](electronic-quote-audit/references/schema.md)
- 离线脚本：[scripts/audit_quotes.py](electronic-quote-audit/scripts/audit_quotes.py)
- 完全自造的样例：[assets/synthetic](electronic-quote-audit/assets/synthetic)
- [版本说明](release-info.md)与[术语来源](electronic-quote-audit/references/sources.md)

运行需要 Python 3.12+ 标准库及系统 IANA 时区数据库；脚本不需要模型 API、联网或 pip 依赖。下载文件不会自动安装或运行技能。宿主加载方式须按所用工具的文档配置，本版本未验证技能自动发现。

## 手动试用合成样例

从仓库根目录进入技能目录，确认输出目录尚不存在，再手动运行：

```bash
cd electronic-quote-audit
python3 scripts/audit_quotes.py --root "$PWD/assets/synthetic" --config config.json --output sample-review-01
```

输出位于 `electronic-quote-audit/assets/synthetic/sample-review-01/`：`comparison.csv`、`clarifications.csv`、`audit.json` 和 `review.html`。先读 HTML，再核对逐项待确认与 JSON 原始证据。正常退出也可能有未决条件，不代表可以采购。

所有样例身份、需求、报价和授权标志均为虚构。样例中的 `true` 仅确认自造事实，不能照抄为真实客户授权。处理真实文件前须由用户确认文件范围、列映射、单位、税运费、汇率及评估时点，详见技能入口和 schema。

## 范围和已知限制

只支持 CSV 与纯值 XLSX；任何工作表包含公式时拒绝整本 XLSX 和本次运行，不能用公式缓存继续。宏或外链也会拒绝。不支持 PDF、OCR、ERP 集成或采购执行。

不理解自由备注、不推断替代关系、不自动查询汇率或转换税费；未知条件列为待确认。货值与含运可比金额分列，多需求行订单运费不自动分摊。不排名、不提供采购建议或下单，不保证省钱、最低价、报价真实性或交期。

CSV 在 Excel 默认打开时可能丢失料号前导零，应按文本导入，并以 JSON / HTML 原值核对。证据输出可能较大；未作上限规模吞吐保证。

## 验证范围

发布前已完成 25 项工程检查；独立验证覆盖 38 个案例、60 次 CLI 调用、114 条断言；另有两项显式 skill 使用验证通过。这些是已报告的发布前验证，本公开包不包含内部评审日志或恶意测试附件。

尚未验证宿主自动发现、真实 Microsoft Excel 宿主打开、模型对照效果、真实供应商全面适配、市场付费价值及生产适用性。这是免费测试版，不是成熟商业产品。反馈请仅提供最小脱敏复现，不要上传采购原件、商业秘密或凭据。

## 发行与许可

Copyright (c) 2026 8npyvz5bd8-lang。MIT 适用于本仓库原创技能、脚本、文档和合成样例。未打包第三方技能、ERP 源码或运行库；外部链接不改变第三方材料的许可。

`candidate-manifest.json` 记录获准公开候选包中各文件的校验值及候选阶段状态；其 `published: false` 和待定许可字段描述原始候选时点。当前公开发行的许可与状态以本 README 和 LICENSE 为准。运行技能、核心源码、参考和样例保持候选原字节。
