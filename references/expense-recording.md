# Expense Recording (费用记录)

This user records daily project expenses in a Feishu Base named **费用明细**.

## Base Info

- Token: `K8UfbguBhamlIWsoGCgctQXFnwc`
- URL: `https://my.feishu.cn/base/K8UfbguBhamlIWsoGCgctQXFnwc`

## Table: 费用记录

| Field | Type | Notes |
|:------|:-----|:------|
| 服务/平台 | text | What the expense was for (e.g., "刻公章") |
| 日期 | text | Format: `M.D` (e.g., "6.25" for June 25) |
| 金额 | number | Decimal with 2 precision |
| 支付方式 | select | Options: 蓝胖子, 欧易, 招商信用卡, 支付宝 |
| 货币 | select | Options: ¥, $ |
| 备注 | text | Optional context |
| 续费截止日 | text | Only for subscription-type expenses |
| 月份 | formula | Auto-calculated from 日期 |

## Dashboard

- **费用分析看板** (blk0Kp5dqpjtLPNv) — visual analytics on expenses

## Write Command

```bash
lark-cli base +record-upsert \
  --base-token K8UfbguBhamlIWsoGCgctQXFnwc \
  --table-id tble0xOe54aUAYRZ \
  --as user \
  --json '{"服务/平台":"<item>","日期":"<M.D>","金额":<number>,"支付方式":"<option>","货币":"¥"}' \
  --format json
```

## Examples

Record a one-time expense:
```
服务/平台: "刻公章"
日期: "6.25"
金额: 120
支付方式: "蓝胖子"
货币: "¥"
```

Record a subscription expense:
```
服务/平台: "MiniMax"
日期: "5.2"
金额: 70
支付方式: "蓝胖子"
货币: "¥"
续费截止日: "每月2日"
```

## Triggers

Use this reference when the user says:
- "记账" / "记一笔" / "记录费用"
- Mentions spending money on project-related items
- References an expense they want tracked

## Interaction Pattern

1. If payment method is not specified, ask which one (options: 蓝胖子, 欧易, 招商信用卡, 支付宝)
2. Default currency is ¥ unless the expense is in USD
3. Date format is `M.D` (month + period + day, no leading zeros needed)
