# Task Board Integration (简洁版-个人任务看板)

This user's task board is a Feishu Base (多维表格) at:
- **Base URL**: `https://my.feishu.cn/base/CYlobxEfKawmfLsTAfZcKFzMnLZ`
- **Base Token**: `CYlobxEfKawmfLsTAfZcKFzMnLZ`
- **Table**: `个人任务看板` (table_id: `tbl4DAONe5dWzY1h`)

## Fields

| Field ID | Name | Type | Options |
|----------|------|------|---------|
| `fldh5pASnV` | 任务名称 | text | — |
| `fld90K5s8z` | 状态 | select | 未开始 / 进行中 / 已完成 |
| `fldCbeuFIV` | 任务类别 | select | 学习 / 生活 / 工作 |
| `fldH808kko` | 任务详情 | text | — |
| `fldnZSmrZT` | 任务标签 | select | 🔥重要紧急 / 📌重要不紧急 / 🤣不重要紧急 / 👇不重要不紧急 |
| `fldrWtvROn` | 任务创建时间 | created_at | auto |
| `fldSpXZDI4` | 任务更新时间 | updated_at | auto |
| `fldgnQrCdY` | 任务耗时 | formula | auto |

## Workflow Rules

1. **New tasks**: When the user mentions a new task or deliverable, add it to the board via `+record-batch-create` or `+record-upsert`.
2. **Status changes**: When the user says something is "完成" (done), update the record's status to "已完成" via `+record-upsert --record-id <id> --json '{"状态":"已完成"}'`.
3. **Progress updates**: When the user says they're working on something, update status to "进行中".
4. **Priority classification**: Use the four-quadrant labels based on urgency/importance:
   - 🔥重要紧急 = urgent + important (overdue tasks, revenue-critical)
   - 📌重要不紧急 = not urgent but important (strategic work)
   - 🤣不重要紧急 = urgent but not important (maintenance, renewals)
   - 👇不重要不紧急 = neither (nice-to-have, blocked tasks)

## Commands

```bash
# Add a single task
lark-cli base +record-upsert --base-token CYlobxEfKawmfLsTAfZcKFzMnLZ \
  --table-id tbl4DAONe5dWzY1h --as user \
  --json '{"任务名称":"...","状态":"未开始","任务类别":"工作","任务详情":"...","任务标签":"🔥 重要紧急"}'

# Batch create tasks
lark-cli base +record-batch-create --base-token CYlobxEfKawmfLsTAfZcKFzMnLZ \
  --table-id tbl4DAONe5dWzY1h --as user \
  --json '{"fields":["任务名称","状态","任务类别","任务详情","任务标签"],"rows":[[...],[...]]}'

# Mark as complete
lark-cli base +record-upsert --base-token CYlobxEfKawmfLsTAfZcKFzMnLZ \
  --table-id tbl4DAONe5dWzY1h --as user --record-id <id> \
  --json '{"状态":"已完成"}'

# Mark as in progress
lark-cli base +record-upsert --base-token CYlobxEfKawmfLsTAfZcKFzMnLZ \
  --table-id tbl4DAONe5dWzY1h --as user --record-id <id> \
  --json '{"状态":"进行中"}'
```

## Pitfalls

- The `select` field for 状态 only accepts: `未开始`, `进行中`, `已完成`. Other values (like "暂停") will error.
- Always use `--as user` for this board (it's the user's personal board).
- Batch create supports max 200 rows per call.
- When updating, only pass the fields that need to change.
