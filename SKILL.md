---
name: daily-review-skill
version: 1.0.0
description: "每日复盘与时间管理全流程：时间记录→日复盘（Get笔记+飞书表）→明日计划写入任务看板→费用记账→周复盘自动维护。跨境电商卖家时间管理一体化工具。"
tags: [复盘, 日复盘, 时间记录, 记账, 周复盘, 任务管理]
---

# Daily Review Skill — 每日复盘与时间管理

## 何时使用

| 场景 | 触发词 |
|:-----|:-------|
| 记录时间段做了什么 | "记录时间" / "记录" / "时间记录" |
| 做日复盘总结 | "日复盘" / "复盘" |
| 标记任务完成 | "已完成" / "完成" |
| 记账/记开支 | "记账" / "记一笔" / 提及花钱 |
| 管理任务看板 | "任务更新" / "修改任务" |

## 资源配置

### 飞书时间记录表
- **URL**: `https://my.feishu.cn/sheets/CRtSsJXa7hz7X9t7LRCcpDZcnFh`
- **Token**: `CRtSsJXa7hz7X9t7LRCcpDZcnFh`
- **时间记录**表: 日常时段记录 (A-M 共13列)
- **每日复盘**表: 日复盘写入 (A-L 共12列)
- **周复盘**表: 周汇总
- **可视化仪表盘** / **日汇总** / **设置**: 配套看板

### Get笔记
- **每日复盘知识库**: `7JbDMeNY`
- **待办事项知识库**: `rYMN5120`

### 任务看板 (飞书多维表格)
- **Base Token**: `CYlobxEfKawmfLsTAfZcKFzMnLZ`
- **Table ID**: `tbl4DAONe5dWzY1h`
- 字段: 任务名称(text)、状态(select: 未开始/进行中/已完成)、任务类别(select: 学习/生活/工作)、任务详情(text)、任务标签(select: 🔥重要紧急/📌重要不紧急/🤣不重要紧急/👇不重要不紧急)

### 费用记录 (飞书多维表格)
- **Base Token**: `K8UfbguBhamlIWsoGCgctQXFnwc`
- **Table ID**: `tble0xOe54aUAYRZ`
- 字段: 服务/平台(text)、日期(text 格式M.D)、金额(number)、支付方式(select: 蓝胖子/欧易/招商信用卡/支付宝)、货币(select: ¥/$)、备注(text)

### 脚本路径
- `scripts/record_lark_time.py` — 时间记录写入
- `scripts/record_lark_review.py` — 日复盘写入飞书表
- `scripts/create_workbook.py` — 创建工作簿模板

---

## 工作流

### 工作流 1: 时间记录

当用户发送时间记录时（如 `记录，08:00-13:00，出门办事`），执行：

```bash
python3 scripts/record_lark_time.py \
  --date 2026-06-25 \
  --start 08:00 \
  --end 13:00 \
  --task "出门办事" \
  --output "完成了什么" \
  --main 是 \
  --focus 3 \
  --energy 3
```

**时间记录列映射：**
| 列 | 内容 | 说明 |
|:--:|:-----|:-----|
| A | 日期 | `2026-06-25` |
| B | 星期 | 自动 |
| C | 开始 | `08:00` |
| D | 结束 | `13:00` |
| E | 时长 | 小数小时，自动 |
| F | 类别 | 13个下拉选项（详见下方「电商类别」） |
| G | 工作类型 | 自动派生 |
| H | 是否主线 | 下拉: 是/否 |
| I | 任务 | 文本 |
| J | 产出 | 文本 |
| K | 精力/能量 | 下拉: 1-5 |
| L | 专注 | 下拉: 1-5 |
| M | 干扰源 | 下拉: 手机/短视频, 平台消息, 客户消息, 临时想法, 工具问题, 家务/生活, 疲劳拖延, 其他 |

#### ⚠️ 自动分类修正（重要）

脚本自动分类不可靠，**每次写入后必须检查返回的 category**，按以下规则修正 F(类别) 和 G(工作类型)：

| 用户实际做了 | 脚本可能归类为 | 应修正为 |
|:-------------|:---------------|:---------|
| 午休/吃饭/洗漱/上厕所 | 行政/工具/杂事 | 休息/生活 / 休息 |
| 刷手机/看电视 | 内容素材 | 干扰/刷信息 / 干扰 |
| 银行/公户/刻公章 | 财务/记账 | 行政/工具/杂事 / 杂事 |
| 买药/去医院 | 财务/记账 | 休息/生活 / 休息 |
| 工具开发含"详情页"关键词 | 上架/Listing优化 | 内容素材 / 推进型 |
| 普通工具开发 | 行政/工具/杂事 | 内容素材 / 推进型 |
| 陪拍/协助女朋友 | 内容素材 | 休息/生活 / 休息 |

修正命令：
```bash
lark-cli sheets +cells-set --spreadsheet-token CRtSsJXa7hz7X9t7LRCcpDZcnFh --sheet-name "时间记录" --as user --range "'时间记录'!F{N}:G{N}" --cells '[[{"value":"休息/生活"},{"value":"休息"}]]'
```

#### 补充/修改已有记录

当用户说"补充"时 → 不创建新行，而是更新已有行：
1. 按日期+开始时间找到对应行
2. 更新 D(结束时间)、E(重算时长)、J(产出) 等字段
3. 用单格 `+cells-set` 逐列写入

#### 干扰源映射规则

| 用户说 | 下拉选 |
|:-------|:------:|
| 看电视/刷手机/刷短视频/刷抖音/刷视频 | 手机/短视频 |
| 平台消息/群消息/通知/消息打扰 | 平台消息 |
| 回客户/售后/回复消息 | 客户消息 |
| 突然想到/临时做的/女朋友说 | 临时想法 |
| 工具出问题/报错/调试卡住 | 工具问题 |
| 做家务/做饭/家里有事/上厕所 | 家务/生活 |
| 累了/不想动/拖延/睡过头 | 疲劳拖延 |
| 以上都不匹配 | 其他 |

#### 时间输入格式

用户常用格式（逗号分隔，各占一行）：
```
记录，17:30-18:20，用 codex 设计电商带货详情页设计工具
产出:已经分别做出第一版...
是主线
专注/精力:4/4
```

---

### 工作流 2: 日复盘

用户发来复盘内容后，**必须立即**按以下顺序执行，不要等待用户提醒。

#### Step 1 — 保存到 Get笔记

```bash
# 构造复盘内容
cat > /tmp/daily-review.md <<'ENDOFFILE'
## 📅 YYYY-MM-DD 日复盘

### 🎯 今日主线
...

### 📊 完成度
65%

### 🏆 今日产出
1. ...
2. ...

### ⏰ 最大时间黑洞
...

### 📋 明日主线
1. ...
2. ...

### 🧠 专注/精力/满意度
X/Y/Z

### 💬 一句话总结
...
ENDOFFILE

# 保存笔记
getnote save --title "YYYY-MM-DD 日复盘" "$(cat /tmp/daily-review.md)"

# 获取 note_id 并加入知识库
NOTE_ID=$(getnote notes --limit 3 -o json | python3 -c "import sys,json; notes=json.load(sys.stdin)['data']['notes']; print(notes[0]['note_id'])")
getnote kb add 7JbDMeNY "$NOTE_ID"
```

**注意：** `getnote save` 只输出 `✓ Note saved.`，不返回 note_id。必须用 `getnote notes` 查最新笔记的 note_id 后再 `kb add`。

#### Step 2 — 写入飞书每日复盘表

```bash
python3 scripts/record_lark_review.py \
  --date 2026-06-25 \
  --main-task "今日主线" \
  --completion 75% \
  --outputs "产出1；产出2；产出3" \
  --time-sink "最大时间黑洞" \
  --distraction "其他" \
  --tomorrow-main "1. 明日任务1\n2. 明日任务2" \
  --tomorrow-first-step "第一步做什么" \
  --focus 4 \
  --energy 2 \
  --satisfaction 3 \
  --conclusion "一句话总结"
```

**完成度取值规则：** 脚本只接受 0%、25%、50%、75%、100%。用户给的 60%→取 50%，85%→取 75%。

**干扰源必须是下拉选项：** 只能从以下8个选：手机/短视频、平台消息、客户消息、临时想法、工具问题、家务/生活、疲劳拖延、其他。按「干扰源映射规则」表格映射。

**用户实际复盘格式：**
- 用 `两大产出` / `三大产出`（不用 `三个产出`）
- 用 `一句话总结`（不用 `一句话结论`）
- 完成度写 `65%`（带百分号）
- 专注/精力/满意度一行：`4/2/3`

#### 每日复盘列映射（fallback 逐格写入）

| 列 | 内容 | 说明 |
|:--:|:-----|:-----|
| A | 日期 | `2026-06-25` |
| B | 今日主线任务 | 文本 |
| C | 主线完成度 | 下拉: 0%/25%/50%/75%/100% |
| D | 今日3个产出 | 文本 |
| E | 最大时间黑洞 | 文本 |
| F | 最大干扰源 | 下拉（同时间记录M列） |
| G | 明日主线任务 | 文本 |
| H | 明日第一步动作 | 文本 |
| I | 专注评分 | 下拉: 1-5 |
| J | 能量评分 | 下拉: 1-5 |
| K | 满意度 | 下拉: 1-5 |
| L | 一句话结论 | 文本 |

脚本超时时的 fallback：
```bash
lark-cli sheets +cells-set --spreadsheet-token CRtSsJXa7hz7X9t7LRCcpDZcnFh --sheet-name "每日复盘" --as user --range "'每日复盘'!B{N}" --cells '[[{"value":"..."}]]'
```

#### Step 3 — 解析明日主线写入任务看板

从复盘内容中提取「明日主线」列表，逐条写入：

```bash
lark-cli base +record-upsert --base-token CYlobxEfKawmfLsTAfZcKFzMnLZ \
  --table-id tbl4DAONe5dWzY1h --as user \
  --json '{"任务名称":"任务描述","状态":"未开始","任务类别":"工作","任务详情":"...","任务标签":"📌 重要不紧急"}'
```

优先级映射：

| 用户写法 | 标签 |
|:---------|:----:|
| 🔥重要紧急 | `🔥 重要紧急` |
| 📌重要不紧急 | `📌 重要不紧急` |
| 🤣不重要紧急 | `🤣 不重要紧急` |
| 👇不重要不紧急 | `👇 不重要不紧急` |

---

### 工作流 3: 标记任务完成

当用户说「完成」某个任务时：

```bash
# 搜索任务
lark-cli base +record-search --base-token CYlobxEfKawmfLsTAfZcKFzMnLZ \
  --table-id tbl4DAONe5dWzY1h --as user \
  --keyword "<关键词>" --search-field "任务名称" --limit 5

# 标记完成
lark-cli base +record-upsert --base-token CYlobxEfKawmfLsTAfZcKFzMnLZ \
  --table-id tbl4DAONe5dWzY1h --as user \
  --record-id <record_id> --json '{"状态":"已完成"}'
```

当用户说「更新」状态时（如开始做了）：
```bash
lark-cli base +record-upsert --base-token CYlobxEfKawmfLsTAfZcKFzMnLZ \
  --table-id tbl4DAONe5dWzY1h --as user \
  --record-id <record_id> --json '{"状态":"进行中","任务详情":"当前进度说明"}'
```

---

### 工作流 4: 记账 / 费用记录

当用户说「记账」「记一笔」或提到花钱时：

```bash
lark-cli base +record-upsert \
  --base-token K8UfbguBhamlIWsoGCgctQXFnwc \
  --table-id tble0xOe54aUAYRZ \
  --as user \
  --json '{"服务/平台":"刻公章","日期":"6.25","金额":120,"支付方式":"蓝胖子","货币":"¥"}' \
  --format json
```

- 日期格式: `M.D`（如 `6.25`）
- 未指定支付方式时需询问（选项: 蓝胖子、欧易、招商信用卡、支付宝）
- 默认货币是 ¥

费用记录字段：

| 字段 | 类型 | 说明 |
|:-----|:-----|:-----|
| 服务/平台 | text | 费用项目名称 |
| 日期 | text | `M.D` 格式 |
| 金额 | number | 精确到分 |
| 支付方式 | select | 蓝胖子/欧易/招商信用卡/支付宝 |
| 货币 | select | ¥/$ |
| 备注 | text | 可选 |
| 续费截止日 | text | 订阅类费用才填 |
| 月份 | formula | 自动 |

---

### 工作流 5: 周复盘（定时任务已配，也可手动执行）

每周日 12:00 自动执行。手动执行时：

1. 确定本周一 ~ 本周日范围
2. 从「时间记录」表读取本周所有行，按工作类型汇总时长：
   - 推进型小时 = G列="推进型" 的 E列之和
   - 维护型小时 = G列="维护型" 的 E列之和
   - 干扰小时 = G列="干扰" 的 E列之和
3. 从「每日复盘」表读取本周复盘记录：
   - 主线完成均值 = 平均 C列完成度
   - 满意度均值 = 平均 K列满意度
4. 写入「周复盘」表对应行（按 A列周开始日期匹配）

周复盘列映射：A=周开始, B=周结束, C=本周战役, D=本周胜利成果, E=下周要改进, F=推进型小时, G=维护型小时, H=干扰小时, I=主线完成均值, J=满意度均值

写入命令：
```bash
lark-cli sheets +cells-set --spreadsheet-token CRtSsJXa7hz7X9t7LRCcpDZcnFh --sheet-name "周复盘" --as user --range "'周复盘'!F{N}" --cells '[[{"value":12.5}]]'
```

---

## 电商类别

| 类别 | 工作类型 |
|:-----|:--------:|
| 选品/市场调研 | 推进型 |
| 上架/Listing优化 | 推进型 |
| 广告/数据分析 | 推进型 |
| 内容素材 | 推进型 |
| 供应链/采购 | 推进型 |
| 客服/售后 | 维护型 |
| 订单/物流 | 维护型 |
| 财务/记账 | 维护型 |
| 学习/资料整理 | 学习型 |
| 沟通/会议 | 维护型 |
| 行政/工具/杂事 | 杂事 |
| 休息/生活 | 休息 |
| 干扰/刷信息 | 干扰 |

---

## 陷阱 & 注意事项

### 通用
- ⚠️ 所有操作使用 `--as user` 身份，不要用 bot
- ⚠️ 日复盘必须**立即自动保存**到 Get笔记，不等用户提醒
- ⚠️ 用户说「完成」时只改状态为「已完成」，不清除/删除记录
- ⚠️ 未完成的任务保留不动，不清空
- ⚠️ 合并任务时先更新保留任务，再删除多余任务

### 时间记录
- ⚠️ 每次 `record_lark_time.py` 调用后必须检查并修正自动分类
- ⚠️ 干扰源必须从8个下拉选项中选，不能写自由文本
- ⚠️ 连续时间段逐条写入，不要合并为一条
- ⚠️ 时间倒置（开始>结束）通常是用户笔误，按上下文推断修正
- ⚠️ 多格 `+cells-set` 长文本可能超时，改用单格逐列写入

### Get笔记
- ⚠️ `getnote save` 不返回 note_id，必须 `getnote notes` 查后再 `kb add`
- ⚠️ 两步骤缺一不可：save 只存笔记，kb add 才放入知识库

### 周复盘
- 周复盘 cron job 已配置（bc1a0f6be6a6，每周日 12:00），如遇数据为空则对应列留空

### 任务看板
- 状态只接受：未开始 / 进行中 / 已完成
- 任务标签含 emoji 和空格，需精确匹配
