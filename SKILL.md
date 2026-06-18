---
name: turbolink-data
description: |
  Turbolink 活动数据拉取与分析。从 Turbolink 后台拉取营销活动数据并导出 Excel。
  当用户提到以下任意场景时触发此 skill：
  - "拉取数据"、"跑数据"、"导出数据"、"活动数据"、"数据分析"
  - "Turbolink"、"活动报告"、"导出 Excel"
  - "金币回收"、"金币数据"、"活动参与率"、"留存率"
  - 任何涉及 Turbolink 平台活动指标的查询
  即使用户没有明确说"Turbolink"，只要上下文涉及活动数据拉取，也应触发。
---

# Turbolink 活动数据拉取

从 Turbolink 后台拉取营销活动数据，生成格式化的 Excel 报告。

## 工作流程

### 1. 获取凭证

需要用户提供三个信息：
- **Token**：登录 Turbolink 后台后，从浏览器 DevTools → Network → 任意请求的 `Authorization` 请求头中复制
- **项目 ID**：Turbolink 后台 URL 中的 `pjid` 参数值
- **公司名称**：显示在 Excel 公司名列中的名称（如"CSSBuy"、"麦芽"）

如果用户没有主动提供，询问他们。

### 2. 询问筛选条件

主动询问用户是否需要筛选：

- **活动类型**：想看哪些活动？（如"赛事竞猜"、"金币大派送"、"骰子大富翁"等）
  - 支持中文名称和 fission_mark（如 `vs`、`coin`）
  - 可指定多个类型
  - 不指定则拉取所有类型
- **活动开始日期**：只看某个日期之后开始的活动？（如"5月份之后的"→ `2025-05-01`）
  - 不指定则不限制

如果用户明确说"全部"或没有提到筛选条件，直接跳过此步。

### 3. 检查 Token 有效性

Token 是 JWT 格式，可以通过解析 payload 中的 `exp` 字段判断是否过期。

```python
import base64, json
from datetime import datetime

token = "Bearer xxx"
payload = token.replace("Bearer ", "").split(".")[1]
payload += "=" * (4 - len(payload) % 4)
data = json.loads(base64.urlsafe_b64decode(payload))
exp = datetime.fromtimestamp(data["exp"])
print(f"过期时间: {exp}")
```

- 如果已过期 → 告知用户并引导获取新 Token
- 如果即将过期（< 24h） → 提醒用户

获取新 Token 的方法：
1. 打开浏览器，登录 `dashboard.turbolink.cc`
2. 打开 DevTools（F12）→ Network 标签
3. 刷新页面，点击任意一个请求
4. 在 Request Headers 中找到 `Authorization` 值，复制完整内容（以 `Bearer ` 开头）

### 4. 运行数据拉取

使用 skill 自带的 `scripts/fetch.py` 脚本。Token 和项目 ID 通过 CLI 参数传入，无需修改脚本文件。

```bash
python3 <skill-path>/scripts/fetch.py \
  --token "Bearer xxx" \
  --project-id "abc123" \
  -t 赛事竞猜 \
  -s 2025-05-01
```

**必填参数：**
| 参数 | 说明 |
|------|------|
| `--token` | Bearer Token（从浏览器 DevTools 获取） |
| `--project-id` | 项目 ID（Turbolink 后台 URL 中的 pjid） |
| `--company` | 公司名称（显示在 Excel 公司名列中） |

**可选参数：**
| 参数 | 默认值 | 说明 |
|------|--------|------|
| `-t, --activity-type` | 全部 | 活动类型（中文名如"赛事竞猜"或 fission_mark 如"vs"），可多次指定 |
| `-s, --start-date` | 无 | 筛选活动开始日期 >= 此值（如 2025-05-01） |
| `--uv-threshold` | 15 | 仅拉取 UV 大于此值的活动 |
| `--coin-cutoff` | 无 | 金币回收统计截止日期（如 2026-06-11），不填则不计算金币回收 |
| `--output` | 活动数据分析.xlsx | 输出文件名 |
| `--search-start` | 2025/01/01 00:00 | 活动搜索起始时间 |
| `--search-end` | 2099/12/31 23:59 | 活动搜索结束时间 |

**Sheet 命名规则：** 如果指定了 `-s`，Sheet 名称会变为 `活动数据_开始>=YYYY-MM-DD`。

**金币大派送活动：** 如果用户提到"金币回收"或活动中有 coin 类型，必须传入 `--coin-cutoff` 参数。默认不计算。

### 5. 输出

脚本会在当前目录生成 Excel 文件。告知用户：
- 文件路径
- 包含的活动数量
- 每个活动的 UV 概览

如果用户需要了解某个指标的计算方式，读取 [references/metrics.md](references/metrics.md) 查看公式。

## 指标查询

当用户问"这个率怎么算的"、"回收金币是什么意思"等问题时，读取 `references/metrics.md` 并返回对应的公式说明。

## 常见问题

| 问题 | 解决方案 |
|------|----------|
| Token 过期 (401) | 引导用户从浏览器获取新 Token |
| 返回 0 个活动 | 检查 UV 阈值是否过高，尝试 `--uv-threshold 0` |
| 依赖缺失 | `pip install requests openpyxl` |
