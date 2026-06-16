# Turbolink 数据拉取 Skill — 安装与使用说明

## 这是什么

一个 AI 编程助手的 Skill，让你通过对话拉取 Turbolink 活动数据并导出 Excel。支持 Claude Code、Codex、Grok 等工具。

GitHub 仓库：https://github.com/fay924/turbolink-data

## 安装

### 方式一：命令行安装（推荐）

```bash
git clone https://github.com/fay924/turbolink-data.git ~/.claude/skills/turbolink-data
```

### 方式二：让 AI 帮你安装

把下面这段话复制粘贴给 AI：

> 帮我安装 turbolink-data 这个 Claude Code skill。
> 执行 git clone https://github.com/fay924/turbolink-data.git ~/.claude/skills/turbolink-data
> 验证：ls ~/.claude/skills/turbolink-data/ 应该看到 SKILL.md、scripts/、references/ 三项

### 方式三：手动下载

1. 从 GitHub 下载仓库（Code → Download ZIP）
2. 解压后放到 skills 目录：
   - Claude Code：`~/.claude/skills/turbolink-data/`
   - Codex：`~/.codex/skills/turbolink-data/`

### 更新

```bash
cd ~/.claude/skills/turbolink-data && git pull
```

## 前置条件

需要安装 Python 依赖：
```bash
pip install requests openpyxl
```

## 使用方法

### 拉取数据

直接对 AI 说：

- "拉取公司名数据"
- "帮我跑一下公司名的活动报告"
- "导出公司名的活动数据到 Excel"
- "拉取所有客户的活动数据"

AI 会：
1. 询问你的 Token 和项目 ID（首次使用时）
2. 检查 Token 是否过期
3. 运行脚本拉取数据
4. 告诉你 Excel 文件保存位置

### 查询指标公式

- "活动参与率怎么算的"
- "回收金币是什么意思"
- "留存率的公式是什么"

## 获取 Token

1. 打开浏览器，登录 `dashboard.turbolink.cc`
2. 按 F12 打开开发者工具
3. 切到 Network（网络）标签
4. 刷新页面，点击任意一个请求
5. 在 Request Headers 中找到 `Authorization`，复制完整值（以 `Bearer ` 开头）

Token 有效期约 7 天，过期后需要重新获取。

## 项目 ID

在 Turbolink 后台 URL 中找到 `pjid` 参数，例如：
```
https://dashboard.turbolink.cc/?pjid=d1qcfa01bc5m7ka0t450
```
其中 `d1qcfa01bc5m7ka0t450` 就是项目 ID。

## 也可以直接运行脚本

不通过 AI 工具，直接命令行运行：

```bash
python3 scripts/fetch.py --token "Bearer 你的token" --project-id "你的项目ID"
```

可选参数：

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--uv-threshold` | 15 | 仅拉取 UV 大于此值的活动 |
| `--coin-cutoff` | 无 | 金币回收截止日期，如 2026-06-11 |
| `--output` | 活动数据分析.xlsx | 输出文件名 |

示例（含金币回收）：
```bash
python3 scripts/fetch.py \
  --token "Bearer xxx" \
  --project-id "abc123" \
  --coin-cutoff 2026-06-11 \
  --output 公司名数据.xlsx
```

## Excel 输出说明

每个活动占多行（一行 per 任务），包含以下列：

| 列 | 内容 |
|----|------|
| A | 活动名称 |
| B | 活动时间 |
| C | UV |
| D | 点击活动主按钮人数 |
| E | 活动参与率 |
| F-I | 任务名称、完成人数、完成率、完成次数 |
| J-K | 次留、7日留 |
| L-M | 回收金币、金币回收比例（仅金币大派送活动） |

## 常见问题

**Q: Token 过期了怎么办？**
A: 按上面"获取 Token"的步骤重新复制一个。

**Q: 返回 0 个活动？**
A: UV 阈值默认 15，太小的活动会被过滤。可以用 `--uv-threshold 0` 查看全部。

**Q: 支持哪些 AI 工具？**
A: Claude Code、Codex、Grok 等支持 skill 的工具都可以。
