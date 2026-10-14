# 多 harness 安装与使用

本 skill 不依赖任何 DSH 专有接口：**纯 Python 3 + openpyxl**，对外只有三件事——读 `.xlsx`、写 `.txt`、调用联网检索。
换 harness 时不需要改代码，只要放对目录、装依赖、必要时改配置。

## 1. 放到目标 harness 的 skill 目录

目录名保持 `news-draft-skill`，`SKILL.md` 的 `name` / `description` 不用改。

| harness | skill 目录 |
|---|---|
| DSH | `~/.dsh/skills/news-draft-skill` |
| Claude Code | `~/.claude/skills/news-draft-skill` |
| OpenCode | `~/.opencode/skills/news-draft-skill` |
| OpenClaw | `~/.openclaw/skills/news-draft-skill` |
| Cursor / Windsurf | `<项目>/.skills/news-draft-skill` |
| 其它平台 | `<该平台的 agent_skill_dir>/news-draft-skill` |

## 2. 装依赖

```bash
python3 -m pip install openpyxl          # Windows 上解释器名是 python
```

只有 openpyxl 一个第三方包，其余全部是标准库（argparse/json/re/pathlib/subprocess/tempfile…）。

## 3. 装 AnySearch（联网补企业资料与行业政策用）

- 官方 skill 装到**同一 harness 的 skill 目录**下，目录名 `anysearch`
  （官方文档列了各平台的目录：`~/.claude/skills/anysearch`、`~/.opencode/skills/anysearch`、
  Cursor 的 `<项目>/.skills/anysearch` 等）。
- Python 版需要 `requests`（`python3 -m pip install requests`）；Node 版零依赖；Windows 还有 PowerShell 版。
- API Key 写到官方 skill 目录下的 `.env`：`ANYSEARCH_API_KEY=...`（没有 key 也能用匿名额度）。

`scripts/anysearch_search.py` 的探测顺序：

1. 环境变量 `NEWS_ANYSEARCH_CMD`（或兼容用的 `ANYSEARCH_CMD`）——直接指定命令行
2. `config.json` 的 `anysearch_cmd`
3. 环境变量 `ANYSEARCH_SKILL_DIR` / `config.json` 的 `anysearch_skill_dir`
4. 各平台常见 skill 目录下的 `anysearch/`：`.dsh`、`.claude`、`.opencode`、`.openclaw`、`.cursor`、
   `.windsurf`、`.codex`、`.agents`，以及当前目录的 `.skills/`、`.claude/skills/`
   （可用 `NEWS_SKILL_ROOTS` 追加自定义根目录，多个用系统路径分隔符隔开）
5. 与本 skill 同级的 `anysearch/`

解释器回退：官方 `runtime.conf` → Python CLI（需 requests）→ Node CLI（零依赖）→
Windows PowerShell CLI → POSIX bash CLI（需 `bash`+`jq`+`curl`）。

自检：

```bash
python3 scripts/anysearch_search.py --show-cmd      # 打印最终用的命令
python3 scripts/anysearch_search.py "中文域名" --max_results 2
```

## 4. 配置（可选）

把 `config.example.json` 复制成 `config.json` 再填。**环境变量优先级高于 config.json。**

| 配置项 | 环境变量 | 作用 |
|---|---|---|
| `anysearch_cmd` | `NEWS_ANYSEARCH_CMD` / `ANYSEARCH_CMD` | 直接指定检索命令（最高优先级） |
| `anysearch_skill_dir` | `ANYSEARCH_SKILL_DIR` | 指定官方 AnySearch skill 目录 |
| `state_dir` | `NEWS_STATE_DIR` | 科普话题状态文件放哪（见第 6 节） |
| — | `NEWS_SKILL_ROOTS` | 追加要扫描的 skill 根目录 |
| — | `NEWS_NODE` | 指定 node 可执行文件 |

## 5. 怎么调用

```bash
python3 scripts/news_draft.py list    --xlsx 表.xlsx
python3 scripts/news_draft.py extract --xlsx 表.xlsx --row 2 --project "某某.网址"
python3 scripts/news_draft.py save    --kind company --xlsx 表.xlsx --row 2 \
        --body-file 企业稿.txt --science-body-file 科普稿.txt
python3 scripts/news_draft.py check   --file 稿件.txt --xlsx 表.xlsx --row 2
python3 scripts/science.py --list
```

`python3` 换成 `python`（Windows）；命令和参数完全一致。

## 6. 平台差异与已知边界

- **换行符**：交付 `.txt` 默认 CRLF（Windows 记事本友好）。macOS/Linux 想要 LF 就加 `--lf`。
- **只读 skill 目录**：有的 harness 把 skills 目录挂成只读。科普话题状态文件会自动按
  `NEWS_STATE_DIR` → `config.json:state_dir` → skill 目录 `.state/` → 用户缓存目录
  （`~/.cache/news-draft-skill`、Windows 为 `%LOCALAPPDATA%\news-draft-skill`）依次挑一个**可写**的位置。
- **文件名净化**：Windows 非法字符替换为 `_`；引号类字符（`“”‘’`）一律去掉——Windows 上
  带引号的文件名会让 PowerShell 截断参数，去掉后两个平台都能正常引用。
- **联网工具名**：brief 里写的是「本 harness 自带的联网检索／网页抓取工具」，不写死具体工具名
  （DSH 是 `web_search`/`web_fetch`，Claude Code 是 `WebSearch`/`WebFetch`）。
- **工作目录**：中间过程文件（brief、草稿、门禁记录）默认放系统临时目录的
  `news-draft/<项目名>/`，交付目录里只留最终 `.txt`；`--work-dir` 可改，`--keep-work` 放回项目目录。

## 7. 不包含什么

**发布流程**。原来负责发搜狐号/头条号/CSDN 的 `media-publish-skill` 已停用并删除，
备份在 `~/.dsh/skill-backups/media-publish-skill-*.zip`。本 skill 只产出 `.txt` 稿件。

## 8. 一次跑通的自检

```bash
python3 scripts/news_draft.py list --xlsx 表.xlsx                 # 能读表
python3 scripts/news_draft.py extract --xlsx 表.xlsx --row 1      # 出两份 brief（企业稿+科普稿）
python3 scripts/science.py --list                                 # 话题状态可读写
python3 scripts/anysearch_search.py --show-cmd                    # 检索链路可达
python3 scripts/news_draft.py check --file 任意稿件.txt \
        --company 某某公司 --domain 某某.网址                      # 门禁可跑
```
