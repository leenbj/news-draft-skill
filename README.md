# news-draft-skill

中文域名新闻稿写作 skill：**读一张 `.xlsx`，一次产出两篇 `.txt`**——一篇企业新闻稿，一篇配套的中文域名科普稿。

只负责写稿，不含发布功能。

## 它做什么

```
list（看状态：只处理未完成的）
  → extract      一次产出两份 brief：企业稿 + 科普稿，并给这一行预占一个科普话题
  → agent 写稿    （先按 human-writing 写实，再按 humanizer-zh 清腔）
  → save         一次过四道门禁 → 企业稿落 <项目目录>/、科普稿落 <项目目录>/科普文章/
                 → 表格状态写「已完成」
```

- **企业稿**：以「某某公司官网启用某某域名」开篇，依次写企业信息与行业位置、中文域名本身、
  启用中文域名的意义，站在行业／产业升级的高度，默认 1000 字左右。
- **科普稿**：话题从 `references/中文域名科普资料库.md` 里挑（16 个话题轮换、用过不重复），
  只宣扬中文域名的正面形象，正文只允许出现 `.网址` 和 `.中国` 两种中文域名后缀。
- **缺企业资料时**：企业介绍为空就联网检索；检索不到该企业任何资料时进入**行业模式**——
  按企业名称推断行业（20 个行业 + 兜底），写行业现状与国家政策，政策只引用核实过的文件名。

## 门禁（脚本强制，不合格不落盘）

| 门禁 | 内容 |
|---|---|
| ① 内容 | 首段三要素、字数、四块结构、行业视角、启用意义、联系方式、平台名、写作过程话语、Markdown 残留 |
| ② 中文后缀 | 正文只允许 `.网址` 和 `.中国`；其他中文后缀硬错 |
| ③ 标题 | 与已交付标题完全相同硬错；句式签名相同提醒换角度 |
| ④ 人化 | 破折号、提示性冒号、29 个硬禁词、模型路标、翻案腔（10 类）硬错 |
| ⑤ 政策 | 引用的《文件名》不在已核实清单里会提醒先核实 |
| ⑥ 科普稿 | 不得出现任何关于中文域名的负面表述；冷水句提醒改写成建设性做法 |

## 快速开始

```bash
# 依赖：只有 openpyxl（Windows 上解释器名是 python）
python3 -m pip install openpyxl

# 看任务（状态列会标出哪些已完成）
python3 scripts/news_draft.py list --xlsx 表.xlsx

# 一次出两份 brief
python3 scripts/news_draft.py extract --xlsx 表.xlsx --row 2 --project "某某.网址"

# 两篇都写好后，一条命令同时落盘
python3 scripts/news_draft.py save --kind company --xlsx 表.xlsx --row 2 \
  --body-file 企业稿.txt --science-body-file 科普稿.txt
```

联网补料用 AnySearch（`https://www.anysearch.com/`）：把官方 skill 装到同一 harness 的 skill 目录下，
或设置 `ANYSEARCH_SKILL_DIR` / `NEWS_ANYSEARCH_CMD`。没有可用检索命令时会提示改用 harness 自带的联网工具。

## 表格契约

表头按关键词识别，列顺序随意；一行一篇企业稿（配一篇科普稿）。

| 字段 | 表头关键词 | 说明 |
|---|---|---|
| 编号 | 编号、序号、No、ID | 可选，缺省用 Excel 行号 |
| 状态 | 状态、稿件状态、进度 | 可选；缺列时首次 save 自动插在「编号」后 |
| 项目名称 | 项目名称、项目 | 可选，决定输出文件夹名；缺省用 xlsx 文件名 |
| 企业名称 | 企业名称、公司名称、企业全称、单位名称 | **必填** |
| 注册域名 | 注册域名、中文域名、官网域名、域名、网址 | **必填** |
| 企业介绍 | 企业信息资料、企业介绍、公司简介、经营范围… | 可选，缺省时联网检索 |
| 简称 | 企业简称、公司简称、简称 | 可选，标题用 |

## 目录结构

```
SKILL.md                      skill 主文档（流程、表格契约、门禁、常见问题）
INSTALL-multi-harness.md      各平台安装与配置（Claude Code / OpenCode / DSH / Cursor…）
scripts/
  news_draft.py               主入口：list / extract / science-brief / save / check / mark / search
  gates.py                    四道门禁的判定规则
  science.py                  科普稿话题库（16 个）与话题轮换
  industry.py                 按企业名称推断行业 + 行业现状 + 可引用政策清单
  anysearch_search.py         AnySearch 接入层（跨平台探测 CLI）
references/
  写作规范.md                  完整写作规范
  中文域名科普资料库.md          科普稿的话题来源
  中文域名背景.md / 网址背景资料.md  中文域名部分的写作参考
  网址销售彩页参考.txt           销售彩页原文（备查）
  human-writing.md            第一阶段「写实」规则
  humanizer-zh.md             第二阶段「清 AI 腔」规则
```

## 多 harness

纯 Python 3 + openpyxl，不用 DSH 专有接口；路径、解释器、状态目录、换行符都可配置。
详见 [INSTALL-multi-harness.md](INSTALL-multi-harness.md)。

## 第三方材料与出处

本仓库的写作纪律与背景资料来自以下来源，版权归原方：

- `references/human-writing.md` — 蒸馏自 [KKKKhazix/human-writing](https://github.com/KKKKhazix/human-writing) 1.1.0（commit `4fda173f`）
- `references/humanizer-zh.md` — 蒸馏自 [op7418/Humanizer-zh](https://github.com/op7418/Humanizer-zh)（commit `91f3d394`）
- `references/网址背景资料.md`、`references/网址销售彩页参考.txt` — 来自用户提供的《中文网址销售彩页》材料
- `references/中文域名科普资料库.md` — 用户提供的科普话题资料

引用与再分发前请自行确认各来源的授权条件。代码部分未附许可证；如需开源请自行补充。
