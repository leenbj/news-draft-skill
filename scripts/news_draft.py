#!/usr/bin/env python3
"""news-draft-skill 主脚本：.xlsx → 材料 → 新闻稿 .txt

子命令
  list     列出表格里的稿件任务（编号 / 企业名称 / 注册域名 / 有无企业介绍 / 项目名称）
  extract  取一行的三条信息（企业名称、注册域名、企业介绍）；企业介绍缺失时用 AnySearch
           联网检索，然后把「材料 + 写作要求 + 落盘命令」写成一份 brief，交给 agent 写稿
  save     校验 agent 写好的稿件（门禁），通过后写成 .txt，存进以项目名称命名的文件夹
  check    只跑门禁，不落盘
  search   临时联网检索（转发给 AnySearch）

表格契约（表头关键词识别，列顺序随意；缺列会明确报错）
  编号 / 序号            可选，缺省用 Excel 行号
  项目名称 / 项目          可选，缺省用 xlsx 文件名；决定输出文件夹名
  企业名称 / 公司名称       必填
  域名 / 中文域名 / 网址     必填，如 品牌.网址
  企业介绍 / 企业信息 / 简介  可选，缺省时联网检索
  简称                   可选，标题用

输出
  <xlsx 所在目录>/<项目名称>/<标题>.txt          交付物（UTF-8，标题 + 空行 + 正文）
  <xlsx 所在目录>/<项目名称>/过程文件/            材料、草稿、门禁结果
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import gates  # noqa: E402
import industry  # noqa: E402
import science  # noqa: E402
from anysearch_search import (AnySearchError, search_company,  # noqa: E402
                              search_industry)

WORK_DIR_NAME = "过程文件"
HEADER_KEYS = {
    "num": ("编号", "序号", "No.", "No", "ID", "id"),
    "project": ("项目名称", "项目名", "项目"),
    "company": ("企业名称", "公司名称", "企业全称", "单位名称", "企业", "公司"),
    "domain": ("注册域名", "中文域名", "官网域名", "域名", "网址"),
    "profile": ("企业信息资料", "企业介绍", "企业信息", "企业资料", "公司介绍",
                "公司简介", "企业简介", "经营范围", "主营业务", "简介"),
    "short": ("企业简称", "公司简称", "简称"),
    "status": ("稿件状态", "写作状态", "完成状态", "状态", "进度"),
}
XLSX_OUT_COLS = ("项目名称", "稿件文件", "稿件字数", "生成时间")
STATUS_COL = "状态"
STATUS_DONE = "已完成"                      # 写完就写这个值
DONE_MARKS = ("已完成", "完成", "已发布", "已生成")
NOT_DONE_MARKS = ("未完成", "待写", "未写", "待处理", "待生成", "不写", "待科普")
RESERVED_NAMES = ({"CON", "PRN", "AUX", "NUL"}
                  | {f"COM{i}" for i in range(1, 10)}
                  | {f"LPT{i}" for i in range(1, 10)})


# ── 输出与文本工具 ────────────────────────────────────────────────
def _newline(args) -> str:
    """交付文本的换行符：默认 CRLF（Windows 记事本友好），--lf 则用 LF。"""
    return "\n" if getattr(args, "lf", False) else "\r\n"


def _utf8_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass


def safe_name(name: str, max_len: int = 90) -> str:
    n = (name or "").strip()
    # 引号类字符在 Windows 文件名里合法，但在 PowerShell/命令行里会截断参数（“标题”.txt 传不进来），
    # 所以文件名一律去掉它们；稿件正文里的引号保持原样。
    n = re.sub(r"[“”‘’″‶]", "", n)
    n = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", n)
    n = re.sub(r"\s+", " ", n).strip(" .")
    if not n:
        n = "未命名稿件"
    if n.split(".")[0].upper() in RESERVED_NAMES:
        n = "_" + n
    return n[:max_len].rstrip(" .") or "未命名稿件"


def norm_domain(raw) -> str:
    d = str(raw or "").strip()
    d = re.sub(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", "", d)
    d = d.split("/")[0].split("?")[0].split("#")[0]
    for ch in ("。", "．", "｡"):
        d = d.replace(ch, ".")
    d = re.sub(r"\s+", "", d)
    return d.strip(".")


def norm_profile(raw) -> str:
    text = str(raw or "")
    text = re.sub(r"\r\n?", "\n", text).strip()
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text


def normalize_body(body: str) -> str:
    body = re.sub(r"[ \t]+$", "", (body or "").replace("\r\n", "\n").replace("\r", "\n"), flags=re.M)
    return re.sub(r"\n{3,}", "\n\n", body).strip()


def split_title_body(text: str, title: str = "") -> tuple[str, str]:
    """模型/agent 常把标题写在正文第一行，这里按需剥掉，避免落盘后重复。"""
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    lines = text.split("\n")
    if lines:
        first = lines[0].strip().lstrip("#").strip()
        first = re.sub(r"^(?:标题|题目|Title)\s*[:：]\s*", "", first, flags=re.I).strip()
        first = first.strip("《》\"'“”")
        want = (title or "").strip().strip("《》\"'“”")
        if want and first and first == want:
            lines = lines[1:]
            while lines and not lines[0].strip():
                lines.pop(0)
    return (title or "").strip(), normalize_body("\n".join(lines))


def looks_like_deliverable(text: str) -> tuple[str, str]:
    """读取已落盘的 .txt：首行标题 + 空行 + 正文。

    判定要保守：首行必须是单行、不超标题上限、不以句末标点收尾，且后面确实还有正文；
    否则一律当成“没有标题行的纯正文”，避免把正文第一段误当标题吃掉。"""
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    parts = re.split(r"\n\s*\n", text, maxsplit=1)
    if len(parts) != 2:
        return "", normalize_body(text)
    head, rest = parts[0].strip(), parts[1]
    if ("\n" in head or not head or len(head) > gates.MAX_TITLE_CHARS
            or re.search(r"[。！？；]$", head) or gates.prose_len(rest) < 50):
        return "", normalize_body(text)
    return head, normalize_body(rest)


# ── 表格读取 ─────────────────────────────────────────────────────
def map_headers(headers: list) -> dict:
    """表头关键词 → 列下标。长关键词优先，一列只归一个字段。"""
    pairs = sorted(((k, f) for f, ks in HEADER_KEYS.items() for k in ks),
                   key=lambda p: -len(p[0]))
    mapping: dict[str, int] = {}
    for idx, h in enumerate(headers):
        if h is None:
            continue
        hs = str(h).strip()
        if not hs:
            continue
        for key, field in pairs:
            if key in hs and field not in mapping:
                mapping[field] = idx
                break
    return mapping


def read_rows(xlsx: Path) -> tuple[list[dict], dict]:
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb.worksheets[0]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    mapping = map_headers(headers)
    missing = [f for f in ("company", "domain") if f not in mapping]
    if missing:
        label = {"company": "企业名称", "domain": "域名"}[missing[0]]
        raise SystemExit(
            f"表格缺少必需的「{label}」列。当前表头：{headers}\n"
            "需要的表头关键词见 SKILL.md 的“表格契约”。")
    default_project = xlsx.stem
    rows: list[dict] = []
    for r in range(2, ws.max_row + 1):
        def cell(field: str) -> str:
            idx = mapping.get(field)
            return "" if idx is None else str(ws.cell(row=r, column=idx + 1).value or "").strip()

        company, domain = cell("company"), norm_domain(cell("domain"))
        if not company and not domain:
            continue
        num = cell("num") or str(r)
        rows.append({
            "excel_row": r,
            "num": num,
            "project": cell("project") or default_project,
            "company": company,
            "domain": domain,
            "profile": norm_profile(cell("profile")),
            "short": cell("short"),
            "status": cell("status"),
        })
    if not rows:
        raise SystemExit("表格里没有读到任何有效行（企业名称与域名都为空）")
    return rows, {"headers": headers, "mapping": mapping}


def pick_rows(rows: list[dict], row_arg) -> list[dict]:
    if row_arg is None or str(row_arg).strip() == "":
        return rows
    s = str(row_arg).strip()
    hit = [x for x in rows if str(x["num"]).strip() == s]
    if not hit and s.isdigit():
        hit = [x for x in rows if x["excel_row"] == int(s)]
    if not hit:
        known = ", ".join(str(x["num"]) for x in rows)
        raise SystemExit(f"表格里找不到编号或行号 {row_arg}；现有编号：{known}")
    return hit


def excel_col(ws, name: str) -> int:
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(row=1, column=c).value or "").strip() == name:
            return c
    c = ws.max_column + 1
    ws.cell(row=1, column=c, value=name)
    return c


def is_done(status: str) -> bool:
    """状态列是否表示这篇已经写完（“未完成”不能被当成“完成”）。"""
    s = re.sub(r"\s+", "", str(status or ""))
    if not s or any(m in s for m in NOT_DONE_MARKS):
        return False
    return any(m in s for m in DONE_MARKS)


def find_header(ws, keys: tuple) -> int | None:
    """按表头关键词找列号（1 起），返回最左命中列。"""
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=1, column=c).value or "").strip()
        if h and any(k in h for k in keys):
            return c
    return None


def write_tracking(xlsx: Path, excel_row: int, values: dict, *,
                   status: str = "", status_pos: str = "after-num") -> str:
    """回写表格：状态列不存在就先插出来，再写状态与其它跟踪列。

    status_pos="after-num"：状态列插在「编号」列后面（默认）；
    status_pos="end"：追加到表尾（表格还交给别的按列号读表的工具时用这个）。"""
    try:
        wb = openpyxl.load_workbook(xlsx)
        ws = wb.worksheets[0]
        if status:
            col = find_header(ws, HEADER_KEYS["status"])
            if col is None:
                pos = None
                if status_pos == "after-num":
                    anchor = find_header(ws, HEADER_KEYS["num"])
                    pos = anchor + 1 if anchor else None
                if pos is None:
                    pos = ws.max_column + 1
                ws.insert_cols(pos)
                ws.cell(row=1, column=pos, value=STATUS_COL)
                col = pos
            ws.cell(row=excel_row, column=col, value=status)
        for name, value in values.items():
            ws.cell(row=excel_row, column=excel_col(ws, name), value=value)
        wb.save(xlsx)
        return ""
    except Exception as exc:  # noqa: BLE001
        return (f"回写表格失败（{type(exc).__name__}: {exc}）。稿件本身已保存；"
                "表格若是打开状态，关掉后重跑 save 即可补写状态。")


# ── 路径 ────────────────────────────────────────────────────────
def project_dir_of(xlsx: Path | None, out_root: Path | None, project: str) -> Path:
    root = Path(out_root) if out_root else (Path(xlsx).resolve().parent if xlsx else Path.cwd())
    return root / safe_name(project, 60)


def work_dir_of(pdir: Path, args=None) -> Path:
    """中间过程文件（材料 brief / 草稿 / 门禁记录）放哪。

    默认放系统临时目录：交付目录里只留最终 .txt。
    --work-dir 指定目录；--keep-work 放回 <项目目录>/过程文件（老行为）。"""
    if args is not None:
        if getattr(args, "work_dir", None):
            return Path(args.work_dir).expanduser()
        if getattr(args, "keep_work", False):
            return pdir / WORK_DIR_NAME
    return Path(tempfile.gettempdir()) / "news-draft" / safe_name(pdir.name, 60)


# ── brief（材料 + 写作要求）───────────────────────────────────────
SCIENCE_TARGET_CHARS = 1200      # 科普稿默认目标字数（比企业稿长一点，便于把话题讲透）


def resolve_length(args, default_target: int = gates.DEFAULT_TARGET_CHARS) -> tuple[int, int, int]:
    """把「字数要求」解析成 (下限, 上限, 目标)。

    默认：1000 字左右 → 下限 900、上限 1100（上限只提醒，不拦截落盘）。
    --target-chars 800 → 800 字左右（720-880）。
    --min-chars/--max-chars 直接写死区间；只给一个时另一边不限。
    """
    target = getattr(args, "target_chars", None)
    lo = getattr(args, "min_chars", None)
    hi = getattr(args, "max_chars", None)
    if lo is None and hi is None:
        target = int(target) if target else default_target
        lo = round(target * gates.TARGET_MIN_RATIO)
        hi = round(target * gates.TARGET_MAX_RATIO)
    else:
        lo = int(lo) if lo else 0
        hi = int(hi) if hi else 0
        target = int(target) if target else 0
    return lo, hi, target


URL_MATERIAL = Path(__file__).resolve().parent.parent / "references" / "网址背景资料.md"

# 标题角度菜单：brief 让 agent 每次换一种，别都写成「X官网启用Y，Z有了中文入口」
TITLE_ANGLES = (
    "入口式：{{简称}}官网启用“{{域名}}”，{{品类}}企业有了中文入口（这条最容易写重复，用过一次就别再用）",
    "直述式：{{企业全称}}官网启用“{{域名}}”",
    "趋势式：{{行业}}企业的线上入口在变，{{简称}}启用“{{域名}}”",
    "区域式：{{地区或产业带}}的{{行业}}企业，{{简称}}把官网地址写成了中文",
    "政策式：{{政策方向}}推进中的{{行业}}，{{简称}}启用“{{域名}}”",
    "品牌式：“{{域名}}”上线，{{简称}}把企业名做成了官网地址",
    "数字化式：{{行业}}的数字门面在变，{{简称}}官网启用“{{域名}}”",
    "场景式：客户要在手机上找到厂，{{简称}}启用“{{域名}}”",
)


def url_material_context(limit: int = 4000) -> str:
    """读 media-publish-skill 带过来的 .网址 背景资料（中文域名部分专用）。"""
    try:
        return URL_MATERIAL.read_text(encoding="utf-8")[:limit]
    except OSError:
        return ""


def used_titles(pdir: Path, exclude_name: str = "") -> list[str]:
    """已交付稿件的标题，用于标题去重。

    项目名往往就是域名，一篇一个文件夹，只看自己那个文件夹等于没去重。
    所以同时扫两层：本项目文件夹，以及它的上级目录（同一张表的稿件根目录）和
    根目录下的各个项目文件夹。"""
    pdir = Path(pdir)
    out, seen = [], set()
    targets = []
    for root in (pdir, pdir.parent):
        try:
            targets += sorted(root.glob("*.txt"))
            targets += sorted(root.glob("*/*.txt"))
        except OSError:
            continue
    for p in targets:
        if exclude_name and p.name == exclude_name:
            continue
        try:
            first = ""
            for line in p.read_text(encoding="utf-8-sig", errors="replace").splitlines():
                if line.strip():
                    first = line.strip()
                    break
        except OSError:
            continue
        if first and len(first) <= gates.MAX_TITLE_CHARS and first not in seen:
            seen.add(first)
            out.append(first)
    return out


def writing_requirements(company: str, domain: str, policy: str,
                         min_chars: int = gates.MIN_BODY_CHARS,
                         max_chars: int = 0, target_chars: int = 0,
                         industry_mode: bool = False) -> str:
    if target_chars and max_chars:
        length_rule = (f"正文写到 {target_chars} 字左右（不含标题；可接受 {min_chars}-{max_chars} 字）。"
                       "写短了只能补充可核验事实，写长了要删重复、删套话、删形容词，"
                       "不能为了压字数删掉已确认的事实。")
    elif max_chars:
        length_rule = (f"正文控制在 {min_chars}-{max_chars} 字（不含标题）。"
                       "写短了只能补充可核验事实，写长了要删重复、删套话、删形容词，不能删掉已确认的事实。")
    else:
        length_rule = (f"正文不少于 {min_chars} 字（不含标题）。"
                       "材料不足时补充可核验事实，不要靠重复和形容词凑字数。")
    industry_block_rule = (
        "\n   **这一篇没有企业资料**：第二块改成「行业与政策」，篇幅放大到正文的四成左右。"
        "按 brief 第三节之二给出的行业名称、行业现状、国家政策来写，讲清这个行业正在经历什么、"
        "政策往哪个方向推、企业的信息入口在这个背景下意味着什么。行业素材只能写成行业层面的普遍情况，"
        "**不得安到这家企业头上**；政策只写「文件名称 + 发布单位 + 年份」，落笔前用 AnySearch 核实，"
        "核不到就不写政策，只写行业现状。"
        if industry_mode else "")
    suffix_note = ("**中文后缀审核**：全篇只允许出现 .网址 和 .中国 两种中文域名后缀；"
                   ".公司/.网络/.商城/.商标/.在线/.网店/.集团/.中文网 等其他中文后缀一律不得出现，"
                   "也不得出现 .com/.cn 等英文后缀和“英文域名／国际域名”这类说法。"
                   if policy != "any-zh" else
                   "不得出现 .com/.cn 等英文后缀，也不得出现“英文域名／国际域名”这类说法。")
    return f"""### 篇幅与结构（四块，按顺序写，不加编号小标题）

1. **核心内容开篇（第一段）**：第一句直接写「{company}官网启用“{domain}”」，紧接一句说清这对官网识别和用户访问入口的直接意义。
   不铺垫行业背景、不设问、不抒情，不要先写“随着……”再转弯。
2. **企业信息与行业位置**：先站在所在行业的高度，写这类企业在数字化、线上获客、招投标、资料发布、客户核验等场景里的共性问题或变化，
   再落到该企业的少量公开事实（主营、典型场景、所在地区）。企业自身介绍只占很小一部分；
   注册资本、成立时间、法定代表人、信用代码、注册地址、股东这类工商公开信息一律不写。{industry_block_rule}
3. **中文域名本身**：讲清中文域名由中文词语构成、怎么写怎么读、在浏览器地址栏如何输入、与线下名称和品牌写法如何对应，
   属于识别与访问层面的常识。这一块是给不熟悉中文域名的读者看的，不要写成企业宣传。
4. **该企业启用中文域名的意义**：结合企业名称和品牌写法，说明入口识别、访问路径统一、资料集中带来的实际作用，
   可以落到行业层面（信息入口规范化、品牌资产一致性、产业数字化的基础设施）。只写可观察的使用价值。

末段回到已确认的事实和待核验边界，简短收束。

### 硬要求

- {length_rule}
- 站在行业／产业升级的高度写，但每一句都要落在事实或可观察的作用上，不写宏大空话。
- {suffix_note}
- 企业事实的来源只写“企业公开资料”；不得出现任何网站、平台、数据库、招聘网站的名称。
- **不得使用任何工商公开信息**：注册资本、实缴资本、成立时间、法定代表人、统一社会信用代码、注册地址、
  股东与股权、经营状态（存续/吊销/注销）一律不写；联网检索材料里只有业务相关内容可以用。
- 全文不得出现联系方式（电话、手机号、邮箱、微信、QQ）和法定代表人或联系人姓名。
- 不承诺流量、排名、销量、转化、平台效果或法律结果。
- 不虚构引语、现场、客户、数字或未来计划；没有材料的细节宁可不写。
- 不出现“本文、本稿、稿件、检索结果、信息边界、提示词”这类写作过程话语。
- 纯文本：不要 Markdown 标记（#、**、列表符号），不要 HTML 标签。

### 语言

白话打底，主语和动作尽早出现，句长有变化。不用翻案腔（“不是……而是……”）、三项同构排比、破折号、
提示性冒号，不用“赋能、抓手、闭环、底座、降本增效”这类汇报黑话，不用“未来可期”这类万能结尾。
**两阶段处理（必做）**：先按 `references/human-writing.md` 把稿子写实（每段推进新事实、材料分身份，
不编现场/引语/数字），再按 `references/humanizer-zh.md` 清一遍 AI 腔（删套话、夸大意义、模糊归因、
假金句、万能结尾）。没做第二遍就交，save 时会被「人化门禁」按硬规则拦下：
破折号、提示性冒号、硬禁词（赋能/抓手/闭环/底层逻辑…）、模型路标（值得注意的是…）、
翻案腔（不是…而是…）都会判硬错。
中文域名背景可参考 `references/中文域名背景.md`；完整规范见 `references/写作规范.md`。"""


def build_brief(row: dict, pdir: Path, search_text: str, search_errors: list[str],
                search_state: str, min_chars: int = gates.MIN_BODY_CHARS,
                max_chars: int = 0, target_chars: int = 0,
                wdir: Path | None = None, search_notes: list[str] | None = None,
                industry_mode: bool = False, industry_text: str = "",
                industry_notes: list[str] | None = None,
                used: list[str] | None = None) -> str:
    wdir = Path(wdir) if wdir else work_dir_of(pdir)
    policy = "zh-only" if row["domain"].endswith(".网址") else "relaxed"
    if row["profile"]:
        profile_block = f"{row['profile']}"
        profile_src = "表格提供（写作以此为准，检索材料只作补充；两者冲突时以表格为准）"
    else:
        profile_src = ("表格未提供 → 已用 AnySearch 联网检索，检索结果见下一节（只作补充，"
                       "读不到的内容宁可不写）")
        profile_block = "（表格未填写企业介绍）"
    if search_state == "searched" and search_text:
        search_block = (
            "⚠ 检索材料只作补充，用之前先过两道筛子：\n"
            "1）核对是不是同一家企业（企业全称要能对上）；名字相近但不是同一家的，整条丢掉。\n"
            "2）**只准用企业业务相关的内容**：主营产品、服务、应用场景、行业、客户类型。"
            "注册资本、实缴资本、成立时间、法定代表人、统一社会信用代码、注册地址、股东与股权、"
            "经营状态（存续/吊销/注销）这类工商公开信息一律不得写进正文，读到也不用。\n"
            + ("".join(f"{n}\n" for n in (search_notes or [])))
            + "\n" + search_text)
    elif search_state == "searched-empty":
        search_block = "（AnySearch 检索没有返回可用材料）"
    elif search_state == "errors":
        search_block = ("（AnySearch 检索失败：" + "；".join(search_errors[:3])
                        + "）\n请用本 harness 自带的联网检索／网页抓取工具补两条企业公开资料"
                          "（DSH 里是 web_search/web_fetch，Claude Code 里是 WebSearch/WebFetch），"
                          "再把可核验的事实写进稿件。")
    elif search_state == "skipped-by-flag":
        search_block = "（按 --search never 跳过联网检索；材料不足时用本 harness 自带的联网检索工具补充）"
    else:
        search_block = "（表格已提供企业介绍，未联网检索；需要补充时用 extract --search always）"

    # ── 行业模式：企业资料为零时，第二块改写行业与政策 ──
    ind = industry.industry_material(row["company"])
    industry_section = ""
    if industry_mode:
        policy_lines = "\n".join(f"  - {p}" for p in ind["policy"])
        status_lines = "\n".join(f"  - {s}" for s in ind["status"])
        note_lines = "".join(f"{n}\n" for n in (industry_notes or []))
        industry_section = f"""
## 三之二、行业与政策素材（本篇没有企业资料，这里是主要写作来源）

- 从企业名称推断的行业：**{ind['name']}**{f"（依据名称中的“{ind['matched_keyword']}”）" if ind["inferred"] else "（名称里没有行业词，按通用制造业处理）"}
- 常见场景与客户：{ind['scene']}
- 行业现状（**行业层面**的普遍情况，不得说成这家企业的事实）：
{status_lines}
- 可引用的国家政策（**落笔前必须用 AnySearch 核实**，只写「名称 + 发布单位 + 年份」）：
{policy_lines}
- 数字化抓手（行业共性的表述，别写成效承诺）：{"；".join(ind['digital'])}
{note_lines}
写作口径：这一块要写到正文的四成左右，讲清行业处境、正在发生的变化、政策方向，再自然接到
这家企业启用中文域名这件事上。行业素材不要安到具体企业头上；不写没有出处的数字、市场份额、
排名和案例；政策名称核不到就不写，只写行业现状。
"""
    url_section = ""
    ctx = url_material_context()
    if ctx:
        url_section = f"""
### 中文域名背景资料（来自 media-publish-skill 的 .网址 资料，第三块直接用）

{ctx}

用法：这块资料只用于写「中文域名本身」和「启用的意义」，不得当成这家企业的经营事实；
不抄销售页里的夸张宣传、具体企业背书和联系方式；不引入 `.网址`、`.中国` 以外的域名后缀。
"""
    used = used or []
    title_section = f"""
## 四之二、标题要求（每篇都要不一样）

本项目已经用过的标题（**不许重复，也不要同一种句式**）：

{chr(10).join(f"  - {t}" for t in used) if used else "  - （这是本项目第一篇）"}

要求：标题要出现企业名称或简称，说明这篇的信息点，不以标点收尾。
**换角度写**，下面这些句式挑一个没被用过的（括号里是写法示意）：

{chr(10).join(f"  - {a}" for a in TITLE_ANGLES)}

save 时会比对已用标题的「句式签名」（是否含 官网启用/启用/上线/有了中文入口/产业/行业 等标记的组合）：
标题与已有标题完全相同会判硬错；句式签名与某一篇相同会提醒换角度。
"""
    draft = wdir / f"{safe_name(row['num'], 20)}-草稿.txt"
    py = sys.executable
    script = Path(__file__).resolve()
    xlsx = row.get("_xlsx", "")
    return f"""# 稿件材料：{row['company']}（编号 {row['num']}）

- 项目名称：{row['project']}
- 交付目录：{pdir}
- 企业名称：{row['company']}
- 注册域名：{row['domain']}
- 企业简称：{row['short'] or "（未提供）"}
- 企业介绍来源：{profile_src}
- 字数要求：{"约 " + str(target_chars) + " 字（可接受 " + str(min_chars) + "-" + str(max_chars) + " 字）" if target_chars and max_chars else ("不少于 " + str(min_chars) + " 字" if not max_chars else str(min_chars) + "-" + str(max_chars) + " 字")}
- 后缀策略：{policy}（{"只允许 .网址" if policy == "zh-only" else "允许其他中文后缀，禁止英文后缀"}）

## 一、必写的核心事实（开篇第一句）

{row['company']}官网启用“{row['domain']}”。

## 二、企业介绍

{profile_block}

## 三、联网检索材料（仅作补充）

{search_block}
{industry_section}
## 四、写作要求

{writing_requirements(row['company'], row['domain'], policy, min_chars, max_chars, target_chars, industry_mode)}
{url_section}
{title_section}
## 五、写完怎么落盘

把标题和正文写成 UTF-8 文本（首行标题，空行，然后正文），保存到：

    {draft}

然后运行（路径已按当前环境填好，可直接执行）：

    "{py}" "{script}" save --xlsx "{xlsx}" --row "{row['num']}" --min-chars {min_chars} --max-chars {max_chars} --body-file "{draft}"

标题直接写在草稿首行即可（脚本读首行当标题，并剥掉正文里重复的那一行）；想覆盖再加 `--title 标题`。
门禁通过后，稿件会写成 {pdir}<首行标题>.txt；门禁不过会逐条列出硬错，照着改完再跑一次即可。
"""


# ── 子命令 ──────────────────────────────────────────────────────
def cmd_list(args) -> int:
    xlsx = Path(args.xlsx)
    rows, meta = read_rows(xlsx)
    print(f"表格：{xlsx}")
    print(f"表头：{meta['headers']}")
    print(f"{'编号':<8}{'状态':<10}{'企业名称':<28}{'注册域名':<20}{'企业介绍':<12}项目名称")
    print("-" * 106)
    for r in rows:
        prof = f"有（{len(r['profile'])}字）" if r["profile"] else "无（需检索）"
        mark = r["status"] or "待写"
        print(f"{str(r['num']):<8}{mark[:8]:<10}{r['company'][:26]:<28}"
              f"{r['domain'][:18]:<20}{prof:<12}{r['project']}")
    print("-" * 106)
    done = sum(1 for r in rows if is_done(r["status"]))
    print(f"共 {len(rows)} 行：已完成 {done} 行（写作时跳过）、待处理 {len(rows) - done} 行；"
          f"缺企业介绍 {sum(1 for r in rows if not r['profile'])} 行。")
    if done:
        print("要重写已完成的行：extract 加 --force（或 --row 指定后加 --force）。")
    return 0


def cmd_extract(args) -> int:
    xlsx = Path(args.xlsx)
    rows, _ = read_rows(xlsx)
    todo = pick_rows(rows, args.row)
    # 状态列优先：已完成的行不再写。--force 才重做。
    skipped = [r for r in todo if is_done(r["status"])]
    if args.force:
        rows = todo
        for r in skipped:
            print(f"[{r['num']}] {r['company']} 状态「{r['status']}」→ --force 重做")
    else:
        rows = [r for r in todo if not is_done(r["status"])]
        for r in skipped:
            print(f"[{r['num']}] {r['company']} 状态「{r['status']}」→ 跳过（要重写加 --force）")
    if skipped and not rows:
        print(f"\n这一批全部已完成，没有需要写作的任务。")
        return 0
    created = []
    for r in rows:
        r["_xlsx"] = str(xlsx.resolve())
        project = args.project or r["project"]
        r["project"] = project          # brief 里显示生效的项目名，和交付目录一致
        pdir = project_dir_of(xlsx, args.out_root, project)
        wdir = work_dir_of(pdir, args)   # 过程文件不进交付目录
        wdir.mkdir(parents=True, exist_ok=True)

        search_text, search_errors, search_notes = "", [], []   # 三个都要初始化：不检索时也要有值
        if args.search == "always" or (args.search == "auto" and not r["profile"]):
            print(f"[{r['num']}] {'按要求' if r['profile'] else '企业介绍缺失，'}AnySearch 检索：{r['company']}")
            search_text, search_errors, search_notes = search_company(
                r["company"], r["profile"], args.max_results)
            if search_text:
                print(f"      命中 {len(search_text)} 字材料")
            for n in search_notes:
                print(f"      · {n}")
            for e in search_errors:
                print(f"      ! {e}")
            search_state = "searched" if search_text else ("errors" if search_errors else "searched-empty")
        else:
            search_state = "skipped-by-flag" if args.search == "never" else "skipped-has-profile"

        # 企业资料为零 → 行业模式：按名称推行业，并补行业与国家政策素材
        industry_mode = not r["profile"]
        industry_text, industry_notes = "", []
        if industry_mode:
            ind_name = industry.industry_material(r["company"])["name"]
            print(f"      企业资料为空 → 按名称推断行业：{ind_name}，补行业与政策素材")
            industry_text, errors2, industry_notes = search_industry(
                ind_name, args.max_results)
            for n in industry_notes:
                print(f"      · {n}")
            for e in errors2:
                print(f"      ! {e}")

        lo, hi, target = resolve_length(args)
        brief = build_brief(r, pdir, search_text, search_errors, search_state,
                            lo, hi, target, wdir=wdir, search_notes=search_notes,
                            industry_mode=industry_mode, industry_text=industry_text,
                            industry_notes=industry_notes,
                            used=used_titles(pdir))
        brief_md = wdir / f"{safe_name(r['num'], 20)}-材料.md"
        brief_md.write_text(brief, encoding="utf-8")
        (wdir / f"{safe_name(r['num'], 20)}-材料.json").write_text(
            json.dumps({k: v for k, v in r.items() if not k.startswith("_")}
                       | {"xlsx": str(xlsx.resolve()), "project_dir": str(pdir),
                          "brief": str(brief_md), "search_state": search_state,
                          "min_chars": lo, "max_chars": hi, "target_chars": target,
                          "search_errors": search_errors},
                       ensure_ascii=False, indent=2), encoding="utf-8")
        # 同一行同时产出科普稿 brief（话题预占，多行提取不会撞话题）
        sci_md = None
        if not getattr(args, "no_science", False):
            row_key = f"{safe_name(project, 60)}/{r['num']}"
            topic = science.reserve(row_key, getattr(args, "topic", "") or "")
            slo, shi, starget = resolve_length(args, SCIENCE_TARGET_CHARS)
            sci_brief, _, _ = build_science_brief(
                topic, pdir, wdir, safe_name(r["num"], 20), slo, shi, starget,
                xlsx, r["num"], project)
            sci_md = wdir / f"{safe_name(r['num'], 20)}-科普材料.md"
            sci_md.write_text(sci_brief, encoding="utf-8")
        created.append((r, brief_md, sci_md))

    print()
    for item in created:
        r, brief_md = item[0], item[1]
        sci_md = item[2]
        print(f"[{r['num']}] {r['company']} → {r['domain']}")
        print(f"      企业稿材料：{brief_md}")
        if sci_md:
            print(f"      科普稿材料：{sci_md}")
    print(f"\n下一步：这一行要出两篇——企业稿和中文域名科普稿。两份 brief 都读一遍，"
          f"各自写成「-草稿.txt」「-科普草稿.txt」，然后用一条 save 命令同时落盘"
          f"（--body-file + --science-body-file）。")
    return 0


def _context_from_args(args) -> tuple[str, str, str, str, Path | None, int | None]:
    """返回 (company, domain, short, project, xlsx, excel_row)。"""
    if args.xlsx:
        xlsx = Path(args.xlsx)
        rows, _ = read_rows(xlsx)
        if args.row is None:
            if len(rows) != 1:
                raise SystemExit(
                    f"表格有 {len(rows)} 行，请用 --row 指定编号：" +
                    ", ".join(str(r['num']) for r in rows))
            r = rows[0]
        else:
            r = pick_rows(rows, args.row)[0]
        project = args.project or r["project"]
        return (r["company"], r["domain"], r["short"], project, xlsx, r["excel_row"],
                not r["profile"])
    if not (args.company and args.domain):
        raise SystemExit("要么给 --xlsx（可加 --row），要么同时给 --company 与 --domain")
    project = args.project or "未命名项目"
    return (args.company, norm_domain(args.domain), args.short or "", project, None, None,
            not (getattr(args, "profile", "") or ""))


def _report_and_save(args, company, domain, short, project, xlsx, excel_row,
                     industry_mode: bool = False) -> int:
    body_file = Path(args.body_file)
    if not body_file.exists():
        raise SystemExit(f"找不到稿件文件：{body_file}")
    raw = body_file.read_text(encoding="utf-8-sig", errors="replace")
    if args.title:
        title, body = split_title_body(raw, args.title)
    else:
        title, body = looks_like_deliverable(raw)
    if not title:
        raise SystemExit(
            "无法确定标题：把标题写在稿件文件首行（其后空一行），或用 --title 指定。"
            "标题里带中英文引号时，建议直接写在首行，避免命令行转义问题。")

    policy = args.suffix_policy
    if policy == "auto":
        policy = "zh-only" if domain.endswith(".网址") else "relaxed"
    lo, hi, _ = resolve_length(args)
    pdir_for_titles = project_dir_of(xlsx, Path(args.out_root) if args.out_root else None, project)
    rep = gates.check_text(title, body, company, domain, short, policy,
                           min_chars=lo, max_chars=hi,
                           allow_registration=args.allow_registration_facts,
                           humanize_gate=not args.no_humanize_gate,
                           allow_any_zh_suffix=(policy == "any-zh"),
                           used_titles=used_titles(pdir_for_titles,
                                                   exclude_name=f"{safe_name(title)}.txt"),
                           industry_mode=industry_mode)
    if not rep.ok or rep.warnings:          # 通过且无提醒时不啰嗦，只报结果
        print(f"门禁：{company} / {domain}（{policy}）\n标题：{title}")
        print(rep.render())

    wdir = work_dir_of(project_dir_of(xlsx, Path(args.out_root) if args.out_root else None, project), args)
    wdir.mkdir(parents=True, exist_ok=True)
    tag = safe_name(args.row if args.row is not None else (excel_row or "稿"), 20)
    (wdir / f"{tag}-门禁.json").write_text(json.dumps(
        {"title": title, "company": company, "domain": domain, "policy": policy,
         "ok": rep.ok, "errors": rep.errors, "warnings": rep.warnings,
         "stats": rep.stats, "checked_at": datetime.now().isoformat(timespec="seconds")},
        ensure_ascii=False, indent=2), encoding="utf-8")

    if not rep.ok:
        print("\n✗ 有硬错，未落盘。按上面的 ✗ 逐条修改后重新运行 save。")
        return 2
    if args.dry_run:
        print("\n(--dry-run：门禁通过，未写文件)")
        return 0

    pdir = project_dir_of(xlsx, Path(args.out_root) if args.out_root else None, project)
    pdir.mkdir(parents=True, exist_ok=True)
    out = pdir / f"{safe_name(title)}.txt"
    if out.exists() and not args.overwrite:
        raise SystemExit(f"文件已存在：{out}\n要覆盖请加 --overwrite")
    nl = _newline(args)
    text = title + nl + nl + body.replace("\n", nl) + nl
    out.write_text(text, encoding="utf-8", newline="")
    chars = gates.prose_len(body)
    print(f"\n✓ 已保存：{out}\n  正文 {chars} 字，{len(gates.paragraphs(body))} 段")

    sci_file = getattr(args, "science_body_file", None)
    status = STATUS_DONE if sci_file else "已完成，待科普"     # 没配科普稿就先不算写完
    if xlsx and excel_row and not args.no_status:
        extra = ({"项目名称": project, "稿件文件": out.name, "稿件字数": chars,
                  "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M")}
                 if args.update_xlsx else {})
        err = write_tracking(Path(xlsx), excel_row, extra,
                             status=status, status_pos=args.status_pos)
        print(f"  表格状态已更新为「{status}」。" if not err else f"  ! {err}")

    if not sci_file:
        print("\n→ 本行还缺一篇中文域名科普稿：按 <编号>-科普材料.md 写好后，"
              "用 save --science-body-file 补上（或下次 extract 会再给一次话题）。")
        return 0

    print("\n=== 接着落盘同一行的中文域名科普稿 ===")
    if not getattr(args, "topic", ""):        # 没显式指定就取该行预留的话题
        args.topic = science.reserved_for(
            f"{safe_name(project, 60)}/{args.row if args.row is not None else ''}")
    rc = _save_science_core(args, str(sci_file), project, xlsx,
                            getattr(args, "science_title", "") or "")
    if rc == 0 and xlsx and excel_row and not args.no_status:
        err = write_tracking(Path(xlsx), excel_row, {}, status=STATUS_DONE,
                             status_pos=args.status_pos)
        print(f"  两篇齐了，表格状态更新为「{STATUS_DONE}」。" if not err else f"  ! {err}")
    return rc


def build_science_brief(topic: dict, pdir: Path, wdir: Path, tag: str,
                        lo: int, hi: int, target: int, xlsx: Path | None,
                        row, project: str) -> tuple[str, Path, Path]:
    """组装科普稿 brief。返回 (brief 文本, 草稿路径, 交付目录)。"""
    xlsx = Path(xlsx) if xlsx else None
    out_dir = pdir / "科普文章"
    draft = wdir / f"{tag}-科普草稿.txt"
    py = sys.executable
    script = Path(__file__).resolve()
    row_arg = f' --row "{row}"' if row is not None and str(row) != "" else ""
    xlsx_arg = f' --xlsx "{xlsx}"{row_arg}' if xlsx else ""
    brief = f"""# 中文域名科普稿材料（话题 {topic['id']}）

- 话题点：**{topic['title']}**
- 资料库：{science.MATERIAL}
- 落点：{topic['focus']}
- 交付目录：{out_dir}
- 字数要求：{target} 字左右（可接受 {lo}-{hi} 字）

## 一、这个话题要说的几个要点

{chr(10).join(f"- {p}" for p in topic["points"])}

## 二、资料库里的对应内容（话题取材处，可再往下读全文）

{science.excerpt(topic)}

## 三、写作要求

- **这是独立的中文域名科普文章**，与企业稿同时产出、但各自成篇，不要写成某家公司的宣传；
  可以举行业场景，不必出现具体企业名。
- **只宣扬中文域名的正面形象**：不出现任何关于中文域名的负面表述（争议、质疑、缺点、不值得、没必要、
  不建议使用等一律不写）；也不写“不会自动带来流量”“无法保证”这类免责式冷水。要说局限，就写成
  “企业可以怎样搭配使用”的建设性做法。
- **正文只允许出现 `.网址` 和 `.中国` 两种中文域名后缀**；`.公司/.网络/.商城/.商标/.在线/.网店` 等
  其他中文后缀一律不得出现。提到英文域名、`.com`/`.cn` 可以（讲清差异要用），但不贬低中文域名。
- 要有可核验的依据：政策文件、国家标准、ICANN/IANA 的公开说明、行业论坛等，引用只写
  「名称 + 发布单位 + 年份」，落笔前用 `news_draft.py search` 核实，核不到就换个说法。
- 不写没有出处的数字、份额、排名；不承诺流量、排名、销量。
- 结构建议：先点明这个话题里常见的片面看法，再用政策与技术依据正面回应，最后给企业可执行的做法。
- **两阶段处理**：先按 `references/human-writing.md` 写实，再按 `references/humanizer-zh.md` 清 AI 腔；
  破折号、提示性冒号、硬禁词、模型路标、翻案腔都会被门禁判硬错。
- 纯文本交付，不要 Markdown 标记。

## 四、写完怎么落盘

推荐与企业稿**一次落盘**（两篇一起过门禁）：

    "{py}" "{script}" save --kind company{xlsx_arg} --project "{project}" --topic {topic['id']} --body-file "<企业稿草稿路径>" --science-body-file "{draft}"

也可以单独落盘：

    "{py}" "{script}" save --kind science --topic {topic['id']}{xlsx_arg} --project "{project}" --body-file "{draft}"

门禁通过后会写成 {out_dir} 下的「首行标题.txt」，并把话题 {topic['id']} 标记为已完成。
"""
    return brief, draft, out_dir


def cmd_science_brief(args) -> int:
    """生成中文域名科普稿的 brief：话题从资料库挑，用过的跳过。"""
    xlsx = Path(args.xlsx) if getattr(args, "xlsx", None) else None
    project = args.project
    if xlsx and not project:
        rows, _ = read_rows(xlsx)
        r = pick_rows(rows, args.row)[0] if args.row is not None else rows[0]
        project = r["project"]
    project = project or "未命名项目"
    pdir = project_dir_of(xlsx, Path(args.out_root) if args.out_root else None, project)
    wdir = work_dir_of(pdir, args)
    wdir.mkdir(parents=True, exist_ok=True)

    row_key = f"{safe_name(project, 60)}/{args.row if getattr(args, 'row', None) is not None else ''}"
    topic = science.reserve(row_key, getattr(args, "topic", "") or "")
    brief, draft, out_dir = build_science_brief(topic, pdir, wdir, tag, lo, hi, target,
                                                xlsx, args.row, project)
    brief_md = wdir / f"{tag}-科普材料.md"
    brief_md.write_text(brief, encoding="utf-8")
    print(f"话题 {topic['id']}：{topic['title']}")
    print(f"材料与写作要求：{brief_md}")
    print(f"\n下一步：按这份 brief 写一篇中文域名科普稿，存成 {draft.name}，再执行里面的 save 命令。")
    return 0


def cmd_mark(args) -> int:
    """只更新表格状态列，不写稿：补记/回退已完成的行时用。"""
    xlsx = Path(args.xlsx)
    rows, _ = read_rows(xlsx)
    for r in pick_rows(rows, args.row):
        err = write_tracking(xlsx, r["excel_row"], {}, status=args.status,
                             status_pos=args.status_pos)
        line = f"[{r['num']}] {r['company']} → 状态「{args.status}」"
        print(line if not err else f"{line}\n  ! {err}")
    return 0


def cmd_save(args) -> int:
    if getattr(args, "kind", "company") == "science":
        return _save_science(args)
    company, domain, short, project, xlsx, excel_row, industry_mode = _context_from_args(args)
    return _report_and_save(args, company, domain, short, project, xlsx, excel_row,
                            industry_mode)


def _save_science(args) -> int:
    """单独落盘科普稿（save --kind science）。"""
    xlsx = Path(args.xlsx) if getattr(args, "xlsx", None) else None
    project = args.project
    if xlsx and not project:
        rows, _ = read_rows(xlsx)
        project = (pick_rows(rows, args.row)[0] if args.row is not None else rows[0])["project"]
    return _save_science_core(args, args.body_file, project or "未命名项目", xlsx,
                              args.title or "")


def _save_science_core(args, body_file, project: str, xlsx: Path | None,
                       title_override: str = "") -> int:
    """科普稿的落盘主体：过科普门禁、写进 <项目目录>/科普文章/、标记话题已完成。"""
    body_file = Path(body_file)
    if not body_file.exists():
        raise SystemExit(f"找不到稿件文件：{body_file}")
    raw = body_file.read_text(encoding="utf-8-sig", errors="replace")
    title, body = (split_title_body(raw, title_override) if title_override
                   else looks_like_deliverable(raw))
    if not title:
        raise SystemExit("无法确定标题：把标题写在稿件文件首行（其后空一行），或用 --title 指定。")

    pdir = project_dir_of(xlsx, Path(args.out_root) if args.out_root else None, project)
    out_dir = pdir / "科普文章"
    lo, hi, _ = resolve_length(args, SCIENCE_TARGET_CHARS)

    rep = gates.check_science(title, body, "",
                              min_chars=lo, max_chars=hi,
                              humanize_gate=not args.no_humanize_gate,
                              used_titles=used_titles(pdir, exclude_name=f"{safe_name(title)}.txt"))
    if not rep.ok or rep.warnings:
        print(f"科普门禁｜标题：{title}")
        print(rep.render())

    wdir = work_dir_of(pdir, args)
    wdir.mkdir(parents=True, exist_ok=True)
    tag = safe_name(getattr(args, "topic", "") or "科普", 20)
    (wdir / f"{tag}-科普门禁.json").write_text(json.dumps(
        {"kind": "science", "title": title, "ok": rep.ok, "errors": rep.errors,
         "warnings": rep.warnings, "stats": rep.stats,
         "checked_at": datetime.now().isoformat(timespec="seconds")},
        ensure_ascii=False, indent=2), encoding="utf-8")

    if not rep.ok:
        print("\n✗ 有硬错，未落盘。按上面的 ✗ 逐条修改后重新运行 save。")
        return 2
    if args.dry_run:
        print("\n(--dry-run：门禁通过，未写文件)")
        return 0

    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"{safe_name(title)}.txt"
    if out.exists() and not args.overwrite:
        raise SystemExit(f"文件已存在：{out}\n要覆盖请加 --overwrite")
    nl = _newline(args)
    out.write_text(title + nl + nl + body.replace("\n", nl) + nl,
                   encoding="utf-8", newline="")
    chars = gates.prose_len(body)
    print(f"\n✓ 已保存：{out}\n  正文 {chars} 字，{len(gates.paragraphs(body))} 段")
    if getattr(args, "topic", ""):
        science.mark_used(args.topic, title, str(out))
        print(f"  话题 {args.topic} 已标记为已用")
    return 0


def cmd_check(args) -> int:
    path = Path(args.file)
    if not path.exists():
        raise SystemExit(f"找不到文件：{path}")
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    if args.title:
        title, body = split_title_body(raw, args.title)
    else:
        title, body = looks_like_deliverable(raw)
    if args.xlsx:
        rows, _ = read_rows(Path(args.xlsx))
        r = (pick_rows(rows, args.row)[0] if args.row is not None
             else rows[0] if len(rows) == 1 else None)
        if r is None:
            raise SystemExit("表格有多行，请用 --row 指定编号")
        company, domain, short = r["company"], r["domain"], r["short"]
        industry_mode = not r["profile"]
    else:
        if not (args.company and args.domain):
            raise SystemExit("检查已有文件时，要么给 --xlsx（可加 --row），要么给 --company 与 --domain")
        company, domain, short = args.company, norm_domain(args.domain), args.short or ""
        industry_mode = False
    policy = args.suffix_policy
    if policy == "auto":
        policy = "zh-only" if domain.endswith(".网址") else "relaxed"
    lo, hi, _ = resolve_length(args)
    rep = gates.check_text(title, body, company, domain, short, policy,
                           min_chars=lo, max_chars=hi,
                           allow_registration=args.allow_registration_facts,
                           humanize_gate=not args.no_humanize_gate,
                           allow_any_zh_suffix=(policy == "any-zh"),
                           used_titles=used_titles(path.parent, exclude_name=path.name),
                           industry_mode=industry_mode)
    print(f"文件：{path}\n企业：{company}\n域名：{domain}\n标题：{title or '（未识别到标题行）'}")
    print(rep.render())
    return 0 if rep.ok else 2


def cmd_search(args) -> int:
    try:
        if getattr(args, "industry", None):
            text, errors, notes = search_industry(args.industry, args.max_results)
            if text:
                print(text)
            for n in notes:
                print(f"({n})", file=sys.stderr)
            for e in errors:
                print(f"(检索失败：{e})", file=sys.stderr)
            return 0 if text else 2
        if args.company:
            text, errors, notes = search_company(args.company, args.profile, args.max_results)
            if text:
                print(text)
            for n in notes:
                print(f"({n})", file=sys.stderr)
            for e in errors:
                print(f"(检索失败：{e})", file=sys.stderr)
            return 0 if text else 2
        from anysearch_search import search
        print(search(args.query, args.max_results))
        return 0
    except AnySearchError as exc:
        print(f"(检索不可用：{exc})", file=sys.stderr)
        return 2


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="xlsx → 中文域名新闻稿 .txt")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("list", help="列出表格里的稿件任务")
    p.add_argument("--xlsx", required=True)
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("extract", help="取企业名称/域名/企业介绍，缺企业介绍时联网检索，产出 brief")
    p.add_argument("--xlsx", required=True)
    p.add_argument("--row", help="编号（或 Excel 行号）；不给则处理全部行")
    p.add_argument("--project", help="覆盖项目名称（默认取表格“项目名称”列，再默认取 xlsx 文件名）")
    p.add_argument("--out-root", "--out_root", help="项目文件夹的上级目录（默认为 xlsx 所在目录）")
    p.add_argument("--search", choices=("auto", "always", "never"), default="auto",
                   help="auto：仅当企业介绍为空时联网检索（默认）")
    p.add_argument("--max-results", "--max_results", type=int, default=5)
    p.add_argument("--target-chars", "--target_chars", type=int, default=None,
                   help=f"目标字数（默认 {gates.DEFAULT_TARGET_CHARS} 字左右："
                        f"{round(gates.DEFAULT_TARGET_CHARS * gates.TARGET_MIN_RATIO)}-"
                        f"{round(gates.DEFAULT_TARGET_CHARS * gates.TARGET_MAX_RATIO)} 字）；"
                        "如“800 字左右”就写 --target-chars 800")
    p.add_argument("--min-chars", "--min_chars", type=int, default=None,
                   help="正文硬下限；给了它或 --max-chars 就按显式区间走")
    p.add_argument("--max-chars", "--max_chars", type=int, default=None,
                   help="正文上限，只提醒不拦截（只给一个时另一边不限）")
    p.add_argument("--work-dir", "--work_dir",
                   help="中间过程文件目录（默认系统临时目录，交付目录只留最终 .txt）")
    p.add_argument("--keep-work", "--keep_work", action="store_true",
                   help="把过程文件放回 <项目目录>/过程文件")
    p.add_argument("--force", action="store_true",
                   help="连状态为「已完成」的行一起重做（默认跳过）")
    p.add_argument("--no-science", "--no_science", action="store_true",
                   help="只出企业稿 brief，不配科普稿")
    p.add_argument("--topic", default="",
                   help="指定科普话题编号（如 S05）；默认给这一行预占一个没用过的")
    p.set_defaults(func=cmd_extract)

    p = sub.add_parser("science-brief", help="生成中文域名科普稿的 brief（话题从资料库挑）")
    p.add_argument("--xlsx", help="可选：用表格某一行的项目名")
    p.add_argument("--row")
    p.add_argument("--project", help="项目名（决定交付目录）")
    p.add_argument("--out-root", "--out_root")
    p.add_argument("--topic", default="", help="指定话题编号，如 S05；不给则取未用过的")
    p.add_argument("--target-chars", "--target_chars", type=int, default=None,
                   help=f"目标字数（默认 {SCIENCE_TARGET_CHARS} 字左右）")
    p.add_argument("--min-chars", "--min_chars", type=int, default=None)
    p.add_argument("--max-chars", "--max_chars", type=int, default=None)
    p.add_argument("--work-dir", "--work_dir")
    p.add_argument("--keep-work", "--keep_work", action="store_true")
    p.set_defaults(func=cmd_science_brief)

    p = sub.add_parser("mark", help="只更新表格「状态」列（补记/回退，不写稿）")
    p.add_argument("--xlsx", required=True)
    p.add_argument("--row", help="编号（或 Excel 行号）；不给则全部行")
    p.add_argument("--status", default=STATUS_DONE, help=f"默认 {STATUS_DONE}")
    p.add_argument("--status-pos", "--status_pos", choices=("after-num", "end"),
                   default="after-num", help="状态列不存在时的插入位置")
    p.set_defaults(func=cmd_mark)

    p = sub.add_parser("save", help="跑门禁并落盘为 .txt（企业稿或科普稿）")
    p.add_argument("--kind", choices=("company", "science"), default="company",
                   help="company=企业稿（默认，写到项目目录）；science=科普稿（写到 <项目目录>/科普文章/）")
    p.add_argument("--topic", default="", help="科普稿用：话题编号，落盘后标记为已完成")
    p.add_argument("--science-body-file", "--science_body_file",
                   help="同一行的科普稿草稿：给了它，一次 save 就把企业稿和科普稿都落盘")
    p.add_argument("--science-title", "--science_title", default="",
                   help="科普稿标题（不给则取科普草稿首行）")
    p.add_argument("--xlsx")
    p.add_argument("--row")
    p.add_argument("--company")
    p.add_argument("--domain")
    p.add_argument("--short", default="")
    p.add_argument("--project")
    p.add_argument("--out-root", "--out_root")
    p.add_argument("--title", default="",
                   help="标题；不给则取稿件文件首行（其后空一行）")
    p.add_argument("--body-file", "--body_file", required=True)
    p.add_argument("--suffix-policy", "--suffix_policy",
                   choices=("auto", "zh-only", "relaxed", "any-zh"), default="auto",
                   help="默认只允许 .网址/.中国；any-zh 放开中文后缀限制（不推荐）")
    p.add_argument("--no-humanize-gate", "--no_humanize_gate", action="store_true",
                   help="跳过人化门禁（默认按 human-writing/humanizer-zh 硬规则审核）")
    p.add_argument("--target-chars", "--target_chars", type=int, default=None,
                   help=f"目标字数（默认 {gates.DEFAULT_TARGET_CHARS} 字左右："
                        f"{round(gates.DEFAULT_TARGET_CHARS * gates.TARGET_MIN_RATIO)}-"
                        f"{round(gates.DEFAULT_TARGET_CHARS * gates.TARGET_MAX_RATIO)} 字）；"
                        "如“800 字左右”就写 --target-chars 800")
    p.add_argument("--min-chars", "--min_chars", type=int, default=None,
                   help="正文硬下限；给了它或 --max-chars 就按显式区间走")
    p.add_argument("--max-chars", "--max_chars", type=int, default=None,
                   help="正文上限，只提醒不拦截（只给一个时另一边不限）")
    p.add_argument("--work-dir", "--work_dir",
                   help="中间过程文件目录（默认系统临时目录，交付目录只留最终 .txt）")
    p.add_argument("--keep-work", "--keep_work", action="store_true",
                   help="把过程文件放回 <项目目录>/过程文件")
    p.add_argument("--allow-registration-facts", "--allow_registration_facts",
                   action="store_true",
                   help="允许写工商登记类信息（默认禁止：企业资料只能用业务相关内容）")
    p.add_argument("--lf", action="store_true",
                   help="交付文本用 LF 换行（默认 CRLF，照顾 Windows 记事本）")
    p.add_argument("--overwrite", action="store_true")
    p.add_argument("--update-xlsx", "--update_xlsx", action="store_true",
                   help="除状态外，把项目名称/稿件文件/字数/生成时间也回写表格")
    p.add_argument("--no-status", "--no_status", action="store_true",
                   help="不碰表格（默认写完会把状态列更新为「已完成」）")
    p.add_argument("--status-pos", "--status_pos", choices=("after-num", "end"),
                   default="after-num",
                   help="状态列位置：after-num 插在「编号」列后（默认）｜end 追加到表尾")
    p.add_argument("--dry-run", "--dry_run", action="store_true")
    p.set_defaults(func=cmd_save)

    p = sub.add_parser("check", help="只跑门禁")
    p.add_argument("--file", required=True)
    p.add_argument("--xlsx")
    p.add_argument("--row")
    p.add_argument("--company")
    p.add_argument("--domain")
    p.add_argument("--short", default="")
    p.add_argument("--title", default="")
    p.add_argument("--suffix-policy", "--suffix_policy",
                   choices=("auto", "zh-only", "relaxed", "any-zh"), default="auto",
                   help="默认只允许 .网址/.中国；any-zh 放开中文后缀限制（不推荐）")
    p.add_argument("--no-humanize-gate", "--no_humanize_gate", action="store_true",
                   help="跳过人化门禁（默认按 human-writing/humanizer-zh 硬规则审核）")
    p.add_argument("--target-chars", "--target_chars", type=int, default=None,
                   help=f"目标字数（默认 {gates.DEFAULT_TARGET_CHARS} 字左右："
                        f"{round(gates.DEFAULT_TARGET_CHARS * gates.TARGET_MIN_RATIO)}-"
                        f"{round(gates.DEFAULT_TARGET_CHARS * gates.TARGET_MAX_RATIO)} 字）；"
                        "如“800 字左右”就写 --target-chars 800")
    p.add_argument("--min-chars", "--min_chars", type=int, default=None,
                   help="正文硬下限；给了它或 --max-chars 就按显式区间走")
    p.add_argument("--max-chars", "--max_chars", type=int, default=None,
                   help="正文上限，只提醒不拦截（只给一个时另一边不限）")
    p.add_argument("--allow-registration-facts", "--allow_registration_facts",
                   action="store_true",
                   help="允许写工商登记类信息（默认禁止）")
    p.set_defaults(func=cmd_check)

    p = sub.add_parser("search", help="临时联网检索（AnySearch）")
    p.add_argument("query", nargs="?")
    p.add_argument("--company")
    p.add_argument("--industry", help="行业名：查行业现状与国家政策素材")
    p.add_argument("--profile", default="")
    p.add_argument("--max-results", "--max_results", type=int, default=5)
    p.set_defaults(func=cmd_search)
    return ap


def main() -> int:
    _utf8_stdout()
    args = build_parser().parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
