#!/usr/bin/env python3
"""AnySearch 接入层：把联网检索包成一个函数，供 news_draft.py 和 agent 使用。

优先复用已安装的 AnySearch 官方 skill（~/.dsh/skills/anysearch/scripts/anysearch_cli.py），
找不到就按 Node.js 版、PowerShell 版依次回退；全都不可用时返回明确提示，
由 agent 改用自带的 web_search 工具补齐材料。

环境变量：
  NEWS_ANYSEARCH_CMD  自定义检索命令（argv 字符串，优先于一切探测）
  ANYSEARCH_CMD       兼容 media-publish-skill 的旧变量
  ANYSEARCH_SKILL_DIR 指定 anysearch skill 目录
  NEWS_NODE           node 可执行文件路径（不走 PATH 时用）
"""
from __future__ import annotations

import html
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
from functools import lru_cache
from pathlib import Path

SKILL_HOME = Path(__file__).resolve().parent.parent
CONFIG = SKILL_HOME / "config.json"


class AnySearchError(RuntimeError):
    pass


def _read_config() -> dict:
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8")) or {}
    except Exception:  # noqa: BLE001
        return {}


def _skill_roots() -> list[Path]:
    """各 agent harness 放 skill 的常见目录（按存在与否过滤）。"""
    home = Path.home()
    roots = [
        Path(os.environ.get("NEWS_SKILL_ROOTS", "").split(os.pathsep)[0]) if os.environ.get("NEWS_SKILL_ROOTS") else None,
        home / ".dsh" / "skills",
        home / ".claude" / "skills",
        home / ".opencode" / "skills",
        home / ".openclaw" / "skills",
        home / ".cursor" / "skills",
        home / ".windsurf" / "skills",
        home / ".codex" / "skills",
        home / ".agents" / "skills",
        Path.cwd() / ".skills",
        Path.cwd() / ".claude" / "skills",
    ]
    extra = os.environ.get("NEWS_SKILL_ROOTS", "")
    if extra:
        roots += [Path(x) for x in extra.split(os.pathsep) if x.strip()]
    out, seen = [], set()
    for r in roots:
        if r is None:
            continue
        try:
            rp = r.expanduser().resolve()
        except OSError:
            continue
        if rp in seen or not rp.is_dir():
            continue
        seen.add(rp)
        out.append(rp)
    return out


def _anysearch_dir() -> Path | None:
    """AnySearch 官方 skill 目录：环境变量 > config.json > 各平台常见位置 > 与本 skill 同级。"""
    for raw in (os.environ.get("ANYSEARCH_SKILL_DIR", ""),
                str(_read_config().get("anysearch_skill_dir", ""))):
        if raw.strip():
            p = Path(raw.strip()).expanduser()
            if p.is_dir():
                return p
    for root in _skill_roots():
        for name in ("anysearch", "anysearch-skill"):
            cand = root / name
            if cand.is_dir():
                return cand
    sibling = SKILL_HOME.parent / "anysearch"
    return sibling if sibling.is_dir() else None


def _node() -> str | None:
    """node 可执行文件：环境变量 > PATH > 常见安装位置（DSH 自带路径只作兜底）。"""
    raw = os.environ.get("NEWS_NODE", "")
    if raw and Path(raw).exists():
        return raw
    found = shutil.which("node")
    if found:
        return found
    cands = [
        Path("/usr/local/bin/node"), Path("/usr/bin/node"), Path("/opt/homebrew/bin/node"),
        Path(r"C:\Program Files\nodejs\node.exe"),
        Path.home() / "AppData" / "Roaming" / "npm" / "node.exe",
        (Path.home() / ".dsh" / "dsh-runtimes" / "dsh-primary-runtime"
         / "dependencies" / "node" / "bin" / "node.exe"),
    ]
    for c in cands:
        if c.exists():
            return str(c)
    return None


def _has_requests() -> bool:
    import importlib.util
    return importlib.util.find_spec("requests") is not None


@lru_cache(maxsize=8)
def _exe_has_requests(exe: str) -> bool:
    """探测某个 python 解释器有没有 requests（结果缓存，最多探 8 个）。"""
    try:
        r = subprocess.run([exe, "-c", "import requests"], capture_output=True, timeout=30)
        return r.returncode == 0
    except Exception:  # noqa: BLE001
        return False


def _python_cmd() -> str | None:
    """挑一个能跑官方 Python CLI 的解释器（需要 requests）。"""
    cands: list[str] = []
    if sys.executable:
        cands.append(sys.executable)
    for name in ("python3", "python", "py"):
        found = shutil.which(name)
        if found and found not in cands:
            cands.append(found)
    for exe in cands:
        if exe == sys.executable:
            if _has_requests():
                return exe
        elif _exe_has_requests(exe):
            return exe
    return None


def _posix_shell_ready() -> bool:
    """官方 bash CLI 需要 bash + jq + curl。"""
    return all(shutil.which(x) for x in ("bash", "jq", "curl"))


def _runtime_conf_command(base: Path) -> list[str] | None:
    """官方 skill 可能生成 runtime.conf，里面有现成的 Command 行。"""
    conf = base / "runtime.conf"
    if not conf.exists():
        return None
    try:
        text = conf.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    for line in text.splitlines():
        if re.match(r"\s*(Command|command|cmd)\s*[:=]", line):
            value = re.split(r"[:=]", line, maxsplit=1)[1].strip().strip('"').strip("'")
            if value:
                return shlex.split(value)
    return None


@lru_cache(maxsize=1)
def resolve_cli() -> tuple[str, ...] | None:
    """返回检索命令的 argv 前缀；探测不到返回 None（结果缓存，进程内只探一次）。

    优先级：显式命令 > 官方 skill 的 runtime.conf > Python CLI（需 requests）
    > Node CLI（零依赖）> Windows PowerShell CLI > POSIX bash CLI（需 jq+curl）。"""
    raw = (os.environ.get("NEWS_ANYSEARCH_CMD", "")
           or os.environ.get("ANYSEARCH_CMD", "")
           or str(_read_config().get("anysearch_cmd", ""))).strip()
    if raw:
        return tuple(shlex.split(raw))
    base = _anysearch_dir()
    if base is None:
        return None
    conf = _runtime_conf_command(base)
    if conf:
        return tuple(conf)
    scripts = base / "scripts"
    py = scripts / "anysearch_cli.py"
    if py.exists():
        exe = _python_cmd()
        if exe:
            return (exe, str(py))
    js = scripts / "anysearch_cli.js"
    node = _node()
    if js.exists() and node:
        return (node, str(js))
    ps1 = scripts / "anysearch_cli.ps1"
    if ps1.exists() and (shutil.which("powershell") or shutil.which("pwsh")):
        shell = shutil.which("powershell") or shutil.which("pwsh")
        return (shell, "-ExecutionPolicy", "Bypass", "-File", str(ps1))
    sh = scripts / "anysearch_cli.sh"
    if sh.exists() and _posix_shell_ready():
        return ("bash", str(sh))
    return None


# 工商数据库来源：这类页面整条丢弃（它们的头条信息是工商登记，不是业务资料）
REGISTRY_SOURCE = re.compile(
    r"qcc\.com|aiqicha\.baidu\.com|tianyancha\.com|qixin\.com|shuidi\.cn|"
    r"gsxt\.gov\.cn|qichacha|aiqicha|tianyancha|启信宝|水滴信用|企查查|爱企查|天眼查|"
    r"国家企业信用信息公示系统", re.I)
# 工商登记类字段：就地屏蔽，避免写稿时顺手用上
REGISTRY_FACTS = (
    re.compile(r"注册资本[^，。；\n]{0,24}"),
    re.compile(r"实缴资本[^，。；\n]{0,24}"),
    re.compile(r"法定代表[^，。；\n]{0,16}"),
    re.compile(r"法人代表[^，。；\n]{0,16}"),
    re.compile(r"统一社会信用代码[^，。；\n]{0,26}"),
    re.compile(r"注册(?:地址|地|住所)[^，。；\n]{0,30}"),
    re.compile(r"(?:股东|股权结构|持股比例)[^，。；\n]{0,24}"),
    re.compile(r"经营状态[^，。；\n]{0,12}"),
    re.compile(r"(?:成立于|成立日期|成立时间|注册时间|注册日期)[^，。；\n]{0,18}"),
    re.compile(r"企业类型[^，。；\n]{0,14}"),
)
REGISTRY_MASK = "〔工商登记类信息，按规则已剔除〕"


def filter_business_material(text: str) -> tuple[str, dict]:
    """只留企业业务相关的资料：丢掉工商数据库来源，屏蔽工商登记类字段。"""
    blocks, cur = [], []
    for line in (text or "").splitlines():
        if line.startswith("### ") and cur:
            blocks.append(cur)
            cur = [line]
        else:
            cur.append(line)
    if cur:
        blocks.append(cur)
    kept, dropped, masked = [], 0, 0
    for b in blocks:
        blob = "\n".join(b)
        if REGISTRY_SOURCE.search(blob):
            dropped += 1
            continue
        for pat in REGISTRY_FACTS:
            blob, n = pat.subn(REGISTRY_MASK, blob)
            masked += n
        kept.append(blob)
    return "\n".join(kept).strip(), {"dropped": dropped, "masked": masked}


def clean(text: str) -> str:
    """洗检索结果：去 HTML、去加粗标记、压空行；保留 URL 供写作时核对来源。"""
    text = re.sub(r"<script[\s\S]*?</script>", "", text, flags=re.I)
    text = re.sub(r"<[^>]{1,200}>", "", text)
    text = html.unescape(text)
    text = text.replace("\ufffd", "").replace("**", "")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def search(query: str, max_results: int = 5, timeout: int = 90) -> str:
    """一次 AnySearch 检索，返回 Markdown 文本。命令不可用/失败时抛 AnySearchError。"""
    cli = resolve_cli()
    if not cli:
        raise AnySearchError(
            "找不到 AnySearch 检索命令。把官方 skill 装到本 harness 的 skill 目录（如 "
            "~/.dsh/skills、~/.claude/skills、~/.opencode/skills，或用 ANYSEARCH_SKILL_DIR 指定），"
            "或用 NEWS_ANYSEARCH_CMD 指定命令，或改用本 harness 自带的联网检索工具。")
    argv = list(cli) + ["search", query, "--max_results", str(max(1, min(10, max_results)))]
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    try:
        proc = subprocess.run(argv, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=timeout, env=env)
    except Exception as exc:  # noqa: BLE001
        raise AnySearchError(f"检索命令执行失败：{type(exc).__name__}: {exc}") from exc
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()[-300:]
        raise AnySearchError(f"检索失败（exit {proc.returncode}）：{detail}")
    out = clean(proc.stdout)
    if not out:
        raise AnySearchError("检索返回空结果")
    return out


def build_queries(company: str, profile: str = "") -> list[str]:
    """最多三轮：先用表格资料派生定向查询，再用公司名兜底。"""
    company = (company or "").strip()
    queries: list[str] = []
    snippet = re.sub(r"\s+", "", (profile or "").strip())[:40]
    if snippet:
        queries.append(f"{company} {snippet}")
    queries += [f"{company} 介绍", f"{company} 主营业务"]
    seen, out = set(), []
    for q in queries:
        if q and q not in seen:
            seen.add(q)
            out.append(q)
    return out[:3]


def search_company(company: str, profile: str = "", max_results: int = 5,
                   timeout: int = 90) -> tuple[str, list[str], list[str]]:
    """按要求联网补企业资料。返回 (汇总文本, 错误列表, 提示列表)。

    结果一律过 filter_business_material：只保留业务相关资料。"""
    chunks, errors, notes = [], [], []
    dropped = masked = 0
    for q in build_queries(company, profile):
        try:
            filtered, stats = filter_business_material(search(q, max_results, timeout))
            dropped += stats["dropped"]
            masked += stats["masked"]
            chunks.append(f"【查询：{q}】\n{filtered}")
        except AnySearchError as exc:
            errors.append(f"{q} → {exc}")
    if dropped or masked:
        notes.append(f"已按规则过滤检索材料：丢弃工商数据库来源 {dropped} 条，"
                     f"屏蔽工商登记类字段 {masked} 处（只留业务相关资料）")
    blob = "\n".join(chunks)
    # 判断“有没有命中企业全称”时必须去掉查询词表头，否则查询词里的企业名会让自己永远通过
    results_only = re.sub(r"【查询：[^】]*】", "", blob)
    if chunks and company and company.replace(" ", "") not in results_only.replace(" ", ""):
        notes.append(f"⚠ 检索结果里没有任何一条命中企业全称「{company}」。"
                     "按“名字相近但不是同一家就整条丢掉”的规则，这些材料都不能当成本企业的事实；"
                     "企业自身情况宁可不写或写少，行业与中文域名部分照常写。"
                     "需要企业实况时，改用本 harness 自带的联网检索工具再查，"
                     "或让用户补「企业信息资料」。")
    return blob, errors, notes


def search_industry(industry_name: str, max_results: int = 5,
                    timeout: int = 90) -> tuple[str, list[str], list[str]]:
    """给没有企业资料的稿件补「行业现状 + 国家政策」素材。"""
    chunks, errors, notes = [], [], []
    dropped = masked = 0
    for q in (f"{industry_name} 行业现状 发展趋势",
              f"{industry_name} 国家政策 规划 数字化转型"):
        try:
            filtered, stats = filter_business_material(search(q, max_results, timeout))
            dropped += stats["dropped"]
            masked += stats["masked"]
            chunks.append(f"【查询：{q}】\n{filtered}")
        except AnySearchError as exc:
            errors.append(f"{q} → {exc}")
    if dropped or masked:
        notes.append(f"行业材料已过滤：丢弃工商数据库来源 {dropped} 条，屏蔽工商登记类字段 {masked} 处")
    return "\n\n".join(chunks), errors, notes


def main() -> int:
    import argparse
    ap = argparse.ArgumentParser(description="AnySearch 检索（新闻稿材料补充用）")
    ap.add_argument("query", nargs="?", help="检索词")
    ap.add_argument("--company", help="企业名称：按其介绍/主营业务自动检索三轮")
    ap.add_argument("--industry", help="行业名：检索行业现状与国家政策素材")
    ap.add_argument("--profile", default="", help="企业介绍片段，用于派生定向查询")
    ap.add_argument("--max_results", type=int, default=5)
    ap.add_argument("--show-cmd", action="store_true", help="只打印探测到的命令")
    args = ap.parse_args()

    if args.show_cmd:
        cli = resolve_cli()
        print(" ".join(cli) if cli else "(未找到检索命令)")
        return 0 if cli else 2

    if args.industry:
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
    if not args.query:
        ap.error("需要 query 或 --company")
    try:
        print(search(args.query, args.max_results))
    except AnySearchError as exc:
        print(f"(检索失败：{exc})", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
