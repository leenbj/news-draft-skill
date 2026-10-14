#!/usr/bin/env python3
"""中文域名科普文章：话题选自 `references/中文域名科普资料库.md`。

规则（用户要求）：
- 每完成一篇企业稿，配一篇独立的中文域名科普文；
- 话题必须从资料库挑选，且**用过的就不再重复**（状态存在 <skill>/.state/science_topics.json）；
- 只宣扬中文域名的正面形象，不出现任何负面内容；
- 正文只允许出现 `.网址` 和 `.中国` 两种中文域名后缀。

CLI：
  python science.py --list                 看话题清单与使用状态
  python science.py --pick                 取下一个未用过的话题
  python science.py --pick --topic S05     指定话题
  python science.py --mark S05             标记话题已用
  python science.py --excerpt S05          打印该话题对应的资料库摘录
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import datetime
from functools import lru_cache
from pathlib import Path

SKILL_HOME = Path(__file__).resolve().parent.parent
CONFIG = SKILL_HOME / "config.json"
MATERIAL = SKILL_HOME / "references" / "中文域名科普资料库.md"


def _read_config() -> dict:
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8")) or {}
    except Exception:  # noqa: BLE001
        return {}


@lru_cache(maxsize=1)
def state_file() -> Path:
    """话题状态文件的位置，按可写性依次挑：

    NEWS_STATE_DIR / config.json 的 state_dir → 本 skill 的 .state/ →
    用户缓存目录（agent 把 skills 目录挂成只读时用这里）。
    """
    cands: list[Path] = []
    raw = (os.environ.get("NEWS_STATE_DIR", "")
           or str(_read_config().get("state_dir", ""))).strip()
    if raw:
        cands.append(Path(raw).expanduser())
    cands.append(SKILL_HOME / ".state")
    if os.name == "nt":
        cands.append(Path(os.environ.get("LOCALAPPDATA", Path.home())) / "news-draft-skill")
    else:
        cands.append(Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache"))
                     / "news-draft-skill")
    for d in cands:
        try:
            d.mkdir(parents=True, exist_ok=True)
            probe = d / ".write-test"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink()
            return d / "science_topics.json"
        except OSError:
            continue
    return cands[-1] / "science_topics.json"

# 已核实过出处的依据文件（稿件里引用这一批不会触发「先核实」提醒）
POLICY_DOCS = (
    "互联网域名管理办法",
    "中华人民共和国商标法",
    "2026年IDN实施与普遍适用性采用报告",
)


def policy_whitelist() -> tuple[str, ...]:
    return POLICY_DOCS

# 话题清单：全部来自资料库，title 是话题点，match 用来在资料库里定位对应段落。
TOPICS = (
    {"id": "S01", "title": "中小企业为什么比知名品牌更需要中文域名",
     "match": "中小企业更可以借助中文域名主动建设品牌",
     "points": ("知名企业已有品牌认知，中小企业首先要解决让消费者认识并记住品牌",
                "把与商标一致的中文域名用在官网、包装、门店和广告上，可形成统一的品牌印象",
                "越在品牌创立阶段，越有条件统一规划商标、品牌名称和互联网入口"),
     "focus": "落点放在主动建设品牌，而不是被动防御"},
    {"id": "S02", "title": "传统英文域名资源日趋成熟，中小企业该怎样重新选择域名",
     "match": "全球顶级域名注册总量达到4.016亿个",
     "points": ("DNIB 2026 年第二季度报告显示，全球域名注册总量约 4.016 亿个",
                "优质短域名大多已被注册，企业要么高价买二手，要么用冗长组合",
                "与品牌名称一致的中文域名让企业不必迁就复杂网址"),
     "focus": "落点是域名为品牌服务，品牌不必迁就网址"},
    {"id": "S03", "title": "中文域名的政策依据：域名管理办法怎么规定的",
     "match": "中文域名是中国互联网域名体系的重要组成部分",
     "points": ("《互联网域名管理办法》第一条把推动中文域名发展应用列为立法目的之一",
                "第六条明确中文域名是中国互联网域名体系的重要组成部分，国家鼓励技术研究和推广应用",
                "《互联网域名管理办法》第二十六条确立先申请先注册原则"),
     "focus": "落点是中文域名有明确的制度基础"},
    {"id": "S04", "title": "中文域名的国际技术基础：ICANN 与国际化域名体系",
     "match": "151个国际化顶级域名",
     "points": ("ICANN《2026年IDN实施与普遍适用性采用报告》显示，全球已委派 151 个国际化顶级域名",
                "覆盖 37 种语言、23 种文字体系，二级国际化域名注册量约 430 万个",
                "“.网址”是 IANA 根区数据库中已正式委派的通用顶级域名"),
     "focus": "落点是中文域名是国际互联网多语言发展的正式组成部分"},
    {"id": "S05", "title": "从防御性注册走向主动品牌传播",
     "match": "不应只作为被动保护资源",
     "points": ("防御性注册只是功能之一，不是全部价值",
                "中文域名可以作为面向消费者的品牌入口，出现在菜单、包装、海报、名片和视频广告上",
                "注册后长期闲置，等于没有发挥域名的使用价值"),
     "focus": "落点是注册之后要投入使用"},
    {"id": "S06", "title": "中文域名可以直接作为官网入口",
     "match": "中文域名可以直接参与官网建设",
     "points": ("国际化域名已纳入全球互联网域名体系，不存在不能作为正式网站地址的技术禁令",
                "301 重定向只是一种地址管理方式，不是中文域名必须遵守的规则",
                "企业也可以采用双域名运营，按场景选择规范网址"),
     "focus": "落点是按业务需要决定入口形式"},
    {"id": "S07", "title": "中文域名让品牌名称和网络入口保持一致",
     "match": "保持品牌名称和对外网址的一致性",
     "points": ("商标、包装、广告和官网网址使用同一个中文品牌名称",
                "减少在中文名称与拼音域名之间建立额外对应关系的需要",
                "对中文消费者来说，读得出来、写得出来的入口更容易记住"),
     "focus": "落点是命名一致性本身就是品牌设计"},
    {"id": "S08", "title": "中文域名在门店、包装和视频广告里的展示价值",
     "match": "不需要额外理解一组较长的拼音或英文字符",
     "points": ("消费者看到网址就能识别品牌名称，不必记忆字符顺序",
                "即使最终通过二维码或搜索进入官网，中文域名仍在品牌识别上发挥作用",
                "不同行业的用法不同，餐饮、制造、地方服务各有展示场景"),
     "focus": "落点是传播场景里的直观性"},
    {"id": "S09", "title": "中文域名与商标的协同管理",
     "match": "商标与域名的同步注册、同步保护、同步管理",
     "points": ("行业论坛建议企业实现商标与域名的同步注册、同步保护、同步管理",
                "两个体系不同，但都服务于品牌的持续使用",
                "协同规划可以降低因管理脱节带来的品牌资产风险"),
     "focus": "落点是品牌资产统一管理，不越界谈确权结论"},
    {"id": "S10", "title": "核心域名为什么值得考虑一次注册十年",
     "match": "域名通常可以注册1至10年",
     "points": ("ICANN 说明域名可注册 1 至 10 年，未及时续费可能失去域名",
                "人员变动、付款账户变化、年度预算安排不及时都会带来管理风险",
                "《商标法》规定注册商标有效期为十年，可参照这一周期管理核心域名"),
     "focus": "落点是降低管理风险、提高连续性，不说成投资升值"},
    {"id": "S11", "title": "靠平台获客的企业，同样需要独立的品牌入口",
     "match": "平台账号不等于企业全部的数字品牌资产",
     "points": ("第三方平台解决流量，企业域名承担跨平台品牌识别与官方信息展示",
                "两者互补，不冲突",
                "渠道调整时，统一的品牌域名能保持连续性"),
     "focus": "落点是互补关系，不贬低平台"},
    {"id": "S12", "title": "外贸企业怎样用中英文双域名服务不同市场",
     "match": "中英文双域名策略",
     "points": ("面向不同市场的客户，域名语言习惯不同",
                "同时服务国内供应商、经销商或华语客户的出口企业，可以用双域名分工",
                "合理配置网站结构与多语言内容即可兼顾"),
     "focus": "落点是按目标市场配置，不下绝对结论"},
    {"id": "S13", "title": "域名投资和企业品牌建设不是一回事",
     "match": "并不是为了将来转卖",
     "points": ("企业注册与自身品牌一致的域名，目的是长期使用和维护入口",
                "商标注册也是为了保护和经营品牌，而不是等待转售",
                "评价标准是能否支撑未来十年的品牌建设"),
     "focus": "落点是长期经营视角"},
    {"id": "S14", "title": "先申请先注册：品牌起步阶段就要布局中文域名",
     "match": "先申请先注册",
     "points": ("《互联网域名管理办法》第二十六条确立先申请先注册原则",
                "理想域名不会因为企业还在创业阶段就自动保留",
                "品牌创建阶段同步做商标检索、域名查询与数字品牌规划，更主动"),
     "focus": "落点是提前规划，不渲染抢注恐慌"},
    {"id": "S15", "title": "怎样评估中文域名的投入：看长期品牌价值而不是只看年费",
     "match": "更应关注它能否成为企业长期稳定使用的品牌资产",
     "points": ("不同后缀差异化定价是域名行业的正常现象",
                "需要比较的是品牌匹配度、获取成本、长期续费与传播价值",
                "核心品牌域名属于长期品牌建设投入"),
     "focus": "落点是评价口径，不纠缠具体价格数字"},
    {"id": "S16", "title": "中文域名在人工智能与数字品牌建设中的探索方向",
     "match": "中文域名与人工智能搜索、品牌实体识别相结合",
     "points": ("行业论坛把中文域名与数字品牌建设、人工智能应用放在一起讨论",
                "中文域名语义直观、文化标识鲜明，便于品牌实体表达",
                "这些属于行业探索方向，需要持续完善应用环境"),
     "focus": "落点是探索方向，不声称平台给了特殊权重"},
)


def _state() -> dict:
    """状态文件：{"done": [...], "reserved": {行标识: 话题编号}}。

    兼容最早的 {"used": [...]} 写法（老文件里的记录当作已完成）。"""
    try:
        data = json.loads(state_file().read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {"done": [], "reserved": {}}
    if "done" not in data:
        data = {"done": data.get("used", []), "reserved": {}}
    data.setdefault("reserved", {})
    return data


def _save_state(data: dict) -> None:
    path = state_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def used_ids() -> list[str]:
    """已完成 + 已预占的话题编号（预占是给某一行留的，别的行不能再挑）。"""
    data = _state()
    ids = [u.get("id", "") for u in data.get("done", [])]
    ids += [v.get("id", "") for v in data.get("reserved", {}).values()]
    return ids


def reserved_for(row_key: str) -> str:
    return (_state().get("reserved", {}).get(row_key) or {}).get("id", "")


def reserve(row_key: str, topic_id: str = "") -> dict:
    """给某一行预留一个话题：该行已预留就复用，否则挑一个没被占用的并记下来。"""
    data = _state()
    data.setdefault("reserved", {})
    cur = reserved_for(row_key)
    if cur and not topic_id:
        topic_id = cur
    topic = pick_topic(topic_id)
    data["reserved"][row_key] = {"id": topic["id"],
                                 "at": datetime.now().isoformat(timespec="seconds")}
    _save_state(data)
    return topic


def release(key: str) -> str:
    """释放预留：key 可以是行标识或话题编号。"""
    data = _state()
    data.setdefault("reserved", {})
    hit = ""
    if key in data["reserved"]:
        hit = data["reserved"].pop(key).get("id", "")
    else:
        for rk, v in list(data["reserved"].items()):
            if v.get("id", "").upper() == (key or "").upper():
                data["reserved"].pop(rk)
                hit = v.get("id", "")
    _save_state(data)
    return hit


def get_topic(topic_id: str) -> dict | None:
    for t in TOPICS:
        if t["id"].upper() == (topic_id or "").upper():
            return t
    return None


def pick_topic(topic_id: str = "") -> dict:
    """取话题：指定优先；否则按顺序取第一个没用过的；全用过就取用得最早的那个。"""
    if topic_id:
        t = get_topic(topic_id)
        if not t:
            raise SystemExit(f"没有这个编号：{topic_id}；现有：{', '.join(t['id'] for t in TOPICS)}")
        return t
    used = used_ids()
    for t in TOPICS:
        if t["id"] not in used:
            return t
    state = _state()
    done = sorted(state.get("done", []), key=lambda u: u.get("used_at", ""))
    if done:
        return get_topic(done[0]["id"]) or TOPICS[0]
    return TOPICS[0]


def mark_used(topic_id: str, title: str = "", out_file: str = "") -> None:
    """标记话题已完成，并清掉它的预留。"""
    data = _state()
    data.setdefault("done", [])
    data.setdefault("reserved", {})
    for rk, v in list(data["reserved"].items()):
        if v.get("id") == topic_id:
            data["reserved"].pop(rk)
    if not any(u.get("id") == topic_id for u in data["done"]):
        data["done"].append({
            "id": topic_id,
            "title": title or (get_topic(topic_id) or {}).get("title", ""),
            "file": out_file,
            "used_at": datetime.now().isoformat(timespec="seconds")})
    _save_state(data)


def excerpt(topic: dict, limit: int = 3000) -> str:
    """取资料库里与该话题对应的段落（找不到就用开头部分）。"""
    try:
        text = MATERIAL.read_text(encoding="utf-8")
    except OSError:
        return ""
    key = topic.get("match", "")
    pos = text.find(key) if key else -1
    if pos < 0:
        return text[:limit]
    start = max(0, pos - limit // 3)
    return text[start:start + limit]


def main() -> int:
    ap = argparse.ArgumentParser(description="中文域名科普文章的话题管理")
    ap.add_argument("--list", action="store_true", help="列出话题与使用状态")
    ap.add_argument("--pick", action="store_true", help="取一个未用过的话题")
    ap.add_argument("--topic", default="", help="指定话题编号，如 S05")
    ap.add_argument("--mark", help="标记某话题已完成")
    ap.add_argument("--reserve", metavar="ROW_KEY", help="给某一行预留一个话题")
    ap.add_argument("--release", metavar="KEY", help="释放预留（行标识或话题编号）")
    ap.add_argument("--excerpt", help="打印某话题在资料库中的摘录")
    args = ap.parse_args()
    used = used_ids()

    if args.list:
        data = _state()
        done = {u.get("id") for u in data.get("done", [])}
        resv = {v.get("id"): rk for rk, v in data.get("reserved", {}).items()}
        print(f"资料库：{MATERIAL}")
        print(f"状态文件：{state_file()}")
        print(f"已完成 {len(done)} / 预占 {len(resv)} / 共 {len(TOPICS)}")
        for t in TOPICS:
            mark = ("✓ 已完成" if t["id"] in done
                    else (f"· 预占（{resv[t['id']]}）" if t["id"] in resv else "  待用"))
            print(f"  {mark:<16}{t['id']}  {t['title']}")
        return 0
    if args.reserve:
        t = reserve(args.reserve, args.topic)
        print(f"{args.reserve} → {t['id']}\t{t['title']}")
        return 0
    if args.release:
        hit = release(args.release)
        print(f"已释放 {hit or args.release}")
        return 0
    if args.mark:
        mark_used(args.mark)
        print(f"已标记 {args.mark} 为已用")
        return 0
    if args.excerpt:
        t = get_topic(args.excerpt) or pick_topic(args.excerpt)
        print(excerpt(t))
        return 0
    t = pick_topic(args.topic)
    print(f"{t['id']}\t{t['title']}")
    for p in t["points"]:
        print(f"  - {p}")
    print(f"  落点：{t['focus']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
