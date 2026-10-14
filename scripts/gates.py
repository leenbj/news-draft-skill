#!/usr/bin/env python3
"""新闻稿门禁：把「写作要求」拆成可判定的硬规则。

设计原则：
- 只判断，不修改稿件。硬错（exit 2）必须清零才允许落盘；其余只提醒，交回写稿的 agent 回看。
- 判断只看稿件自身要素（企业全称、注册域名、四块结构），不依赖任何预设文案。
- 断定“有问题”的规则要能指名道姓地说出问题在第几处，方便直接改。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

# 标题句式标记：用来判断两篇标题是不是同一种写法
TITLE_MARKERS = ("官网启用", "正式启用", "启用", "上线", "官网地址", "有了中文入口", "中文入口",
                 "转向", "落地", "提速", "在变", "产业", "行业", "制造业", "数字化", "门槛",
                 "识别", "已经不是", "成了")
# 行业模式（企业资料为零）下，正文应该写到行业与政策
INDUSTRY_DEPTH_WORDS = ("政策", "规划", "指导意见", "行动方案", "十四五", "主管部门",
                        "产业", "行业", "集群", "产业带")

MIN_BODY_CHARS = 1000   # 正文最低字数（不含标题，不含空白）
DEFAULT_TARGET_CHARS = 1000   # 默认要求“1000 字左右”
TARGET_MIN_RATIO = 0.90       # “左右”的下限：0.90x
TARGET_MAX_RATIO = 1.10       # “左右”的上限：1.10x（只提醒，不拦截）
MIN_PARAGRAPHS = 4      # 开篇 / 企业信息 / 中文域名 / 启用意义
WARN_TITLE_CHARS = 34   # 标题偏长的提醒线
MAX_TITLE_CHARS = 60    # 标题硬上限

# ── 红线：中文域名报道里不该出现的东西 ────────────────────────────────
# 英文/外文后缀
EN_SUFFIX = re.compile(
    r"[\w-]+\.(?:com|cn|net|org|top|vip|shop|xyz|info|biz|cc|tv|io|ai|me|co|"
    r"club|site|online|store|ltd|name|pro|mobi|asia|hk|tw|gov|edu)\b", re.I)
# “英文域名 / 国际域名”这类说法
EN_WORDING = re.compile(r"英文域名|国际域名|外文域名|英文后缀|纯英文域")
# 其他中文后缀：仅当本文域名本身就是 .网址 时拦截，避免分散主推后缀
ZH_SUFFIX_OTHER = re.compile(
    r"\.(?:中国|公司|商城|商标|在线|网店|集团|网络|信息|手机|招聘|购物|餐厅|中文网|我爱你|广东|佛山)")
# 贬低中文域名的表述
ZH_DISPARAGE = re.compile(
    r"中文域名[^。；！？\n]{0,24}(?:争议|质疑|缺点|不足|局限|风险|鸡肋|没用|不实用)")
# 联系方式与法定代表人姓名
CONTACT = re.compile(
    r"(?:联系|咨询)?电话\s*[:：]|座机\s*[:：]|传真\s*[:：]|邮箱\s*[:：]|"
    r"微信(?:号)?\s*[:：]|QQ\s*[:：]|公众号\s*[:：]|"
    r"1[3-9]\d{9}|0\d{2,3}-\d{7,8}|[\w.+-]+@[\w-]+\.[\w.]+|"
    r"(?:法定代表人|法人代表|联系人)\s*[:：]?\s*[\u4e00-\u9fa5]{2,4}")
# 企业信息平台／招聘网站名（来源只写“企业公开资料”）
SOURCE_SITE = re.compile(
    r"企查查|爱企查|天眼查|启信宝|水滴信用|顺企网|黄页|BOSS直聘|职友集|猎聘|"
    r"智联招聘|58同城|脉脉|领英|招投标网|千里马招标")
SOURCE_SITE_WARN = re.compile(r"百科|快懂百科|搜狗百科|维基百科|国家企业信用信息公示系统")
# 工商登记类信息：企业资料只允许业务相关内容，工商公开信息一律不写
REGISTRY_FACTS_HARD = (
    ("注册资本/实缴资本", re.compile(r"注册资本|实缴资本|认缴")),
    ("法定代表人", re.compile(r"法定代表人|法人代表")),
    ("统一社会信用代码", re.compile(r"统一社会信用代码|信用代码")),
    ("注册地址", re.compile(r"注册地址|注册地|注册住所")),
    ("股东或股权", re.compile(r"股东|股权结构|持股比例")),
)
REGISTRY_FACTS_WARN = (
    ("成立时间", re.compile(r"成立于|成立时间|成立日期|注册时间|\d{4}\s*年(?:成立|创立)")),
    ("法人代表/企业类型", re.compile(r"法人代表|企业类型")),
    ("经营范围", re.compile(r"经营范围")),
    ("经营状态", re.compile(r"经营状态|存续|在业|吊销|注销|出资")),
)
REGISTRY_SOURCING = re.compile(r"公开工商登记信息|工商登记信息|工商公开信息|工商资料")
# 写作过程／提示词泄漏
META_LEAK = re.compile(
    r"本文|本稿|本篇稿件|稿件中|提示词|提示内容|写作要求|检索结果|检索页面|公开检索|信息边界|"
    r"上述资料|上述材料|本次报道|素材来源|以上材料|写入本文|作为企业情况的来源|"
    r"AI\s*生成|人工智能生成|大模型生成")
# 排版残留
FORMAT_BAD = (
    ("Markdown 标题", re.compile(r"(?m)^\s{0,3}#{1,6}\s")),
    ("Markdown 加粗", re.compile(r"\*\*|__")),
    ("Markdown 列表", re.compile(r"(?m)^\s{0,3}(?:[-*+]|\d+\.)\s+")),
    ("HTML 标签", re.compile(r"</?[a-zA-Z][^>]{0,100}>")),
    ("编码替换符", re.compile("\ufffd")),
    ("控制字符", re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")),
)

# ── 结构要素（用户明确要求的写法）────────────────────────────────────
# 更高角度：行业／产业／数字化等
INDUSTRY_WORDS = ("行业", "产业", "领域", "赛道", "数字化", "智能化", "转型", "升级",
                  "生态", "供应链", "集群", "结构", "高质量", "规划", "标准")
# 中文域名本身的常识：后缀、地址栏、浏览器输入、读写
DOMAIN_KNOWLEDGE_WORDS = ("后缀", "顶级域", "地址栏", "浏览器", "输入", "构成",
                          "中文词语", "中文词", "怎么读", "读音", "写法")
# 启用意义／价值
VALUE_WORDS = ("意义", "价值", "作用", "意味着", "有助于", "带来", "提升", "降低",
               "便利", "效率", "一致性", "识别", "门槛", "成本", "统一")

# ── 风格提醒（只报告，不拦截）────────────────────────────────────────
STYLE_WARN = (
    ("万能结尾", re.compile(r"未来可期|值得期待|让我们拭目以待|大有可为|行稳致远")),
    ("三项同构排比", re.compile(r"(?:[\u4e00-\u9fa5]{2,6}、){2}[\u4e00-\u9fa5]{2,6}(?=[，。])")),
)

# ══ 人化硬规则：来自 human-writing 1.1.0（KKKKhazix）与 Humanizer-zh（op7418）══
# 与 media-publish-skill/scripts/check_humanized.py 同一套判定，合并到本门禁里，
# 这样一次 save 同时过「内容门禁」和「人化门禁」，不合格的稿子落不了盘。
HARD_TERMS_HUMAN = (
    "说白了", "说穿了", "先说结论", "赋能", "抓手", "商业闭环", "价值闭环",
    "能力沉淀", "拉通", "底层逻辑", "顶层设计", "认知跃迁", "价值释放",
    "能力建设", "降本增效", "内容矩阵", "全链路", "组合拳", "打开想象空间",
    "结构性机会", "关键命题", "深层逻辑", "技术底座", "公共底座", "技术主权",
    "单点风险", "主脊柱", "材料锚点", "认知增量", "迭代闭环",
)
ROAD_SIGNS_HUMAN = (
    "更微妙的是", "还有一层", "只说对了一半", "值得注意的是",
    "需要指出的是", "从某种意义上说", "希望这对您有帮助",
)
PIVOT_PATTERNS_HUMAN = (
    re.compile(r"(?:并)?不是[^。！？\n]{0,90}而是"),
    re.compile(r"不仅(?:仅仅)?是[^。！？\n]{0,90}而是"),
    re.compile(r"并非[^。！？\n]{0,90}而是"),
    re.compile(r"不在于[^。！？\n]{0,90}而在于"),
    re.compile(r"与其说[^。！？\n]{0,90}(?:不如|毋宁|倒不如)"),
    re.compile(r"表面(?:上)?[^。！？\n]{0,90}(?:其实|实际|实则)"),
    re.compile(r"看似[^。！？\n]{0,90}(?:其实|实际|实则)"),
    re.compile(r"回头(?:看|一看)?才(?:发现|明白|知道)"),
    re.compile(r"答案(?:是否定的|恰恰相反)|恰恰相反"),
    re.compile(r"[^，。！？\n]{1,12}不重要，(?:重要|要紧)的是"),
)
NOMINALIZATION_PATTERNS = (
    re.compile(r"进行(?:了|一次|一场|着)?[^。，！？\n]{0,12}(?:调整|优化|升级|分析|讨论|沟通|梳理|复盘|迭代|探索|尝试|思考|规划|布局)"),
    re.compile(r"实现了?[^。，！？\n]{0,16}的?[^。，！？\n]{0,8}(?:提升|增长|突破|转变|跃升|落地)"),
    re.compile(r"完成了?对[^。！？\n]{0,18}的"),
    re.compile(r"起到了?[^。，！？\n]{0,14}的?作用"),
)
CONJUNCTIONS_HUMAN = ("因为", "所以", "但是", "然而", "同时", "此外", "而且", "并且", "因此", "不仅")

# ══ 中文后缀审核：只允许 .网址 和 .中国 ══
ZH_SUFFIX_ALLOWED = (".网址", ".中国")
ZH_SUFFIX_KNOWN = (
    "中国", "中國", "网址", "公司", "网络", "網絡", "网店", "商城", "商标", "在线",
    "中文网", "集团", "广东", "佛山", "信息", "手机", "招聘", "购物", "餐厅", "我爱你",
    "游戏", "娱乐", "企业", "移动", "健康", "世界", "政务", "政府", "公益", "测试",
    "香港", "台湾", "台灣", "澳门", "澳門", "新加坡", "时尚", "新闻", "汽车", "手表",
    "珠宝", "保险", "证券", "商店", "书店", "饭店", "招聘", "八卦", "网址",
)
# 命中已知中文后缀（且不是 .网址/.中国）→ 硬错
ZH_SUFFIX_HIT = re.compile(
    r"[.．]\s*(" + "|".join(sorted(set(ZH_SUFFIX_KNOWN), key=len, reverse=True)) + r")")
# 兜底：点号后跟中文词的通用形态 → 提醒人工确认（可能是没收录的后缀）
ZH_SUFFIX_GENERIC = re.compile(
    r"(?<![\u4e00-\u9fa5A-Za-z0-9])([A-Za-z0-9\u4e00-\u9fa5-]{1,24})[.．]([\u4e00-\u9fa5]{2,10})")


@dataclass
class Report:
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    stats: dict = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return not self.errors

    def render(self) -> str:
        lines = []
        st = self.stats
        if st:
            lines.append(
                f"[统计] 正文 {st.get('body_chars', 0)} 字 / {st.get('paragraphs', 0)} 段 · "
                f"域名出现 {st.get('domain_hits', 0)} 次 · 中文域名 {st.get('zh_hits', 0)} 次")
        for e in self.errors:
            lines.append(f"✗ {e}")
        for w in self.warnings:
            lines.append(f"! {w}")
        if self.ok and not self.warnings:
            lines.append("✓ 门禁通过")
        elif self.ok:
            lines.append("✓ 硬规则通过（提醒项请人工回看）")
        return "\n".join(lines)


def prose_len(text: str) -> int:
    """字数：去掉空白后的字符数（中文按字计，标点也算，和常用字数口径一致）。"""
    return len(re.sub(r"\s+", "", text or ""))


def paragraphs(body: str) -> list[str]:
    return [p.strip() for p in re.split(r"\n\s*\n", (body or "").strip()) if p.strip()]


def _mask(text: str, *values: str) -> str:
    """把本文自己的域名/企业名抠掉，避免红线扫到它们本身。"""
    out = text or ""
    for v in values:
        if v:
            out = re.sub(re.escape(v), "", out, flags=re.I)
    return out


def _hit(pattern: re.Pattern, text: str) -> list[str]:
    return [m.group(0) for m in pattern.finditer(text)]


def _context(text: str, pattern: re.Pattern, width: int = 14) -> list[str]:
    out = []
    for m in pattern.finditer(text):
        out.append("…" + text[max(0, m.start() - width):m.end() + width].replace("\n", " ") + "…")
    return out


try:                      # 政策清单＝industry.py 收录的行业政策 + science.py 已核实的依据
    import industry as _industry
    import science as _science
    POLICY_WHITELIST = set(_industry.policy_whitelist()) | set(_science.policy_whitelist())
except Exception:         # noqa: BLE001
    POLICY_WHITELIST = set()


# ── 科普稿（中文域名科普文章）专用：只讲正面，只允许 .网址 和 .中国 ──
SCIENCE_NEGATIVE_HARD = (
    re.compile(r"中文域名[^。；！？\n]{0,24}(?:争议|质疑|缺点|不实用|鸡肋|骗局|炒作|贬值|"
               r"不值得|没必要|没有价值|不推荐|风险大|后悔|不可靠|不靠谱)"),
    re.compile(r"(?:不建议|不必|无需|放弃|别用|慎用)(?:注册|使用|购买|选择|引入)?中文域名"),
    re.compile(r"中文域名[^。；！？\n]{0,20}(?:没用|白花钱|会失败|行不通)"),
    re.compile(r"(?:不如|比不上|远不如)[^。；！？\n]{0,14}(?:中文域名)"),
)
SCIENCE_NEGATIVE_WARN = (
    ("免责式冷水", re.compile(r"不会自动(?:为|给|带来)[^。；\n]{0,12}流量|无法保证|不能保证|不保证")),
    ("价格对比的负面口径", re.compile(r"价格(?:明显)?(?:偏高|过高|昂贵|太贵)|价格远高于")),
    ("兼容性问题的直述", re.compile(r"兼容性不足|尚未实现全面兼容|尚未普及")),
)
SCIENCE_BASIS_WORDS = ("政策", "办法", "规划", "标准", "规定", "ICANN", "IDN", "国际化域名",
                       "IANA", "论坛", "商标法", "报告")


def check_science(title: str, body: str, domain: str = "",
                  min_chars: int = MIN_BODY_CHARS, max_chars: int = 0,
                  humanize_gate: bool = True,
                  used_titles: list[str] | None = None) -> Report:
    """中文域名科普稿的门禁。

    与企业稿的差别：不要求企业三要素，不限制行业词，允许为了讲清差异提到英文域名与 .com/.cn；
    但中文域名后缀仍只允许 .网址 和 .中国，且不得出现任何关于中文域名的负面表述。"""
    rep = Report()
    body = (body or "").strip()
    paras = paragraphs(body)
    rep.stats = {
        "kind": "science",
        "body_chars": prose_len(body),
        "paragraphs": len(paras),
        "min_chars": min_chars,
        "max_chars": max_chars,
        "zh_hits": body.count("中文域名"),
    }
    if not paras:
        rep.errors.append("正文为空")
        return rep

    # 篇幅与层次
    if rep.stats["body_chars"] < min_chars:
        rep.errors.append(f"正文只有 {rep.stats['body_chars']} 字，不足 {min_chars} 字")
    if max_chars and rep.stats["body_chars"] > max_chars:
        rep.warnings.append(
            f"正文 {rep.stats['body_chars']} 字，超出 {max_chars} 字的上限，压缩要删重复不删依据")
    if len(paras) < 3:
        rep.warnings.append(f"正文只有 {len(paras)} 段，科普稿建议分段把话题讲透")

    # 主题聚焦与依据
    if body.count("中文域名") < 3:
        rep.errors.append("正文提到「中文域名」不足 3 次，主题不够聚焦")
    if not any(w in body for w in SCIENCE_BASIS_WORDS):
        rep.warnings.append(
            "没有写到任何政策、标准或行业依据（政策/办法/规划/ICANN/IDN/论坛等），"
            "科普稿要有可核验的依据")

    # 中文后缀审核：只允许 .网址 和 .中国
    bad = [m for m in ZH_SUFFIX_HIT.findall(body) if f".{m}" not in ZH_SUFFIX_ALLOWED]
    if bad:
        rep.errors.append("出现不允许的中文域名后缀：" + "、".join(f".{b}" for b in dict.fromkeys(bad))
                          + "（只允许 .网址 和 .中国）")
    for m in ZH_SUFFIX_GENERIC.finditer(body):
        suffix = m.group(2)
        if f".{suffix}" in ZH_SUFFIX_ALLOWED or suffix in ZH_SUFFIX_KNOWN:
            continue
        rep.warnings.append(f"疑似其他中文域名后缀「{m.group(0)}」：只允许 .网址 和 .中国")

    # 正面性：不得出现任何关于中文域名的负面表述
    for pat in SCIENCE_NEGATIVE_HARD:
        for m in pat.finditer(body):
            rep.errors.append(f"科普稿只能讲正面，这里出现负面表述：“{m.group()[:40]}”")
    for label, pat in SCIENCE_NEGATIVE_WARN:
        for m in pat.finditer(body):
            rep.warnings.append(
                f"{label}：“{m.group()[:32]}” —— 科普稿落点要放在正面价值与做法上，"
                "这类冷水句建议改写或删掉")

    # 通用红线（联系方式／平台名／写作过程话语／贬低中文域名）
    masked = _mask(body, domain.strip())
    for label, pat, hard in (
        ("联系方式或联系人姓名", CONTACT, True),
        ("企业信息平台／招聘网站名称", SOURCE_SITE, True),
        ("写作过程话语", META_LEAK, True),
        ("贬低中文域名的表述", ZH_DISPARAGE, True),
        ("百科或公示系统名称", SOURCE_SITE_WARN, False),
    ):
        hits = _context(masked, pat)
        if not hits:
            continue
        (rep.errors if hard else rep.warnings).append(f"出现{label}：" + "；".join(hits[:3]))

    # 政策名称核对（与企业稿同一套清单）
    for m in re.finditer(r"《([^》]{4,60})》", body):
        name = m.group(1)
        if name not in POLICY_WHITELIST:
            rep.warnings.append(
                f"引用了《{name}》，不在已收录的政策清单里：先用 AnySearch 核实名称与发布单位")

    # 排版
    for label, pat in FORMAT_BAD:
        if pat.search(body):
            rep.errors.append(f"正文含{label}，只交付纯文本段落")

    # 人化门禁
    if humanize_gate:
        _humanized_checks(body, rep)

    # 标题
    t_err, t_warn = check_title(title, "", "")
    rep.errors.extend(f"标题：{e}" for e in t_err if "企业名称" not in e)
    rep.warnings.extend(f"标题：{w}" for w in t_warn if "企业名称" not in w)
    v_err, v_warn = check_title_variety(title, used_titles)
    rep.errors.extend(f"标题：{e}" for e in v_err)
    rep.warnings.extend(f"标题：{w}" for w in v_warn)
    return rep


def _humanized_checks(body: str, rep: Report) -> None:
    """human-writing / humanizer-zh 的硬规则，直接并进主门禁。

    对应 media-publish-skill/scripts/check_humanized.py：破折号、提示性冒号、
    硬禁词、模型路标、翻案腔、写作过程话语是硬错；名词化、连词密度、通用收束是提醒。"""
    text = body or ""
    lines = text.split("\n")

    def where(pos: int) -> int:
        return text.count("\n", 0, pos) + 1

    for symbol, label in (("—", "破折号"), ("–", "连接号式破折号")):
        for m in re.finditer(re.escape(symbol), text):
            rep.errors.append(f"人化｜{label}：第 {where(m.start())} 行")
    for symbol, label in (("：", "中文冒号"), (":", "英文冒号")):
        for m in re.finditer(re.escape(symbol), text):
            tail = text[m.end():m.end() + 2].lstrip()
            if tail[:1] not in ("「", "『", "“", "‘", '"'):
                rep.errors.append(f"人化｜提示性{label}：第 {where(m.start())} 行")
    for term in HARD_TERMS_HUMAN:
        for m in re.finditer(re.escape(term), text):
            rep.errors.append(f"人化｜硬禁词“{term}”：第 {where(m.start())} 行")
    for term in ROAD_SIGNS_HUMAN:
        for m in re.finditer(re.escape(term), text):
            rep.errors.append(f"人化｜模型路标“{term}”：第 {where(m.start())} 行")
    for pat in PIVOT_PATTERNS_HUMAN:
        for m in pat.finditer(text):
            rep.errors.append(
                f"人化｜翻案腔：第 {where(m.start())} 行“{re.sub(r'\\s+', ' ', m.group())[:44]}”")

    nominal = sum(len(p.findall(text)) for p in NOMINALIZATION_PATTERNS)
    if nominal:
        rep.warnings.append(f"人化提醒｜名词化句式 {nominal} 处，改成直接动词并确认有实际结果")
    han = len(re.findall(r"[\u4e00-\u9fff]", text))
    conj = sum(text.count(t) for t in CONJUNCTIONS_HUMAN)
    if han >= 600 and conj * 1000 / han > 7:
        rep.warnings.append(f"人化提醒｜连词密度偏高：每千字约 {conj * 1000 // han} 个")
    if re.search(r"(?:此外|同时|因此|综上|总的来说).{0,12}(?:未来|新篇章|新征程|未来可期)", text):
        rep.warnings.append("人化提醒｜结尾可能是通用积极收束，回到当前事实或已知动作")


def title_signature(title: str) -> tuple[str, ...]:
    """标题的句式签名：按固定标记的组合判断两篇标题是否同一写法。"""
    t = title or ""
    return tuple(m for m in TITLE_MARKERS if m in t)


def check_title_variety(title: str, used: list[str] | None) -> tuple[list[str], list[str]]:
    """标题去重与去同构。used 是已交付稿件的标题列表。"""
    errors, warnings = [], []
    t = (title or "").strip()
    if not t or not used:
        return errors, warnings
    norm = lambda s: re.sub(r"[\s“”‘’《》\"']", "", s)
    for old in used:
        if norm(old) == norm(t):
            errors.append(f"标题与已交付的《{old}》完全相同，换一个说法")
            return errors, warnings
    sig = title_signature(t)
    for old in used:
        if sig and title_signature(old) == sig:
            warnings.append(
                f"标题句式与已交付的《{old}》相同（都用到了 {'、'.join(sig)}），"
                "换一个切入角度，别让一组稿件看起来像一个模子")
            break
    return errors, warnings


def check_title(title: str, company: str, short: str = "") -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    t = (title or "").strip()
    if not t:
        errors.append("标题为空")
        return errors, warnings
    if "\n" in t or "\r" in t:
        errors.append("标题里不能有换行")
    if len(t) > MAX_TITLE_CHARS:
        errors.append(f"标题 {len(t)} 字，超过 {MAX_TITLE_CHARS} 字上限（平台列表页会截断）")
    elif len(t) > WARN_TITLE_CHARS:
        warnings.append(f"标题 {len(t)} 字，偏长（建议 {WARN_TITLE_CHARS} 字以内）")
    if re.search(r"[，、：；,;:（(【\[“\"']$", t):
        errors.append(f"标题结尾突兀（以标点收尾）：{t[-12:]}")
    if company and t == company.strip():
        errors.append("标题不能只是企业名称本身，需要给出稿件的信息点")
    names = [n for n in (company.strip(), (short or "").strip()) if n]
    if names and not any(n in t for n in names):
        warnings.append(f"标题没有出现企业名称或其简称（现在：{t}）")
    if re.search(r"\.txt$|\.docx?$", t, re.I):
        errors.append("标题里不要带文件扩展名")
    return errors, warnings


def check_text(title: str, body: str, company: str, domain: str,
               short: str = "", suffix_policy: str = "auto",
               min_chars: int = MIN_BODY_CHARS, max_chars: int = 0,
               allow_registration: bool = False,
               humanize_gate: bool = True,
               allow_any_zh_suffix: bool = False,
               used_titles: list[str] | None = None,
               industry_mode: bool = False) -> Report:
    """suffix_policy：auto（域名是 .网址 时按 .网址 规则）｜zh-only｜relaxed。

    min_chars：正文硬下限（不给就用 MIN_BODY_CHARS）。
    max_chars：正文上限，只提醒不拦截（0 表示不限）。要求“800 字左右”时用
    min_chars 700 / max_chars 950 这类组合，别把软要求做成硬拦截。
    allow_registration：默认 False，正文不得出现工商公开信息（注册资本/成立时间/法定代表人/
    信用代码/注册地址/股东/经营状态）；显式放开时降级为提醒。"""
    rep = Report()
    body = (body or "").strip()
    paras = paragraphs(body)
    is_url_suffix = domain.strip().endswith(".网址")
    policy = suffix_policy
    if policy == "auto":
        policy = "zh-only" if is_url_suffix else "relaxed"

    rep.stats = {
        "body_chars": prose_len(body),
        "paragraphs": len(paras),
        "min_chars": min_chars,
        "max_chars": max_chars,
        "domain_hits": len(re.findall(re.escape(domain.strip()), body, flags=re.I)) if domain else 0,
        "zh_hits": body.count("中文域名"),
        "suffix_policy": policy,
    }

    # 1) 开篇核心事实
    if not paras:
        rep.errors.append("正文为空")
        return rep
    lead = paras[0]
    if company and company.strip() not in lead:
        rep.errors.append(f"第一段没有出现企业全称「{company.strip()}」")
    if "官网启用" not in lead:
        rep.errors.append("第一段没有出现「官网启用」这一核心事实")
    if domain and domain.strip().lower() not in lead.lower():
        rep.errors.append(f"第一段没有出现注册域名「{domain.strip()}」")
    if company and domain and f"{company.strip()}官网启用" not in lead.replace(" ", ""):
        pos = len(lead)
        rep.warnings.append(
            f"第一段里「{company.strip()}官网启用」不是连着的（第 {pos} 字处才有开篇段），"
            "读起来像铺垫而不是开篇")

    # 2) 篇幅
    if rep.stats["body_chars"] < min_chars:
        rep.errors.append(
            f"正文只有 {rep.stats['body_chars']} 字，不足 {min_chars} 字"
            "（材料不足就补充可核验事实，不要靠重复凑字数）")
    if max_chars and rep.stats["body_chars"] > max_chars:
        rep.warnings.append(
            f"正文 {rep.stats['body_chars']} 字，超出 {max_chars} 字的上限："
            "要求是控制在多少字左右时，压缩要删重复和套话，不能删掉已确认的事实")
    if len(paras) < MIN_PARAGRAPHS:
        rep.warnings.append(
            f"正文只有 {len(paras)} 段，四块内容（开篇 / 企业信息 / 中文域名 / 启用意义）可能被合并了")

    # 3) 四块结构
    if body.count("中文域名") < 2:
        rep.errors.append("正文提到「中文域名」不足 2 次，第三块（中文域名本身）没写到位")
    if not any(w in body for w in DOMAIN_KNOWLEDGE_WORDS):
        rep.errors.append(
            "正文没有介绍中文域名本身（后缀构成、地址栏输入、读写方式等），第三块缺失")
    if not any(w in body for w in INDUSTRY_WORDS):
        rep.errors.append(
            "正文没有行业／产业层面的视角（行业、产业、数字化、转型等），未达到“站在更高角度”的要求")
    if not any(w in body for w in VALUE_WORDS):
        rep.errors.append("正文没有写企业启用中文域名的意义／价值，第四块缺失")
    if domain and rep.stats["domain_hits"] < 2:
        rep.warnings.append(f"注册域名在正文只出现 {rep.stats['domain_hits']} 次，建议开篇与意义段各出现一次")

    # 4) 红线（先抠掉本文自己的域名与企业名）
    masked = _mask(body, domain.strip(), company.strip())
    for label, pat, hard in (
        ("其他英文后缀或域名", EN_SUFFIX, True),
        ("“英文域名／国际域名”类说法", EN_WORDING, True),
        ("贬低中文域名的表述", ZH_DISPARAGE, True),
        ("联系方式或联系人姓名", CONTACT, True),
        ("企业信息平台／招聘网站名称", SOURCE_SITE, True),
        ("写作过程话语", META_LEAK, True),
        ("百科或公示系统名称", SOURCE_SITE_WARN, False),
    ):
        hits = _context(masked, pat)
        if not hits:
            continue
        if hard:
            rep.errors.append(f"出现{label}：" + "；".join(hits[:3]))
        else:
            rep.warnings.append(f"出现{label}，来源只写“企业公开资料”：" + "；".join(hits[:2]))
    # 4.1) 中文后缀审核：除 .网址 和 .中国，其他中文域名后缀一律不允许
    if not allow_any_zh_suffix:
        bad = [_m for _m in ZH_SUFFIX_HIT.findall(masked)
               if f".{_m}" not in ZH_SUFFIX_ALLOWED]
        if bad:
            rep.errors.append(
                "出现不允许的中文域名后缀：" + "、".join(f".{b}" for b in dict.fromkeys(bad))
                + "（只允许 .网址 和 .中国）")
        for m in ZH_SUFFIX_GENERIC.finditer(masked):
            suffix = m.group(2)
            if f".{suffix}" in ZH_SUFFIX_ALLOWED or suffix in ZH_SUFFIX_KNOWN:
                continue
            rep.warnings.append(
                f"疑似其他中文域名后缀「{m.group(0)}」：只允许 .网址 和 .中国，"
                "确认不是域名后缀再放过")

    # 4.5) 企业资料只允许业务相关内容
    if REGISTRY_SOURCING.search(masked):
        rep.errors.append(
            "出现“工商登记信息”这类来源表述：来源只写“企业公开资料”")
    for label, pat in REGISTRY_FACTS_HARD:
        if pat.search(masked):
            msg = (f"出现工商登记类信息（{label}）：企业资料只能用业务相关内容，"
                   "工商公开信息一律不写")
            (rep.warnings if allow_registration else rep.errors).append(msg)
    for label, pat in REGISTRY_FACTS_WARN:
        if pat.search(masked):
            rep.warnings.append(f"出现工商登记类表述（{label}），建议删掉，只留业务相关内容")

    # 5) 排版残留
    for label, pat in FORMAT_BAD:
        if pat.search(body):
            rep.errors.append(f"正文含{label}，只交付纯文本段落")

    # 5.5) 人化门禁：human-writing → humanizer-zh 后的硬规则
    if humanize_gate:
        _humanized_checks(body, rep)

    # 6) 风格提醒
    for label, pat in STYLE_WARN:
        n = len(pat.findall(body))
        if n:
            rep.warnings.append(f"{label} × {n}，建议改写")

    # 6.5) 行业模式：正文要写到行业与政策
    if industry_mode and not any(w in body for w in INDUSTRY_DEPTH_WORDS):
        rep.warnings.append(
            "本篇没有企业资料，正文却没写到行业与政策（政策、规划、产业、行业等），"
            "第二块应该按 brief 的行业素材放大篇幅")

    # 6.6) 政策名称核对：不在已收录清单里的《文件》要人工核实
    for m in re.finditer(r"《([^》]{4,60})》", body):
        name = m.group(1)
        if name not in POLICY_WHITELIST:
            rep.warnings.append(
                f"引用了《{name}》，不在已收录的政策清单里：先用 AnySearch 核实名称与发布单位，"
                "核不到就删掉，不要编文件名")

    # 7) 标题
    t_err, t_warn = check_title(title, company, short)
    rep.errors.extend(f"标题：{e}" for e in t_err)
    rep.warnings.extend(f"标题：{w}" for w in t_warn)
    v_err, v_warn = check_title_variety(title, used_titles)
    rep.errors.extend(f"标题：{e}" for e in v_err)
    rep.warnings.extend(f"标题：{w}" for w in v_warn)
    return rep
