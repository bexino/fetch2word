#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
title_naming.py —— 文件名规则：yyyy.mm.dd(简化标题).docx

简化标题（short_title）只吃「标题字符串」，**不读正文、不联网、不调模型**：
按标点切片段 → 丢掉口号/抒情片段 → 剥掉机构名与"赴/走进"这类引导动作 →
剥掉尾巴上的煽情词 → 裁到 NAME_TITLE_MAX 字。

例：
    情暖重阳，爱满新才！机电与建筑工程学院志愿者赴青山路街道新才社区传递敬老温情
        → 青山路街道敬老
    我为同学办实事|清洗空调滤网，守护夏日“呼吸健康”
        → 清洗空调滤网

规则给的是「还过得去」的底线结果；AI 认为不够好时，用 --names 映射表覆写
（见 shorten_titles.py），映射表优先级永远高于规则。
"""

import csv
import os
import re

# 文件名里标题部分的字数上限：短标题才有意义，长标题等于没简化
NAME_TITLE_MAX = 12
# 整个文件名的安全上限（含日期与括号）
NAME_MAX = 60

# 口号 / 抒情 / 表态度：出现即说明这段不承载信息
CLICHE_WORDS = (
    "情暖", "爱满", "温情", "温暖", "暖心", "暖意", "真情", "深情", "敬意", "致敬",
    "礼赞", "共筑", "共话", "共赴", "共育", "奋进", "启航", "筑梦", "逐梦", "初心",
    "使命", "担当", "新篇", "新程", "再出发", "砥砺", "笃行", "聚力", "凝心", "铸魂",
    "守护", "护航", "点亮", "绽放", "出彩", "风采", "华章", "赞歌", "致敬", "礼献",
    "青春", "绽放", "温度", "情怀", "同行", "相伴", "接力", "传承", "弘扬", "践行",
)

# 主体 / 机构：交代「谁」，不是「干了什么」
ORG_WORDS = (
    "学院", "大学", "学校", "中学", "小学", "校区", "团委", "团总支", "学生会",
    "研究生会", "党支部", "党总支", "党委", "工会", "妇联", "协会", "社团", "中心",
    "基地", "工作室", "服务队", "实践团", "志愿队", "师生", "同学", "志愿者",
    "学生", "党员", "团员", "代表", "一行", "团队", "小分队",
)

# 事件 / 动作：交代「干了什么」，是简化标题最该保留的部分
ACTION_WORDS = (
    "开展", "举办", "召开", "组织", "参加", "参与", "走进", "走访", "赴", "前往",
    "来到", "深入", "联合", "携手", "发布", "公示", "通知", "公告", "培训", "讲座",
    "会议", "活动", "比赛", "竞赛", "评选", "表彰", "招新", "招聘", "检查", "调研",
    "慰问", "清洗", "维修", "答辩", "汇报", "演练", "签约", "揭牌", "仪式", "启动",
    "上线", "获奖", "荣获", "入选", "录取", "评审", "验收", "支教", "宣讲", "义诊",
    "捐赠", "志愿服务", "座谈", "分享", "出征", "开班", "结业", "毕业", "迎新",
)

# 片段切分：中文标点 + 竖线 + 空白 + 波浪号等排版噪声
SPLIT_RE = re.compile(r"[，。！？；：、,.!?;:|\s—\-~·…“”\"'（）()《》【】\[\]]+")

# 前导机构名（非贪婪，尽量短地吃掉「XX学院/XX团委」；不跨越 的/在/而 等虚词）
ORG_HEAD_RE = re.compile(
    r"^[^，。的而在为、]{0,14}?(?:学院|大学|学校|中学|小学|校区|团委|团总支|"
    r"学生会|研究生会|党支部|党总支|党委|工会|中心|基地|服务队|实践团)(?!生|们)(?=.)"
)
# 前导「谁 + 赴/走进」：剥掉主体与引导动作，留下真正的地点/事由
SUBJ_VERB_RE = re.compile(
    r"^(?:(?:全体|部分|广大|我院|我校|本院)?(?:师生|同学|学子|志愿者|学生|党员|团员|"
    r"代表|团队|一行|队员们?|干事|干部)*)(?:一行)?(?:在|于)?"
    r"(?:赴|走进|前往|来到|深入|走入)"
)
# 前导主语（师生/学子/志愿者…），顺带吃掉紧跟的「在/于」
SUBJ_LEAD_RE = re.compile(
    r"^(?:全体|部分|广大)?(?:师生员工|师生|学子|学生|同学|志愿者|党员|团员|代表|"
    r"队员|团队|干事|干部)(?:们)?(?:在|于)?"
)
# 前导主语（我院/我校…）
SUBJ_HEAD_RE = re.compile(r"^(?:我院|我校|本院|我系|系部|学院|学校|学部|团委|学生会)")
# 前导公文腔
LEAD_VERB_RE = re.compile(r"^(?:关于|对于)?(?:开展|举办|组织|召开|举行|做好|推进|加强|深化)")
# 前导时间
TIME_HEAD_RE = re.compile(
    r"^(?:近日|日前|近期|今天|昨日|上午|下午|晚上|\d{1,4}年|\d{1,2}月\d{0,2}日?|"
    r"\d{1,2}月|\d{1,2}日|第[一二三四五六七八九十]+届|\d{1,4}届|\d{2,4}级)"
)
# 街道 + 社区连写时，社区是冗余的
STREET_COMMUNITY_RE = re.compile(r"([\u4e00-\u9fa5]{2,6}街道)[\u4e00-\u9fa5]{2,8}社区")
# 中间的语气性动词（留着也不加信息，删了更干净）
FILLER_VERB_RE = re.compile(
    r"(?:传递|送去|送上|传承|弘扬|践行|彰显|展现|共话|共筑|护航|守护|助力|赋能)")
# 尾巴上的煽情词
TAIL_EMOTION_RE = re.compile(
    r"(?:温情|温暖|暖心|真情|深情|关怀|暖意|温度|力量|风采|新篇|华章|赞歌)+$")


def _clauses(text):
    return [c for c in SPLIT_RE.split(text) if c]


def _score(clause):
    """片段信息量打分：动作+地点/事由 > 中性 > 主体介绍 > 口号抒情。"""
    s = 0.0
    has_action = any(w in clause for w in ACTION_WORDS)
    if has_action:
        s += 3
    if any(w in clause for w in CLICHE_WORDS):
        s -= 6
    if any(w in clause for w in ORG_WORDS) and not has_action:
        s -= 2
    elif any(w in clause for w in ORG_WORDS):
        s -= 1
    # 长度：4~12 字最像「要点」，过长过短都扣分
    n = len(clause)
    if n <= 3:
        s -= 1.5
    elif n > 16:
        s -= 1.0
    return s


def _clean(clause):
    t = clause.strip()
    for _ in range(4):
        before = t
        t = ORG_HEAD_RE.sub("", t)
        t = SUBJ_HEAD_RE.sub("", t)
        t = SUBJ_LEAD_RE.sub("", t)
        t = SUBJ_VERB_RE.sub("", t)
        t = LEAD_VERB_RE.sub("", t)
        t = re.sub(r"^(?:关于|对于)", "", t)
        t = TIME_HEAD_RE.sub("", t)
        t = re.sub(r"^(?:在|于)", "", t)
        if t == before:
            break
    t = STREET_COMMUNITY_RE.sub(r"\1", t)
    t = FILLER_VERB_RE.sub("", t)
    t = TAIL_EMOTION_RE.sub("", t)
    t = re.sub(r"[的了及与和、]+$", "", t)
    t = re.sub(r"的(?:通知|公告|公示|启事|函|方案|意见)$", "", t)
    t = FILLER_VERB_RE.sub("", t)
    return t.strip()


# 超长时才动的「瘦身」规则：按信息量从低到高依次丢弃
COMPRESS_RULES = (
    re.compile(r"^(?:关于|对于)"),
    re.compile(r"^\d{2,4}年"),
    re.compile(r"(?:春季|秋季)?学期"),
    re.compile(r"^(?:在|于)"),
    re.compile(r"(?:圆满|顺利|成功|隆重)?(?:举行|举办|召开|落幕|收官|结束)$"),
    re.compile(r"的通知$|的公告$|的公示$|的启事$|的函$"),
    re.compile(r"(?:有关工作|相关工作|工作|事宜|事项)$"),
    re.compile(r"(?:中|上)?(?:荣获|斩获|喜获|获得|取得|再获)(?:佳绩|优异成绩|丰硕成果|"
               r"金奖|银奖|铜奖|一等奖|二等奖|三等奖)?$"),
    re.compile(r"(?:专题|主题)?(?:学习会|报告会|座谈会|培训会|动员会|部署会|推进会|"
               r"交流会|宣讲会|讲座|开幕式|闭幕式)$"),
    re.compile(r"^(?:深入|认真|扎实)?(?:学习|贯彻|传达|落实)(?:贯彻)?"),
    re.compile(r"^党的"),
    re.compile(r"^(?:荣获|斩获|喜获|获得|取得|再获)"),
    re.compile(r"^(?:全国|全省|全市|全校|全院)"),
)


def _compress(text, limit):
    """超长标题逐条丢掉低信息量成分，直到能塞进 limit。"""
    t = text
    if len(t) > limit and "暨" in t:  # 「A暨B」只留主项
        left = t.split("暨", 1)[0]
        if len(left) >= 4:
            t = left
    for pat in COMPRESS_RULES:
        while len(t) > limit:
            new = pat.sub("", t)
            if new == t or len(new) < 3:
                break
            t = new
    return t


def _truncate(text, limit=NAME_TITLE_MAX):
    if len(text) <= limit:
        return text
    t = _compress(text, limit)
    if len(t) <= limit:
        return t
    # 尽量切在虚词/连接符后面，别把词切一半
    cut = t[:limit]
    m = None
    for m2 in re.finditer(r"[的了和及暨、·\-]", cut):
        if m2.start() >= 5:
            m = m2
    if m:
        cut = cut[:m.start()]
    cut = re.sub(r"[A-Za-z0-9]+$", "", cut)
    return cut.strip() or t[:limit]


def short_title(title, limit=NAME_TITLE_MAX):
    """从标题文字本身提取要点，产出 ≤ limit 字的文件名用标题。"""
    base = re.sub(r"\|+", "|", (title or "").strip())
    if "||" in base:
        base = base.split("||", 1)[1]
    elif "|" in base:
        parts = [p for p in base.split("|") if p.strip()]
        base = parts[-1] if len(parts) > 1 else base
    base = re.sub(r"\s+", " ", base).strip()
    if not base:
        return ""

    cls = _clauses(base)
    if not cls:
        return _truncate(base, limit)

    # 选信息量最高的片段，同分取靠前的
    best_i = max(range(len(cls)), key=lambda i: (_score(cls[i]), -i))
    picked = _clean(cls[best_i])

    # 清洗后掏空（比如整段都是机构名/口号），退回首段原文
    if len(picked) < 2:
        for c in cls:
            cand = _clean(c)
            if len(cand) >= 2:
                picked = cand
                break

    if len(picked) < 2:
        picked = base

    picked = re.sub(r'[\\/:*?"<>|\r\n\t]', "", picked)
    picked = re.sub(r"\s+", " ", picked).strip()
    return _truncate(picked, limit) or _truncate(base, limit)


def safe_name(date, title, short=None, limit=NAME_TITLE_MAX):
    """yyyy.mm.dd(简化标题).docx。

    short=None    → 按规则从 title 简化（默认行为）
    short="原文"  → 直接用给定标题（--no-summarize 走这里，此时 limit 传 NAME_MAX）
    """
    y, m, d = date.split("-")
    if short is None:
        base = short_title(title, limit)
    else:
        base = (short or "").strip()
        if limit and len(base) > limit:
            base = base[:limit]
    base = re.sub(r'[\\/:*?"<>|\r\n\t]', "", base or "")
    base = re.sub(r"\s+", " ", base).strip()
    if not base:
        base = short_title(title, limit) or "未命名"
    name = "%s.%s.%s(%s).docx" % (y, m, d, base)
    if len(name) > NAME_MAX:  # 极端情况最后兜底
        name = name[:NAME_MAX - len(".docx")] + ".docx"
    return name


def load_name_map(path):
    """读 AI 覆写映射：两列（原文标题 或 URL / 简化标题），Tab 或逗号分隔。"""
    m = {}
    if not path:
        return m
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        sample = fh.read(4096)
        fh.seek(0)
        delim = "\t" if sample.count("\t") >= sample.count(",") else ","
        for row in csv.reader(fh, delimiter=delim):
            if not row or len(row) < 2:
                continue
            key = (row[0] or "").strip().strip('"')
            val = (row[1] or "").strip().strip('"')
            if not key or not val or key.startswith("#"):
                continue
            if key in ("标题", "原文标题", "URL", "url", "链接"):
                continue  # 表头
            m[key] = val
    return m


def lookup_name(mapping, url, title):
    """映射查找：先按 URL，再按完整标题，再按去掉栏目前缀的标题。"""
    if not mapping:
        return ""
    if url and url in mapping:
        return mapping[url]
    if title in mapping:
        return mapping[title]
    base = title.split("||", 1)[1] if "||" in (title or "") else (title or "")
    return mapping.get(base, "")
