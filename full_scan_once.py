#!/usr/bin/env python3
"""
一次性抓取并筛选实习僧上海校招/全职岗位。

用途：
1. 抓取实习僧列表页中“上海 + 校招”的岗位 ID。
2. 与上次扫描保存的岗位 ID 快照比对，只展示上次以来新增的岗位。
3. 逐个打开新增岗位详情页，提取岗位名称、公司、薪资、规模、工作描述等字段。
4. 先按薪资和公司规模做硬筛选，再按“明确社科/社会学匹配”和
   “近似匹配：可迁移社会学能力”两类筛选。
4. 输出 CSV 和 JSON，供后续生成 Excel 或交给其他 AI 继续分析。

运行示例：
    python3 full_scan_once.py --pages 20 --delay 0.15 --outdir outputs

列表页默认每次重新抓取，以便发现新增岗位；详情页会优先读取缓存。
如果你明确想重新抓取详情页，可加 --refresh。
"""
import argparse
import csv
import hashlib
import html
import json
import re
import sqlite3
import time
from datetime import datetime
from pathlib import Path
from urllib.error import URLError
from urllib.parse import urlencode
from urllib.request import ProxyHandler, Request, build_opener


BASE_LIST_URL = "https://resume.shixiseng.com/interns"
BASE_DETAIL_URL = "https://www.shixiseng.com/intern/{uuid}"
SCAN_STATE_KEY = "shixiseng_school_shanghai"

# 第一类：岗位描述中直接出现这些词，就归入“直接匹配”。
# 这里保留“社科”这种短词，是因为招聘文案常写“人文社科类专业”。
# “用户研究/用研”按用户需求视为社会学博士可直接对应的岗位方向。
EXPLICIT_TERMS = [
    "社会学",
    "社会科学",
    "人文社科",
    "社科",
    "用户",
    "用户研究",
    "用户洞察",
    "用研",
]

# 第二类：近似匹配规则。
#
# 每条规则格式为：
#   (关键词, 分数权重, 能力类别)
#
# 设计思路：
# - 社会学博士不只匹配“社会学”字面词，也匹配研究方法、用户洞察、
#   行业/政策分析、报告写作、社会议题等可迁移能力。
# - “用户研究”“定性研究”等强相关词给高分。
# - “数据分析”“消费者”“文化”等泛词给低分，必须与其他信号叠加后才会入选。
# - 想放宽或收紧筛选，优先改这里的关键词和分数。
APPROX_RULES = [
    ("用户研究", 8, "用户/消费者洞察"),
    ("用户洞察", 8, "用户/消费者洞察"),
    ("消费者洞察", 8, "用户/消费者洞察"),
    ("用户画像", 7, "用户/消费者洞察"),
    ("用研", 7, "用户/消费者洞察"),
    ("用户运营", 8, "用户/消费者洞察"),
    ("用户增长", 7, "用户/消费者洞察"),
    ("用户体验", 6, "用户/消费者洞察"),
    ("用户行为", 7, "用户/消费者洞察"),
    ("用户需求", 6, "用户/消费者洞察"),
    ("市场调研", 7, "调研方法"),
    ("定性研究", 7, "调研方法"),
    ("定量研究", 7, "调研方法"),
    ("深度访谈", 7, "调研方法"),
    ("民族志", 8, "社会学/人类学方法"),
    ("田野", 8, "社会学/人类学方法"),
    ("人类学", 7, "社会学/人类学方法"),
    ("访谈", 5, "调研方法"),
    ("问卷", 5, "调研方法"),
    ("调研", 5, "调研方法"),
    ("桌面研究", 5, "研究分析"),
    ("案头研究", 5, "研究分析"),
    ("文献", 4, "研究分析"),
    ("研究报告", 6, "研究写作"),
    ("报告撰写", 5, "研究写作"),
    ("行业研究", 6, "研究分析"),
    ("市场研究", 6, "研究分析"),
    ("政策研究", 7, "政策/公共议题"),
    ("政策分析", 7, "政策/公共议题"),
    ("公共政策", 6, "政策/公共议题"),
    ("舆情", 5, "传播/社会观察"),
    ("竞品分析", 4, "研究分析"),
    ("数据分析", 3, "定量分析"),
    ("统计分析", 4, "定量分析"),
    ("质性", 6, "调研方法"),
    ("量化", 4, "定量分析"),
    ("咨询项目", 4, "研究/咨询"),
    ("管理咨询", 4, "研究/咨询"),
    ("战略咨询", 4, "研究/咨询"),
    ("公益", 5, "社会议题"),
    ("社区", 4, "社会议题"),
    ("性别", 6, "社会议题"),
    ("女性", 3, "社会议题"),
    ("可持续", 5, "社会议题"),
    ("ESG", 5, "社会议题"),
    ("SDG", 5, "社会议题"),
    ("社会责任", 5, "社会议题"),
    ("教育", 3, "社会议题"),
    ("文化", 3, "文化/传播"),
    ("传播策略", 4, "文化/传播"),
    ("消费者", 3, "用户/消费者洞察"),
    ("人群", 3, "用户/消费者洞察"),
    ("AI行为评估", 9, "AI/人文评估"),
    ("AI评估", 8, "AI/人文评估"),
    ("评估体系", 6, "AI/人文评估"),
    ("测试集", 5, "AI/人文评估"),
    ("真实语境", 6, "AI/人文评估"),
    ("真实情绪", 6, "AI/人文评估"),
    ("表达方式", 6, "AI/人文评估"),
    ("行为准则", 6, "AI/人文评估"),
    ("价值观", 6, "AI/人文评估"),
    ("人文视角", 8, "AI/人文评估"),
    ("跨文化理解", 8, "AI/人文评估"),
    ("心理学", 5, "AI/人文评估"),
    ("哲学", 5, "AI/人文评估"),
    ("批判性思维", 6, "AI/人文评估"),
    ("文字表达", 5, "研究写作"),
    ("文字功底", 5, "研究写作"),
    # ── 市场营销 / 品牌 ──
    ("市场营销", 5, "市场营销/品牌"),
    ("品牌营销", 5, "市场营销/品牌"),
    ("数字营销", 5, "市场营销/品牌"),
    ("整合营销", 5, "市场营销/品牌"),
    ("品牌", 3, "市场营销/品牌"),
    ("品牌管理", 5, "市场营销/品牌"),
    ("品牌策略", 5, "市场营销/品牌"),
    ("品牌建设", 5, "市场营销/品牌"),
    ("快消", 4, "市场营销/品牌"),
    ("FMCG", 4, "市场营销/品牌"),
    ("营销策划", 5, "市场营销/品牌"),
    ("营销活动", 4, "市场营销/品牌"),
    ("营销策略", 5, "市场营销/品牌"),
    ("市场部", 4, "市场营销/品牌"),
    ("市场推广", 4, "市场营销/品牌"),
    ("产品定位", 5, "市场营销/品牌"),
    ("消费者行为", 6, "市场营销/品牌"),
    ("消费者研究", 7, "市场营销/品牌"),
    ("人群洞察", 6, "市场营销/品牌"),
    ("Campaign", 5, "市场营销/品牌"),
    ("客户洞察", 6, "市场营销/品牌"),
    # ── 社媒 / 内容运营 ──
    ("社交媒体", 5, "社媒/内容运营"),
    ("社媒", 5, "社媒/内容运营"),
    ("新媒体", 5, "社媒/内容运营"),
    ("内容生态", 5, "社媒/内容运营"),
    ("内容运营", 5, "社媒/内容运营"),
    ("内容策略", 5, "社媒/内容运营"),
    ("内容营销", 5, "社媒/内容运营"),
    ("种草", 5, "社媒/内容运营"),
    ("达人", 5, "社媒/内容运营"),
    ("KOL", 6, "社媒/内容运营"),
    ("KOC", 5, "社媒/内容运营"),
    ("小红书", 5, "社媒/内容运营"),
    ("抖音", 5, "社媒/内容运营"),
    ("短视频", 4, "社媒/内容运营"),
    ("直播", 4, "社媒/内容运营"),
    ("私域", 5, "社媒/内容运营"),
    ("公域", 4, "社媒/内容运营"),
    ("流量", 3, "社媒/内容运营"),
    ("涨粉", 4, "社媒/内容运营"),
    ("爆款", 4, "社媒/内容运营"),
    ("创意", 3, "社媒/内容运营"),
    ("创意输出", 5, "社媒/内容运营"),
    ("文案", 4, "社媒/内容运营"),
    ("策划", 3, "社媒/内容运营"),
    # ── 电商 / 增长 ──
    ("电商", 4, "市场营销/品牌"),
    ("生意参谋", 5, "市场营销/品牌"),
    ("京东商智", 5, "市场营销/品牌"),
    ("增长", 3, "市场营销/品牌"),
    ("转化", 3, "市场营销/品牌"),
    ("投放", 4, "市场营销/品牌"),
    ("GMV", 4, "市场营销/品牌"),
    ("ROI", 4, "市场营销/品牌"),
    # ── AI / 大模型 / 智能体 ──
    ("人工智能", 5, "AI/大模型"),
    ("AI", 3, "AI/大模型"),
    ("大模型", 6, "AI/大模型"),
    ("LLM", 6, "AI/大模型"),
    ("智能体", 7, "AI/大模型"),
    ("Agent", 5, "AI/大模型"),
    ("AI Agent", 7, "AI/大模型"),
    ("机器学习", 5, "AI/大模型"),
    ("深度学习", 5, "AI/大模型"),
    ("自然语言处理", 5, "AI/大模型"),
    ("NLP", 5, "AI/大模型"),
    ("提示词", 5, "AI/大模型"),
    ("Prompt", 5, "AI/大模型"),
    ("AIGC", 5, "AI/大模型"),
    ("ChatGPT", 5, "AI/大模型"),
    ("Claude", 5, "AI/大模型"),
    ("Copilot", 4, "AI/大模型"),
    ("Codex", 4, "AI/大模型"),
    ("RAG", 5, "AI/大模型"),
    ("模型训练", 5, "AI/大模型"),
    ("模型微调", 5, "AI/大模型"),
    ("数据标注", 4, "AI/大模型"),
    ("AI工具", 5, "AI/大模型"),
    ("AI应用", 5, "AI/大模型"),
    ("大语言模型", 6, "AI/大模型"),
    ("生成式", 5, "AI/大模型"),
    ("GenAI", 5, "AI/大模型"),
    ("自动化", 3, "AI/大模型"),
    ("工作流", 3, "AI/大模型"),
    ("Vibe Coding", 6, "AI/大模型"),
    ("Skill", 2, "AI/大模型"),
    ("MCP", 5, "AI/大模型"),
    ("工具调用", 4, "AI/大模型"),
    ("Function Call", 5, "AI/大模型"),
]

# 至少命中一个核心类别，近似匹配才会保留。
# 这可以防止只有“文化”“女性”等泛词的岗位被误收太多。
CORE_CATEGORIES = {
    "用户/消费者洞察",
    "调研方法",
    "社会学/人类学方法",
    "研究分析",
    "研究写作",
    "政策/公共议题",
    "社会议题",
    "AI/人文评估",
    "市场营销/品牌",
    "社媒/内容运营",
    "AI/大模型",
    "传播/社会观察",
    "文化/传播",
    "定量分析",
    "研究/咨询",
}

# CSV / JSON 的统一字段顺序。
# build_excel_report.py 也会读取这些字段并生成 Excel。
FIELDS = [
    "category",
    "match_level",
    "score",
    "is_new",
    "signature",
    "title",
    "company",
    "industry",
    "city",
    "address",
    "salary",
    "min_salary",
    "max_salary",
    "days_per_week",
    "months",
    "degree",
    "refresh_time",
    "end_time",
    "url",
    "matched_keywords",
    "fit_reason",
    "evidence",
    "description",
    "parse_status",
    "parse_warning",
    "scale",
    "list_tags",
    "filter_reasons",
    "uuid",
]


def fetch(url, timeout=40, retries=5, backoff=2.0):
    """下载网页 HTML。

    ProxyHandler({}) 用来忽略当前 shell 中可能存在但不可用的代理环境变量。
    这次环境里曾出现过 127.0.0.1:6518 代理不可连接的问题，所以这里显式禁用代理。

    临时性的网络错误（SSL 握手被对方掐断、连接超时、对方限流等）会自动重试，
    每次失败后等待时间翻倍（backoff），重试 retries 次仍失败才真正抛出异常。
    """
    request = Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/125.0 Safari/537.36"
            ),
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        },
    )
    opener = build_opener(ProxyHandler({}))
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            with opener.open(request, timeout=timeout) as response:
                return response.read().decode("utf-8", errors="replace")
        except (URLError, TimeoutError, ConnectionError) as error:
            last_error = error
            if attempt < retries:
                wait = backoff * (2 ** (attempt - 1))
                print(
                    f"  请求失败（第 {attempt}/{retries} 次）：{error}；{wait:.1f}s 后重试 {url}",
                    flush=True,
                )
                time.sleep(wait)
    raise last_error


def list_url(page):
    """拼出实习僧列表页 URL。

    当前筛选条件：
    - type=school：校招/全职
    - publishTime 为空：不按发布时间限制
    - city=上海
    """
    query = {
        "page": page,
        "type": "school",
        "keyword": "",
        "area": "",
        "months": "",
        "days": "",
        "degree": "",
        "official": "",
        "enterprise": "",
        "salary": "-0",
        "publishTime": "",
        "sortType": "",
        "city": "上海",
        "internExtend": "",
    }
    return f"{BASE_LIST_URL}?{urlencode(query)}"


def unique_uuids(text):
    """从列表页 HTML 中提取岗位 ID，并保持网页出现顺序去重。"""
    seen = set()
    uuids = []
    for uuid in re.findall(r"inn_[a-z0-9]+", text):
        if uuid not in seen:
            seen.add(uuid)
            uuids.append(uuid)
    return uuids


def split_top_level(text):
    """按顶层逗号切分 JavaScript 参数或数组元素。"""
    parts = []
    start = 0
    depth = 0
    quote = ""
    escaped = False
    for index, char in enumerate(text):
        if quote:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                quote = ""
            continue
        if char in {'"', "'"}:
            quote = char
        elif char in "([{":
            depth += 1
        elif char in ")]}":
            depth -= 1
        elif char == "," and depth == 0:
            parts.append(text[start:index].strip())
            start = index + 1
    tail = text[start:].strip()
    if tail:
        parts.append(tail)
    return parts


def js_literal_value(token):
    """把简单 JavaScript 字面量转成 Python 值。"""
    token = token.strip()
    if not token or token == "void 0":
        return ""
    if token in {"true", "false"}:
        return token == "true"
    if token == "null":
        return ""
    if token.startswith('"') and token.endswith('"'):
        try:
            return json.loads(token)
        except json.JSONDecodeError:
            return html.unescape(token[1:-1])
    if re.fullmatch(r"-?\d+(?:\.\d+)?", token):
        return int(token) if "." not in token else float(token)
    return None


def nuxt_arg_map(text):
    """解析 window.__NUXT__ IIFE 的形参和实参映射。

    列表页为了压缩体积，会把公司名、行业、规模等值写成变量；例如
    `cname:c`，真正的 `c` 在 IIFE 末尾参数里才是 `"阿里巴巴"`。
    """
    match = re.search(
        r"window\.__NUXT__=\(function\((.*?)\)\{return .*?\}\((.*)\)\);</script>",
        text,
        flags=re.S,
    )
    if not match:
        return {}
    names = [name.strip() for name in match.group(1).split(",")]
    values = split_top_level(match.group(2))
    mapping = {}
    for name, token in zip(names, values):
        value = js_literal_value(token)
        mapping[name] = "" if value is None else value
    return mapping


def resolve_js_value(token, mapping):
    """解析列表页对象里的字段值，支持变量、字符串、数字和简单数组。"""
    token = token.strip()
    literal = js_literal_value(token)
    if literal is not None:
        return literal
    if token in mapping:
        return mapping[token]
    if token.startswith("[") and token.endswith("]"):
        values = []
        for item in split_top_level(token[1:-1]):
            value = resolve_js_value(item, mapping)
            if value not in ("", None, False):
                values.append(str(value))
        return values
    return ""


def object_fragment_around_uuid(text, uuid):
    """找到列表页里包含某个 uuid 的岗位对象片段。"""
    marker = f'uuid:"{uuid}"'
    index = text.find(marker)
    if index < 0:
        return ""
    start = text.rfind("{ad_type", 0, index)
    if start < 0:
        start = text.rfind("{", 0, index)
    if start < 0:
        return ""

    depth = 0
    quote = ""
    escaped = False
    for pos in range(start, len(text)):
        char = text[pos]
        if quote:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                quote = ""
            continue
        if char in {'"', "'"}:
            quote = char
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start : pos + 1]
    return ""


def top_level_object_fields(fragment):
    """解析对象片段的顶层字段，忽略嵌套对象内部字段。"""
    if not fragment.startswith("{") or not fragment.endswith("}"):
        return {}
    fields = {}
    body = fragment[1:-1]
    parts = split_top_level(body)
    for part in parts:
        match = re.match(r"\s*([A-Za-z_][A-Za-z0-9_]*)\s*:(.*)\s*$", part, flags=re.S)
        if match:
            fields[match.group(1)] = match.group(2).strip()
    return fields


def salary_desc_from_bounds(min_salary, max_salary):
    """根据列表页薪资上下限生成展示文本。"""
    try:
        low = int(min_salary or 0)
        high = int(max_salary or 0)
    except (TypeError, ValueError):
        low = high = 0
    if low <= 0 and high <= 0:
        return "薪资面议"
    if low > 0 and high > 0:
        return f"{low}-{high}/月"
    return f"{low or high}/月"


def normalize_signature_part(value):
    """规范化岗位名/公司名，用于快速查重。"""
    value = (value or "").lower()
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"[（）()【】\[\]《》<>·•\-_/|,，.。:：;；]+", "", value)
    return value


def job_signature(row):
    """用岗位名 + 公司名生成跨扫描稳定签名。"""
    title = normalize_signature_part(row.get("title", ""))
    company = normalize_signature_part(row.get("company", ""))
    if not title or not company:
        return row.get("uuid", "")
    return f"{title}::{company}"


def list_row_from_fragment(uuid, fragment, mapping):
    """把列表页岗位对象片段转换为与详情页兼容的 row。"""
    fields = top_level_object_fields(fragment)

    def value(name):
        if name not in fields:
            return ""
        return resolve_js_value(fields[name], mapping)

    min_salary = value("minsalary") or value("minsal")
    max_salary = value("maxsalary") or value("maxsal")
    tag_values = []
    for field in ("skill", "c_tags", "i_tags", "hope_you"):
        field_value = value(field)
        if isinstance(field_value, list):
            tag_values.extend(field_value)
        elif field_value:
            tag_values.append(str(field_value))
    list_tags = ", ".join(dict.fromkeys(tag_values))

    row = {
        "uuid": uuid,
        "title": value("name"),
        "company": value("cname"),
        "industry": value("industry"),
        "city": value("city"),
        "address": value("city"),
        "degree": value("degree"),
        "salary": salary_desc_from_bounds(min_salary, max_salary),
        "min_salary": str(min_salary or ""),
        "max_salary": str(max_salary or ""),
        "days_per_week": str(value("day") or ""),
        "months": str(value("month_num") or ""),
        "refresh_time": value("refresh"),
        "end_time": "",
        "url": BASE_DETAIL_URL.format(uuid=uuid),
        "description": list_tags,
        "parse_status": "list",
        "parse_warning": "",
        "scale": value("scale"),
        "list_tags": list_tags,
    }
    row["signature"] = job_signature(row)
    return row


def list_rows_from_page(text):
    """从列表页 HTML 直接解析岗位卡片信息。"""
    mapping = nuxt_arg_map(text)
    rows = []
    for uuid in unique_uuids(text):
        fragment = object_fragment_around_uuid(text, uuid)
        if not fragment:
            continue
        row = list_row_from_fragment(uuid, fragment, mapping)
        if row.get("title"):
            rows.append(row)
    return rows


def js_string_field(text, field):
    """从详情页 Nuxt 数据块里提取 x.<field> = "..." 形式的字符串字段。

    实习僧详情页服务端 HTML 里含有 window.__NUXT__ 数据，
    多数字段会以 j.iname、q.iname、q.info 等形式出现。
    """
    match = re.search(rf"[a-z]\.{re.escape(field)}=\"((?:\\.|[^\"\\])*)\";", text)
    if not match:
        return ""
    try:
        return json.loads(f'"{match.group(1)}"')
    except json.JSONDecodeError:
        return html.unescape(match.group(1))


def js_number_field(text, field):
    """从详情页 Nuxt 数据块里提取 x.<field> = 123 形式的数字字段。"""
    match = re.search(rf"[a-z]\.{re.escape(field)}=([0-9]+);", text)
    return match.group(1) if match else ""


def html_text_between(text, pattern):
    """按正则从 HTML 片段中提取文本，作为 Nuxt 字段缺失时的兜底。"""
    match = re.search(pattern, text, flags=re.S)
    if not match:
        return ""
    return clean_html(match.group(1))


def clean_html(raw):
    """把岗位描述中的 HTML 转成便于搜索和写入表格的纯文本。"""
    raw = raw.replace("\\/", "/")
    raw = re.sub(r"(?i)<\s*br\s*/?>", "\n", raw)
    raw = re.sub(r"(?i)</\s*p\s*>", "\n", raw)
    raw = re.sub(r"(?is)<script.*?</script>|<style.*?</style>", " ", raw)
    raw = re.sub(r"(?s)<[^>]+>", " ", raw)
    text = html.unescape(raw)
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\n\s+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_description_from_html(text):
    """从服务端渲染 HTML 中兜底提取职位描述正文。"""
    patterns = [
        r'<div class="job_til"[^>]*>\s*职位描述：?\s*</div>\s*'
        r'(?:<!---->\s*)?<div class="job_part"[^>]*>(.*?)</div>\s*</div>',
        r'<div class="job_detail"[^>]*>(.*?)</div>',
    ]
    for pattern in patterns:
        description = html_text_between(text, pattern)
        if description:
            return description
    return ""


def extract_description(text):
    """优先读取 Nuxt 数据，失败时从页面 HTML 兜底解析职位描述。"""
    info_html = js_string_field(text, "info")
    description = clean_html(info_html)
    if description:
        return description, "nuxt", ""

    description = extract_description_from_html(text)
    if description:
        return description, "html_fallback", ""

    has_description_marker = any(
        marker in text
        for marker in ("职位描述", "岗位职责", "职责描述", "职位要求", "任职要求")
    )
    if has_description_marker:
        return "", "missing", "页面含职位描述标记，但程序未能解析出正文，请复核解析规则"
    return "", "missing", "详情页未解析出正文，且页面未发现常见职位描述标记"


def fallback_title_company(text):
    """当 Nuxt 字段缺失时，从 <title> 中兜底解析岗位名和公司名。"""
    title = ""
    company = ""
    match = re.search(r"<title>(.*?)</title>", text, flags=re.S)
    if match:
        title_text = clean_html(match.group(1))
        if "校招-" in title_text:
            title = title_text.split("校招-")[0]
            company_part = title_text.split("校招-", 1)[1]
            company = company_part.split("校园招聘")[0].split("招聘")[0].split("-")[0]
        else:
            title = title_text.split("实习招聘-")[0].replace("实习生实习招聘", "实习生")
            parts = title_text.split("实习招聘-")
            if len(parts) > 1:
                company = parts[1].split("实习生招聘")[0]
    return title, company


def extract_detail(uuid, text):
    """把一个岗位详情页 HTML 解析成结构化 dict。"""
    description, parse_status, parse_warning = extract_description(text)
    title = js_string_field(text, "iname")
    company = js_string_field(text, "cname")
    fallback_title, fallback_company = fallback_title_company(text)
    url = js_string_field(text, "url") or BASE_DETAIL_URL.format(uuid=uuid)
    address = js_string_field(text, "address")
    if not address:
        address = html_text_between(
            text,
            r'<span class="com_position"[^>]*>(.*?)</span>',
        )
    city = js_string_field(text, "city")
    if not city and "上海" in address:
        city = "上海"
    # 公司规模字段在 Nuxt 数据块里可能以 k.scale / j.scale / i.scale 等形式出现，
    # 不同页面的变量前缀不一致，这里匹配任意单字母前缀。
    scale_match = re.search(r'[a-z]\.scale\s*=\s*"((?:\\.|[^"\\])*)";', text)
    scale = ""
    if scale_match:
        try:
            scale = json.loads(f'"{scale_match.group(1)}"')
        except json.JSONDecodeError:
            scale = html.unescape(scale_match.group(1))
    return {
        "uuid": uuid,
        "title": title or fallback_title,
        "company": company or fallback_company,
        "industry": js_string_field(text, "industry"),
        "city": city,
        "address": address,
        "degree": js_string_field(text, "degree"),
        "salary": js_string_field(text, "salary_desc"),
        "min_salary": js_number_field(text, "minsal") or js_number_field(text, "minsalary"),
        "max_salary": js_number_field(text, "maxsal") or js_number_field(text, "maxsalary"),
        "days_per_week": js_number_field(text, "day"),
        "months": js_number_field(text, "month"),
        "refresh_time": js_string_field(text, "refresh"),
        "end_time": js_string_field(text, "endtime"),
        "url": url,
        "description": description,
        "parse_status": parse_status,
        "parse_warning": parse_warning,
        "scale": scale,
    }


def merge_detail_over_list(list_row, detail_row):
    """用详情页字段补强列表页字段，列表页已有的公司/规模/薪资作为兜底保留。"""
    merged = dict(list_row)
    for key, value in detail_row.items():
        if value not in ("", None):
            merged[key] = value
    merged["signature"] = list_row.get("signature") or job_signature(merged)
    merged["list_tags"] = list_row.get("list_tags", "")
    return merged


def init_cache(path):
    """初始化 SQLite 缓存。

    缓存分两张表：
    - list_pages：列表页 HTML，按完整 URL 缓存，避免“一天内”和“一周内”互相污染。
    - details：详情页 HTML，按岗位 uuid 缓存。

    好处：
    - 中途断网可重跑续上。
    - 后续只改关键词规则时，不需要重新访问网站。
    """
    conn = sqlite3.connect(path)
    conn.execute(
        """
        create table if not exists list_pages (
            cache_key text primary key,
            page integer not null,
            url text not null,
            fetched_at text not null,
            html text not null
        )
        """
    )
    conn.execute(
        """
        create table if not exists details (
            uuid text primary key,
            url text not null,
            fetched_at text not null,
            html_sha256 text not null,
            html text not null
        )
        """
    )
    conn.execute(
        """
        create table if not exists scan_state (
            state_key text primary key,
            updated_at text not null,
            uuids_json text not null
        )
        """
    )
    conn.commit()
    return conn


def cleanup_cache(conn, retention_days=7):
    """删除超过 retention_days 天的网页缓存记录，并回收磁盘空间。

    这里只清理 list_pages 和 details，不清理 scan_state；
    因此旧网页 HTML 不会无限增长，每日增量基线也不会受影响。
    """
    import datetime as dt

    cutoff = (dt.datetime.now() - dt.timedelta(days=retention_days)).isoformat(
        timespec="seconds"
    )
    deleted_pages = conn.execute(
        "delete from list_pages where fetched_at < ?", (cutoff,)
    ).rowcount
    deleted_details = conn.execute(
        "delete from details where fetched_at < ?", (cutoff,)
    ).rowcount
    if deleted_pages or deleted_details:
        conn.commit()
        conn.execute("vacuum")
        print(
            f"缓存清理：删除 {deleted_pages} 个列表页、{deleted_details} 个详情页"
            f"（{retention_days} 天前），已回收磁盘空间",
            flush=True,
        )


def load_previous_snapshot(conn, state_key=SCAN_STATE_KEY):
    """读取上次扫描保存的岗位签名快照。"""
    row = conn.execute(
        "select updated_at, uuids_json from scan_state where state_key = ?",
        (state_key,),
    ).fetchone()
    if not row:
        return None, []
    updated_at, raw = row
    try:
        uuids = json.loads(raw)
    except json.JSONDecodeError:
        uuids = []
    return updated_at, [uuid for uuid in uuids if isinstance(uuid, str)]


def save_current_snapshot(conn, signatures, state_key=SCAN_STATE_KEY):
    """保存本轮列表页看到的岗位签名，供下次启动时做增量比对。"""
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        """
        replace into scan_state(state_key, updated_at, uuids_json)
        values (?, ?, ?)
        """,
        (state_key, now, json.dumps(list(dict.fromkeys(signatures)), ensure_ascii=False)),
    )
    conn.commit()
    return now


def cached_page(conn, page, use_cache=False):
    """读取或下载列表页。返回 (html, 是否来自缓存)。"""
    url = list_url(page)
    cache_key = hashlib.sha256(url.encode("utf-8")).hexdigest()
    if use_cache:
        row = conn.execute("select html from list_pages where cache_key = ?", (cache_key,)).fetchone()
        if row:
            return row[0], True
    text = fetch(url)
    conn.execute(
        """
        replace into list_pages(cache_key, page, url, fetched_at, html)
        values (?, ?, ?, ?, ?)
        """,
        (cache_key, page, url, datetime.now().isoformat(timespec="seconds"), text),
    )
    conn.commit()
    return text, False


def cached_detail(conn, uuid, use_cache=True):
    """读取或下载详情页。返回 (html, 是否来自缓存)。"""
    url = BASE_DETAIL_URL.format(uuid=uuid)
    if use_cache:
        row = conn.execute("select html from details where uuid = ?", (uuid,)).fetchone()
        if row:
            return row[0], True
    text = fetch(url)
    digest = hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()
    conn.execute(
        """
        replace into details(uuid, url, fetched_at, html_sha256, html)
        values (?, ?, ?, ?, ?)
        """,
        (uuid, url, datetime.now().isoformat(timespec="seconds"), digest, text),
    )
    conn.commit()
    return text, False


def haystack(row):
    """把需要参与匹配的字段合并成一段文本。"""
    return "\n".join(
        [
            row.get("title", ""),
            row.get("company", ""),
            row.get("industry", ""),
            row.get("list_tags", ""),
            row.get("description", ""),
        ]
    )


def find_terms(text, terms):
    """大小写不敏感地查找关键词。"""
    lower = text.lower()
    return [term for term in terms if term.lower() in lower]


def classify(row):
    """对岗位做分类。

    分类顺序：
    1. 先看是否明确出现社会学/社科类词，命中则直接归为明确匹配——即使标题含"工程师"或"HR"也不排除。
    2. 未命中直接匹配的，排除工程师、HR 等不相关方向。
    3. 剩余岗位按 APPROX_RULES 累积分数。
    """
    text = haystack(row)

    # 第一优先级：直接匹配。命中了就不再做任何排除，直接保留。
    explicit_hits = find_terms(text, EXPLICIT_TERMS)
    if explicit_hits:
        return build_match(row, "直接匹配：社会学/用户研究", explicit_hits, 100, {"直接匹配"})

    # 以下排除只对非直接匹配生效。
    # 标题含"工程师"的岗位属于工科/技术方向，直接排除。
    if "工程师" in row.get("title", ""):
        return None

    # 排除 HR/人力资源/招聘/猎头等相关岗位。
    # 只检查标题和公司名，不检查正文描述，避免误伤提到"招聘"一词的非HR岗位。
    _hr_text = "\n".join([row.get("title", ""), row.get("company", "")]).lower()
    _hr_keywords = [
        "人力资源", "人事", "招聘", "猎头", "hr", "h r",
        "员工关系", "薪酬绩效", "组织发展", "人才发展",
        "hrbp", "ssc", "coe", "人力",
    ]
    if any(kw in _hr_text for kw in _hr_keywords):
        return None

    # 近似匹配。
    matched = []
    categories = []
    score = 0
    lower = text.lower()
    for term, weight, category in APPROX_RULES:
        if term.lower() in lower:
            matched.append(term)
            categories.append(category)
            score += weight

    unique_categories = sorted(set(categories))
    has_core = any(category in CORE_CATEGORIES for category in unique_categories)
    if score < 3 or not has_core:
        return None

    # 宽松模式：不再对长描述做分数封顶，保留原始打分。
    return build_match(row, "近似匹配：可迁移社会学能力", matched, score, unique_categories)


def build_match(row, category, hits, score, categories):
    """把匹配结果补充到原始岗位字段中。"""
    text = haystack(row)
    level = "高" if score >= 14 or category.startswith("直接") else "中" if score >= 9 else "低"
    result = dict(row)
    result["category"] = category
    result["match_level"] = level
    result["score"] = score
    result["matched_keywords"] = ", ".join(dict.fromkeys(hits))
    result["fit_reason"] = fit_reason(categories)
    result["evidence"] = evidence(text, hits)
    return result


def fit_reason(categories):
    """根据命中的能力类别生成中文推荐理由。"""
    categories = set(categories)
    reasons = []
    if "直接匹配" in categories:
        reasons.append("岗位文字直接提到社会学、社科、用户研究或用户洞察。")
    if categories & {"调研方法", "社会学/人类学方法"}:
        reasons.append("需要访谈、问卷、调研、定性/定量等研究方法，可迁移社会学博士训练。")
    if "用户/消费者洞察" in categories:
        reasons.append("涉及用户、消费者或人群洞察，适合用社会学视角做需求与行为分析。")
    if categories & {"研究分析", "研究写作", "研究/咨询"}:
        reasons.append("需要资料搜集、结构化分析和报告写作，匹配博士阶段的研究与表达能力。")
    if "政策/公共议题" in categories:
        reasons.append("涉及政策、公共议题或社会环境分析，社会学背景有解释优势。")
    if "社会议题" in categories:
        reasons.append("涉及社区、公益、性别、可持续等社会议题，社会学背景有内容优势。")
    if "AI/人文评估" in categories:
        reasons.append("涉及AI行为评估、表达训练、人文视角或跨文化理解，适合用社会科学训练判断人机互动、语境与价值问题。")
    if "传播/社会观察" in categories or "文化/传播" in categories:
        reasons.append("涉及传播、文化或舆情观察，可发挥社会分析和文本分析能力。")
    if "定量分析" in categories:
        reasons.append("包含数据/统计分析要求，适合量化研究经验迁移。")
    return " ".join(reasons) or "岗位内容与社会学博士的研究、分析或写作能力存在可迁移空间。"


def evidence(text, hits, width=65):
    """截取命中关键词附近的证据片段，方便人工复核。"""
    snippets = []
    lower = text.lower()
    for hit in hits[:6]:
        index = lower.find(hit.lower())
        if index < 0:
            continue
        start = max(0, index - width)
        end = min(len(text), index + len(hit) + width)
        snippet = text[start:end].replace("\n", " ")
        snippets.append(snippet)
    return " | ".join(snippets)


def parse_money_amount(value, default_k_unit=False):
    """把 10k、10千、1万、10000 等薪资片段统一换算成人民币整数。"""
    number = float(value[0])
    unit = value[1].lower()
    if unit in {"k", "千"} or (default_k_unit and number < 1000):
        number *= 1000
    elif unit == "万":
        number *= 10000
    return int(number)


def salary_floor(row):
    """返回岗位薪资下限；薪资面议或解析失败时返回 None。"""
    for field in ("min_salary", "minsalary"):
        value = row.get(field)
        if value not in (None, ""):
            try:
                parsed = int(value)
            except (TypeError, ValueError):
                continue
            if parsed > 0:
                return parsed

    salary = row.get("salary", "")
    if not salary or "面议" in salary:
        return None
    matches = re.findall(r"(\d+(?:\.\d+)?)\s*([kK千萬万]?)", salary)
    if not matches:
        return None
    default_k_unit = bool(re.search(r"[kK千]", salary))
    amounts = [parse_money_amount(match, default_k_unit=default_k_unit) for match in matches]
    return min(amounts) if amounts else None


def scale_floor(scale):
    """返回公司规模区间下限；未知或无法解析时返回 None。"""
    if not scale:
        return None
    if "少于" in scale or "小于" in scale:
        return 0
    matches = re.findall(r"(\d+(?:\.\d+)?)\s*([萬万]?)", scale)
    if not matches:
        return None
    amounts = []
    for number, unit in matches:
        value = float(number)
        if unit in {"万", "萬"}:
            value *= 10000
        amounts.append(int(value))
    return min(amounts) if amounts else None


def hard_filter_reasons(row, min_salary, min_company_size):
    """返回硬筛剔除原因；空列表表示保留进入匹配分类。"""
    reasons = []
    floor = salary_floor(row)
    if floor is not None and floor < min_salary:
        reasons.append(f"薪资下限 {floor} < {min_salary}")

    company_floor = scale_floor(row.get("scale", ""))
    if company_floor is not None and company_floor < min_company_size:
        reasons.append(f"公司规模下限 {company_floor} < {min_company_size}")
    return reasons


def is_shanghai_detail(row):
    """只接受详情页工作地点明确包含上海的岗位。"""
    location_text = "\n".join([row.get("address", ""), row.get("city", "")])
    return "上海" in location_text


def scan(args):
    """执行完整扫描：列表页快扫 -> 新职位精读 -> 分类。"""
    conn = init_cache(args.cache)
    cleanup_cache(conn, retention_days=args.cache_retention_days)
    previous_snapshot_at, previous_signatures = load_previous_snapshot(conn)
    previous_key_set = set(previous_signatures)
    previous_uuid_set = {key for key in previous_signatures if key.startswith("inn_")}

    list_rows = []
    empty_pages = 0
    for page in range(1, args.pages + 1):
        text, from_cache = cached_page(
            conn,
            page,
            use_cache=args.use_list_cache and not args.refresh,
        )
        page_rows = list_rows_from_page(text)
        if not page_rows:
            # 兜底：如果列表页结构变化，至少保留 uuid，后续仍可进入详情页。
            page_rows = [
                {
                    "uuid": uuid,
                    "title": "",
                    "company": "",
                    "industry": "",
                    "city": "",
                    "address": "",
                    "degree": "",
                    "salary": "",
                    "min_salary": "",
                    "max_salary": "",
                    "days_per_week": "",
                    "months": "",
                    "refresh_time": "",
                    "end_time": "",
                    "url": BASE_DETAIL_URL.format(uuid=uuid),
                    "description": "",
                    "parse_status": "list_uuid_only",
                    "parse_warning": "列表页结构化字段解析失败，仅提取到岗位 ID",
                    "scale": "",
                    "list_tags": "",
                    "signature": uuid,
                }
                for uuid in unique_uuids(text)
            ]
        print(
            f"page {page}/{args.pages}: {len(page_rows)} jobs"
            f"{' cache' if from_cache else ''}",
            flush=True,
        )
        if not page_rows:
            empty_pages += 1
            if empty_pages >= args.stop_after_empty_pages:
                break
        else:
            empty_pages = 0
        list_rows.extend(page_rows)
        if not from_cache:
            time.sleep(args.delay)

    ordered_list_rows = []
    seen_signatures = set()
    for row in list_rows:
        signature = row.get("signature") or job_signature(row)
        if signature in seen_signatures:
            continue
        row["signature"] = signature
        seen_signatures.add(signature)
        ordered_list_rows.append(row)

    current_signatures = [row["signature"] for row in ordered_list_rows]
    new_list_rows = [
        row
        for row in ordered_list_rows
        if row["signature"] not in previous_key_set and row.get("uuid", "") not in previous_uuid_set
    ]
    if args.scan_current:
        new_list_rows = ordered_list_rows
    if args.reset_baseline:
        new_list_rows = []

    print(
        f"snapshot_previous_at={previous_snapshot_at or 'none'} "
        f"current_jobs={len(ordered_list_rows)} new_jobs={len(new_list_rows)} "
        f"seen_before={len(ordered_list_rows) - len(new_list_rows)}"
        f"{' scan_current' if args.scan_current else ''}"
        f"{' baseline_reset' if args.reset_baseline else ''}",
        flush=True,
    )

    all_rows = []
    matched_rows = []
    failed_uuids = []
    parse_warning_rows = []
    filtered_out_rows = []
    list_skipped_rows = []
    location_filtered_rows = []
    detail_attempted = 0
    for index, list_row in enumerate(new_list_rows, start=1):
        uuid = list_row.get("uuid", "")
        row = dict(list_row)
        row["is_new"] = "yes"
        filter_reasons = hard_filter_reasons(
            row,
            min_salary=args.min_salary,
            min_company_size=args.min_company_size,
        )
        row["filter_reasons"] = "; ".join(filter_reasons)

        list_match = None if filter_reasons else classify(row)
        if filter_reasons:
            filtered_out_rows.append(row)
            all_rows.append(row)
            print(
                f"list {index}/{len(new_list_rows)}: {uuid} filtered",
                flush=True,
            )
            continue

        detail_attempted += 1
        # 单个详情页抓取失败（超时、限流、SSL 被掐断等）不应该让整轮崩溃。
        try:
            detail_html, from_cache = cached_detail(conn, uuid, use_cache=not args.refresh)
            detail_row = extract_detail(uuid, detail_html)
            row = merge_detail_over_list(row, detail_row)
        except (URLError, TimeoutError, ConnectionError, OSError) as error:
            failed_uuids.append(uuid)
            row["parse_status"] = "list_fallback"
            row["parse_warning"] = f"详情页抓取失败，使用列表页字段：{error}"
            from_cache = False

        if not is_shanghai_detail(row):
            row["filter_reasons"] = "详情页工作地点不含上海"
            location_filtered_rows.append(row)
            all_rows.append(row)
            print(
                f"detail {detail_attempted}: {uuid} "
                f"{'cache ' if from_cache else ''}location-filtered",
                flush=True,
            )
            if not from_cache:
                time.sleep(args.delay)
            continue

        if row.get("parse_warning"):
            parse_warning_rows.append(row)
        row["is_new"] = "yes"
        row["filter_reasons"] = ""
        all_rows.append(row)
        matched = classify(row) or list_match
        if matched:
            matched["is_new"] = "yes"
            matched["signature"] = row.get("signature", "")
            matched["list_tags"] = row.get("list_tags", "")
            matched["filter_reasons"] = ""
            matched_rows.append(matched)
        print(
            f"detail {detail_attempted}: {uuid} "
            f"{'cache ' if from_cache else ''}"
            f"{'matched' if matched else 'skip'}",
            flush=True,
        )
        if not from_cache:
            time.sleep(args.delay)

    if failed_uuids:
        print(
            f"本轮有 {len(failed_uuids)} 个详情页抓取失败已跳过："
            f"{', '.join(failed_uuids)}",
            flush=True,
        )
    if parse_warning_rows:
        warning_uuids = ", ".join(row["uuid"] for row in parse_warning_rows[:20])
        more = "..." if len(parse_warning_rows) > 20 else ""
        print(
            f"本轮有 {len(parse_warning_rows)} 个详情页未能解析正文，"
            f"请复核：{warning_uuids}{more}",
            flush=True,
        )
    if args.reset_baseline:
        snapshot_signatures = current_signatures
    else:
        snapshot_signatures = list(dict.fromkeys(previous_signatures + current_signatures))
    current_snapshot_at = save_current_snapshot(conn, snapshot_signatures)
    conn.close()
    metadata = {
        "previous_snapshot_at": previous_snapshot_at,
        "current_snapshot_at": current_snapshot_at,
        "current_jobs": len(ordered_list_rows),
        "previous_jobs": len(previous_signatures),
        "snapshot_jobs": len(snapshot_signatures),
        "new_jobs": len(new_list_rows),
        "seen_before": len(ordered_list_rows) - len(new_list_rows),
        "filtered_out": len(filtered_out_rows),
        "list_skipped": len(list_skipped_rows),
        "location_filtered": len(location_filtered_rows),
        "detail_attempted": detail_attempted,
        "failed_details": len(failed_uuids),
        "min_salary": args.min_salary,
        "min_company_size": args.min_company_size,
        "baseline_reset": args.reset_baseline,
        "scan_current": args.scan_current,
    }
    return all_rows, matched_rows, metadata, ordered_list_rows


def write_csv(rows, path):
    """写 UTF-8 with BOM 的 CSV，方便 Excel 直接打开中文不乱码。"""
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def write_json(data, path):
    """写 JSON，方便其他程序或 AI 继续处理。"""
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pages", type=int, default=500, help="要抓取的列表页页数上限；连续空页会提前停止")
    parser.add_argument("--delay", type=float, default=0.15, help="每次新请求后的等待秒数，避免访问过快")
    parser.add_argument("--cache", type=Path, default=Path("shixiseng_school_cache.sqlite3"), help="SQLite 缓存和增量快照文件")
    parser.add_argument("--outdir", type=Path, default=Path("outputs"), help="输出目录")
    parser.add_argument("--refresh", action="store_true", help="忽略缓存，重新下载网页")
    parser.add_argument("--use-list-cache", action="store_true", help="调试用：允许复用列表页缓存；日常扫描不建议开启")
    parser.add_argument("--detail-all-new", action="store_true", help="兼容旧参数；当前默认已对所有新增且通过硬筛的岗位打开详情页确认地点")
    parser.add_argument("--scan-current", action="store_true", help="筛选当前列表中的全部岗位，同时保留累计基线用于明天对比")
    parser.add_argument("--reset-baseline", action="store_true", help="只用当前列表建立增量基线，不把当前既有岗位当作新增展示")
    parser.add_argument("--stop-after-empty-pages", type=int, default=2, help="连续空列表页达到该数量后提前停止")
    parser.add_argument("--cache-retention-days", type=int, default=2, help="网页缓存保留天数，超过的自动清理；不影响基线快照（默认 2 天）")
    parser.add_argument("--min-salary", type=int, default=10000, help="已知薪资下限低于该值的岗位会被剔除（默认 10000）")
    parser.add_argument("--min-company-size", type=int, default=1000, help="已知公司规模区间下限低于该值的岗位会被剔除（默认 1000）")
    args = parser.parse_args()

    started = datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir = args.outdir / started
    outdir.mkdir(parents=True, exist_ok=True)

    print(f"started={started}", flush=True)
    all_rows, matched_rows, metadata, list_rows = scan(args)

    explicit = [r for r in matched_rows if r["category"].startswith("直接")]
    approximate = [r for r in matched_rows if r["category"].startswith("近似")]
    approximate.sort(key=lambda r: (-int(r["score"]), r["refresh_time"], r["title"]))

    write_csv(list_rows, outdir / "current_list_jobs.csv")
    write_csv(all_rows, outdir / "new_jobs.csv")
    write_csv(explicit, outdir / "explicit_matches.csv")
    write_csv(approximate, outdir / "approximate_matches.csv")
    write_csv(explicit + approximate, outdir / "all_matches.csv")
    write_json(
        {
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "pages": args.pages,
            "list_url_page_1": list_url(1),
            "scanned_jobs": len(all_rows),
            "scope": "上海校招/全职；列表页快速查重和薪资/规模硬筛，详情页确认上海地点后输出",
            **metadata,
            "explicit_matches": len(explicit),
            "approximate_matches": len(approximate),
            "explicit": explicit,
            "approximate": approximate,
        },
        outdir / "matches.json",
    )
    print(
        f"new_scanned_jobs={len(all_rows)} explicit={len(explicit)} "
        f"approximate={len(approximate)} filtered_out={metadata['filtered_out']} "
        f"detail_attempted={metadata['detail_attempted']} "
        f"outdir={outdir}",
        flush=True,
    )


if __name__ == "__main__":
    main()
