#!/usr/bin/env python3
import argparse
import csv
import html
import json
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import ProxyHandler, Request, build_opener


BASE_LIST_URL = "https://www.shixiseng.com/interns"
BASE_DETAIL_URL = "https://www.shixiseng.com/intern/{uuid}"

KEYWORDS = [
    "社会学",
    "社会科学",
    "人类学",
    "民族志",
    "田野",
    "定性研究",
    "访谈",
    "问卷",
    "调研",
    "用户研究",
    "用户洞察",
    "消费者洞察",
    "用户画像",
    "市场研究",
    "行业研究",
    "数据分析",
    "社区",
    "公益",
    "性别",
    "SDG",
    "可持续发展",
]


def fetch(url, timeout=20):
    request = Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/125.0 Safari/537.36"
            )
        },
    )
    opener = build_opener(ProxyHandler({}))
    with opener.open(request, timeout=timeout) as response:
        return response.read().decode("utf-8", errors="replace")


def list_url(page):
    query = {
        "page": page,
        "type": "intern",
        "keyword": "",
        "area": "",
        "months": "",
        "days": "",
        "degree": "",
        "official": "",
        "enterprise": "",
        "salary": "-0",
        "publishTime": "wek",
        "sortType": "",
        "city": "上海",
        "internExtend": "",
    }
    return f"{BASE_LIST_URL}?{urlencode(query)}"


def unique_uuids(text):
    seen = set()
    uuids = []
    for uuid in re.findall(r"inn_[a-z0-9]+", text):
        if uuid not in seen:
            seen.add(uuid)
            uuids.append(uuid)
    return uuids


def js_string_field(text, field):
    match = re.search(rf"j\.{re.escape(field)}=\"((?:\\.|[^\"\\])*)\";", text)
    if not match:
        return ""
    try:
        return json.loads(f'"{match.group(1)}"')
    except json.JSONDecodeError:
        return html.unescape(match.group(1))


def js_number_field(text, field):
    match = re.search(rf"j\.{re.escape(field)}=([0-9]+);", text)
    return match.group(1) if match else ""


def clean_html(raw):
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


def extract_detail(uuid, text):
    info_html = js_string_field(text, "info")
    description = clean_html(info_html)
    title = js_string_field(text, "iname")
    company = js_string_field(text, "cname")
    url = js_string_field(text, "url") or BASE_DETAIL_URL.format(uuid=uuid)
    return {
        "uuid": uuid,
        "title": title,
        "company": company,
        "industry": js_string_field(text, "industry"),
        "city": js_string_field(text, "city"),
        "address": js_string_field(text, "address"),
        "degree": js_string_field(text, "degree"),
        "salary": js_string_field(text, "salary_desc"),
        "days_per_week": js_number_field(text, "day"),
        "months": js_number_field(text, "month"),
        "refresh_time": js_string_field(text, "refresh"),
        "end_time": js_string_field(text, "endtime"),
        "url": url,
        "description": description,
    }


def find_matches(row):
    haystack = "\n".join(
        [
            row.get("title", ""),
            row.get("company", ""),
            row.get("industry", ""),
            row.get("description", ""),
        ]
    )
    return [keyword for keyword in KEYWORDS if keyword.lower() in haystack.lower()]


def evidence(text, hits, width=55):
    snippets = []
    for hit in hits[:4]:
        index = text.lower().find(hit.lower())
        if index < 0:
            continue
        start = max(0, index - width)
        end = min(len(text), index + len(hit) + width)
        snippet = text[start:end].replace("\n", " ")
        snippets.append(snippet)
    return " | ".join(snippets)


def scan(pages, delay):
    uuids = []
    for page in range(1, pages + 1):
        text = fetch(list_url(page))
        page_uuids = unique_uuids(text)
        print(f"page {page}: {len(page_uuids)} unique ids", flush=True)
        uuids.extend(page_uuids)
        time.sleep(delay)

    seen = set()
    ordered_uuids = []
    for uuid in uuids:
        if uuid not in seen:
            seen.add(uuid)
            ordered_uuids.append(uuid)

    rows = []
    for index, uuid in enumerate(ordered_uuids, start=1):
        detail_html = fetch(BASE_DETAIL_URL.format(uuid=uuid))
        row = extract_detail(uuid, detail_html)
        hits = find_matches(row)
        if hits:
            text = "\n".join([row["title"], row["industry"], row["description"]])
            row["matched_keywords"] = ", ".join(hits)
            row["evidence"] = evidence(text, hits)
            rows.append(row)
        print(f"detail {index}/{len(ordered_uuids)}: {uuid} hits={len(hits)}", flush=True)
        time.sleep(delay)

    return ordered_uuids, rows


def write_csv(rows, path):
    fields = [
        "matched_keywords",
        "title",
        "company",
        "industry",
        "city",
        "address",
        "salary",
        "days_per_week",
        "months",
        "degree",
        "refresh_time",
        "end_time",
        "url",
        "evidence",
        "description",
        "uuid",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pages", type=int, default=3)
    parser.add_argument("--delay", type=float, default=0.3)
    parser.add_argument("--out", default="shixiseng_sociology_matches_3pages.csv")
    args = parser.parse_args()

    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"started: {started_at}")
    uuids, rows = scan(args.pages, args.delay)
    output = Path(args.out)
    write_csv(rows, output)
    print(f"scanned_jobs={len(uuids)} matched_jobs={len(rows)} output={output}")


if __name__ == "__main__":
    main()
