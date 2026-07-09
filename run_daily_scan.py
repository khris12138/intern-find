#!/usr/bin/env python3
"""
每日扫描入口。

执行内容：
1. 抓取实习僧“上海 + 校招/全职”岗位。
2. 与上次扫描结果比对，只展示上次以来新增的岗位。
3. 打开新增岗位详情页，按薪资、公司规模和能力匹配筛选适合社会学博士的岗位。
3. 生成 CSV、JSON 和 Excel。

推荐由 launchd 在每天 11:00 调用：
    python3 run_daily_scan.py
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def run(command):
    """在项目目录运行子命令，失败时直接抛出清晰错误。"""
    print("+ " + " ".join(str(part) for part in command), flush=True)
    subprocess.run(command, cwd=ROOT, check=True)


def newest_run_dir(outputs):
    """返回 outputs 下最新的时间戳扫描目录。"""
    candidates = [
        path
        for path in outputs.iterdir()
        if path.is_dir() and (path / "matches.json").exists()
    ]
    if not candidates:
        raise RuntimeError("没有找到扫描输出目录")
    return max(candidates, key=lambda path: path.name)


def unique_excel_path(outputs, timestamp):
    """生成不会覆盖既有结果的 Excel 路径。"""
    base = outputs / f"上海校招新增岗位筛选_{timestamp}.xlsx"
    if not base.exists():
        return base
    counter = 2
    while True:
        candidate = outputs / f"上海校招新增岗位筛选_{timestamp}_{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def cleanup_old_outputs(outdir: Path, today: str):
    """删除今天之前的时间戳目录和带时间戳的 Excel 文件。"""
    deleted = []
    for path in outdir.iterdir():
        if path.name.startswith("."):
            continue
        m = re.match(r"^(\d{8})_\d{6}$", path.name)
        if m and path.is_dir() and m.group(1) < today:
            shutil.rmtree(path)
            deleted.append(path.name)
            continue
        m2 = re.match(r"^.*_(\d{8})_\d{6}.*\.xlsx$", path.name)
        if m2 and path.is_file() and m2.group(1) < today:
            path.unlink()
            deleted.append(path.name)
    if deleted:
        print(f"已清理旧文件：{deleted}", flush=True)


def open_urls_in_browser(run_dir: Path):
    """在默认浏览器中打开直接匹配和高等级近似匹配职位链接。"""
    data = json.loads((run_dir / "matches.json").read_text())
    selected_jobs = list(data.get("explicit", []))
    selected_jobs.extend(
        job
        for job in data.get("approximate", [])
        if job.get("match_level") == "高"
    )

    urls = list(dict.fromkeys(job["url"] for job in selected_jobs if job.get("url")))
    if not urls:
        print("没有直接匹配或近似匹配-高的职位链接。", flush=True)
        return
    print(f"在浏览器中打开 {len(urls)} 个直接匹配和近似匹配-高链接...", flush=True)
    subprocess.run(["open"] + urls)


def update_latest_copy(source):
    """尝试更新 latest 副本；失败不影响本次带时间戳的正式结果。"""
    latest = source.parent / "上海校招新增岗位筛选_latest.xlsx"
    try:
        shutil.copy2(source, latest)
    except OSError as exc:
        print(f"warning: latest copy not updated: {exc}", flush=True)
    return latest


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pages", type=int, default=50, help="列表页扫描上限")
    parser.add_argument("--delay", type=float, default=0.3, help="每次新请求后的等待秒数")
    parser.add_argument("--outdir", type=Path, default=ROOT / "outputs", help="输出根目录")
    parser.add_argument("--cache", type=Path, default=ROOT / "shixiseng_school_cache.sqlite3", help="校招增量扫描缓存和快照")
    parser.add_argument("--refresh", action="store_true", help="忽略缓存，重新下载网页")
    parser.add_argument("--open-urls", action="store_true", help="完成后在默认浏览器打开直接匹配和近似匹配-高链接")
    args = parser.parse_args()

    scan_command = [
        sys.executable,
        str(ROOT / "full_scan_once.py"),
        "--pages",
        str(args.pages),
        "--delay",
        str(args.delay),
        "--outdir",
        str(args.outdir),
        "--cache",
        str(args.cache),
    ]
    if args.refresh:
        scan_command.append("--refresh")

    started = datetime.now().strftime("%Y%m%d_%H%M%S")
    run(scan_command)

    run_dir = newest_run_dir(args.outdir)
    excel_path = unique_excel_path(args.outdir, started)
    run(
        [
            sys.executable,
            str(ROOT / "build_excel_report.py"),
            str(run_dir),
            str(excel_path),
        ]
    )
    latest_path = update_latest_copy(excel_path)
    print(f"scan_dir={run_dir}", flush=True)
    print(f"excel={excel_path}", flush=True)
    print(f"latest={latest_path}", flush=True)

    cleanup_old_outputs(args.outdir, datetime.now().strftime("%Y%m%d"))
    if args.open_urls:
        open_urls_in_browser(run_dir)


if __name__ == "__main__":
    main()
