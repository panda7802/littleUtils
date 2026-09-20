#!/usr/bin/env python3
"""Export all public videos from a Bilibili UP account to XLSX."""

from __future__ import annotations

import argparse
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, Response, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from social_excel import write_social_xlsx


VIDEO_API_MARKER = "/x/space/wbi/arc/search"


def console_safe(value: Any) -> str:
    text = str(value)
    encoding = sys.stdout.encoding or "utf-8"
    return text.encode(encoding, errors="replace").decode(encoding, errors="replace")


def normalize_cover(url: Any) -> str:
    value = str(url or "")
    return "https:" + value if value.startswith("//") else value


def parse_mid(account: str, profile_url: str | None) -> tuple[str, str]:
    candidate = profile_url or account
    match = re.search(r"space\.bilibili\.com/(\d+)", candidate)
    if match:
        mid = match.group(1)
        return mid, f"https://space.bilibili.com/{mid}/upload/video"
    if account.isdigit():
        return account, f"https://space.bilibili.com/{account}/upload/video"
    raise ValueError("B站账号请传 UID（纯数字）或使用 --profile-url 传空间链接。")


def video_to_row(video: dict[str, Any], mid: str) -> dict[str, Any]:
    aid = str(video.get("aid") or "")
    bvid = str(video.get("bvid") or "")
    created = video.get("created") or video.get("pubdate")
    published_at = datetime.fromtimestamp(created) if isinstance(created, (int, float)) else None
    video_id = bvid or (f"av{aid}" if aid else "")
    return {
        "bvid": bvid,
        "aid": aid,
        "author": str(video.get("author") or "").strip(),
        "mid": str(video.get("mid") or mid),
        "published_at": published_at,
        "title": str(video.get("title") or "").strip(),
        "description": str(video.get("description") or "").strip(),
        "duration": str(video.get("length") or ""),
        "play_count": video.get("play") if isinstance(video.get("play"), int) else "",
        "danmaku_count": video.get("video_review") if isinstance(video.get("video_review"), int) else "",
        "comment_count": video.get("comment") if isinstance(video.get("comment"), int) else "",
        "favorite_count": video.get("favorites") if isinstance(video.get("favorites"), int) else "",
        "video_url": f"https://www.bilibili.com/video/{video_id}" if video_id else "",
        "cover_url": normalize_cover(video.get("pic")),
    }


def collect_videos(page: Page, profile_url: str, mid: str, max_videos: int, login_wait: int) -> list[dict[str, Any]]:
    videos: dict[str, dict[str, Any]] = {}
    total_count: int | None = None
    last_change = time.monotonic()

    def on_response(response: Response) -> None:
        nonlocal total_count, last_change
        if VIDEO_API_MARKER not in response.url:
            return
        try:
            payload = response.json()
        except Exception:
            return
        data = payload.get("data") or {}
        listing = data.get("list") or {}
        items = listing.get("vlist") or data.get("vlist") or []
        page_info = data.get("page") or {}
        count = page_info.get("count")
        if isinstance(count, int):
            total_count = count
        before = len(videos)
        for item in items:
            if not isinstance(item, dict):
                continue
            key = str(item.get("bvid") or item.get("aid") or "")
            if key:
                videos[key] = item
        if len(videos) > before:
            last_change = time.monotonic()
            total_text = f"/{total_count}" if total_count is not None else ""
            print(f"      已获取 {len(videos)}{total_text} 条视频")

    page.on("response", on_response)
    print(f"[1/3] 打开B站空间：{profile_url}")
    page.goto(profile_url, wait_until="domcontentloaded", timeout=60_000)

    deadline = time.monotonic() + login_wait
    while not videos and time.monotonic() < deadline:
        page.wait_for_timeout(2_000)
        if time.monotonic() + 3 < deadline and not videos:
            print("      若页面要求登录或验证，请在浏览器中完成", end="\r")
    print()
    if not videos:
        page.remove_listener("response", on_response)
        raise RuntimeError("没有收到投稿列表，可能需要登录/验证，或该账号没有公开投稿。")

    print("[2/3] 翻页采集公开视频 ...")
    stalled_pages = 0
    while len(videos) < max_videos and (total_count is None or len(videos) < total_count):
        before = len(videos)
        clicked = False
        for selector in (
            "button:has-text('下一页')",
            "a:has-text('下一页')",
            ".be-pager-next",
            ".vui_pagenation--btn-side:last-child",
        ):
            locator = page.locator(selector).last
            try:
                if locator.count() and locator.is_visible() and locator.is_enabled():
                    locator.click()
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        wait_deadline = time.monotonic() + 12
        while len(videos) == before and time.monotonic() < wait_deadline:
            page.wait_for_timeout(1_000)
        if len(videos) == before:
            stalled_pages += 1
            if stalled_pages >= 2 or time.monotonic() - last_change > 20:
                break
        else:
            stalled_pages = 0

    page.remove_listener("response", on_response)
    rows = [video_to_row(video, mid) for video in videos.values()]
    rows.sort(key=lambda row: row["published_at"] or datetime.min, reverse=True)
    return rows[:max_videos]


def export_videos(rows: list[dict[str, Any]], output_dir: Path, mid: str, page_title: str) -> Path:
    author = next((row["author"] for row in rows if row.get("author")), "")
    if not author and page_title:
        author = re.split(r"的个人空间|_哔哩哔哩", page_title, maxsplit=1)[0].strip()
    columns = [
        ("bvid", "BV号", 18),
        ("aid", "AV号", 16),
        ("author", "UP主昵称", 18),
        ("mid", "UID", 16),
        ("published_at", "发布时间", 20),
        ("title", "标题", 36),
        ("description", "简介", 60),
        ("duration", "时长", 12),
        ("play_count", "播放数", 14),
        ("danmaku_count", "弹幕数", 14),
        ("comment_count", "评论数", 14),
        ("favorite_count", "收藏数", 14),
        ("video_url", "视频链接", 42),
        ("cover_url", "封面链接", 42),
    ]
    return write_social_xlsx(
        rows,
        output_dir,
        author,
        mid,
        columns,
        sheet_name="视频列表",
        table_name="BilibiliVideos",
        date_keys={"published_at"},
        hyperlink_keys={"video_url", "cover_url"},
        wrap_keys={"title", "description"},
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="采集B站UP主的公开视频并导出Excel")
    parser.add_argument("account", help="B站UID或空间链接")
    parser.add_argument("--profile-url", help="UP主空间链接")
    parser.add_argument("--max-videos", type=int, default=10_000, help="最多采集视频数")
    parser.add_argument("--login-wait", type=int, default=120, help="等待手动登录/验证的秒数")
    parser.add_argument("--output", type=Path, default=Path("output/bilibili"), help="输出目录")
    parser.add_argument("--headless", action="store_true", help="无头运行；首次使用不建议开启")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.max_videos < 1 or args.login_wait < 0:
        print("--max-videos 必须 >= 1，--login-wait 必须 >= 0", file=sys.stderr)
        return 2
    try:
        mid, profile_url = parse_mid(args.account, args.profile_url)
        with sync_playwright() as playwright:
            context = playwright.chromium.launch_persistent_context(
                str(Path(".bilibili-browser").resolve()),
                headless=args.headless,
                channel="chrome",
                viewport={"width": 1440, "height": 900},
                locale="zh-CN",
            )
            page = context.pages[0] if context.pages else context.new_page()
            wait = min(args.login_wait, 12) if args.headless else args.login_wait
            rows = collect_videos(page, profile_url, mid, args.max_videos, wait)
            title = page.title()
            context.close()
        output_path = export_videos(rows, args.output, mid, title)
    except (PlaywrightTimeoutError, ValueError, RuntimeError) as exc:
        print(f"采集失败：{exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"采集失败：{exc}", file=sys.stderr)
        return 1
    print(f"[3/3] 完成，共 {len(rows)} 条")
    print(f"      Excel: {console_safe(output_path.resolve())}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
