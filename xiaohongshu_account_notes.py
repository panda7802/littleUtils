#!/usr/bin/env python3
"""Export all public notes from a Xiaohongshu account to XLSX."""

from __future__ import annotations

import argparse
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable
from urllib.parse import quote, urlencode, urlparse

from playwright.sync_api import Page, Response, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from social_excel import write_social_xlsx


POSTED_API_MARKER = "/api/sns/web/v1/user_posted"
DETAIL_API_MARKER = "/api/sns/web/v1/feed"


def first_text(*values: Any) -> str:
    return next((str(value).strip() for value in values if value), "")


def console_safe(value: Any) -> str:
    """Keep emoji filenames from crashing legacy Windows GBK consoles."""
    text = str(value)
    encoding = sys.stdout.encoding or "utf-8"
    return text.encode(encoding, errors="replace").decode(encoding, errors="replace")


def is_browser_closed_error(exc: Exception) -> bool:
    message = str(exc).lower()
    return any(marker in message for marker in (
        "target page, context or browser has been closed",
        "browser has been closed",
        "target closed",
        "page has been closed",
    ))


def timestamp_to_datetime(value: Any) -> datetime | None:
    if not isinstance(value, (int, float)):
        return None
    seconds = value / 1000 if value > 10_000_000_000 else value
    return datetime.fromtimestamp(seconds)


def goto_resilient(page: Page, url: str, timeout: int = 60_000) -> None:
    """Navigate without requiring Xiaohongshu's document to finish loading."""
    expected_host = urlparse(url).netloc
    try:
        page.goto(url, wait_until="commit", timeout=min(timeout, 20_000))
    except PlaywrightTimeoutError:
        print("      页面未及时响应，改用异步导航继续等待 ...")
        try:
            page.goto("about:blank", wait_until="commit", timeout=5_000)
            page.evaluate("target => window.location.replace(target)", url)
        except Exception as exc:
            current_host = urlparse(page.url).netloc
            if current_host != expected_host:
                raise RuntimeError(f"无法打开小红书页面：{url}") from exc
        deadline = time.monotonic() + min(timeout / 1000, 15)
        while urlparse(page.url).netloc != expected_host and time.monotonic() < deadline:
            page.wait_for_timeout(500)
        if urlparse(page.url).netloc != expected_host:
            raise RuntimeError(f"无法连接小红书页面：{url}")
    page.wait_for_timeout(2_500)


def profile_note_to_row(note: dict[str, Any]) -> dict[str, Any]:
    note_id = str(note.get("note_id") or note.get("id") or "")
    token = str(note.get("xsec_token") or "")
    user = note.get("user") if isinstance(note.get("user"), dict) else {}
    cover = note.get("cover") if isinstance(note.get("cover"), dict) else {}
    query = urlencode({"xsec_token": token, "xsec_source": "pc_user"}) if token else ""
    note_url = f"https://www.xiaohongshu.com/explore/{note_id}"
    if query:
        note_url += "?" + query
    return {
        "note_id": note_id,
        "author": first_text(user.get("nickname"), note.get("nickname")),
        "user_id": first_text(user.get("user_id"), user.get("id")),
        "published_at": timestamp_to_datetime(note.get("time")),
        "title": first_text(note.get("display_title"), note.get("title")),
        "content": "",
        "note_type": first_text(note.get("type"), "normal"),
        "liked_count": "",
        "collected_count": "",
        "comment_count": "",
        "share_count": "",
        "note_url": note_url,
        "cover_url": first_text(cover.get("url_default"), cover.get("url_pre"), cover.get("url")),
        "media_url": "",
        "image_urls": [],
        "xsec_token": token,
    }


def extract_video_url(card: dict[str, Any]) -> str:
    video = card.get("video") if isinstance(card.get("video"), dict) else {}
    media = video.get("media") if isinstance(video.get("media"), dict) else {}
    stream = media.get("stream") if isinstance(media.get("stream"), dict) else {}
    for codec in ("h264", "h265", "av1"):
        variants = stream.get(codec)
        if not isinstance(variants, list):
            continue
        for variant in variants:
            if isinstance(variant, dict):
                url = first_text(variant.get("master_url"), variant.get("backup_urls", [""])[0] if variant.get("backup_urls") else "")
                if url:
                    return url
    return first_text(video.get("url"))


def update_from_detail(row: dict[str, Any], card: dict[str, Any]) -> None:
    user = card.get("user") if isinstance(card.get("user"), dict) else {}
    interact = card.get("interact_info") if isinstance(card.get("interact_info"), dict) else {}
    images: list[str] = []
    for item in card.get("image_list") or []:
        if not isinstance(item, dict):
            continue
        info_list = item.get("info_list") if isinstance(item.get("info_list"), list) else []
        info_url = ""
        if info_list and isinstance(info_list[0], dict):
            info_url = first_text(info_list[0].get("url"))
        url = first_text(item.get("url_default"), item.get("url_pre"), info_url)
        if url:
            images.append(url)
    row.update({
        "author": first_text(user.get("nickname"), row.get("author")),
        "user_id": first_text(user.get("user_id"), row.get("user_id")),
        "published_at": timestamp_to_datetime(card.get("time")) or row.get("published_at"),
        "title": first_text(card.get("title"), row.get("title")),
        "content": first_text(card.get("desc"), row.get("content")),
        "note_type": first_text(card.get("type"), row.get("note_type")),
        "liked_count": interact.get("liked_count", row.get("liked_count", "")),
        "collected_count": interact.get("collected_count", row.get("collected_count", "")),
        "comment_count": interact.get("comment_count", row.get("comment_count", "")),
        "share_count": interact.get("share_count", row.get("share_count", "")),
        "media_url": extract_video_url(card),
        "image_urls": images or row.get("image_urls", []),
    })


def find_profile_url(page: Page, account: str, login_wait: int) -> str:
    if account.startswith(("http://", "https://")) and "/user/profile/" in account:
        return account
    print(f"[1/4] 搜索小红书账号：{account}")
    goto_resilient(
        page,
        f"https://www.xiaohongshu.com/search_result?keyword={quote(account)}&source=web_search_result_users",
    )
    deadline = time.monotonic() + login_wait
    while True:
        page.wait_for_timeout(2_000)
        anchors = page.locator('a[href*="/user/profile/"]')
        for index in range(min(anchors.count(), 50)):
            anchor = anchors.nth(index)
            href = anchor.get_attribute("href") or ""
            try:
                context = anchor.evaluate("el => (el.parentElement?.parentElement || el).innerText")
            except Exception:
                context = ""
            if account in (context or ""):
                return "https://www.xiaohongshu.com" + href if href.startswith("/") else href
        if time.monotonic() >= deadline:
            break
        print("      若页面要求登录或验证，请在浏览器中完成", end="\r")
    raise RuntimeError("搜索结果中未找到完全匹配账号，请复制用户主页 URL 作为参数。")


def collect_notes(
    page: Page,
    profile_url: str,
    max_notes: int,
    max_idle: int,
    include_details: bool,
    checkpoint: Callable[[list[dict[str, Any]]], None] | None = None,
) -> list[dict[str, Any]]:
    notes: dict[str, dict[str, Any]] = {}
    last_change = time.monotonic()

    def on_response(response: Response) -> None:
        nonlocal last_change
        if POSTED_API_MARKER not in response.url and DETAIL_API_MARKER not in response.url:
            return
        try:
            payload = response.json()
        except Exception:
            return
        data = payload.get("data") or {}
        if POSTED_API_MARKER in response.url:
            items = data.get("notes") or data.get("note_list") or []
            before = len(notes)
            for item in items:
                if not isinstance(item, dict):
                    continue
                row = profile_note_to_row(item)
                if row["note_id"]:
                    notes[row["note_id"]] = {**notes.get(row["note_id"], {}), **row}
            if len(notes) > before:
                last_change = time.monotonic()
                print(f"      已获取 {len(notes)} 条笔记")
        else:
            for item in data.get("items") or []:
                if not isinstance(item, dict):
                    continue
                card = item.get("note_card") if isinstance(item.get("note_card"), dict) else {}
                note_id = str(item.get("id") or card.get("note_id") or "")
                if note_id in notes:
                    update_from_detail(notes[note_id], card)

    page.on("response", on_response)
    print(f"[2/4] 打开主页并滚动采集：{profile_url}")
    goto_resilient(page, profile_url)
    page.wait_for_timeout(5_000)
    unchanged_height_rounds = 0
    previous_height = 0
    while len(notes) < max_notes:
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(1_500)
            height = page.evaluate("document.body.scrollHeight")
        except Exception as exc:
            if notes and is_browser_closed_error(exc):
                print("      浏览器已关闭，使用当前已获取的笔记继续导出")
                break
            raise
        unchanged_height_rounds = unchanged_height_rounds + 1 if height == previous_height else 0
        previous_height = height
        if time.monotonic() - last_change >= max_idle and unchanged_height_rounds >= 3:
            break
    if not notes:
        try:
            page.remove_listener("response", on_response)
        except Exception:
            pass
        raise RuntimeError("没有收到公开笔记列表，请确认已登录、已完成验证且主页链接正确。")

    selected = list(notes.values())[:max_notes]
    if checkpoint:
        checkpoint(selected)
        print("      已保存笔记列表检查点")
    if include_details:
        print(f"[3/4] 逐条补齐正文（共 {len(selected)} 条）...")
        for index, row in enumerate(selected, start=1):
            try:
                goto_resilient(page, row["note_url"], timeout=45_000)
                if not row.get("content"):
                    meta = page.locator('meta[name="description"]')
                    if meta.count():
                        row["content"] = first_text(meta.first.get_attribute("content"))
                if not row.get("title"):
                    title_meta = page.locator('meta[property="og:title"]')
                    if title_meta.count():
                        row["title"] = first_text(title_meta.first.get_attribute("content"))
            except PlaywrightTimeoutError:
                pass
            except Exception as exc:
                if is_browser_closed_error(exc):
                    print("\n      浏览器已关闭，停止补正文并导出当前进度")
                    break
                print(f"\n      第 {index} 条正文读取失败，已跳过：{exc}")
            if checkpoint and (index % 10 == 0 or index == len(selected)):
                checkpoint(selected)
            print(f"      正文进度 {index}/{len(selected)}", end="\r")
        print()
    else:
        print("[3/4] 已跳过逐条正文采集")

    try:
        page.remove_listener("response", on_response)
    except Exception:
        pass
    selected.sort(key=lambda row: row.get("published_at") or datetime.min, reverse=True)
    return selected


def export_notes(rows: list[dict[str, Any]], output_dir: Path, account: str, page_title: str) -> Path:
    author = next((row["author"] for row in rows if row.get("author")), "")
    if not author and page_title:
        author = re.split(r" - 小红书|_小红书", page_title, maxsplit=1)[0].strip()
    columns = [
        ("note_id", "笔记ID", 24),
        ("author", "作者昵称", 18),
        ("user_id", "用户ID", 24),
        ("published_at", "发布时间", 20),
        ("title", "标题", 36),
        ("content", "正文", 60),
        ("note_type", "笔记类型", 12),
        ("liked_count", "点赞数", 14),
        ("collected_count", "收藏数", 14),
        ("comment_count", "评论数", 14),
        ("share_count", "分享数", 14),
        ("note_url", "笔记链接", 48),
        ("cover_url", "封面链接", 42),
        ("media_url", "视频地址（临时）", 42),
        ("image_urls", "图片地址", 48),
    ]
    return write_social_xlsx(
        rows,
        output_dir,
        author,
        account,
        columns,
        sheet_name="笔记列表",
        table_name="XiaohongshuNotes",
        date_keys={"published_at"},
        hyperlink_keys={"note_url", "cover_url", "media_url"},
        wrap_keys={"title", "content", "image_urls"},
        transforms={
            "note_type": lambda value: {"normal": "图文", "video": "视频"}.get(str(value), value),
            "image_urls": lambda value: "\n".join(str(url) for url in (value or [])),
        },
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="采集小红书账号的公开笔记并导出Excel")
    parser.add_argument("account", help="小红书用户主页URL，或用于搜索的昵称/小红书号")
    parser.add_argument("--profile-url", help="小红书用户主页URL（推荐）")
    parser.add_argument("--max-notes", type=int, default=10_000, help="最多采集笔记数")
    parser.add_argument("--max-idle", type=int, default=12, help="多少秒无新数据后停止滚动")
    parser.add_argument("--login-wait", type=int, default=120, help="等待手动登录/验证的秒数")
    parser.add_argument("--skip-details", action="store_true", help="不逐条打开笔记补正文，可显著加快速度")
    parser.add_argument("--output", type=Path, default=Path("output/xiaohongshu"), help="输出目录")
    parser.add_argument("--headless", action="store_true", help="无头运行；首次使用不建议开启")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.max_notes < 1 or args.max_idle < 3 or args.login_wait < 0:
        print("参数范围不正确", file=sys.stderr)
        return 2
    try:
        with sync_playwright() as playwright:
            context = playwright.chromium.launch_persistent_context(
                str(Path(".xiaohongshu-browser").resolve()),
                headless=args.headless,
                channel="chrome",
                viewport={"width": 1440, "height": 900},
                locale="zh-CN",
            )
            page = context.pages[0] if context.pages else context.new_page()
            wait = min(args.login_wait, 12) if args.headless else args.login_wait
            profile_url = args.profile_url or find_profile_url(page, args.account, wait)
            checkpoint_path: Path | None = None

            def save_checkpoint(current_rows: list[dict[str, Any]]) -> None:
                nonlocal checkpoint_path
                checkpoint_path = export_notes(current_rows, args.output, args.account, "")

            rows = collect_notes(
                page,
                profile_url,
                args.max_notes,
                args.max_idle,
                not args.skip_details,
                checkpoint=save_checkpoint,
            )
            try:
                title = page.title()
            except Exception:
                title = ""
            try:
                context.close()
            except Exception:
                pass
        output_path = export_notes(rows, args.output, args.account, title)
    except (PlaywrightTimeoutError, RuntimeError) as exc:
        print(f"采集失败：{exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"采集失败：{exc}", file=sys.stderr)
        return 1
    print(f"[4/4] 完成，共 {len(rows)} 条")
    print(f"      Excel: {console_safe(output_path.resolve())}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
