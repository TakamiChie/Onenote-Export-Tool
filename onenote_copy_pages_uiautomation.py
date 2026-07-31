#!/usr/bin/env python3
import ctypes
import html
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import uiautomation as auto

VK_ESCAPE = 0x1B
VK_CONTROL = 0x11
VK_RIGHT = 0x27
VK_A = 0x41
VK_APPS = 0x5D
KEYEVENTF_KEYUP = 0x0002
COPY_ERROR_MESSAGE = "現在、コンテンツをコピーできません。後でもう一度お試しください。"
COPY_ERROR_WAIT_SECONDS = 1.5

onenote = auto.WindowControl(searchDepth=1, ClassName="Framework::CFrame")
list = onenote.ListControl()


def log_message(message: str) -> None:
    repo_root = Path(__file__).resolve().parent
    log_path = repo_root / "export_log.txt"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(f"[{timestamp}] {message}\n")
    print(message)


def get_clipboard_preview() -> Optional[str]:
    try:
        import win32clipboard
    except ImportError:
        return None

    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                if isinstance(data, str):
                    return (
                        data.splitlines()[0].strip()
                        if data.splitlines()
                        else data.strip()
                    )
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
                data = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
                if isinstance(data, bytes):
                    text = data.decode("utf-8", errors="ignore")
                    return (
                        text.splitlines()[0].strip()
                        if text.splitlines()
                        else text.strip()
                    )
                if isinstance(data, str):
                    return (
                        data.splitlines()[0].strip()
                        if data.splitlines()
                        else data.strip()
                    )
            cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
            if win32clipboard.IsClipboardFormatAvailable(cf_html):
                data = win32clipboard.GetClipboardData(cf_html)
                if isinstance(data, bytes):
                    text = data.decode("utf-8", errors="ignore")
                    return (
                        text.splitlines()[0].strip()
                        if text.splitlines()
                        else text.strip()
                    )
                if isinstance(data, str):
                    return (
                        data.splitlines()[0].strip()
                        if data.splitlines()
                        else data.strip()
                    )
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return None

    return None


def press_key(key: int) -> None:
    ctypes.windll.user32.keybd_event(key, 0, 0, 0)
    ctypes.windll.user32.keybd_event(key, 0, KEYEVENTF_KEYUP, 0)
    time.sleep(0.05)


def press_ctrl_a() -> None:
    ctypes.windll.user32.keybd_event(VK_CONTROL, 0, 0, 0)
    press_key(VK_A)
    ctypes.windll.user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)


def clear_clipboard() -> bool:
    try:
        import win32clipboard

        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception:
        return False


def get_clipboard_content(
    wait_seconds: float = 1.0,
) -> Optional[tuple[str, Optional[bytes]]]:
    try:
        import win32clipboard
    except ImportError:
        return None

    deadline = time.monotonic() + wait_seconds
    while time.monotonic() < deadline:
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(
                    win32clipboard.CF_UNICODETEXT
                ):
                    text = win32clipboard.GetClipboardData(
                        win32clipboard.CF_UNICODETEXT
                    )
                    if isinstance(text, str):
                        cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
                        html_data = None
                        if win32clipboard.IsClipboardFormatAvailable(cf_html):
                            value = win32clipboard.GetClipboardData(cf_html)
                            if isinstance(value, bytes):
                                html_data = value
                            elif isinstance(value, str):
                                html_data = value.encode("utf-8")
                        return text.strip(), html_data
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            pass
        time.sleep(0.1)

    return None


def copy_focused_selection(
    page_name: str,
) -> Optional[tuple[str, Optional[bytes]]]:
    if not clear_clipboard():
        return None

    press_key(VK_APPS)
    time.sleep(0.1)
    menu = auto.MenuControl(searchDepth=2)
    copy_item = menu.MenuItemControl(Name="コピー")
    if not copy_item.Exists(1.0, 0.1):
        press_key(VK_ESCAPE)
        return None

    copy_item.Click()
    if close_copy_error_dialog(page_name, wait_seconds=0.5):
        return None
    return get_clipboard_content()


def extract_html_fragment(data: bytes) -> Optional[str]:
    """CF_HTMLから選択範囲のHTML断片を取り出す。"""
    header_match = re.search(
        rb"StartFragment:(\d+).*?EndFragment:(\d+)",
        data[:4096],
        flags=re.DOTALL,
    )
    if header_match:
        start = int(header_match.group(1))
        end = int(header_match.group(2))
        if 0 <= start < end <= len(data):
            fragment_data = data[start:end]
        else:
            return None
    else:
        fragment_data = data

    for encoding in ("utf-8-sig", "utf-8", "cp932"):
        try:
            text = fragment_data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        return None

    start_marker = "<!--StartFragment-->"
    end_marker = "<!--EndFragment-->"
    start = text.find(start_marker)
    end = text.find(end_marker)
    if start >= 0 and end > start:
        return text[start + len(start_marker) : end]

    return text


def set_recovered_page_clipboard(
    title: str,
    page_date: str,
    page_time: str,
    body: str,
    body_html: Optional[str],
) -> bool:
    try:
        import win32clipboard
    except ImportError:
        return False

    body_fragment = body_html
    if not body_fragment:
        body_fragment = "".join(
            f"<p>{html.escape(line)}</p>" for line in body.splitlines()
        )
    fragment = (
        f'<p style="font-size:20pt">{html.escape(title)}</p>'
        f"<p>{html.escape(page_date)}</p>"
        f"<p>{html.escape(page_time)}</p>"
        f"{body_fragment}"
    )
    html_document = (
        "<html><head>"
        '<meta name="Generator" content="Microsoft OneNote">'
        "</head><body><!--StartFragment-->"
        f"{fragment}"
        "<!--EndFragment--></body></html>"
    )

    try:
        cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(
                win32clipboard.CF_UNICODETEXT,
                "\n".join((title, page_date, page_time, body)),
            )
            win32clipboard.SetClipboardData(cf_html, html_document.encode("utf-8"))
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception:
        return False


def recover_page_text(item, page_name: str) -> bool:
    """ページ内の各領域を個別にコピーし、ページ全体のHTMLを再構成する。"""
    item.RightClick()
    menu = auto.MenuControl(searchDepth=2)
    rename_item = menu.MenuItemControl(Name="名前の変更")
    if not rename_item.Exists(1.0, 0.1):
        press_key(VK_ESCAPE)
        log_message(
            f"コピーエラーからの復旧失敗: {page_name}: 名前の変更がありません。"
        )
        return False

    rename_item.Click()
    time.sleep(0.2)

    title_content = copy_focused_selection(page_name)
    if not title_content or not title_content[0]:
        log_message(
            f"コピーエラーからの復旧失敗: {page_name}: タイトルを取得できません。"
        )
        return False
    title = title_content[0]

    print(f"title: {title}")
    press_key(VK_RIGHT)
    press_key(VK_RIGHT)
    page_date_content = copy_focused_selection(page_name)
    page_date = page_date_content[0] if page_date_content else None
    print(f"page_date: {page_date}")
    press_key(VK_RIGHT)
    page_time_content = copy_focused_selection(page_name)
    page_time = page_time_content[0] if page_time_content else None
    print(f"page_time: {page_time}")
    press_key(VK_RIGHT)

    for _ in range(3):
        press_ctrl_a()
        time.sleep(0.1)
    body_content = copy_focused_selection(page_name)
    body = body_content[0] if body_content else None
    body_html = (
        extract_html_fragment(body_content[1])
        if body_content and body_content[1]
        else None
    )
    print(f"page_body: {body}")

    if not page_date or not page_time or body is None:
        log_message(
            f"コピーエラーからの復旧失敗: {page_name}: "
            "日付、時刻、または本文を取得できません。"
        )
        return False

    if not set_recovered_page_clipboard(
        title,
        page_date,
        page_time,
        body,
        body_html,
    ):
        log_message(
            f"コピーエラーからの復旧失敗: {page_name}: "
            "クリップボードを再構成できません。"
        )
        return False

    log_message(f"コピーエラーからページテキストを復旧しました: {title}")
    return True


def save_clipboard_as_markdown() -> None:
    repo_root = Path(__file__).resolve().parent
    script_path = repo_root / "onenote_to_markdown.py"
    output_dir = repo_root / "output"

    preview = get_clipboard_preview()
    if preview:
        log_message(f"処理中: {preview}")
    else:
        log_message("処理中: クリップボード内容の先頭行を取得できませんでした。")

    result = subprocess.run(
        [
            sys.executable,
            str(script_path),
            "--clipboard",
            "--overwrite",
            "--output-dir",
            str(output_dir),
        ],
        cwd=str(repo_root),
        check=False,
    )

    if result.returncode != 0:
        log_message("Markdown保存に失敗しました。")


def close_copy_error_dialog(
    page_name: str,
    wait_seconds: float = COPY_ERROR_WAIT_SECONDS,
) -> bool:
    """コピー失敗ダイアログを閉じ、検出した場合は True を返す。"""
    error_text = auto.TextControl(
        searchDepth=8,
        Name=COPY_ERROR_MESSAGE,
    )
    deadline = time.monotonic() + wait_seconds

    while time.monotonic() < deadline:
        if error_text.Exists(0, 0):
            dialog = error_text.GetTopLevelControl()
            ok_button = dialog.ButtonControl(Name="OK")

            if ok_button.Exists(0.5, 0.1):
                ok_button.Click()
            else:
                dialog.SetFocus()
                auto.SendKeys("{Esc}")

            page_label = page_name or "ページ名不明"
            log_message(
                f"コピーエラー（無視して続行）: {page_label}: " f"{COPY_ERROR_MESSAGE}"
            )
            return True

        time.sleep(0.1)

    return False


def search_loop() -> None:
    while True:
        for item in list.GetChildren():
            if not process(item):
                return

        pattern = list.GetScrollPattern()

        old = pattern.VerticalScrollPercent

        pattern.SetScrollPercent(pattern.HorizontalScrollPercent, min(old + 10, 100))

        time.sleep(0.3)

        if pattern.VerticalScrollPercent == old:
            return


def process(item) -> bool:
    page_name = item.Name
    item.RightClick()
    menu = auto.MenuControl(searchDepth=2)
    mi = menu.MenuItemControl(Name="コピー")
    mi.Click()
    if close_copy_error_dialog(page_name):
        if recover_page_text(item, page_name):
            save_clipboard_as_markdown()
        return True
    save_clipboard_as_markdown()
    if bool(ctypes.windll.user32.GetAsyncKeyState(VK_ESCAPE) & 0x8000):
        return False
    return True


search_loop()
