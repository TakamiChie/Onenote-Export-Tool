#!/usr/bin/env python3
import ctypes
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import uiautomation as auto

VK_ESCAPE = 0x1B

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
    item.RightClick()
    menu = auto.MenuControl(searchDepth=2)
    mi = menu.MenuItemControl(Name="コピー")
    mi.Click()
    time.sleep(0.5)
    save_clipboard_as_markdown()
    if bool(ctypes.windll.user32.GetAsyncKeyState(VK_ESCAPE) & 0x8000):
        return False
    return True


search_loop()
