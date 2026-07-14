#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Optional

from bs4 import BeautifulSoup, Tag
from dateutil import parser as date_parser
from dateutil.parser import ParserError
from markdownify import markdownify as md

try:
    import win32clipboard
except ImportError:
    win32clipboard = None


CF_HTML_HEADER_RE = re.compile(
    rb"StartFragment:(\d+).*?EndFragment:(\d+)",
    re.DOTALL,
)

WEEK_TITLE_RE = re.compile(r"^\s*(?P<month>\d{1,2})/(?P<day>\d{1,2})週\s*$")

MONTH_TITLE_RE = re.compile(r"^\s*(?P<month>\d{1,2})\s*月\s*$")

FONT_SIZE_RE = re.compile(
    r"font-size\s*:\s*(?P<size>\d+(?:\.\d+)?)pt",
    re.IGNORECASE,
)

JAPANESE_DATE_RE = re.compile(
    r"(?P<year>\d{4})\s*年\s*" r"(?P<month>\d{1,2})\s*月\s*" r"(?P<day>\d{1,2})\s*日"
)

JAPANESE_WEEKDAY_RE = re.compile(r"^\s*(?:月|火|水|木|金|土|日)曜日\s*[,、]?\s*")

JAPANESE_AM_PM_REPLACEMENTS = (
    ("午前", "AM"),
    ("午後", "PM"),
)


@dataclass(frozen=True)
class ParsedPage:
    title: str
    page_date: date
    fragment_html: str
    markdown: str
    output_filename: str


def make_soup(value: str) -> BeautifulSoup:
    """
    html5libが利用可能なら優先し、未導入時は標準html.parserを使う。
    OneNoteの崩れたHTMLを安定して補正するにはhtml5libを推奨する。
    """
    try:
        return BeautifulSoup(value, "html5lib")
    except Exception:
        return BeautifulSoup(value, "html.parser")


def read_html_file(path: Path) -> bytes:
    """CF_HTML形式または通常HTMLをバイト列で読み込む。"""
    return path.read_bytes()


def read_cf_html_from_clipboard() -> bytes:
    """WindowsクリップボードからHTML Formatを取得する。"""
    if win32clipboard is None:
        raise RuntimeError(
            "クリップボード入力には pywin32 が必要です。"
            " `pip install pywin32` を実行してください。"
        )

    cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")

    win32clipboard.OpenClipboard()
    try:
        if not win32clipboard.IsClipboardFormatAvailable(cf_html):
            raise ValueError(
                "クリップボードに HTML Format がありません。"
                " OneNoteのページ一覧で右クリックし、コピーを実行してください。"
            )
        data = win32clipboard.GetClipboardData(cf_html)
    finally:
        win32clipboard.CloseClipboard()

    if isinstance(data, bytes):
        return data
    if isinstance(data, str):
        return data.encode("utf-8")

    raise TypeError(f"HTML Formatの型が想定外です: {type(data).__name__}")


def decode_html_bytes(data: bytes) -> str:
    """
    HTMLバイト列を文字列へ変換する。

    OneNoteのCF_HTMLは通常UTF-8だが、古いページを考慮して
    UTF-8 BOM、UTF-8、CP932の順で試す。
    """
    encodings = ("utf-8-sig", "utf-8", "cp932")

    for encoding in encodings:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue

    raise UnicodeDecodeError(
        "unknown",
        data,
        0,
        min(len(data), 1),
        "HTMLデータをUTF-8またはCP932としてデコードできませんでした。",
    )


def extract_cf_html_fragment(data: bytes) -> str:
    """
    CF_HTMLからStartFragment～EndFragmentを抽出する。

    通常HTMLの場合はStartFragmentコメントを使用し、
    コメントも存在しない場合はHTML全体を返す。
    """
    header_match = CF_HTML_HEADER_RE.search(data[:4096])
    if header_match:
        start = int(header_match.group(1))
        end = int(header_match.group(2))

        if not 0 <= start < end <= len(data):
            raise ValueError("CF_HTMLのStartFragmentまたはEndFragmentが不正です。")

        return decode_html_bytes(data[start:end])

    text = decode_html_bytes(data)

    start_marker = "<!--StartFragment-->"
    end_marker = "<!--EndFragment-->"

    start_index = text.find(start_marker)
    end_index = text.find(end_marker)

    if start_index >= 0 or end_index >= 0:
        if start_index < 0 or end_index < 0 or end_index <= start_index:
            raise ValueError("StartFragmentまたはEndFragmentコメントの対応が不正です。")

        start_index += len(start_marker)
        return text[start_index:end_index]

    return text


def validate_onenote_html(data: bytes) -> None:
    """Generatorに「Microsoft OneNote」が含まれることを確認する。"""
    text = decode_html_bytes(data)
    soup = make_soup(text)

    generator = soup.find(
        "meta",
        attrs={"name": lambda value: value and value.lower() == "generator"},
    )

    if generator is None:
        raise ValueError(
            "Generatorメタ情報が見つかりません。"
            " OneNoteからコピーしたHTMLではない可能性があります。"
        )

    content = str(generator.get("content", ""))

    if "Microsoft OneNote" not in content:
        raise ValueError("Generatorに「Microsoft OneNote」が含まれていません。")


def normalize_text(value: str) -> str:
    """連続空白やNBSPを整理する。"""
    value = html.unescape(value)
    value = value.replace("\u00a0", " ")
    value = re.sub(r"[ \t\r\f\v]+", " ", value)
    return value.strip()


def parse_font_size(style: str) -> Optional[float]:
    match = FONT_SIZE_RE.search(style or "")
    if not match:
        return None
    return float(match.group("size"))


def extract_title_date_time_tags(
    fragment_html: str,
) -> tuple[str, Tag, Tag, Optional[Tag]]:
    """
    タイトル、タイトル要素、日付要素、時刻要素を取得する。

    タイトルはp要素のうち最大font-sizeのものを優先する。
    日付はタイトルの次に現れる空でないp要素とする。
    時刻は日付の次に現れる空でないp要素が時刻として解析可能な場合だけ返す。
    """
    soup = make_soup(fragment_html)
    paragraphs = soup.find_all("p")

    if not paragraphs:
        raise ValueError("本文内にp要素が見つかりません。")

    title_tag = None
    title_size = -1.0

    for paragraph in paragraphs:
        text = normalize_text(paragraph.get_text(" ", strip=True))
        if not text:
            continue

        size = parse_font_size(str(paragraph.get("style", "")))
        if size is not None and size > title_size:
            title_tag = paragraph
            title_size = size

    if title_tag is None:
        title_tag = next(
            (
                paragraph
                for paragraph in paragraphs
                if normalize_text(paragraph.get_text(" ", strip=True))
            ),
            None,
        )

    if title_tag is None:
        raise ValueError("ページタイトルを取得できません。")

    title = normalize_text(title_tag.get_text(" ", strip=True))
    if not title:
        raise ValueError("ページタイトルが空です。")

    title_index = paragraphs.index(title_tag)

    date_tag = None
    date_index = None

    for index, paragraph in enumerate(
        paragraphs[title_index + 1 :],
        start=title_index + 1,
    ):
        text = normalize_text(paragraph.get_text(" ", strip=True))
        if text:
            date_tag = paragraph
            date_index = index
            break

    if date_tag is None or date_index is None:
        raise ValueError("タイトル直後の二行目の日付が見つかりません。")

    time_tag = None
    for paragraph in paragraphs[date_index + 1 :]:
        text = normalize_text(paragraph.get_text(" ", strip=True))
        if not text:
            continue

        if looks_like_time(text):
            time_tag = paragraph
        break

    return title, title_tag, date_tag, time_tag


def preprocess_date_text(value: str) -> str:
    """
    dateutilへ渡す前に、日本語表記や曜日表記を正規化する。
    """
    value = normalize_text(value)
    value = JAPANESE_WEEKDAY_RE.sub("", value)

    for source, replacement in JAPANESE_AM_PM_REPLACEMENTS:
        value = value.replace(source, replacement)

    japanese_match = JAPANESE_DATE_RE.search(value)
    if japanese_match:
        return (
            f"{japanese_match.group('year')}-"
            f"{japanese_match.group('month')}-"
            f"{japanese_match.group('day')}"
        )

    value = re.sub(
        r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日",
        r"\1-\2-\3",
        value,
    )
    value = re.sub(r"\s+", " ", value).strip()

    return value


def parse_flexible_date(date_line: str) -> date:
    """
    python-dateutilを使い、可能な限り多くの日付形式を解析する。

    日本語の「YYYY年M月D日」は事前にISO風へ正規化する。
    OneNoteの「曜日, M月 D, YYYY」もdateutilで解析できる形に整える。
    """
    normalized = preprocess_date_text(date_line)

    attempts = (
        {"yearfirst": True, "dayfirst": False, "fuzzy": True},
        {"yearfirst": False, "dayfirst": False, "fuzzy": True},
        {"yearfirst": False, "dayfirst": True, "fuzzy": True},
    )

    errors = []

    for options in attempts:
        try:
            parsed = date_parser.parse(normalized, **options)
            return parsed.date()
        except (ParserError, ValueError, OverflowError) as exc:
            errors.append(str(exc))

    raise ValueError(
        "二行目の日付を解析できません。"
        f" 実際: {date_line!r} / 正規化後: {normalized!r}"
        f" / 詳細: {' | '.join(errors)}"
    )


def looks_like_time(value: str) -> bool:
    """日付直後の行が時刻だけを表しているか判定する。"""
    normalized = normalize_text(value)

    patterns = (
        r"^(?:午前|午後)?\s*\d{1,2}:\d{2}(?::\d{2})?\s*(?:午前|午後)?$",
        r"^\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?$",
        r"^(?:AM|PM|am|pm)\s*\d{1,2}:\d{2}(?::\d{2})?$",
    )

    return any(re.fullmatch(pattern, normalized) for pattern in patterns)


def build_output_filename(title: str, page_date: date) -> str:
    """
    タイトルと二行目の日付から保存ファイル名を決定する。

    - M/DD週 → YYYY-MM-DD週.md
    - M月    → YYYY-MM月.md
    - その他 → YYYY-MM-DD.md
    """
    week_match = WEEK_TITLE_RE.fullmatch(title)
    if week_match:
        month = int(week_match.group("month"))
        day = int(week_match.group("day"))

        try:
            weekly_date = date(page_date.year, month, day)
        except ValueError as exc:
            raise ValueError(f"週タイトルの日付が不正です: {title!r}") from exc

        return f"{weekly_date:%Y-%m-%d}週.md"

    month_match = MONTH_TITLE_RE.fullmatch(title)
    if month_match:
        month = int(month_match.group("month"))
        if not 1 <= month <= 12:
            raise ValueError(f"月タイトルが不正です: {title!r}")

        return f"{page_date.year:04d}-{month:02d}月.md"

    return f"{page_date:%Y-%m-%d}.md"


def normalize_onenote_lists(root: Tag) -> None:
    """
    OneNoteが出力する、liの兄弟として置かれたネストulを
    直前のliの子要素へ移動する。

    例:
      <ul>
        <li>親</li>
        <ul><li>子</li></ul>
      </ul>

    を次へ修正する:
      <ul>
        <li>親<ul><li>子</li></ul></li>
      </ul>
    """
    for list_tag in root.find_all(["ul", "ol"]):
        for child in list(list_tag.children):
            if not isinstance(child, Tag):
                continue
            if child.name not in ("ul", "ol"):
                continue

            previous = child.find_previous_sibling()
            while previous is not None and not isinstance(previous, Tag):
                previous = previous.previous_sibling

            if previous is not None and previous.name == "li":
                previous.append(child.extract())


def increase_heading_levels(root: Tag) -> None:
    """
    元HTMLの見出しを1レベル下げる。

    ページタイトルをMarkdownのh1として追加するため、
    元のh1→h2、h2→h3 ... h5→h6 とする。
    h6はMarkdown上これ以上下げられないためh6のままとする。
    """
    for level in range(5, 0, -1):
        for heading in root.find_all(f"h{level}"):
            heading.name = f"h{level + 1}"


def clean_fragment_for_markdown(
    fragment_html: str,
    title: str,
) -> str:
    """
    日付・時刻・元タイトルを除去し、
    Markdown変換用にページタイトルと本文構造を整理する。
    """
    soup = make_soup(fragment_html)
    body = soup.body or soup

    _, title_tag, date_tag, time_tag = extract_title_date_time_tags(fragment_html)

    # 別のBeautifulSoupインスタンス上の要素なので、本文側でも再特定する。
    body_paragraphs = body.find_all("p")
    body_title = None
    body_title_size = -1.0

    for paragraph in body_paragraphs:
        text = normalize_text(paragraph.get_text(" ", strip=True))
        if not text:
            continue
        size = parse_font_size(str(paragraph.get("style", "")))
        if size is not None and size > body_title_size:
            body_title = paragraph
            body_title_size = size

    if body_title is None:
        raise ValueError("Markdown変換時にタイトル要素を再特定できません。")

    title_index = body_paragraphs.index(body_title)
    body_date = next(
        (
            paragraph
            for paragraph in body_paragraphs[title_index + 1 :]
            if normalize_text(paragraph.get_text(" ", strip=True))
        ),
        None,
    )

    if body_date is None:
        raise ValueError("Markdown変換時に日付要素を再特定できません。")

    date_index = body_paragraphs.index(body_date)
    body_time = None

    for paragraph in body_paragraphs[date_index + 1 :]:
        text = normalize_text(paragraph.get_text(" ", strip=True))
        if not text:
            continue
        if looks_like_time(text):
            body_time = paragraph
        break

    body_title.decompose()
    body_date.decompose()
    if body_time is not None:
        body_time.decompose()

    for tag in body.find_all(["script", "style", "meta", "link"]):
        tag.decompose()

    normalize_onenote_lists(body)
    increase_heading_levels(body)

    for tag in body.find_all(True):
        tag.attrs = {}

    for paragraph in list(body.find_all("p")):
        text = normalize_text(paragraph.get_text(" ", strip=True))
        if not text:
            paragraph.decompose()

    for span in body.find_all("span"):
        span.unwrap()

    title_heading = soup.new_tag("h1")
    title_heading.string = title
    body.insert(0, title_heading)

    return str(body)


def postprocess_markdown(markdown: str) -> str:
    """Markdownify後のOneNote向け後処理。"""
    markdown = html.unescape(markdown)
    markdown = markdown.replace("\u00a0", " ")
    markdown = markdown.replace("\r\n", "\n").replace("\r", "\n")

    markdown = re.sub(r"[ \t]+\n", "\n", markdown)
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)
    markdown = re.sub(r"(?m)^(#{1,6})([^\s#])", r"\1 \2", markdown)

    return markdown.strip() + "\n"


def convert_fragment_to_markdown(
    fragment_html: str,
    title: str,
) -> str:
    cleaned_html = clean_fragment_for_markdown(fragment_html, title)

    markdown = md(
        cleaned_html,
        heading_style="ATX",
        bullets="-",
        strip=["html", "body"],
    )

    return postprocess_markdown(markdown)


def parse_onenote_page(data: bytes) -> ParsedPage:
    validate_onenote_html(data)

    fragment_html = extract_cf_html_fragment(data)
    title, _, date_tag, _ = extract_title_date_time_tags(fragment_html)
    date_line = normalize_text(date_tag.get_text(" ", strip=True))
    page_date = parse_flexible_date(date_line)
    output_filename = build_output_filename(title, page_date)
    markdown = convert_fragment_to_markdown(fragment_html, title)

    return ParsedPage(
        title=title,
        page_date=page_date,
        fragment_html=fragment_html,
        markdown=markdown,
        output_filename=output_filename,
    )


def save_parsed_page(
    page: ParsedPage,
    output_dir: Path,
    overwrite: bool,
    debug_html: bool,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = output_dir / page.output_filename

    if output_path.exists() and not overwrite:
        raise FileExistsError(
            f"出力先がすでに存在します: {output_path}"
            " 上書きする場合は --overwrite を指定してください。"
        )

    output_path.write_text(page.markdown, encoding="utf-8")

    if debug_html:
        debug_path = output_path.with_suffix(".fragment.html")
        debug_path.write_text(page.fragment_html, encoding="utf-8")

    return output_path


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=("OneNoteのCF_HTMLまたは保存済みHTMLをMarkdownへ変換します。")
    )

    source_group = parser.add_mutually_exclusive_group(required=True)
    source_group.add_argument(
        "html_file",
        nargs="?",
        type=Path,
        help="デバッグ用のCF_HTMLまたはHTMLファイル",
    )
    source_group.add_argument(
        "--clipboard",
        action="store_true",
        help="WindowsクリップボードのHTML Formatから読み込む",
    )

    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Markdownの出力先ディレクトリ（既定値: output）",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="同名ファイルが存在する場合に上書きする",
    )
    parser.add_argument(
        "--debug-html",
        action="store_true",
        help="抽出後のHTML断片も保存する",
    )

    return parser


def main() -> int:
    parser = build_argument_parser()
    args = parser.parse_args()

    if args.clipboard:
        data = read_cf_html_from_clipboard()
    else:
        if args.html_file is None:
            raise ValueError("HTMLファイルが指定されていません。")
        data = read_html_file(args.html_file)

    page = parse_onenote_page(data)

    output_path = save_parsed_page(
        page=page,
        output_dir=args.output_dir,
        overwrite=args.overwrite,
        debug_html=args.debug_html,
    )

    print(f"タイトル: {page.title}")
    print(f"日付: {page.page_date:%Y-%m-%d}")
    print(f"保存先: {output_path}")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        raise
