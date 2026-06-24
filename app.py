import io
import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple
from urllib.parse import parse_qs, unquote, urlparse

import cv2
import fitz  # PyMuPDF
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
except Exception:
    pyzbar_decode = None


REQUIRED_COLUMNS = ["페이지", "카드번호", "제목", "유튜브링크"]
ISSUE_ORDER = [
    "TITLE_MISMATCH",
    "TITLE_CHECK",
    "MISSING_QR",
    "QR_MISMATCH",
    "QR_URL_DETAIL_DIFF",
    "MISSING_LINK",
    "LINK_MISMATCH",
    "LINK_URL_DETAIL_DIFF",
]


@dataclass
class PdfLink:
    page: int
    x0: float
    y0: float
    x1: float
    y1: float
    url: str
    kind: str
    target_page: Optional[int] = None


@dataclass
class PdfQr:
    page: int
    x0: float
    y0: float
    x1: float
    y1: float
    url: str


def read_csv_flexible(file_bytes: bytes) -> pd.DataFrame:
    """CSV 인코딩이 UTF-8, UTF-8 BOM, CP949, EUC-KR이어도 읽을 수 있게 처리합니다."""
    last_error = None
    for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
        except Exception as exc:
            last_error = exc
    raise ValueError(f"CSV를 읽지 못했습니다. 인코딩 또는 파일 형식을 확인해 주세요. 원인: {last_error}")


def normalize_text(value: str) -> str:
    value = "" if value is None else str(value)
    value = unicodedata.normalize("NFKC", value)
    value = value.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    value = value.replace("–", "-").replace("—", "-")
    value = re.sub(r"\s+", "", value)
    return value.lower().strip()


def normalize_url(value: str) -> str:
    if not value:
        return ""
    return unquote(str(value).strip())


def parse_url_parts(url: str) -> Dict[str, Optional[str]]:
    """YouTube URL은 영상 ID / playlist ID / 채널 핸들을 분리해서 비교합니다."""
    url = normalize_url(url)
    out = {
        "raw": url,
        "domain": None,
        "type": "other",
        "video_id": None,
        "playlist_id": None,
        "channel": None,
        "index": None,
    }
    if not url:
        return out

    try:
        p = urlparse(url)
        domain = p.netloc.lower().replace("www.", "")
        out["domain"] = domain
        qs = parse_qs(p.query)
        out["index"] = (qs.get("index") or [None])[0]
        out["playlist_id"] = (qs.get("list") or [None])[0]

        if "youtube.com" in domain:
            if p.path == "/watch":
                out["type"] = "video"
                out["video_id"] = (qs.get("v") or [None])[0]
            elif p.path.startswith("/shorts/"):
                out["type"] = "video"
                out["video_id"] = p.path.split("/shorts/", 1)[1].split("/")[0]
            elif p.path.startswith("/playlist"):
                out["type"] = "playlist"
            elif p.path.startswith("/@"):
                out["type"] = "channel"
                out["channel"] = p.path.strip("/")
            elif p.path.startswith("/channel/") or p.path.startswith("/c/") or p.path.startswith("/user/"):
                out["type"] = "channel"
                out["channel"] = p.path.strip("/")
        elif "youtu.be" in domain:
            out["type"] = "video"
            out["video_id"] = p.path.strip("/").split("/")[0]
    except Exception:
        pass
    return out


def compare_urls(expected: str, actual: str) -> Tuple[str, str, bool]:
    """return: 비교상태, 비고, URL 세부값 차이 여부"""
    if not actual:
        return "MISSING", "URL 없음", False

    e = parse_url_parts(expected)
    a = parse_url_parts(actual)

    if normalize_url(expected) == normalize_url(actual):
        return "EXACT", "URL 전체 일치", False

    if e["video_id"] and a["video_id"] and e["video_id"] == a["video_id"]:
        details = []
        if e.get("playlist_id") != a.get("playlist_id"):
            details.append(f"playlist/list 다름: CSV {e.get('playlist_id')} vs PDF {a.get('playlist_id')}")
        if e.get("index") != a.get("index"):
            details.append(f"index 다름: CSV {e.get('index')} vs PDF {a.get('index')}")
        note = "영상 ID 일치" + (" / " + " / ".join(details) if details else " / URL 세부값 다름")
        return "VIDEO_ID_MATCH", note, True

    if e["playlist_id"] and a["playlist_id"] and e["playlist_id"] == a["playlist_id"]:
        return "PLAYLIST_MATCH", "플레이리스트 ID 일치 / URL 세부값 다름", True

    if e["channel"] and a["channel"] and normalize_text(e["channel"]) == normalize_text(a["channel"]):
        return "CHANNEL_MATCH", "채널 일치 / URL 세부값 다름", True

    return "MISMATCH", "기대 URL과 다름", False


def is_external_video_url(url: str) -> bool:
    if not url:
        return False
    parts = parse_url_parts(url)
    return parts["type"] in {"video", "playlist", "channel"}


def load_pdf(pdf_bytes: bytes) -> fitz.Document:
    return fitz.open(stream=pdf_bytes, filetype="pdf")


def extract_links(doc: fitz.Document) -> List[PdfLink]:
    links: List[PdfLink] = []
    for page_index, page in enumerate(doc, start=1):
        for link in page.get_links():
            rect = link.get("from")
            if not rect:
                continue
            uri = link.get("uri") or ""
            kind = "external" if uri else "internal"
            target_page = None
            if not uri and link.get("page") is not None:
                target_page = int(link.get("page")) + 1
            links.append(
                PdfLink(
                    page=page_index,
                    x0=float(rect.x0),
                    y0=float(rect.y0),
                    x1=float(rect.x1),
                    y1=float(rect.y1),
                    url=uri,
                    kind=kind,
                    target_page=target_page,
                )
            )
    return links


def extract_qrs(doc: fitz.Document, zoom: float = 2.0, expected_counts: Optional[Dict[int, int]] = None) -> List[PdfQr]:
    qrs: List[PdfQr] = []
    detector = cv2.QRCodeDetector()
    matrix = fitz.Matrix(zoom, zoom)
    seen = set()

    def add_qr(page_index: int, data: str, xs, ys):
        data = str(data).strip()
        if not data:
            return
        x0, y0, x1, y1 = float(min(xs)), float(min(ys)), float(max(xs)), float(max(ys))
        key = (page_index, data, round(x0, 1), round(y0, 1))
        if key in seen:
            return
        seen.add(key)
        qrs.append(PdfQr(page=page_index, x0=x0, y0=y0, x1=x1, y1=y1, url=data))

    for page_index, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
        bgr = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        ok, decoded_info, points, _ = detector.detectAndDecodeMulti(bgr)
        if ok and points is not None:
            for data, pts in zip(decoded_info, points):
                if not data:
                    continue
                xs = pts[:, 0] / zoom
                ys = pts[:, 1] / zoom
                add_qr(page_index, data, xs, ys)

        current_page_count = sum(1 for item in qrs if item.page == page_index)
        need_extra_scan = False
        if expected_counts and page_index in expected_counts:
            need_extra_scan = current_page_count < int(expected_counts.get(page_index, 0))

        if pyzbar_decode is not None and need_extra_scan:
            try:
                pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                for item in pyzbar_decode(pil_img):
                    data = item.data.decode("utf-8", errors="replace")
                    rect = item.rect
                    xs = [rect.left / zoom, (rect.left + rect.width) / zoom]
                    ys = [rect.top / zoom, (rect.top + rect.height) / zoom]
                    add_qr(page_index, data, xs, ys)
            except Exception:
                pass

    qrs.sort(key=lambda r: (r.page, r.x0, r.y0))
    return qrs


def page_text_lines(page: fitz.Page) -> List[Dict]:
    lines = []
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = [span for span in line.get("spans", []) if span.get("text", "").strip()]
            if not spans:
                continue
            text = " ".join(span.get("text", "").strip() for span in spans).strip()
            x0 = min(span.get("bbox")[0] for span in spans)
            y0 = min(span.get("bbox")[1] for span in spans)
            x1 = max(span.get("bbox")[2] for span in spans)
            y1 = max(span.get("bbox")[3] for span in spans)
            lines.append({"text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1})
    lines.sort(key=lambda s: (round(s["y0"] / 4) * 4, s["x0"]))
    return lines


def page_has_text(page: fitz.Page, expected_title: str) -> bool:
    text = page.get_text("text") or ""
    return normalize_text(expected_title) in normalize_text(text)


def similarity(a: str, b: str) -> float:
    na, nb = normalize_text(a), normalize_text(b)
    if not na or not nb:
        return 0.0
    base = SequenceMatcher(None, na, nb).ratio()
    if na in nb or nb in na:
        # 짧은 일부 단어만 잡혔는데 100% 일치로 처리되는 것을 방지합니다.
        coverage = min(len(na), len(nb)) / max(len(na), len(nb))
        return max(base, coverage)
    return base


def best_title_in_region(
    page: fitz.Page,
    expected_title: str,
    x0: float,
    x1: float,
    y_min: float,
    y_max: float,
    auto_expand: bool = True,
) -> Tuple[str, float, str]:
    """카드 영역 안에서 CSV 제목과 가장 유사한 텍스트를 찾습니다."""
    all_lines = page_text_lines(page)

    def pick_lines(local_y_min: float, local_y_max: float) -> List[Dict]:
        selected = []
        for line in all_lines:
            cx = (line["x0"] + line["x1"]) / 2
            cy = (line["y0"] + line["y1"]) / 2
            if x0 <= cx <= x1 and local_y_min <= cy <= local_y_max:
                selected.append(line)
        selected.sort(key=lambda s: (round(s["y0"] / 4) * 4, s["x0"]))
        return selected

    def score_candidates(lines: List[Dict]) -> Tuple[str, float]:
        best_text = ""
        best_score = 0.0
        # 제목이 1~3줄로 나뉘는 경우가 많아 인접 줄을 묶어서 비교합니다.
        for i in range(len(lines)):
            for n in [1, 2, 3, 4, 5]:
                chunk = lines[i : i + n]
                if not chunk:
                    continue
                # 세로 위치가 너무 떨어진 줄은 하나의 제목으로 묶지 않습니다.
                if len(chunk) > 1 and max(c["y1"] for c in chunk) - min(c["y0"] for c in chunk) > 130:
                    continue
                text = " ".join(c["text"] for c in chunk).strip()
                score = similarity(expected_title, text)
                if score > best_score:
                    best_text, best_score = text, score
        return best_text, best_score

    lines = pick_lines(y_min, y_max)
    best_text, best_score = score_candidates(lines)
    source = "설정 Y범위"

    if auto_expand and best_score < 0.72:
        # 레이아웃이 조금 바뀌어도 제목을 찾을 수 있게 카드 중간 영역을 넓게 재탐색합니다.
        expanded_min = page.rect.height * 0.25
        expanded_max = page.rect.height * 0.72
        expanded_lines = pick_lines(expanded_min, expanded_max)
        expanded_text, expanded_score = score_candidates(expanded_lines)
        if expanded_score > best_score:
            best_text, best_score = expanded_text, expanded_score
            source = "자동 확장 탐색"

    return best_text, best_score, source


def group_by_page(items):
    grouped: Dict[int, List] = {}
    for item in items:
        grouped.setdefault(item.page, []).append(item)
    for page in grouped:
        grouped[page].sort(key=lambda r: (r.x0, r.y0))
    return grouped


def card_x_region(page: fitz.Page, card_count: int, card_no: int) -> Tuple[float, float]:
    width = page.rect.width
    # 좌우 여백이 있어도 카드 번호별 중심이 안정적으로 잡히도록 전체 폭을 카드 수로 나눕니다.
    card_count = max(card_count, 1)
    x0 = (card_no - 1) * width / card_count
    x1 = card_no * width / card_count
    return x0, x1


def item_center_x(item) -> float:
    return (item.x0 + item.x1) / 2


def item_key(item) -> Tuple[int, float, float, str]:
    url = getattr(item, "url", "") or ""
    return (item.page, round(item.x0, 2), round(item.y0, 2), url)


def find_item_for_card(items: List, page: fitz.Page, card_count: int, card_no: int, used_keys: set):
    """카드번호의 x영역에 들어오는 QR/링크를 우선 매칭하고, 실패하면 기존처럼 순서 기준으로 보조 매칭합니다."""
    if not items:
        return None
    x0, x1 = card_x_region(page, card_count, card_no)
    candidates = [item for item in items if item_key(item) not in used_keys and x0 <= item_center_x(item) <= x1]
    if candidates:
        candidates.sort(key=lambda r: (r.x0, r.y0))
        return candidates[0]

    remaining = [item for item in items if item_key(item) not in used_keys]
    if remaining:
        remaining.sort(key=lambda r: (r.x0, r.y0))
        # 보조: 카드번호 순서에 맞는 남은 항목을 선택합니다.
        if card_no - 1 < len(items):
            fallback = items[card_no - 1]
            if item_key(fallback) not in used_keys:
                return fallback
        return remaining[0]
    return None


def title_status_from_score(score: float, page_match: bool, ok_threshold: float, check_threshold: float) -> str:
    if score >= ok_threshold:
        return "OK"
    if score >= check_threshold or page_match:
        return "PAGE_ONLY"
    return "MISMATCH"


def validate_video_cards(
    doc: fitz.Document,
    csv_df: pd.DataFrame,
    qrs: List[PdfQr],
    links: List[PdfLink],
    title_y_min: float,
    title_y_max: float,
    semantic_match_ok: bool = True,
    auto_title_search: bool = True,
    title_ok_threshold: float = 0.88,
    title_check_threshold: float = 0.68,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = csv_df.copy()
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            raise ValueError(f"CSV에 '{col}' 컬럼이 없습니다. 현재 컬럼: {list(df.columns)}")

    df["페이지"] = pd.to_numeric(df["페이지"], errors="coerce").astype("Int64")
    df["카드번호"] = pd.to_numeric(df["카드번호"], errors="coerce").astype("Int64")
    df = df.dropna(subset=["페이지", "카드번호"]).copy()
    df["페이지"] = df["페이지"].astype(int)
    df["카드번호"] = df["카드번호"].astype(int)

    qrs_by_page = group_by_page([q for q in qrs if is_external_video_url(q.url)])
    links_by_page = group_by_page([l for l in links if l.url and is_external_video_url(l.url)])
    card_count_by_page = df.groupby("페이지")["카드번호"].max().to_dict()

    matched_qr_keys = set()
    matched_link_keys = set()
    results = []

    for _, r in df.sort_values(["페이지", "카드번호"]).iterrows():
        page_num = int(r["페이지"])
        card_no = int(r["카드번호"])
        expected_title = str(r["제목"]).strip()
        expected_url = str(r["유튜브링크"]).strip()

        if page_num < 1 or page_num > len(doc):
            results.append(
                {
                    "상태": "ISSUE",
                    "문제유형": "PAGE_NOT_FOUND",
                    "페이지": page_num,
                    "카드번호": card_no,
                    "CSV 제목": expected_title,
                    "PDF 카드영역 텍스트": "",
                    "제목유사도": 0,
                    "제목검수": "MISMATCH",
                    "CSV URL": expected_url,
                    "QR URL": "",
                    "QR검수": "MISSING",
                    "하이퍼링크 URL": "",
                    "하이퍼링크검수": "MISSING",
                    "비고": "CSV의 페이지 번호가 PDF 페이지 수를 벗어남",
                }
            )
            continue

        page = doc[page_num - 1]
        page_qrs = qrs_by_page.get(page_num, [])
        page_links = links_by_page.get(page_num, [])
        card_count = int(card_count_by_page.get(page_num, 4))

        qr = find_item_for_card(page_qrs, page, card_count, card_no, matched_qr_keys)
        link = find_item_for_card(page_links, page, card_count, card_no, matched_link_keys)

        if qr:
            matched_qr_keys.add(item_key(qr))
        if link:
            matched_link_keys.add(item_key(link))

        x0, x1 = card_x_region(page, card_count, card_no)
        pdf_title_region, title_score, title_source = best_title_in_region(
            page,
            expected_title,
            x0,
            x1,
            title_y_min,
            title_y_max,
            auto_expand=auto_title_search,
        )
        title_page_match = page_has_text(page, expected_title)
        title_status = title_status_from_score(
            title_score,
            title_page_match,
            ok_threshold=title_ok_threshold,
            check_threshold=title_check_threshold,
        )

        qr_status, qr_note, qr_detail_diff = compare_urls(expected_url, qr.url if qr else "")
        link_status, link_note, link_detail_diff = compare_urls(expected_url, link.url if link else "")

        issues = []
        detail_notes = []

        if title_status == "MISMATCH":
            issues.append("TITLE_MISMATCH")
        elif title_status == "PAGE_ONLY":
            issues.append("TITLE_CHECK")
            detail_notes.append(f"제목 확인 필요: {title_source}, 유사도 {title_score:.2f}")

        if qr_status == "MISSING":
            issues.append("MISSING_QR")
        elif qr_status == "MISMATCH":
            issues.append("QR_MISMATCH")
        elif qr_detail_diff and not semantic_match_ok:
            issues.append("QR_URL_DETAIL_DIFF")
        elif qr_status != "EXACT":
            detail_notes.append(f"QR: {qr_note}")

        if link_status == "MISSING":
            issues.append("MISSING_LINK")
        elif link_status == "MISMATCH":
            issues.append("LINK_MISMATCH")
        elif link_detail_diff and not semantic_match_ok:
            issues.append("LINK_URL_DETAIL_DIFF")
        elif link_status != "EXACT":
            detail_notes.append(f"LINK: {link_note}")

        if any(i in {"TITLE_MISMATCH", "MISSING_QR", "QR_MISMATCH", "MISSING_LINK", "LINK_MISMATCH"} for i in issues):
            overall = "ISSUE"
        elif issues:
            overall = "CHECK"
        else:
            overall = "OK"

        results.append(
            {
                "상태": overall,
                "문제유형": ", ".join(issues) if issues else "",
                "페이지": page_num,
                "카드번호": card_no,
                "CSV 제목": expected_title,
                "PDF 카드영역 텍스트": pdf_title_region,
                "제목유사도": round(float(title_score), 3),
                "제목검수": title_status,
                "CSV URL": expected_url,
                "QR URL": qr.url if qr else "",
                "QR검수": qr_status,
                "하이퍼링크 URL": link.url if link else "",
                "하이퍼링크검수": link_status,
                "비고": " / ".join(detail_notes),
            }
        )

    extra_rows = []
    for qr in qrs:
        key = item_key(qr)
        if key not in matched_qr_keys:
            extra_rows.append(
                {
                    "구분": "EXTRA_QR",
                    "페이지": qr.page,
                    "x0": round(qr.x0, 1),
                    "y0": round(qr.y0, 1),
                    "URL/대상": qr.url,
                    "비고": "CSV 영상 카드 기준에 매칭되지 않은 QR",
                }
            )

    for link in links:
        key = item_key(link)
        if key not in matched_link_keys:
            if link.kind == "internal":
                kind = "INTERNAL_LINK"
                target = f"page {link.target_page}" if link.target_page else "internal"
                note = "PDF 내부 페이지 이동 링크"
            else:
                kind = "EXTRA_LINK"
                target = link.url
                note = "CSV 영상 카드 기준에 매칭되지 않은 외부 링크"
            extra_rows.append(
                {
                    "구분": kind,
                    "페이지": link.page,
                    "x0": round(link.x0, 1),
                    "y0": round(link.y0, 1),
                    "URL/대상": target,
                    "비고": note,
                }
            )

    return pd.DataFrame(results), pd.DataFrame(extra_rows)


def filter_video_df(video_df: pd.DataFrame, mode: str, keyword: str) -> pd.DataFrame:
    df = video_df.copy()
    if mode == "문제/확인 필요만":
        df = df[df["상태"].isin(["CHECK", "ISSUE"])]
    elif mode == "ISSUE만":
        df = df[df["상태"] == "ISSUE"]
    elif mode == "CHECK만":
        df = df[df["상태"] == "CHECK"]
    elif mode == "OK만":
        df = df[df["상태"] == "OK"]
    elif mode == "제목 문제":
        df = df[df["문제유형"].str.contains("TITLE", na=False)]
    elif mode == "QR 문제":
        df = df[df["문제유형"].str.contains("QR|MISSING_QR", na=False)]
    elif mode == "하이퍼링크 문제":
        df = df[df["문제유형"].str.contains("LINK|MISSING_LINK", na=False)]

    if keyword.strip():
        pattern = re.escape(keyword.strip())
        mask = pd.Series(False, index=df.index)
        for col in ["CSV 제목", "PDF 카드영역 텍스트", "CSV URL", "QR URL", "하이퍼링크 URL", "문제유형", "비고"]:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.contains(pattern, case=False, na=False)
        df = df[mask]
    return df


def issue_type_counts(video_df: pd.DataFrame) -> pd.DataFrame:
    counts = {key: 0 for key in ISSUE_ORDER}
    for value in video_df.get("문제유형", pd.Series(dtype=str)).fillna(""):
        for part in [p.strip() for p in str(value).split(",") if p.strip()]:
            counts[part] = counts.get(part, 0) + 1
    rows = [{"문제유형": k, "개수": v} for k, v in counts.items() if v > 0]
    return pd.DataFrame(rows)


def render_page_preview(pdf_bytes: bytes, page_num: int, qrs: List[PdfQr], links: List[PdfLink], draw_boxes: bool = True, zoom: float = 1.25) -> Image.Image:
    doc = load_pdf(pdf_bytes)
    page = doc[page_num - 1]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    if draw_boxes:
        draw = ImageDraw.Draw(img)
        for qr in [q for q in qrs if q.page == page_num]:
            box = [qr.x0 * zoom, qr.y0 * zoom, qr.x1 * zoom, qr.y1 * zoom]
            draw.rectangle(box, outline=(20, 140, 70), width=4)
            draw.text((box[0], max(0, box[1] - 18)), "QR", fill=(20, 140, 70))
        for link in [l for l in links if l.page == page_num]:
            box = [link.x0 * zoom, link.y0 * zoom, link.x1 * zoom, link.y1 * zoom]
            color = (220, 120, 20) if link.kind == "external" else (80, 80, 80)
            draw.rectangle(box, outline=color, width=3)
            draw.text((box[0], max(0, box[1] - 18)), "LINK", fill=color)
    return img


def autosize_sheet(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(len(value), 70))
        ws.column_dimensions[col_letter].width = max(min(max_len + 2, 60), 10)


def write_dataframe(ws, df: pd.DataFrame):
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    if df.empty:
        for c_idx, col in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col)
            cell.font = Font(bold=True, color="1F2A44")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        autosize_sheet(ws)
        return

    for c_idx, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col)
        cell.font = Font(bold=True, color="1F2A44")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    status_colors = {
        "OK": "E2F0D9",
        "CHECK": "FFF2CC",
        "ISSUE": "FCE4D6",
        "EXTRA_QR": "EADCF8",
        "EXTRA_LINK": "EADCF8",
        "INTERNAL_LINK": "E7E6E6",
    }

    status_col = df.columns[0] if len(df.columns) else None
    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        status_value = str(row.get(status_col, "")) if status_col else ""
        row_fill = PatternFill("solid", fgColor=status_colors.get(status_value, "FFFFFF"))
        for c_idx, col in enumerate(df.columns, start=1):
            value = row[col]
            if pd.isna(value):
                value = ""
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if c_idx == 1:
                cell.font = Font(bold=True)
            cell.fill = row_fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_sheet(ws)


def style_summary(ws):
    header_fill = PatternFill("solid", fgColor="1F2A44")
    title_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_idx in [8, 9]:
        if ws.cell(row=row_idx, column=1).value:
            ws.cell(row=row_idx, column=1).fill = title_fill
            ws.cell(row=row_idx, column=1).font = Font(bold=True)
    autosize_sheet(ws)


def build_excel_report(video_df: pd.DataFrame, extra_df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "요약"

    total = len(video_df)
    ok = int((video_df["상태"] == "OK").sum()) if total else 0
    check = int((video_df["상태"] == "CHECK").sum()) if total else 0
    issue = int((video_df["상태"] == "ISSUE").sum()) if total else 0
    extra_count = len(extra_df)
    qr_count = int(video_df["QR URL"].astype(str).str.len().gt(0).sum()) if total else 0
    link_count = int(video_df["하이퍼링크 URL"].astype(str).str.len().gt(0).sum()) if total else 0

    summary_rows = [
        ["항목", "개수", "비고"],
        ["CSV 기준 영상 카드", total, "CSV의 페이지/카드번호 기준"],
        ["OK", ok, "바로 통과 가능"],
        ["CHECK", check, "확인 권장"],
        ["ISSUE", issue, "수정 필요 가능성 높음"],
        ["인식된 영상 카드 QR", qr_count, "CSV 기준 카드에 매칭된 QR"],
        ["인식된 영상 카드 하이퍼링크", link_count, "CSV 기준 카드에 매칭된 외부 링크"],
        ["기타 QR/링크", extra_count, "CSV 영상 카드에 매칭되지 않은 QR/링크"],
        ["", "", ""],
        ["문제유형", "개수", "설명"],
    ]
    descriptions = {
        "TITLE_MISMATCH": "PDF 카드 제목과 CSV 제목이 다름",
        "TITLE_CHECK": "제목이 페이지에는 있으나 카드 영역 정확도 확인 필요",
        "MISSING_QR": "QR을 찾지 못함",
        "QR_MISMATCH": "QR URL이 CSV URL과 다름",
        "QR_URL_DETAIL_DIFF": "영상/플레이리스트 ID는 같지만 URL 세부값이 다름",
        "MISSING_LINK": "하이퍼링크를 찾지 못함",
        "LINK_MISMATCH": "하이퍼링크 URL이 CSV URL과 다름",
        "LINK_URL_DETAIL_DIFF": "영상/플레이리스트 ID는 같지만 URL 세부값이 다름",
    }
    counts_df = issue_type_counts(video_df)
    for row in summary_rows:
        ws_summary.append(row)
    for _, row in counts_df.iterrows():
        ws_summary.append([row["문제유형"], int(row["개수"]), descriptions.get(row["문제유형"], "")])
    style_summary(ws_summary)

    ws_video = wb.create_sheet("영상카드_검수")
    write_dataframe(ws_video, video_df)

    ws_issue = wb.create_sheet("문제항목만")
    issue_only = video_df[video_df["상태"].isin(["CHECK", "ISSUE"])].copy()
    if issue_only.empty:
        issue_only = pd.DataFrame(columns=video_df.columns)
    write_dataframe(ws_issue, issue_only)

    ws_extra = wb.create_sheet("기타_QR_링크")
    if extra_df.empty:
        extra_df = pd.DataFrame(columns=["구분", "페이지", "x0", "y0", "URL/대상", "비고"])
    write_dataframe(ws_extra, extra_df)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def main():
    st.set_page_config(page_title="PDF QR 링크 자동 검수기", page_icon="✅", layout="wide")
    st.title("PDF QR · 하이퍼링크 자동 검수기")
    st.caption("CSV의 페이지/카드번호/제목/유튜브링크를 기준으로 PDF 안의 영상 제목, QR, 클릭 링크를 비교합니다.")

    with st.sidebar:
        st.header("검수 옵션")
        zoom = st.slider("QR 인식 해상도", min_value=1.0, max_value=4.0, value=2.0, step=0.5)
        st.caption("QR 인식 실패가 있으면 2.5 또는 3.0으로 올려보세요. 값이 높을수록 느려질 수 있습니다.")
        title_y_min = st.number_input("제목 탐색 Y 시작", min_value=0, max_value=1080, value=520, step=5)
        title_y_max = st.number_input("제목 탐색 Y 끝", min_value=0, max_value=1080, value=600, step=5)
        auto_title_search = st.checkbox("제목 위치 자동 보정", value=True)
        st.caption("디자인이 조금 바뀌어도 카드 영역 안에서 CSV 제목과 가장 비슷한 텍스트를 찾습니다.")
        semantic_match_ok = st.checkbox("영상 ID가 같으면 OK 처리", value=True)
        st.caption("체크 해제 시 index/list 등 URL 세부값이 다르면 CHECK로 표시됩니다.")
        show_preview_boxes = st.checkbox("미리보기에서 QR/링크 박스 표시", value=True)

    pdf_file = st.file_uploader("PDF 파일 업로드", type=["pdf"])
    csv_file = st.file_uploader("CSV 기준 데이터 업로드", type=["csv"])

    if not pdf_file or not csv_file:
        st.info("PDF와 CSV를 모두 업로드한 뒤 검수를 실행해 주세요.")
        return

    if st.button("검수 실행", type="primary"):
        try:
            pdf_bytes = pdf_file.read()
            csv_bytes = csv_file.read()
            csv_df = read_csv_flexible(csv_bytes)

            missing = [c for c in REQUIRED_COLUMNS if c not in csv_df.columns]
            if missing:
                st.error(f"CSV 필수 컬럼이 없습니다: {missing}")
                st.stop()

            with st.spinner("PDF에서 QR과 하이퍼링크를 추출하는 중..."):
                doc = load_pdf(pdf_bytes)
                expected_counts = (
                    csv_df.groupby("페이지")["카드번호"].max().dropna().astype(int).to_dict()
                    if "페이지" in csv_df.columns and "카드번호" in csv_df.columns
                    else None
                )
                qrs = extract_qrs(doc, zoom=zoom, expected_counts=expected_counts)
                links = extract_links(doc)
                video_df, extra_df = validate_video_cards(
                    doc=doc,
                    csv_df=csv_df,
                    qrs=qrs,
                    links=links,
                    title_y_min=float(title_y_min),
                    title_y_max=float(title_y_max),
                    semantic_match_ok=semantic_match_ok,
                    auto_title_search=auto_title_search,
                )

            st.session_state["result_video_df"] = video_df
            st.session_state["result_extra_df"] = extra_df
            st.session_state["result_pdf_bytes"] = pdf_bytes
            st.session_state["result_qrs"] = qrs
            st.session_state["result_links"] = links
            st.success("검수가 완료되었습니다.")

        except Exception as exc:
            st.exception(exc)

    if "result_video_df" not in st.session_state:
        return

    video_df = st.session_state["result_video_df"]
    extra_df = st.session_state["result_extra_df"]
    pdf_bytes = st.session_state["result_pdf_bytes"]
    qrs = st.session_state["result_qrs"]
    links = st.session_state["result_links"]

    ok = int((video_df["상태"] == "OK").sum())
    check = int((video_df["상태"] == "CHECK").sum())
    issue = int((video_df["상태"] == "ISSUE").sum())

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("CSV 영상 카드", len(video_df))
    col2.metric("OK", ok)
    col3.metric("CHECK", check)
    col4.metric("ISSUE", issue)
    col5.metric("기타 QR/링크", len(extra_df))

    counts_df = issue_type_counts(video_df)
    if not counts_df.empty:
        with st.expander("문제유형별 요약 보기", expanded=False):
            st.dataframe(counts_df, use_container_width=True, hide_index=True)

    tab1, tab2, tab3 = st.tabs(["영상 카드 검수", "기타 QR/하이퍼링크", "페이지 미리보기"])
    with tab1:
        fcol1, fcol2 = st.columns([1, 2])
        with fcol1:
            filter_mode = st.selectbox(
                "표시 필터",
                ["전체", "문제/확인 필요만", "ISSUE만", "CHECK만", "OK만", "제목 문제", "QR 문제", "하이퍼링크 문제"],
            )
        with fcol2:
            keyword = st.text_input("검색", placeholder="제목, URL, 문제유형 등")
        filtered_df = filter_video_df(video_df, filter_mode, keyword)
        st.caption(f"표시 중: {len(filtered_df)}건 / 전체 {len(video_df)}건")
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    with tab2:
        if extra_df.empty:
            st.info("CSV 영상 카드 기준에 매칭되지 않은 기타 QR/하이퍼링크가 없습니다.")
        else:
            extra_mode = st.selectbox("기타 항목 필터", ["전체", "EXTRA_QR", "EXTRA_LINK", "INTERNAL_LINK"])
            display_extra = extra_df if extra_mode == "전체" else extra_df[extra_df["구분"] == extra_mode]
            st.dataframe(display_extra, use_container_width=True, hide_index=True)

    with tab3:
        issue_pages = sorted(video_df.loc[video_df["상태"].isin(["CHECK", "ISSUE"]), "페이지"].dropna().astype(int).unique().tolist())
        all_pages = list(range(1, len(load_pdf(pdf_bytes)) + 1))
        page_options = issue_pages if issue_pages else all_pages
        pcol1, pcol2 = st.columns([1, 4])
        with pcol1:
            selected_page = st.selectbox("미리보기 페이지", page_options)
            st.caption("초록: QR / 주황: 외부 링크 / 회색: 내부 링크")
        with pcol2:
            preview_img = render_page_preview(pdf_bytes, int(selected_page), qrs, links, draw_boxes=show_preview_boxes, zoom=1.25)
            st.image(preview_img, use_container_width=True)

    excel_bytes = build_excel_report(video_df, extra_df)
    st.download_button(
        label="엑셀 검수표 다운로드",
        data=excel_bytes,
        file_name="pdf_qr_link_validation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.caption("업로드한 PDF/CSV는 앱 실행 중 메모리에서만 처리됩니다. 이 코드에서는 서버 디스크에 별도로 저장하지 않습니다.")


if __name__ == "__main__":
    main()
