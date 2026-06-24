import io
import re
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from urllib.parse import parse_qs, unquote, urlparse

import cv2
import fitz  # PyMuPDF
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
except Exception:
    pyzbar_decode = None


REQUIRED_COLUMNS = ["페이지", "카드번호", "제목", "유튜브링크"]


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
    """CSV 인코딩이 UTF-8, UTF-8 BOM, CP949여도 읽을 수 있게 처리."""
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
    value = value.replace("’", "'").replace("–", "-").replace("—", "-")
    value = re.sub(r"\s+", "", value)
    return value.lower().strip()


def normalize_url(value: str) -> str:
    if not value:
        return ""
    return unquote(str(value).strip())


def parse_url_parts(url: str) -> Dict[str, Optional[str]]:
    """YouTube URL은 영상 ID / playlist ID / 채널 핸들을 분리해서 비교."""
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


def compare_urls(expected: str, actual: str) -> Tuple[str, str]:
    if not actual:
        return "MISSING", "URL 없음"

    e = parse_url_parts(expected)
    a = parse_url_parts(actual)

    if normalize_url(expected) == normalize_url(actual):
        return "EXACT", "URL 전체 일치"

    if e["video_id"] and a["video_id"] and e["video_id"] == a["video_id"]:
        note = "영상 ID 일치"
        if e.get("index") != a.get("index"):
            note += f" / index 다름: CSV {e.get('index')} vs PDF {a.get('index')}"
        return "VIDEO_ID_MATCH", note

    if e["playlist_id"] and a["playlist_id"] and e["playlist_id"] == a["playlist_id"]:
        return "PLAYLIST_MATCH", "플레이리스트 ID 일치"

    if e["channel"] and a["channel"] and normalize_text(e["channel"]) == normalize_text(a["channel"]):
        return "CHANNEL_MATCH", "채널 일치"

    return "MISMATCH", "기대 URL과 다름"


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

        # 1차: OpenCV QR 인식
        ok, decoded_info, points, _ = detector.detectAndDecodeMulti(bgr)
        if ok and points is not None:
            for data, pts in zip(decoded_info, points):
                xs = pts[:, 0] / zoom
                ys = pts[:, 1] / zoom
                add_qr(page_index, data, xs, ys)

        # 2차: pyzbar 보조 인식. 모든 페이지에 돌리면 느릴 수 있어,
        # CSV 기준 예상 개수보다 적게 잡힌 페이지에만 추가 실행합니다.
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

def page_text_spans(page: fitz.Page) -> List[Dict]:
    spans = []
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text:
                    continue
                x0, y0, x1, y1 = span.get("bbox")
                spans.append({"text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1})
    return spans


def extract_title_text_for_region(page: fitz.Page, x0: float, x1: float, y_min: float, y_max: float) -> str:
    spans = page_text_spans(page)
    selected = []
    for sp in spans:
        cx = (sp["x0"] + sp["x1"]) / 2
        cy = (sp["y0"] + sp["y1"]) / 2
        if x0 <= cx <= x1 and y_min <= cy <= y_max:
            selected.append(sp)
    selected.sort(key=lambda s: (round(s["y0"] / 4) * 4, s["x0"]))
    return " ".join(s["text"] for s in selected).strip()


def page_has_text(page: fitz.Page, expected_title: str) -> bool:
    text = page.get_text("text") or ""
    return normalize_text(expected_title) in normalize_text(text)


def group_by_page(items):
    grouped: Dict[int, List] = {}
    for item in items:
        grouped.setdefault(item.page, []).append(item)
    for page in grouped:
        grouped[page].sort(key=lambda r: (r.x0, r.y0))
    return grouped


def row_card_rect(page: fitz.Page, card_count: int, card_no: int, link: Optional[PdfLink], qr: Optional[PdfQr]) -> Tuple[float, float, float, float]:
    if link:
        return link.x0, link.y0, link.x1, link.y1
    if qr:
        width = page.rect.width
        approx_width = width / max(card_count, 1)
        cx = (qr.x0 + qr.x1) / 2
        return max(0, cx - approx_width / 2), 0, min(width, cx + approx_width / 2), page.rect.height

    width = page.rect.width
    margin = width * 0.04
    usable = width - margin * 2
    card_w = usable / max(card_count, 1)
    x0 = margin + (card_no - 1) * card_w
    x1 = margin + card_no * card_w
    return x0, 0, x1, page.rect.height


def validate_video_cards(
    doc: fitz.Document,
    csv_df: pd.DataFrame,
    qrs: List[PdfQr],
    links: List[PdfLink],
    title_y_min: float,
    title_y_max: float,
    semantic_match_ok: bool = True,
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

        page = doc[page_num - 1]
        page_qrs = qrs_by_page.get(page_num, [])
        page_links = links_by_page.get(page_num, [])
        qr = page_qrs[card_no - 1] if card_no - 1 < len(page_qrs) else None
        link = page_links[card_no - 1] if card_no - 1 < len(page_links) else None

        if qr:
            matched_qr_keys.add((qr.page, round(qr.x0, 2), round(qr.y0, 2), qr.url))
        if link:
            matched_link_keys.add((link.page, round(link.x0, 2), round(link.y0, 2), link.url))

        card_count = int(card_count_by_page.get(page_num, 4))
        x0, _, x1, _ = row_card_rect(page, card_count, card_no, link, qr)
        pdf_title_region = extract_title_text_for_region(page, x0, x1, title_y_min, title_y_max)
        title_region_match = normalize_text(expected_title) in normalize_text(pdf_title_region)
        title_page_match = page_has_text(page, expected_title)

        qr_status, qr_note = compare_urls(expected_url, qr.url if qr else "")
        link_status, link_note = compare_urls(expected_url, link.url if link else "")

        title_status = "OK" if title_region_match else ("PAGE_ONLY" if title_page_match else "MISMATCH")
        issues = []
        detail_notes = []

        if title_status == "MISMATCH":
            issues.append("TITLE_MISMATCH")
        elif title_status == "PAGE_ONLY":
            issues.append("TITLE_CHECK")
            detail_notes.append("제목은 페이지 전체에는 있으나 카드 영역에서 정확히 잡히지 않음")

        if qr_status == "MISSING":
            issues.append("MISSING_QR")
        elif qr_status == "MISMATCH":
            issues.append("QR_MISMATCH")
        elif qr_status != "EXACT":
            detail_notes.append(f"QR: {qr_note}")

        if link_status == "MISSING":
            issues.append("MISSING_LINK")
        elif link_status == "MISMATCH":
            issues.append("LINK_MISMATCH")
        elif link_status != "EXACT":
            detail_notes.append(f"LINK: {link_note}")

        if issues:
            overall = "CHECK" if all(i in {"TITLE_CHECK"} for i in issues) else "ISSUE"
        else:
            semantic_ok_values = {"EXACT", "VIDEO_ID_MATCH", "PLAYLIST_MATCH", "CHANNEL_MATCH"}
            if semantic_match_ok:
                overall = "OK" if qr_status in semantic_ok_values and link_status in semantic_ok_values else "CHECK"
            else:
                overall = "OK" if qr_status == "EXACT" and link_status == "EXACT" else "CHECK"

        results.append(
            {
                "상태": overall,
                "문제유형": ", ".join(issues) if issues else "",
                "페이지": page_num,
                "카드번호": card_no,
                "CSV 제목": expected_title,
                "PDF 카드영역 텍스트": pdf_title_region,
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
        key = (qr.page, round(qr.x0, 2), round(qr.y0, 2), qr.url)
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
        key = (link.page, round(link.x0, 2), round(link.y0, 2), link.url)
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


def autosize_sheet(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(len(value), 80))
        ws.column_dimensions[col_letter].width = max(max_len + 2, 12)


def write_dataframe(ws, df: pd.DataFrame):
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

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

    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        status_value = str(getattr(row, df.columns[0], "")) if len(df.columns) else ""
        row_fill = PatternFill("solid", fgColor=status_colors.get(status_value, "FFFFFF"))
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if c_idx == 1:
                cell.font = Font(bold=True)
            cell.fill = row_fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
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

    summary_rows = [
        ["항목", "개수"],
        ["CSV 기준 영상 카드", total],
        ["OK", ok],
        ["CHECK", check],
        ["ISSUE", issue],
        ["기타 QR/링크", extra_count],
    ]
    for row in summary_rows:
        ws_summary.append(row)
    ws_summary["A1"].font = Font(bold=True)
    ws_summary["B1"].font = Font(bold=True)
    autosize_sheet(ws_summary)

    ws_video = wb.create_sheet("영상카드_검수")
    write_dataframe(ws_video, video_df)

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
        st.caption("QR 인식 실패가 있으면 2.5 또는 3.0으로 올려보세요.")
        title_y_min = st.number_input("제목 탐색 Y 시작", min_value=0, max_value=1080, value=520, step=5)
        title_y_max = st.number_input("제목 탐색 Y 끝", min_value=0, max_value=1080, value=600, step=5)
        st.caption("현재 1920x1080 PDF 기준 기본값입니다. 제목 위치가 바뀌면 조정하세요.")
        semantic_match_ok = st.checkbox("영상 ID가 같으면 OK 처리", value=True)
        st.caption("체크 해제 시 index 등 URL 세부값이 다르면 CHECK로 표시됩니다.")

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
                )

            ok = int((video_df["상태"] == "OK").sum())
            check = int((video_df["상태"] == "CHECK").sum())
            issue = int((video_df["상태"] == "ISSUE").sum())

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("CSV 영상 카드", len(video_df))
            col2.metric("OK", ok)
            col3.metric("CHECK", check)
            col4.metric("ISSUE", issue)
            col5.metric("기타 QR/링크", len(extra_df))

            tab1, tab2 = st.tabs(["영상 카드 검수", "기타 QR/하이퍼링크"])
            with tab1:
                st.dataframe(video_df, use_container_width=True, hide_index=True)
            with tab2:
                st.dataframe(extra_df, use_container_width=True, hide_index=True)

            excel_bytes = build_excel_report(video_df, extra_df)
            st.download_button(
                label="엑셀 검수표 다운로드",
                data=excel_bytes,
                file_name="pdf_qr_link_validation_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as exc:
            st.exception(exc)


if __name__ == "__main__":
    main()
