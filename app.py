from __future__ import annotations

import re
import tempfile
from dataclasses import fields
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from engine.batch import collect_pptx_uploads, pack_outputs_zip
from engine.excel_loader import format_iso_date, parse_date
from engine.pipeline import PipelineReport, run_batch
from engine.ppt_convert import PptConvertError

DROP_HINT = "파일을 드래그하여 놓거나, 클릭하여 선택하세요"


st.set_page_config(page_title="역별 현황 최신화", layout="wide")
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.4rem; max-width: 1180px; }
    div[data-testid="stFileUploader"] section {
        border: 2px dashed #3b82f6 !important;
        border-radius: 14px !important;
        background: #f8fbff !important;
        padding: 0.9rem 0.85rem !important;
        min-height: 5.6rem;
        position: relative !important;
        cursor: pointer;
    }
    div[data-testid="stFileUploader"] section:hover {
        border-color: #1d4ed8 !important;
        background: #eef4ff !important;
    }
    /* Streamlit 기본 input이 0x0이라 아이콘 근처만 드롭됨. 점선 박스 전체를 드롭 영역으로 덮음. */
    [data-testid="stFileUploaderDropzoneInput"] {
        position: absolute !important;
        inset: 0 !important;
        width: 100% !important;
        height: 100% !important;
        opacity: 0 !important;
        cursor: pointer !important;
        z-index: 3 !important;
        margin: 0 !important;
        display: block !important;
    }
    div[data-testid="stFileUploaderDropzoneInstructions"] {
        pointer-events: none;
    }
    div[data-testid="stFileChips"],
    div[data-testid="stFileChip"],
    button[data-testid="stFileChipDeleteBtn"] {
        position: relative;
        z-index: 4;
    }
    div[data-testid="stFileUploaderDropzoneInstructions"] span {
        font-size: 0.92rem;
    }
    div[data-testid="stFileUploaderDropzoneInstructions"] small {
        font-size: 0.78rem !important;
        color: #64748b !important;
    }
    .korail-table { overflow: auto; max-height: 28rem; margin: 0.35rem 0 0.6rem; }
    .korail-table table { width: 100%; border-collapse: collapse; font-size: 0.9rem; }
    .korail-table th, .korail-table td {
        border: 1px solid #e2e8f0;
        padding: 0.4rem 0.55rem;
        text-align: left;
        vertical-align: top;
        white-space: pre-wrap;
    }
    .korail-table th { background: #f8fafc; font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("역별 현황 최신화")
st.caption(
    "엑셀을 기준으로 표·텍스트만 고칩니다. 사진은 그대로 둡니다. "
    "추가·수정은 빨간색, 삭제는 슬라이드 하단 메모와 노트에 사유가 남습니다."
)

with st.container(border=True):
    st.subheader("1. 기준일")
    st.caption("이 날짜 기준으로 마스터에서 영업중인 매장만 남깁니다. (운영단계·영업개시일·영업종료일)")
    as_of = st.date_input(
        "기준일",
        value=date.today(),
        help="마스터의 운영단계(영업중)와 영업개시일~영업종료일로 현재 운영 매장을 가립니다. 자판기는 기본계약일이 비어 있어도 영업중이면 남깁니다.",
    )

with st.container(border=True):
    st.markdown(
        f"### 2. 엑셀 파일&nbsp;&nbsp;<span style='font-size:0.95rem;font-weight:500;color:#475569'>"
        f"{DROP_HINT}</span>",
        unsafe_allow_html=True,
    )
    excel_cols = st.columns(3)
    with excel_cols[0]:
        sales_file = st.file_uploader(
            "매장별 년간 매출내역",
            type=["xlsx"],
            accept_multiple_files=False,
            key="sales_xlsx",
        )
        st.caption("필수 · .xlsx")
    with excel_cols[1]:
        master_file = st.file_uploader(
            "매장마스터 현황 목록",
            type=["xlsx"],
            accept_multiple_files=False,
            key="master_xlsx",
        )
        st.caption("필수 · .xlsx")
    with excel_cols[2]:
        year_file = st.file_uploader(
            "연도별조회 · 승하차",
            type=["xlsx"],
            accept_multiple_files=False,
            key="year_xlsx",
        )
        st.caption("선택 · 기준일 최근 4개년 · 1년치면 PPT 과거 3년을 유지하고 올해만 갱신")

with st.container(border=True):
    st.subheader("3. PPT 파일")
    st.caption("여러 개의 .ppt / .pptx 또는 그 파일이 들어 있는 .zip을 올릴 수 있습니다. 구버전 .ppt는 자동으로 변환합니다.")
    ppt_files = st.file_uploader(
        "PPT 파일 또는 압축본",
        type=["ppt", "pptx", "zip"],
        accept_multiple_files=True,
        key="ppt_uploads",
        help="ppt, pptx를 여러 개 올리거나, 그 파일이 들어 있는 zip을 올려 주세요.",
    )

ready = bool(ppt_files and sales_file and master_file)
missing = []
if not as_of:
    missing.append("기준일")
if not sales_file:
    missing.append("매출내역")
if not master_file:
    missing.append("매장마스터")
if not ppt_files:
    missing.append("PPT")

with st.container(border=True):
    st.subheader("4. 실행")
    status_cols = st.columns(3)
    status_cols[0].metric("기준일", format_iso_date(as_of) if as_of else "-")
    status_cols[1].metric("엑셀", "준비됨" if sales_file and master_file else "미완료")
    status_cols[2].metric("PPT", f"{len(ppt_files)}개" if ppt_files else "없음")
    if missing:
        st.caption("아직 필요한 항목: " + " · ".join(missing))
    run = st.button("최신화 실행", type="primary", disabled=not ready, width="stretch")


def _as_report(item: dict) -> PipelineReport:
    names = {f.name for f in fields(PipelineReport)}
    return PipelineReport(**{k: v for k, v in item.items() if k in names})


def _change_table(changes: list[dict]) -> pd.DataFrame:
    rows = []
    for item in changes:
        slide = item.get("슬라이드") or f"슬라이드 {item.get('슬라이드번호') or '-'}"
        part = item.get("부분") or ""
        content = (item.get("내용") or "").strip()
        bits = [bit.strip() for bit in re.split(r"\s·\s", content) if bit.strip()]
        if not bits:
            bits = [content or "-"]
        num = int(item.get("슬라이드번호") or 0)
        for bit in bits:
            rows.append({"_no": num, "슬라이드": slide, "구분": part, "변경 내용": bit})
    if not rows:
        return pd.DataFrame(columns=["슬라이드", "구분", "변경 내용"])
    return (
        pd.DataFrame(rows)
        .sort_values(["_no", "구분"], kind="stable")
        .drop(columns=["_no"])
        .reset_index(drop=True)
    )


def _show_df(df: pd.DataFrame) -> None:
    if df is None or df.empty:
        st.caption("표시할 행이 없습니다.")
        return
    html = df.fillna("").astype(str).to_html(index=False, escape=True, border=0)
    st.markdown(f'<div class="korail-table">{html}</div>', unsafe_allow_html=True)


def _show_report(report: PipelineReport) -> None:
    if report.error:
        st.error(f"{report.source_name}: {report.error}")
        return
    as_of_text = format_iso_date(parse_date(report.as_of)) if report.as_of else "-"
    st.success(
        f"{report.source_name} → {report.station}역 · 기준일 {as_of_text} · "
        f"운영 매장 {report.excel_store_count}개 · 계약기간 외 {report.inactive_count}개"
    )
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("개소코드 치환", report.code_replacements)
    m2.metric("매장 표 행", report.shops_written)
    m3.metric("자판기 표 행", report.vendings_written)
    m4.metric("이름 매칭", len(report.matches))
    with st.expander("매칭·수정 상세", expanded=False):
        st.caption("엑셀 기준으로 바뀐 내용을 슬라이드·구분별로 표로 봅니다.")
        change_df = _change_table(list(report.changes or []))
        if change_df.empty:
            st.caption("이번 실행에서 표시할 수정 내용이 없습니다.")
        else:
            slides = ["전체"] + list(dict.fromkeys(change_df["슬라이드"].tolist()))
            picked = st.selectbox("슬라이드", slides, key=f"chg_slide_{report.source_name}")
            view = change_df if picked == "전체" else change_df[change_df["슬라이드"] == picked]
            _show_df(view)
            st.caption(f"{len(view)}건")
        tabs = st.tabs(["개소코드 매칭", "PPT만 있던 매장(삭제)", "엑셀만 있는 매장(추가)"])
        with tabs[0]:
            if report.matches:
                match_df = pd.DataFrame(report.matches).rename(
                    columns={
                        "old_code": "기존 개소코드",
                        "new_code": "엑셀 매장코드",
                        "ppt_name": "PPT 매장명",
                        "excel_name": "엑셀 매장명",
                        "score": "점수",
                        "confidence": "신뢰도",
                        "aliases": "별칭코드",
                    }
                )
                _show_df(match_df)
            else:
                st.caption("매칭된 매장이 없습니다.")
        with tabs[1]:
            _show_df(
                pd.DataFrame(report.unmatched_ppt).rename(
                    columns={"code": "개소코드", "name": "매장명", "brand": "브랜드"}
                )
                if report.unmatched_ppt
                else pd.DataFrame(columns=["개소코드", "매장명", "브랜드"])
            )
        with tabs[2]:
            _show_df(
                pd.DataFrame(report.unmatched_excel).rename(
                    columns={"code": "매장코드", "name": "매장명", "biz_type": "사업구분"}
                )
                if report.unmatched_excel
                else pd.DataFrame(columns=["매장코드", "매장명", "사업구분"])
            )


if run:
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        tmp_path = Path(tmp)
        uploads = [(f.name, f.getvalue()) for f in ppt_files]
        pptx_paths: list[Path] = []
        convert_failed = False
        try:
            pptx_paths = collect_pptx_uploads(uploads, tmp_path / "inputs")
        except PptConvertError as exc:
            st.error(str(exc))
            convert_failed = True
        if not pptx_paths:
            if not convert_failed:
                st.error("처리할 PPT 파일이 없습니다. .ppt / .pptx 또는 그 파일이 들어 있는 zip인지 확인해 주세요.")
            st.session_state.pop("batch_reports", None)
            st.session_state.pop("batch_ok", None)
            st.session_state.pop("batch_zip", None)
        else:
            sales_path = tmp_path / "sales.xlsx"
            master_path = tmp_path / "master.xlsx"
            sales_path.write_bytes(sales_file.getvalue())
            master_path.write_bytes(master_file.getvalue())
            year_path = None
            if year_file:
                year_path = tmp_path / "year.xlsx"
                year_path.write_bytes(year_file.getvalue())

            reports = run_batch(
                pptx_paths=pptx_paths,
                sales_path=sales_path,
                master_path=master_path,
                ridership_path=year_path,
                output_dir=tmp_path / "out",
                station=None,
                as_of=as_of,
            )

            ok_pairs: list[tuple[str, bytes]] = []
            for report in reports:
                if report.error or not report.output_path:
                    continue
                data = Path(report.output_path).read_bytes()
                ok_pairs.append((Path(report.output_path).name, data))

            zip_bytes = None
            if len(ok_pairs) >= 2:
                zip_path = tmp_path / "updated_ppts.zip"
                pack_outputs_zip(ok_pairs, zip_path)
                zip_bytes = zip_path.read_bytes()

            st.session_state["batch_ok"] = ok_pairs
            st.session_state["batch_zip"] = zip_bytes
            st.session_state["batch_reports"] = [r.to_dict() for r in reports]
            st.session_state["batch_total"] = len(pptx_paths)

if st.session_state.get("batch_reports") is not None:
    ok_pairs = st.session_state.get("batch_ok") or []
    zip_bytes = st.session_state.get("batch_zip")
    total = st.session_state.get("batch_total", len(ok_pairs))
    st.info(f"PPT {total}개 중 {len(ok_pairs)}개 처리 완료")
    if zip_bytes:
        st.download_button(
            "최신 PPT 묶음 다운로드 (zip)",
            data=zip_bytes,
            file_name="매장현황_최신_묶음.zip",
            mime="application/zip",
            key="dl_zip",
        )
    elif len(ok_pairs) == 1:
        name, data = ok_pairs[0]
        st.download_button(
            "최신 PPT 다운로드",
            data=data,
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="dl_one",
        )
    for item in st.session_state["batch_reports"]:
        _show_report(_as_report(item))

