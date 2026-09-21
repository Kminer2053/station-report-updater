from __future__ import annotations

import re
from copy import deepcopy
from io import BytesIO
from dataclasses import dataclass

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import _Relationship
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

CODE_RE = re.compile(r"(?<!\d)(\d{5,7})(?!\d)")
DATA_FONT = "굴림"
DATA_SIZE_PT = 9.0
ROW_H_DEFAULT = 291465  # 0.319 in
REVIEW_RED = RGBColor(0xFF, 0x00, 0x00)
REVIEW_RED_HEX = "FF0000"


def emu_to_in(value) -> float:
    return round(value / 914400, 2) if value is not None else 0.0


def iter_shapes(container):
    for shape in container.shapes:
        yield shape
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from iter_shapes(shape)
        except Exception:
            continue


def set_run_text_preserve(paragraph, new_text: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.text = new_text


def apply_run_font(run, font_name: str, size_pt: float, bold: bool | None = None) -> None:
    run.font.size = Pt(size_pt)
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font_name)
    if bold is not None:
        run.font.bold = bold


def _font_from_rpr(rPr) -> tuple[str | None, float | None]:
    if rPr is None:
        return None, None
    pt = None
    sz = rPr.get("sz")
    if sz:
        try:
            pt = int(sz) / 100.0
        except ValueError:
            pass
    name = None
    for tag in ("a:ea", "a:latin", "a:cs"):
        el = rPr.find(qn(tag))
        if el is not None and el.get("typeface"):
            name = el.get("typeface")
            break
    return name, pt


def _bold_from_rpr(rPr) -> bool | None:
    if rPr is None:
        return None
    b = rPr.get("b")
    if b is None:
        return None
    return b not in ("0", "false", "off")


def sample_run_bold(run) -> bool | None:
    if run.font.bold is not None:
        return bool(run.font.bold)
    return _bold_from_rpr(run._r.find(qn("a:rPr")))


def sample_cell_bold(cell) -> bool | None:
    try:
        tf = cell.text_frame
        for para in tf.paragraphs:
            for run in para.runs:
                bold = sample_run_bold(run)
                if bold is not None:
                    return bold
            pPr = para._p.find(qn("a:pPr"))
            if pPr is not None:
                bold = _bold_from_rpr(pPr.find(qn("a:defRPr")))
                if bold is not None:
                    return bold
            bold = _bold_from_rpr(para._p.find(qn("a:endParaRPr")))
            if bold is not None:
                return bold
    except Exception:
        pass
    return None


def sample_table_data_bold(table, fallback: bool = True) -> bool:
    """데이터 행의 굵기. 칸마다 다르면 서울역처럼 굵게를 쓴다."""
    votes: list[bool] = []
    for ri, row in enumerate(table.rows):
        if ri == 0:
            continue
        for cell in row.cells:
            bold = sample_cell_bold(cell)
            if bold is not None:
                votes.append(bold)
    if not votes:
        return fallback
    if True in votes and False in votes:
        return True
    return votes[0]


def sample_run_font(run) -> tuple[str | None, float | None]:
    name = run.font.name
    pt = float(run.font.size.pt) if run.font.size else None
    if name and pt:
        return name, pt
    xml_name, xml_pt = _font_from_rpr(run._r.find(qn("a:rPr")))
    return name or xml_name, pt or xml_pt


def sample_cell_font(cell) -> tuple[str | None, float | None]:
    try:
        tf = cell.text_frame
        for para in tf.paragraphs:
            for run in para.runs:
                name, pt = sample_run_font(run)
                if name or pt:
                    return name, pt
            pPr = para._p.find(qn("a:pPr"))
            if pPr is not None:
                name, pt = _font_from_rpr(pPr.find(qn("a:defRPr")))
                if name or pt:
                    return name, pt
            name, pt = _font_from_rpr(para._p.find(qn("a:endParaRPr")))
            if name or pt:
                return name, pt
    except Exception:
        pass
    return None, None


def sample_table_data_font(table, fallback: tuple[str, float] = (DATA_FONT, 10.0)) -> tuple[str, float]:
    """헤더를 제외한 데이터 행에서 글꼴을 읽어, 빈 행 추가 시에도 같은 크기를 쓴다."""
    for ri, row in enumerate(table.rows):
        if ri == 0:
            continue
        for cell in row.cells:
            name, pt = sample_cell_font(cell)
            if pt:
                return name or fallback[0], pt
    return fallback


def norm_text(value: str | None) -> str:
    return re.sub(r"\s+", "", (value or "").strip())


def _paint_r_red(r) -> None:
    rPr = r.find(qn("a:rPr"))
    if rPr is None:
        rPr = etree.SubElement(r, qn("a:rPr"))
    for tag in ("a:solidFill", "a:gradFill", "a:noFill"):
        el = rPr.find(qn(tag))
        if el is not None:
            rPr.remove(el)
    solid = etree.Element(qn("a:solidFill"))
    srgb = etree.SubElement(solid, qn("a:srgbClr"))
    srgb.set("val", REVIEW_RED_HEX)
    # OOXML: ln → fill → effect → latin/ea. Fill after effectLst is ignored by PowerPoint.
    insert_at = 0
    for i, child in enumerate(list(rPr)):
        local = etree.QName(child).localname
        if local == "ln":
            insert_at = i + 1
        elif local in {"effectLst", "effectDag", "highlight", "latin", "ea", "cs", "sym"}:
            break
    rPr.insert(insert_at, solid)


def apply_review_red(run) -> None:
    _paint_r_red(run._r)


def mark_cell_red(cell) -> None:
    try:
        tf = cell.text_frame
    except Exception:
        return
    for para in tf.paragraphs:
        for run in para.runs:
            if (run.text or "").strip():
                apply_review_red(run)


def write_paragraph_segments(paragraph, segments: list[tuple[str, bool]]) -> None:
    """기존 런의 폰트·굵기를 복사한 뒤, is_red 구간만 빨간색으로 쓴다."""
    parts = [(text, is_red) for text, is_red in segments if text]
    if not parts:
        set_run_text_preserve(paragraph, "")
        return
    p = paragraph._p
    runs_xml = [child for child in p if child.tag == qn("a:r")]
    if not runs_xml:
        paragraph.text = "".join(text for text, _ in parts)
        if any(is_red for _, is_red in parts):
            for run in paragraph.runs:
                apply_review_red(run)
        return
    template = deepcopy(runs_xml[0])
    end = p.find(qn("a:endParaRPr"))
    for child in list(runs_xml):
        p.remove(child)
    for text, is_red in parts:
        r = deepcopy(template)
        t = r.find(qn("a:t"))
        if t is None:
            t = etree.SubElement(r, qn("a:t"))
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        if is_red:
            _paint_r_red(r)
        if end is not None:
            end.addprevious(r)
        else:
            p.append(r)


def labeled_review_segments(new_text: str, old_text: str) -> list[tuple[str, bool]] | None:
    if new_text == old_text:
        return None
    match = re.match(r"^(.*?:\s*)(.*)$", new_text, flags=re.S)
    if match and match.group(1).strip() and match.group(2):
        return [(match.group(1), False), (match.group(2), True)]
    if (new_text or "").strip():
        return [(new_text, True)]
    return [(new_text, False)]


def diff_code_segments(text: str, mapping: dict[str, str]) -> list[tuple[str, bool]] | None:
    if not text or not mapping:
        return None
    items = [(old, new) for old, new in mapping.items() if old and new and old != new]
    if not items:
        return None
    items.sort(key=lambda kv: len(kv[0]), reverse=True)
    pattern = "|".join(rf"(?<!\d){re.escape(old)}(?!\d)" for old, _ in items)
    cmap = {old: new for old, new in items}
    segs: list[tuple[str, bool]] = []
    pos = 0
    changed = False
    for match in re.finditer(pattern, text):
        if match.start() > pos:
            segs.append((text[pos : match.start()], False))
        old = match.group(0)
        new = cmap.get(old, old)
        segs.append((new, new != old))
        changed = changed or new != old
        pos = match.end()
    if not changed:
        return None
    if pos < len(text):
        segs.append((text[pos:], False))
    return segs or None


def _longest_line_len(text: str) -> int:
    lines = (text or "").replace(" ", "").split("\n")
    return max((len(line) for line in lines), default=0)


def fit_size(text: str, base_pt: float, col_kind: str = "") -> float:
    """칸 너비 안에서 줄이 더 생기지 않게 글자 크기를 줄인다.

    개소명 열(~1.5in, 9pt)은 한글 약 12자. 그보다 긴 줄은 표 행을 밀어
    슬라이드 밖으로 넘긴다.
    """
    n = _longest_line_len(text)
    if col_kind == "title":
        total = len((text or "").replace(" ", "").replace("\n", ""))
        if total <= 6:
            return 18.0
        if total <= 8:
            return 14.0
        if total <= 12:
            return 12.0
        return 10.0
    if col_kind == "name":
        if n >= 16:
            return 6.0
        if n >= 14:
            return 6.5
        if n >= 12:
            return 7.0
        return base_pt
    if col_kind in {"brand", "location", "owner"}:
        if n >= 16:
            return 6.0
        if n >= 14:
            return 6.5
        if n >= 11:
            return 7.0
        if n >= 9:
            return 7.5
        return base_pt
    total = len((text or "").replace(" ", "").replace("\n", ""))
    if total >= 20:
        return max(7.0, base_pt - 2)
    if total >= 14:
        return max(7.5, base_pt - 1)
    return base_pt


def apply_cell_run_font(
    cell,
    font_name: str,
    size_pt: float,
    bold: bool | None = None,
) -> None:
    """칸의 글자 런에 글꼴·크기·굵기를 맞춘다. 개행 뒤 빈 문단은 1pt로 남겨 행 높이를 안 민다."""
    try:
        tf = cell.text_frame
    except Exception:
        return
    paras = list(tf.paragraphs)
    nonempty = any((para.text or "").strip() for para in paras)
    for para in paras:
        has_text = bool((para.text or "").strip())
        pt = size_pt if has_text or not nonempty else 1.0
        if not para.runs:
            if not has_text and nonempty:
                continue
            run = para.add_run()
            apply_run_font(run, font_name, pt, bold=bold)
            continue
        for run in para.runs:
            apply_run_font(run, font_name, pt, bold=bold)


def unify_table_data_style(table, font_name: str, size_pt: float, bold: bool) -> None:
    """헤더를 제외한 데이터 칸을 한 크기·한 굵기로 맞춘다."""
    for ri, row in enumerate(table.rows):
        if ri == 0:
            continue
        for cell in row.cells:
            name, _pt = sample_cell_font(cell)
            apply_cell_run_font(cell, name or font_name, size_pt, bold)


def set_cell_text_preserve(
    cell,
    value: str,
    font_name: str | None = None,
    size_pt: float | None = None,
    review_red: bool | None = None,
    bold: bool | None = None,
) -> None:
    old = get_cell_text(cell)
    tf = cell.text_frame
    saved_aligns = [para.alignment for para in tf.paragraphs] if tf.paragraphs else []
    saved_anchor = None
    try:
        saved_anchor = cell.vertical_anchor
    except Exception:
        saved_anchor = None
    try:
        tf.word_wrap = True
    except Exception:
        pass
    lines = (value or "").replace("\r\n", "\n").split("\n")
    if len(lines) > 1 and lines[-1] == "":
        lines = lines[:-1]
    if not lines:
        lines = [""]
    if not tf.paragraphs:
        cell.text = "\n".join(lines)
        return
    while len(tf.paragraphs) < len(lines):
        para = tf.add_paragraph()
        if saved_aligns and saved_aligns[0] is not None:
            para.alignment = saved_aligns[0]
    for i, line in enumerate(lines):
        para = tf.paragraphs[i]
        set_run_text_preserve(para, line)
        if not para.runs:
            run = para.add_run()
            run.text = line
        if i < len(saved_aligns) and saved_aligns[i] is not None:
            para.alignment = saved_aligns[i]
        elif saved_aligns and saved_aligns[0] is not None:
            para.alignment = saved_aligns[0]
    for para in tf.paragraphs[len(lines) :]:
        set_run_text_preserve(para, "")
    if font_name is not None or size_pt is not None or bold is not None:
        name = font_name or DATA_FONT
        size = size_pt if size_pt is not None else DATA_SIZE_PT
        for para in tf.paragraphs[: len(lines)]:
            try:
                para.space_before = Pt(0)
                para.space_after = Pt(0)
            except Exception:
                pass
            if not para.runs:
                run = para.add_run()
                run.text = para.text or ""
                apply_run_font(run, name, size, bold=bold)
            else:
                for run in para.runs:
                    apply_run_font(run, name, size, bold=bold)
        for para in tf.paragraphs[len(lines) :]:
            try:
                para.space_before = Pt(0)
                para.space_after = Pt(0)
            except Exception:
                pass
            for run in para.runs:
                apply_run_font(run, name, 1.0, bold=bold)
    if review_red is None:
        review_red = bool((value or "").strip()) and norm_text(old) != norm_text(value)
    for para, align in zip(tf.paragraphs, saved_aligns):
        if align is not None:
            para.alignment = align
    if saved_anchor is not None:
        try:
            cell.vertical_anchor = saved_anchor
        except Exception:
            pass
    if review_red and (value or "").strip():
        mark_cell_red(cell)



def last_cell(row):
    cells = row.cells
    if not cells:
        return None
    return cells[len(cells) - 1]


def get_cell_text(cell) -> str:
    return (cell.text or "").strip()


def replace_codes_in_text(text: str, mapping: dict[str, str]) -> str:
    if not text or not mapping:
        return text
    # 긴 코드부터 치환 (7자리 오타가 6자리에 먹히지 않게)
    items = sorted(mapping.items(), key=lambda kv: len(kv[0]), reverse=True)

    def repl(match: re.Match) -> str:
        old = match.group(1)
        return mapping.get(old, old)

    # 매핑된 코드만 정확 치환
    out = text
    for old, new in items:
        out = re.sub(rf"(?<!\d){re.escape(old)}(?!\d)", new, out)
    return out


def replace_codes_in_paragraph(paragraph, mapping: dict[str, str]) -> bool:
    """런 단위로 코드만 치환해 동그라미 숫자 등 기존 글꼴을 유지한다."""
    if not mapping:
        return False
    changed = False
    runs = list(paragraph.runs)
    if not runs:
        full = paragraph.text or ""
        new = replace_codes_in_text(full, mapping)
        if new != full:
            paragraph.text = new
            return True
        return False
    for run in runs:
        old = run.text or ""
        new = replace_codes_in_text(old, mapping)
        if new != old:
            run.text = new
            changed = True
    return changed


def replace_codes_in_shape(shape, mapping: dict[str, str]) -> int:
    changed = 0
    try:
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for para in cell.text_frame.paragraphs:
                        if replace_codes_in_paragraph(para, mapping):
                            changed += 1
    except Exception:
        pass
    try:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                if replace_codes_in_paragraph(para, mapping):
                    changed += 1
    except Exception:
        pass
    return changed


def header_key(text: str) -> str:
    return re.sub(r"\s+", "", (text or "").replace("\n", ""))


def classify_table(rows: list[list[str]]) -> str:
    if not rows:
        return "unknown"
    header = "".join(header_key(c) for c in rows[0])
    blob = header_key("".join("".join(r) for r in rows[:4]))
    if "월평균매출액" in header and "브랜드" in header and "개소코드" in header:
        return "store_detail"
    if "월평균매출액" in header and "개소코드" in header:
        return "vending_detail"
    if "개소수" in header or "업종" in header:
        return "count_summary"
    if "승차" in header or "승차" in blob or "하차" in header:
        return "ridership"
    if "담당소속" in blob or "운영현황" in blob:
        return "ops_info"
    if "개소코드" in blob and len(rows[0]) <= 3:
        return "photo_card"
    return "unknown"


def table_rows_text(table) -> list[list[str]]:
    return [[get_cell_text(c) for c in row.cells] for row in table.rows]


def find_header_columns(header: list[str]) -> dict[str, int]:
    mapping = {}
    for idx, raw in enumerate(header):
        h = header_key(raw)
        if "개소코드" in h:
            mapping["code"] = idx
        elif "개소명" in h:
            mapping["name"] = idx
        elif "브랜드" in h:
            mapping["brand"] = idx
        elif "대표자" in h:
            mapping["owner"] = idx
        elif h == "위치":
            mapping["location"] = idx
        elif "면적" in h:
            mapping["area"] = idx
        elif "최초계약" in h:
            mapping["start"] = idx
        elif "만료" in h:
            mapping["end"] = idx
        elif "월평균매출" in h:
            mapping["sales"] = idx
        elif "용량" in h:
            mapping["kw"] = idx
        elif h in {"구분", "no", "No"}:
            mapping["no"] = idx
    return mapping


def is_empty_row(values: list[str]) -> bool:
    return not any(v.strip() for v in values)


def delete_table_row(table, row_idx: int) -> None:
    tbl = table._tbl
    tr = table.rows[row_idx]._tr
    tbl.remove(tr)


def set_row_height(table, row_idx: int, emu: int) -> None:
    table._tbl.tr_lst[row_idx].set("h", str(int(emu)))


def add_table_row(table, template_idx: int = 1):
    """서식 있는 데이터 행을 복사해 추가. 빈 마지막 행을 복사하지 않는다."""
    rows_xml = table._tbl.tr_lst
    idx = template_idx if template_idx < len(rows_xml) else 0
    src = rows_xml[idx]
    new_tr = deepcopy(src)
    table._tbl.append(new_tr)
    new_row = table.rows[len(table.rows) - 1]
    for cell in new_row.cells:
        set_cell_text_preserve(cell, "", DATA_FONT, DATA_SIZE_PT)
    return new_row


def insert_table_row_before(table, before_idx: int, template_idx: int = 1):
    rows_xml = table._tbl.tr_lst
    if before_idx < 0 or before_idx >= len(rows_xml):
        return add_table_row(table, template_idx=template_idx)
    idx = template_idx if template_idx < len(rows_xml) else 0
    new_tr = deepcopy(rows_xml[idx])
    rows_xml[before_idx].addprevious(new_tr)
    new_row = table.rows[before_idx]
    for cell in new_row.cells:
        set_cell_text_preserve(cell, "", review_red=False)
    return new_row


def equalize_data_row_heights(table, original_data_count: int) -> None:
    """표 전체 높이는 유지하고 데이터 행 높이를 균등하게 맞춘다."""
    trs = table._tbl.tr_lst
    if len(trs) <= 1:
        return
    header_h = int(trs[0].get("h") or ROW_H_DEFAULT)
    orig_n = max(original_data_count, 1)
    # 원본 데이터 영역 높이 (추가 전에 쓰던 값). 현재 행 높이 합을 기준으로 맞춤
    current_data_h = sum(int(tr.get("h") or ROW_H_DEFAULT) for tr in trs[1:])
    n_data = len(trs) - 1
    if n_data <= original_data_count:
        return
    each = max(int(current_data_h / n_data), int(0.22 * 914400))
    # 원본보다 행이 늘었으면 원본 데이터 영역(대략 orig_n * 0.319in)에 맞춘다
    orig_area = orig_n * ROW_H_DEFAULT
    if n_data > orig_n:
        each = max(int(orig_area / n_data), int(0.20 * 914400))
    for i in range(1, len(trs)):
        set_row_height(table, i, each)


@dataclass
class FoundTable:
    slide_index: int
    shape: object
    kind: str
    left_in: float


C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def _c(tag: str) -> str:
    return f"{C_NS}{tag}"


def unlink_external_chart_data(chart) -> None:
    """차트에 연결된 외부 xls 경로를 끊어 PowerPoint 액세스 오류를 막는다."""
    el = chart._element
    for child in list(el):
        if child.tag == _c("externalData"):
            el.remove(child)
    part = chart.part
    for rId in list(part.rels.keys()):
        rel = part.rels[rId]
        if getattr(rel, "is_external", False):
            part.drop_rel(rId)


def sanitize_chart_workbooks(prs) -> None:
    for slide in prs.slides:
        for shape in iter_shapes(slide):
            try:
                if shape.has_chart:
                    unlink_external_chart_data(shape.chart)
            except Exception:
                continue


def _rewrite_cache_values(cache, values: list) -> None:
    count = cache.find(_c("ptCount"))
    if count is not None:
        count.set("val", str(len(values)))
    pts = cache.findall(_c("pt"))
    for i, value in enumerate(values):
        if i < len(pts):
            pt = pts[i]
        else:
            pt = etree.SubElement(cache, _c("pt"))
            pt.set("idx", str(i))
            etree.SubElement(pt, _c("v"))
        pt.set("idx", str(i))
        vnode = pt.find(_c("v"))
        if vnode is None:
            vnode = etree.SubElement(pt, _c("v"))
        vnode.text = str(value)
    for pt in pts[len(values) :]:
        cache.remove(pt)


def patch_category_chart(chart, labels: list[str], series_values: dict[str, list]) -> bool:
    """replace_data 대신 캐시만 고쳐 원본 차트 링크를 깨지 않는다."""
    space = chart._element
    sers = space.findall(f".//{_c('ser')}")
    if not sers:
        return False
    changed = False
    for ser in sers:
        name_el = ser.find(f".//{_c('tx')}//{_c('v')}")
        name = name_el.text if name_el is not None else ""
        vals = None
        for key, numbers in series_values.items():
            if key and key in (name or ""):
                vals = numbers
                break
        if vals is None and series_values:
            if "승" in (name or ""):
                vals = series_values.get("승")
            elif "하" in (name or ""):
                vals = series_values.get("하")
        cat = ser.find(_c("cat"))
        if cat is not None:
            for cache in cat.findall(f".//{_c('strCache')}"):
                _rewrite_cache_values(cache, labels)
                changed = True
        if vals is not None:
            val = ser.find(_c("val"))
            if val is not None:
                for cache in val.findall(f".//{_c('numCache')}"):
                    _rewrite_cache_values(cache, vals)
                    changed = True
    unlink_external_chart_data(chart)
    return changed


def find_tables(prs) -> list[FoundTable]:
    found = []
    for sidx, slide in enumerate(prs.slides):
        for shape in iter_shapes(slide):
            try:
                if not shape.has_table:
                    continue
            except Exception:
                continue
            rows = table_rows_text(shape.table)
            kind = classify_table(rows)
            found.append(
                FoundTable(
                    slide_index=sidx,
                    shape=shape,
                    kind=kind,
                    left_in=emu_to_in(shape.left),
                )
            )
    found.sort(key=lambda t: (t.slide_index, t.left_in))
    return found


def duplicate_slide(prs, index: int) -> int:
    """원본 슬라이드 레이아웃·도형·그림을 복제해 맨 뒤에 추가하고, 새 슬라이드 인덱스를 반환한다."""
    source = prs.slides[index]
    dest = prs.slides.add_slide(source.slide_layout)
    for shape in list(dest.shapes):
        shape.element.getparent().remove(shape.element)
    for shape in source.shapes:
        dest.shapes._spTree.insert_element_before(deepcopy(shape.element), "p:extLst")
    dest_rels = dest.part.rels
    for rId, rel in source.part.rels.items():
        if "notesSlide" in rel.reltype or "slideLayout" in rel.reltype:
            continue
        if rId in dest_rels:
            continue
        target = rel.target_ref if rel.is_external else rel.target_part
        dest_rels._rels[rId] = _Relationship(
            dest_rels._base_uri,
            rId,
            rel.reltype,
            rel._target_mode,
            target,
        )
    return len(prs.slides) - 1


def move_slide(prs, old_index: int, new_index: int) -> None:
    if old_index == new_index:
        return
    sld_id_lst = prs.slides._sldIdLst
    entries = list(sld_id_lst)
    el = entries[old_index]
    sld_id_lst.remove(el)
    sld_id_lst.insert(new_index, el)


# 1x1 white PNG. New photo-form slides retarget blips here so original photos stay intact.
_BLANK_PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00"
    b"\x01\x01\x01\x00\x18\xdd\x8d\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
)
_A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
_R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"


def blank_pictures_on_slide(slide) -> int:
    """이 슬라이드의 그림만 빈 칸으로 바꾼다. 같은 이미지를 쓰던 다른 슬라이드는 건드리지 않는다."""
    sp_tree = slide.shapes._spTree
    blips = sp_tree.findall(f".//{_A_BLIP}")
    if not blips:
        return 0
    try:
        _image_part, r_id = slide.part.get_or_add_image_part(BytesIO(_BLANK_PNG))
    except Exception:
        return 0
    changed = 0
    for blip in blips:
        blip.set(_R_EMBED, r_id)
        changed += 1
    return changed
