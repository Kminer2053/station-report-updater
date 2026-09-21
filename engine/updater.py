from __future__ import annotations

import calendar
import math
import re
from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import date

from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

from engine.excel_loader import RidershipSet, RidershipYear, StoreRecord, format_iso_date, format_won, _to_str
from engine.name_match import MatchResult, greedy_match, score_names, build_code_map, is_usable_ppt_name
from engine.pptx_io import (
    CODE_RE,
    DATA_FONT,
    DATA_SIZE_PT,
    apply_cell_run_font,
    apply_run_font,
    FoundTable,
    classify_table,
    duplicate_slide,
    emu_to_in,
    equalize_data_row_heights,
    find_header_columns,
    find_tables,
    get_cell_text,
    header_key,
    is_empty_row,
    iter_shapes,
    insert_table_row_before,
    labeled_review_segments,
    last_cell,
    move_slide,
    blank_pictures_on_slide,
    norm_text,
    patch_category_chart,
    replace_codes_in_shape,
    sample_cell_font,
    sample_run_font,
    sample_table_data_bold,
    sample_table_data_font,
    set_cell_text_preserve,
    set_run_text_preserve,
    table_rows_text,
    unify_table_data_style,
    write_paragraph_segments,
)


def detect_station(prs) -> str:
    for shape in iter_shapes(prs.slides[0]):
        try:
            if shape.has_text_frame:
                text = shape.text_frame.text or ""
                match = re.search(r"([가-힣A-Za-z0-9]+)\s*역", text)
                if match:
                    return match.group(1)
        except Exception:
            continue
    return "주안"


def _shape_key(slide_index: int, shape) -> tuple:
    """슬라이드가 끼워져도 같은 도형을 가리킨다.

    슬라이드 번호는 표 복제 후 밀리고, XML 요소 id()는 프록시가
    새로 생기면 바뀐다. 슬라이드 part 경로 + shape_id는 둘 다 유지된다.
    """
    sid = int(getattr(shape, "shape_id", 0))
    try:
        partname = str(shape.part.partname)
        if partname:
            return (partname, sid)
    except Exception:
        pass
    return ("idx", slide_index, sid)


def _snap_slide(*snaps: dict | None, key=None) -> int:
    for snap in snaps:
        loc = (snap or {}).get("slides", {}).get(key)
        if loc is not None:
            return loc
    if isinstance(key, tuple) and key:
        try:
            return int(key[0])
        except (TypeError, ValueError):
            return 0
    return 0


def capture_snapshot(prs) -> dict:
    tables: dict[tuple, list[list[str]]] = {}
    texts: dict[tuple, list[str]] = {}
    charts: dict[tuple, dict] = {}
    slides: dict[tuple, int] = {}
    for sidx, slide in enumerate(prs.slides):
        for shape in iter_shapes(slide):
            key = _shape_key(sidx, shape)
            slides[key] = sidx
            try:
                if shape.has_table:
                    tables[key] = table_rows_text(shape.table)
            except Exception:
                pass
            try:
                if shape.has_text_frame and not shape.has_table:
                    texts[key] = [para.text or "" for para in shape.text_frame.paragraphs]
            except Exception:
                pass
            try:
                if shape.has_chart:
                    plot = shape.chart.plots[0]
                    charts[key] = {
                        "cats": [str(c) for c in plot.categories],
                        "series": {s.name: [v for v in s.values] for s in plot.series},
                    }
            except Exception:
                pass
    return {"tables": tables, "texts": texts, "charts": charts, "slides": slides}


def extract_detail_entities(prs) -> list[dict]:
    rows = []
    for found in find_tables(prs):
        if found.kind not in {"store_detail", "vending_detail"}:
            continue
        texts = table_rows_text(found.shape.table)
        cols = find_header_columns(texts[0])
        for row in texts[1:]:
            if is_empty_row(row):
                continue
            name = row[cols["name"]] if "name" in cols and cols["name"] < len(row) else ""
            code = row[cols["code"]] if "code" in cols and cols["code"] < len(row) else ""
            brand = row[cols["brand"]] if "brand" in cols and cols["brand"] < len(row) else ""
            code = re.sub(r"\D", "", code)
            if not name and not code:
                continue
            rows.append(
                {
                    "code": code,
                    "name": name,
                    "brand": brand,
                    "slide": found.slide_index,
                    "src": found.kind,
                }
            )
    return rows


WEAK_CARD_NAMES = {"스토리웨이", "편의점", "전문점", "자판기"}


def _photo_card_code(rows: list[list[str]]) -> str:
    for i, row in enumerate(rows):
        joined = header_key("".join(row))
        if "개소코드" in joined and i + 1 < len(rows):
            nxt = "".join(rows[i + 1])
            codes = CODE_RE.findall(nxt)
            if codes:
                return codes[0]
        codes = [c for c in row if re.fullmatch(r"\d{5,7}", c or "")]
        if codes:
            return codes[0]
    return ""


def extract_photo_entities(prs) -> list[dict]:
    entities = []
    for found in find_tables(prs):
        if found.kind != "photo_card":
            continue
        rows = table_rows_text(found.shape.table)
        code = _photo_card_code(rows)
        name = rows[0][0] if rows and rows[0] else ""
        if not code or header_key(name) in WEAK_CARD_NAMES:
            continue
        entities.append(
            {
                "code": code,
                "name": name,
                "brand": name,
                "slide": found.slide_index,
                "src": "card",
            }
        )
    return entities


def extract_label_entities(prs) -> list[dict]:
    entities = []
    for sidx, slide in enumerate(prs.slides):
        for shape in iter_shapes(slide):
            try:
                if shape.has_table:
                    continue
                if not shape.has_text_frame:
                    continue
            except Exception:
                continue
            text = shape.text_frame.text or ""
            if "[삭제 메모]" in text or getattr(shape, "name", "") == "korail_review_note":
                continue
            codes = CODE_RE.findall(text)
            if not codes:
                continue
            name = CODE_RE.sub("", text)
            name = re.sub(r"NO\s*\d+\.?", "", name, flags=re.I)
            name = " ".join(name.split())
            if len(name) < 2:
                continue
            entities.append(
                {
                    "code": codes[0],
                    "name": name,
                    "brand": "",
                    "slide": sidx,
                    "src": "label",
                }
            )
    return entities


def _is_code_typo(a: str, b: str) -> bool:
    if not a or not b or a == b:
        return a == b
    longer, shorter = (a, b) if len(a) > len(b) else (b, a)
    if len(longer) - len(shorter) == 1:
        for i in range(len(longer)):
            if longer[:i] + longer[i + 1 :] == shorter:
                return True
    return False


def add_code_aliases(matches: list[MatchResult], ppt_rows: list[dict]) -> None:
    mapped_old = {m.old_code for m in matches if m.old_code}
    leftover = [p for p in ppt_rows if p.get("code") and p["code"] not in mapped_old]
    for ppt in leftover:
        best = None
        best_score = 0.0
        for match in matches:
            sc = score_names(ppt.get("name") or "", match.excel_name, ppt.get("brand") or "")
            if sc > best_score:
                best_score = sc
                best = match
        if best and best_score >= 85:
            best.aliases.append(ppt["code"])
            continue
        for match in matches:
            if _is_code_typo(ppt["code"], match.old_code):
                match.aliases.append(ppt["code"])
                break


def collect_all_codes(prs) -> set[str]:
    codes: set[str] = set()
    for slide in prs.slides:
        for shape in iter_shapes(slide):
            try:
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            codes.update(CODE_RE.findall(cell.text or ""))
            except Exception:
                pass
            try:
                if shape.has_text_frame:
                    codes.update(CODE_RE.findall(shape.text_frame.text or ""))
            except Exception:
                pass
    return codes


def expand_typo_mapping(prs, matches: list[MatchResult], mapping: dict[str, str]) -> None:
    new_codes = {m.new_code for m in matches}
    for code in collect_all_codes(prs):
        if code in mapping or code in new_codes:
            continue
        for match in matches:
            if _is_code_typo(code, match.old_code):
                mapping[code] = match.new_code
                if code not in match.aliases:
                    match.aliases.append(code)
                break


def match_presentation(prs, stores: list[StoreRecord]):
    detail = extract_detail_entities(prs)
    labels = extract_label_entities(prs)
    cards = extract_photo_entities(prs)
    by_code: dict[str, dict] = {}
    nameless: list[dict] = []

    def rank(item: dict) -> tuple:
        name = item.get("name") or ""
        brand = item.get("brand") or ""
        return (is_usable_ppt_name(name, brand), len(name) + len(brand), 1 if item.get("src") == "detail" else 0)

    for item in detail + labels + cards:
        code = item.get("code") or ""
        if not code:
            nameless.append(item)
            continue
        prev = by_code.get(code)
        if prev is None or rank(item) > rank(prev):
            by_code[code] = item
    ppt_rows = list(by_code.values()) + nameless
    matches, leftover_ppt, leftover_xls = greedy_match(ppt_rows, stores)
    add_code_aliases(matches, ppt_rows)
    mapping = build_code_map(matches)
    expand_typo_mapping(prs, matches, mapping)
    aliased = {a for m in matches for a in m.aliases}
    leftover_ppt = [
        p
        for p in leftover_ppt
        if (p.get("code") or "") not in mapping and (p.get("code") or "") not in aliased
    ]
    for match in matches:
        match.aliases = list(dict.fromkeys(match.aliases))
    return matches, mapping, leftover_ppt, leftover_xls


def apply_code_mapping(prs, mapping: dict[str, str]) -> int:
    changed = 0
    for slide in prs.slides:
        for shape in iter_shapes(slide):
            changed += replace_codes_in_shape(shape, mapping)
    return changed


def _snapshot_rows_from_texts(texts: list[list[str]]) -> dict[str, list[str]]:
    snap = {}
    if not texts:
        return snap
    cols = find_header_columns(texts[0])
    if "code" not in cols:
        return snap
    for row in texts[1:]:
        code = re.sub(r"\D", "", row[cols["code"]]) if cols["code"] < len(row) else ""
        if code:
            snap[code] = row
    return snap


def _prev_row_for_store(
    snap: dict[str, list[str]],
    store: StoreRecord,
    mapping: dict[str, str],
) -> list[str] | None:
    if store.code in snap:
        return snap[store.code]
    for old, new in mapping.items():
        if new == store.code and old in snap:
            return snap[old]
    return None


def _parse_area(text: str | None) -> float | None:
    raw = (text or "").replace("㎡", "").replace("m2", "").replace(",", "").strip()
    if not raw:
        return None
    try:
        return float(raw)
    except ValueError:
        match = re.search(r"[\d.]+", raw)
        if not match:
            return None
        try:
            return float(match.group())
        except ValueError:
            return None


def _inherit_kw_from_unmatched(
    store: StoreRecord,
    orig_texts: list[list[str]] | None,
    mapping: dict[str, str],
    used_codes: set[str],
) -> str:
    """엑셀에만 있는 신규 매장은 KW가 없다. 같은 위치·비슷한 면적의 삭제 행에서 이어받는다."""
    if not orig_texts:
        return ""
    cols = find_header_columns(orig_texts[0])
    if "kw" not in cols:
        return ""
    loc_i = cols.get("location")
    area_i = cols.get("area")
    kw_i = cols["kw"]
    code_i = cols.get("code")
    store_loc = header_key(store.location_display())
    store_area = _parse_area(store.area)
    candidates: list[tuple[float, str]] = []
    for row in orig_texts[1:]:
        if is_empty_row(row):
            continue
        code = re.sub(r"\D", "", row[code_i]) if code_i is not None and code_i < len(row) else ""
        mapped = mapping.get(code, code) if code else ""
        if (mapped and mapped in used_codes) or (code and code in used_codes):
            continue
        kw = (row[kw_i] if kw_i < len(row) else "").strip()
        if not kw:
            continue
        row_loc = header_key(row[loc_i]) if loc_i is not None and loc_i < len(row) else ""
        row_area = _parse_area(row[area_i]) if area_i is not None and area_i < len(row) else None
        if store_loc and row_loc and store_loc != row_loc:
            continue
        if store_area is not None and row_area is not None:
            diff = abs(store_area - row_area)
        elif store_loc and row_loc and store_loc == row_loc:
            diff = 0.0
        else:
            continue
        candidates.append((diff, kw))
    if not candidates:
        return ""
    candidates.sort(key=lambda item: item[0])
    best_diff, best_kw = candidates[0]
    if store_area is not None and best_diff > 2.0:
        return ""
    return best_kw


SHOP_NAME_BREAKS = ("맞이방(단기)", "맞이방", "타는곳", "대합실", "직영")


def wrap_shop_name(name: str) -> str:
    """원본처럼 위치 / 업종을 두 줄로 나눠 표 높이가 밀리지 않게 한다."""
    raw = (name or "").replace("\r\n", "\n").strip()
    if not raw or "\n" in raw:
        return raw
    for token in SHOP_NAME_BREAKS:
        idx = raw.find(token)
        if idx < 0:
            continue
        cut = idx + len(token)
        rest = raw[cut:].lstrip()
        if rest:
            return f"{raw[:cut]}\n{rest}"
    paren = raw.find("(")
    if paren >= 8:
        return f"{raw[:paren].rstrip()}\n{raw[paren:]}"
    return raw


def _location_from_prev(excel_loc: str, orig_loc: str) -> str:
    excel = (excel_loc or "").replace("\n", " ").strip()
    orig = (orig_loc or "").strip()
    if not orig:
        return excel_loc
    compact_excel = excel.replace(" ", "")
    compact_orig = orig.replace("\n", "").replace(" ", "")
    if compact_excel and compact_excel in compact_orig and len(compact_excel) < len(compact_orig):
        return orig
    return excel_loc


def _write_store_row(
    row,
    store: StoreRecord,
    cols: dict[str, int],
    no: int,
    unit: str,
    prev: list[str] | None,
    is_new: bool,
    kw_fallback: str = "",
    font_name: str = DATA_FONT,
    size_pt: float = DATA_SIZE_PT,
    bold: bool = True,
):
    location = store.location_display()
    if prev and "location" in cols and cols["location"] < len(prev):
        location = _location_from_prev(location, prev[cols["location"]])
        if not location:
            location = prev[cols["location"]]
    values = {
        "no": str(no),
        "code": store.code,
        "name": wrap_shop_name(store.display_name()),
        "brand": store.short_brand() or store.brand,
        "owner": store.owner,
        "location": location,
        "area": store.area,
        "start": store.contract_start,
        "end": store.contract_end,
        "sales": format_won(store.monthly_sales, unit),
    }
    prev_kw = ""
    if prev and "kw" in cols and cols["kw"] < len(prev):
        prev_kw = (prev[cols["kw"]] or "").strip()
    if "kw" in cols:
        values["kw"] = prev_kw or kw_fallback or ""
    for key, col_i in cols.items():
        if key not in values or col_i >= len(row.cells):
            continue
        text = values[key]
        old = prev[col_i] if prev and col_i < len(prev) else ""
        changed = is_new or norm_text(old) != norm_text(text)
        set_cell_text_preserve(
            row.cells[col_i],
            text,
            font_name,
            size_pt,
            review_red=changed,
            bold=bold,
        )


def rebuild_detail_table(
    table,
    stores: list[StoreRecord],
    unit: str,
    orig_texts: list[list[str]] | None = None,
    mapping: dict[str, str] | None = None,
    start_no: int = 1,
) -> None:
    texts = table_rows_text(table)
    cols = find_header_columns(texts[0])
    snap = _snapshot_rows_from_texts(orig_texts or texts)
    mapping = mapping or {}
    used_codes = {s.code for s in stores}
    orig_data_n = max(len(table.rows) - 1, 1)
    n_write = min(len(stores), orig_data_n)
    font_name, size_pt = sample_table_data_font(table, (DATA_FONT, DATA_SIZE_PT))
    bold = sample_table_data_bold(table, True)
    for i in range(n_write):
        store = stores[i]
        prev = _prev_row_for_store(snap, store, mapping)
        kw_fallback = ""
        if "kw" in cols:
            prev_kw = (prev[cols["kw"]] or "").strip() if prev and cols["kw"] < len(prev) else ""
            if not prev_kw:
                kw_fallback = _inherit_kw_from_unmatched(store, orig_texts, mapping, used_codes)
        _write_store_row(
            table.rows[i + 1],
            store,
            cols,
            start_no + i,
            unit,
            prev,
            is_new=prev is None,
            kw_fallback=kw_fallback,
            font_name=font_name,
            size_pt=size_pt,
            bold=bold,
        )
    for i in range(n_write + 1, len(table.rows)):
        row = table.rows[i]
        for cell in row.cells:
            set_cell_text_preserve(
                cell, "", font_name, size_pt, review_red=False, bold=bold
            )
    unify_table_data_style(table, font_name, size_pt, bold)
    equalize_data_row_heights(table, orig_data_n)


def _ppt_codes_in_order(orig_texts: list[list[str]] | None) -> list[str]:
    if not orig_texts:
        return []
    cols = find_header_columns(orig_texts[0])
    if "code" not in cols:
        return []
    codes = []
    for row in orig_texts[1:]:
        if is_empty_row(row):
            continue
        code = re.sub(r"\D", "", row[cols["code"]]) if cols["code"] < len(row) else ""
        if code:
            codes.append(code)
    return codes


def _order_stores_like_ppt(
    tables: list[FoundTable],
    stores: list[StoreRecord],
    snap: dict | None,
    mapping: dict[str, str] | None,
) -> list[StoreRecord]:
    by_code = {s.code: s for s in stores}
    mapping = mapping or {}
    table_snap = (snap or {}).get("tables", {})
    ordered: list[StoreRecord] = []
    seen: set[str] = set()
    for found in tables:
        orig = table_snap.get(_shape_key(found.slide_index, found.shape))
        for code in _ppt_codes_in_order(orig):
            mapped = mapping.get(code, code)
            store = by_code.get(mapped) or by_code.get(code)
            if store and store.code not in seen:
                ordered.append(store)
                seen.add(store.code)
    for store in stores:
        if store.code not in seen:
            ordered.append(store)
    return ordered


def detect_table_unit(prs, found: FoundTable, default: str) -> str:
    blob = ""
    try:
        blob += "".join("".join(r) for r in table_rows_text(found.shape.table)[:1])
    except Exception:
        pass
    try:
        slide = prs.slides[found.slide_index]
        for shape in iter_shapes(slide):
            try:
                if shape.has_text_frame:
                    blob += shape.text_frame.text or ""
            except Exception:
                continue
    except Exception:
        pass
    compact = blob.replace(" ", "")
    if "천원" in compact:
        return "천원"
    if "단위" in compact and "원" in compact:
        return "원"
    return default


def _table_data_capacity(tables: list[FoundTable]) -> int:
    return sum(max(len(found.shape.table.rows) - 1, 0) for found in tables)


def ensure_detail_slides(prs, kind: str, n_rows: int) -> None:
    """원본 표 행 수를 넘기지 않도록, 모자라면 마지막 세부내역 슬라이드를 복제한다."""
    if n_rows <= 0:
        return
    tables = [t for t in find_tables(prs) if t.kind == kind]
    if not tables:
        return
    cap = _table_data_capacity(tables)
    if n_rows <= cap:
        return
    source_idx = tables[-1].slide_index
    per_slide = _table_data_capacity([t for t in tables if t.slide_index == source_idx])
    if per_slide <= 0:
        return
    extra = math.ceil((n_rows - cap) / per_slide)
    insert_at = source_idx + 1
    for _ in range(extra):
        new_idx = duplicate_slide(prs, source_idx)
        move_slide(prs, new_idx, insert_at)
        insert_at += 1


NUMBERED_NAME_RE = re.compile(r"^(?P<prefix>.+?)(?P<num>\d+)\s*(?P<suffix>\([^)]*\))?$")


def _parse_numbered_name(name: str) -> tuple[str, int, str] | None:
    match = NUMBERED_NAME_RE.match((name or "").strip())
    if not match:
        return None
    prefix = match.group("prefix") or ""
    if len(prefix) < 2:
        return None
    return prefix, int(match.group("num")), match.group("suffix") or ""


def collapse_numbered_vending_names(stores: list[StoreRecord]) -> list[StoreRecord]:
    """귀여움1, 귀여움2…처럼 번호만 다른 자판기를 '귀여움1~50(직영)' 한 행으로 묶는다."""
    buckets: dict[tuple[str, str], list[StoreRecord]] = defaultdict(list)
    parsed: dict[str, tuple[str, int, str]] = {}
    for store in stores:
        parsed_name = _parse_numbered_name(store.name)
        if parsed_name is None:
            continue
        prefix, _num, suffix = parsed_name
        parsed[store.code] = parsed_name
        buckets[(prefix, suffix)].append(store)

    seen: set[tuple[str, str]] = set()
    result: list[StoreRecord] = []
    for store in stores:
        parsed_name = parsed.get(store.code)
        if parsed_name is None:
            result.append(store)
            continue
        prefix, _num, suffix = parsed_name
        key = (prefix, suffix)
        members = buckets[key]
        if len(members) < 2:
            result.append(store)
            continue
        if key in seen:
            continue
        seen.add(key)
        members_sorted = sorted(members, key=lambda s: parsed[s.code][1])
        nums = [parsed[s.code][1] for s in members_sorted]
        suffix_text = suffix
        if not suffix_text and any("직영" in f"{s.owner}{s.name}" for s in members_sorted):
            suffix_text = "(직영)"
        grouped_name = f"{prefix}{nums[0]}~{nums[-1]}{suffix_text}"
        sales_vals = [s.monthly_sales for s in members_sorted if s.monthly_sales is not None]
        result.append(
            replace(
                members_sorted[0],
                name=grouped_name,
                monthly_sales=sum(sales_vals) if sales_vals else None,
            )
        )
    return result


def _distribute(
    tables: list[FoundTable],
    stores: list[StoreRecord],
    unit: str,
    snap: dict | None = None,
    mapping: dict[str, str] | None = None,
    prs=None,
) -> None:
    if not tables:
        return
    remaining = _order_stores_like_ppt(tables, stores, snap, mapping)
    table_snap = (snap or {}).get("tables", {})
    start_no = 1
    for found in tables:
        cap = max(len(found.shape.table.rows) - 1, 1)
        chunk = remaining[:cap]
        remaining = remaining[cap:]
        orig = table_snap.get(_shape_key(found.slide_index, found.shape))
        use_unit = detect_table_unit(prs, found, unit) if prs is not None else unit
        rebuild_detail_table(found.shape.table, chunk, use_unit, orig, mapping, start_no=start_no)
        start_no += len(chunk)


def update_detail_tables(
    prs,
    shops: list[StoreRecord],
    vendings: list[StoreRecord],
    snap: dict | None = None,
    mapping: dict[str, str] | None = None,
) -> None:
    vend_rows = collapse_numbered_vending_names(vendings)
    ensure_detail_slides(prs, "store_detail", len(shops))
    ensure_detail_slides(prs, "vending_detail", len(vend_rows))
    found = find_tables(prs)
    store_tables = [t for t in found if t.kind == "store_detail"]
    vending_tables = [t for t in found if t.kind == "vending_detail"]
    _distribute(store_tables, shops, "천원", snap, mapping, prs=prs)
    _distribute(vending_tables, vend_rows, "원", snap, mapping, prs=prs)


def _labeled_value_from_rows(rows: list[list[str]], label_keys: tuple[str, ...]) -> str:
    for i, row in enumerate(rows):
        joined = header_key("".join(row))
        if not any(k in joined for k in label_keys):
            continue
        if i + 1 < len(rows) and rows[i + 1]:
            return (rows[i + 1][-1] or "").strip()
        if len(row) > 1:
            return (row[1] or "").strip()
    return ""


def _set_labeled_value(table, label_keys: tuple[str, ...], value: str, old_value: str | None = None) -> None:
    rows = table_rows_text(table)
    for i, row in enumerate(rows):
        joined = header_key("".join(row))
        if not any(k in joined for k in label_keys):
            continue
        target = None
        next_row = table.rows[i + 1] if i + 1 < len(table.rows) else None
        if next_row is not None and len(next_row.cells) > 0:
            target = next_row.cells[len(next_row.cells) - 1]
        elif len(table.rows[i].cells) > 1:
            target = table.rows[i].cells[1]
        if target is not None:
            old = old_value if old_value is not None else get_cell_text(target)
            changed = bool((value or "").strip()) and norm_text(old) != norm_text(value)
            set_cell_text_preserve(target, value, DATA_FONT, 10, review_red=changed, bold=True)
            return


def _is_new_shop_photo_target(store: StoreRecord) -> bool:
    if not store.is_shop:
        return False
    blob = f"{store.name}{store.brand}{store.biz_type}".replace(" ", "")
    if "매출계리" in blob:
        return False
    return True


def _clear_photo_card(table, *, clear_biz: bool = False) -> None:
    rows = table_rows_text(table)
    if rows and rows[0]:
        set_cell_text_preserve(
            table.rows[0].cells[0], "", DATA_FONT, 18, review_red=False, bold=True
        )
        if clear_biz and len(table.rows[0].cells) > 1:
            set_cell_text_preserve(
                table.rows[0].cells[1], "", DATA_FONT, 18, review_red=False, bold=True
            )
    _set_labeled_value(table, ("개소코드",), "", old_value="")
    _set_labeled_value(table, ("위치",), "", old_value="")
    _set_labeled_value(table, ("면적",), "", old_value="")


def _write_photo_card(
    table,
    store: StoreRecord,
    *,
    keep_title_if_no_brand: bool = False,
    orig: list[list[str]] | None = None,
    live_title: str = "",
    live_biz: str = "",
) -> str:
    excel_brand = store.short_brand()
    if excel_brand:
        title = store.card_title()
    elif keep_title_if_no_brand and (live_title or "").strip():
        title = live_title
    else:
        title = store.card_title()
    old_loc = ""
    old_area = ""
    area = ""
    if table.rows:
        title_name, title_base = sample_cell_font(table.rows[0].cells[0])
        set_cell_text_preserve(
            table.rows[0].cells[0],
            title,
            title_name or DATA_FONT,
            title_base or 18.0,
            review_red=norm_text(live_title) != norm_text(title),
            bold=True,
        )
        if len(table.rows[0].cells) > 1 and store.biz_type:
            old_biz = live_biz or (
                orig[0][1] if orig and orig[0] and len(orig[0]) > 1 else ""
            )
            biz_name, biz_pt = sample_cell_font(table.rows[0].cells[1])
            set_cell_text_preserve(
                table.rows[0].cells[1],
                store.biz_type,
                biz_name or DATA_FONT,
                biz_pt or 18.0,
                review_red=norm_text(old_biz) != norm_text(store.biz_type),
                bold=True,
            )
    old_code = _labeled_value_from_rows(orig or [], ("개소코드",)) or _photo_card_code(orig or [])
    _set_labeled_value(table, ("개소코드",), store.code, old_value=old_code)
    if store.location:
        old_loc = _labeled_value_from_rows(orig or [], ("위치",))
        _set_labeled_value(table, ("위치",), store.location_display(), old_value=old_loc)
    if store.area:
        area = store.area if str(store.area).endswith("㎡") else f"{store.area}㎡"
        old_area = _labeled_value_from_rows(orig or [], ("면적",))
        _set_labeled_value(table, ("면적",), area, old_value=old_area)
    return title


def update_photo_cards(prs, by_code: dict[str, StoreRecord], snap: dict | None = None) -> dict:
    cleared = []
    updated = []
    table_snap = (snap or {}).get("tables", {})
    for found in find_tables(prs):
        if found.kind != "photo_card":
            continue
        table = found.shape.table
        orig = table_snap.get(_shape_key(found.slide_index, found.shape))
        try:
            rows = table_rows_text(table)
            code = _photo_card_code(rows)
        except Exception:
            continue
        store = by_code.get(code)
        live_title = get_cell_text(table.rows[0].cells[0]) if rows and rows[0] else ""
        live_biz = (
            get_cell_text(table.rows[0].cells[1])
            if rows and rows[0] and len(table.rows[0].cells) > 1
            else ""
        )
        if store is None:
            _clear_photo_card(table)
            if code:
                cleared.append(code)
            continue
        old_title = live_title
        old_code = _labeled_value_from_rows(orig or [], ("개소코드",)) or _photo_card_code(orig or [])
        old_loc = _labeled_value_from_rows(orig or [], ("위치",))
        old_area = _labeled_value_from_rows(orig or [], ("면적",))
        title = _write_photo_card(
            table,
            store,
            keep_title_if_no_brand=True,
            orig=orig,
            live_title=live_title,
            live_biz=live_biz,
        )
        area = ""
        if store.area:
            area = store.area if str(store.area).endswith("㎡") else f"{store.area}㎡"
        if (
            norm_text(old_title) != norm_text(title)
            or (old_code or "") != store.code
            or (store.location and norm_text(old_loc) != norm_text(store.location))
            or (store.area and norm_text(old_area) != norm_text(area))
        ):
            updated.append(title or store.code)
    unify_photo_card_fonts(prs)
    return {"photo_updated": updated, "photo_cleared": cleared}


def _strip_review_notes_from_slide(slide) -> None:
    for shape in list(slide.shapes):
        try:
            name = getattr(shape, "name", "") or ""
            text = ""
            if shape.has_text_frame:
                text = shape.text_frame.text or ""
        except Exception:
            continue
        if name == REVIEW_SHAPE_NAME or text.startswith(REVIEW_NOTE_PREFIX):
            try:
                shape.element.getparent().remove(shape.element)
            except Exception:
                pass


def _clear_feature_product_line(slide) -> None:
    for shape in iter_shapes(slide):
        try:
            if not shape.has_text_frame or shape.has_table:
                continue
        except Exception:
            continue
        if "주력" not in (shape.text_frame.text or ""):
            continue
        for para in shape.text_frame.paragraphs:
            text = para.text or ""
            if "주력" not in text:
                continue
            new_text = re.sub(r"(주력\s*품목\s*:\s*)(.*)", r"\1", text)
            if new_text != text:
                set_run_text_preserve(para, new_text)


def add_photo_forms_for_new_shops(prs, leftover_xls: list[StoreRecord]) -> list[str]:
    """엑셀에만 있는 편의점·전문점용으로 빈 사진 폼 슬라이드를 붙인다. 기존 사진은 그대로 둔다."""
    existing: set[str] = set()
    for found in find_tables(prs):
        if found.kind != "photo_card":
            continue
        try:
            code = _photo_card_code(table_rows_text(found.shape.table))
        except Exception:
            continue
        if code:
            existing.add(code)
    new_shops = [
        store
        for store in leftover_xls
        if _is_new_shop_photo_target(store) and store.code not in existing
    ]
    if not new_shops:
        return []
    photo_tables = [found for found in find_tables(prs) if found.kind == "photo_card"]
    if not photo_tables:
        return []
    source_idx = photo_tables[-1].slide_index
    per_slide = max(
        len([found for found in photo_tables if found.slide_index == source_idx]),
        1,
    )
    extra = math.ceil(len(new_shops) / per_slide)
    insert_at = source_idx + 1
    new_indices: list[int] = []
    for _ in range(extra):
        new_idx = duplicate_slide(prs, source_idx)
        move_slide(prs, new_idx, insert_at)
        new_indices.append(insert_at)
        insert_at += 1
    remaining = list(new_shops)
    added: list[str] = []
    for sidx in new_indices:
        slide = prs.slides[sidx]
        blank_pictures_on_slide(slide)
        _strip_review_notes_from_slide(slide)
        _clear_feature_product_line(slide)
        cards = [
            found
            for found in find_tables(prs)
            if found.kind == "photo_card" and found.slide_index == sidx
        ]
        cards.sort(key=lambda found: found.left_in)
        for found in cards:
            if remaining:
                store = remaining.pop(0)
                title = _write_photo_card(table=found.shape.table, store=store)
                added.append(title or store.code)
            else:
                _clear_photo_card(found.shape.table, clear_biz=True)
    unify_photo_card_fonts(prs)
    return added


def unify_photo_card_fonts(prs) -> None:
    """같은 슬라이드의 사진 카드 제목 크기·굵기를 맞춘다. 사진은 건드리지 않는다."""
    by_slide: dict[int, list] = defaultdict(list)
    for found in find_tables(prs):
        if found.kind == "photo_card":
            by_slide[found.slide_index].append(found.shape.table)
    for tables in by_slide.values():
        title_pts: list[float] = []
        for table in tables:
            title = get_cell_text(table.rows[0].cells[0])
            if not title:
                continue
            _name, base = sample_cell_font(table.rows[0].cells[0])
            title_pts.append(base or 18.0)
        title_size = max(title_pts) if title_pts else 18.0
        for table in tables:
            title_cell = table.rows[0].cells[0]
            name, _pt = sample_cell_font(title_cell)
            apply_cell_run_font(title_cell, name or DATA_FONT, title_size, True)
            if len(table.rows[0].cells) > 1:
                biz = table.rows[0].cells[1]
                bname, bpt = sample_cell_font(biz)
                apply_cell_run_font(biz, bname or DATA_FONT, bpt or 18.0, True)
            for ri, row in enumerate(table.rows):
                if ri == 0:
                    continue
                for cell in row.cells:
                    if not (cell.text or "").strip():
                        continue
                    cname, cpt = sample_cell_font(cell)
                    apply_cell_run_font(cell, cname or DATA_FONT, cpt or 10.0, True)


def update_feature_boxes(prs, by_code: dict[str, StoreRecord]) -> int:
    """사진 카드와 같은 쪽의 특이사항 상자에서 계약기간/매출만 치환."""
    changed = 0
    for slide in prs.slides:
        cards_by_side: dict[str, list[str]] = defaultdict(list)
        boxes = []
        for shape in iter_shapes(slide):
            side = "left" if emu_to_in(shape.left) < 4.8 else "right"
            try:
                if shape.has_table:
                    rows = table_rows_text(shape.table)
                    if classify_table(rows) == "photo_card":
                        code = _photo_card_code(rows)
                        if code:
                            cards_by_side[side].append(code)
            except Exception:
                pass
            try:
                if shape.has_text_frame and "월평균매출액" in (shape.text_frame.text or ""):
                    boxes.append((side, shape))
            except Exception:
                pass
        for side, shape in boxes:
            codes = cards_by_side.get(side, [])
            store = by_code.get(codes[0]) if codes else None
            for para in shape.text_frame.paragraphs:
                text = para.text or ""
                new_text = text
                if store is None:
                    new_text = re.sub(r"(월평균매출액\s*:\s*)(.*)", r"\1", new_text)
                    new_text = re.sub(r"(계약기간\s*:\s*)(.*)", r"\1", new_text)
                else:
                    sales = format_won(store.monthly_sales, "천원")
                    period = ""
                    if store.contract_start or store.contract_end:
                        period = f"{store.contract_start}~{store.contract_end}"
                    if sales and "월평균매출액" in new_text:
                        new_text = re.sub(
                            r"(월평균매출액\s*:\s*)[0-9,\-]*\s*(천원)?",
                            rf"\g<1>{sales}천원",
                            new_text,
                            count=1,
                        )
                    if period and "계약기간" in new_text:
                        new_text = re.sub(
                            r"(계약기간\s*:?\s*)([^\n(]*)",
                            rf"\g<1>{period}",
                            new_text,
                            count=1,
                        )
                if new_text != text:
                    segs = labeled_review_segments(new_text, text)
                    if segs:
                        write_paragraph_segments(para, segs)
                    else:
                        set_run_text_preserve(para, new_text)
                    changed += 1
    return changed


SHOP_NO_RE = re.compile(r"(NO\s*)(\d+)(\s*\.)", re.I)
CIRCLED_RE = re.compile(r"[①-⑳]")


def _shop_no_in_text(text: str) -> int | None:
    match = SHOP_NO_RE.search(text or "")
    return int(match.group(2)) if match else None


def clear_unmatched_map_labels(prs, valid_codes: set[str]) -> list[dict]:
    cleared = []
    for sidx, slide in enumerate(prs.slides):
        for shape in iter_shapes(slide):
            try:
                if not shape.has_text_frame or shape.has_table:
                    continue
            except Exception:
                continue
            text = shape.text_frame.text or ""
            codes = CODE_RE.findall(text)
            if not codes:
                continue
            bad = [c for c in codes if c not in valid_codes]
            if not bad:
                continue
            if any(c in valid_codes for c in codes):
                new_text = text
                for code in bad:
                    new_text = re.sub(rf"(?<!\d){re.escape(code)}(?!\d)", "", new_text)
                new_text = re.sub(r"[|/·]\s*[|/·]", " | ", new_text)
                new_text = re.sub(r"(\s*[|/·]\s*)+$", "", new_text)
                new_text = re.sub(r"^(\s*[|/·]\s*)+", "", new_text)
                new_text = re.sub(
                    r"(?:^|[|/·]\s*)(?:커|캔|멀(?:\(소\))?|사|토캡슐|폰충)\s*$",
                    "",
                    new_text,
                )
                new_text = re.sub(r"[ \t]{2,}", " ", new_text).strip(" |·/\t")
                if new_text != text:
                    for para in shape.text_frame.paragraphs:
                        para_new = para.text or ""
                        for code in bad:
                            para_new = re.sub(rf"(?<!\d){re.escape(code)}(?!\d)", "", para_new)
                        para_new = re.sub(r"[|/·]\s*[|/·]", " | ", para_new)
                        para_new = re.sub(r"(\s*[|/·]\s*)+$", "", para_new)
                        para_new = re.sub(r"^(\s*[|/·]\s*)+", "", para_new)
                        para_new = re.sub(
                            r"(?:^|[|/·]\s*)(?:커|캔|멀(?:\(소\))?|사|토캡슐|폰충)\s*$",
                            "",
                            para_new,
                        )
                        para_new = re.sub(r"[ \t]{2,}", " ", para_new)
                        set_run_text_preserve(para, para_new.strip(" |·/\t"))
                    cleared.append(
                        {
                            "slide": sidx,
                            "code": ",".join(bad),
                            "name": ",".join(bad),
                            "shop_no": _shop_no_in_text(text),
                        }
                    )
                continue
            name = " ".join(CODE_RE.sub("", text).split())
            for para in shape.text_frame.paragraphs:
                set_run_text_preserve(para, "")
            cleared.append(
                {
                    "slide": sidx,
                    "code": ",".join(codes),
                    "name": name,
                    "shop_no": _shop_no_in_text(text),
                }
            )
    return cleared


def _subst_shop_no(para, number: int) -> bool:
    for run in para.runs:
        old = run.text or ""
        new, count = SHOP_NO_RE.subn(rf"\g<1>{number}\g<3>", old, count=1)
        if count and new != old:
            run.text = new
            return True
    old = para.text or ""
    new, count = SHOP_NO_RE.subn(rf"\g<1>{number}\g<3>", old, count=1)
    if count and new != old:
        set_run_text_preserve(para, new)
        return True
    return False


def _unify_circled_no_fonts(shape) -> None:
    """위치도 NO 라벨에서 동그라미 숫자·글꼴 없는 런을 NO 접두와 같은 굴림 크기로 맞춘다."""
    sample_name, sample_pt = None, None
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            text = run.text or ""
            if CIRCLED_RE.search(text):
                continue
            name, pt = sample_run_font(run)
            if not pt:
                continue
            sample_name = name or sample_name
            sample_pt = pt
            if re.search(r"NO", text, re.I):
                break
        if sample_pt is not None:
            break
    if sample_pt is None:
        sample_name, sample_pt = DATA_FONT, 8.0
    name = sample_name or DATA_FONT
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            text = run.text or ""
            if text == "":
                continue
            run_name, run_pt = sample_run_font(run)
            if CIRCLED_RE.search(text) or run_pt is None or run_name is None:
                apply_run_font(run, name, sample_pt)


def normalize_map_labels(
    prs,
    shops: list[StoreRecord] | None = None,
    snap: dict | None = None,
    mapping: dict[str, str] | None = None,
    map_cleared: list[dict] | None = None,
) -> int:
    """매장 라벨이 지워진 위치도에서만 남은 NOn. 빈 칸을 메우고, 동그라미 숫자 글꼴을 통일한다."""
    compact_slides = {item["slide"] for item in (map_cleared or []) if item.get("shop_no")}
    changed = 0
    by_slide: dict[int, list[tuple[int, object]]] = defaultdict(list)
    for sidx, slide in enumerate(prs.slides):
        for shape in iter_shapes(slide):
            try:
                if not shape.has_text_frame or shape.has_table:
                    continue
            except Exception:
                continue
            full = "\n".join(para.text or "" for para in shape.text_frame.paragraphs)
            if "NO" not in full.upper():
                continue
            _unify_circled_no_fonts(shape)
            match = SHOP_NO_RE.search(full)
            if match:
                by_slide[sidx].append((int(match.group(2)), shape))
    for sidx, items in by_slide.items():
        if sidx not in compact_slides:
            continue
        items.sort(key=lambda item: item[0])
        for number, shape in enumerate((item[1] for item in items), start=1):
            for para in shape.text_frame.paragraphs:
                if _subst_shop_no(para, number):
                    changed += 1
                    break
    return changed


VEND_KIND_ORDER = ["커피", "캔", "멀티", "사진", "장난감", "보관함", "기타"]
SHOP_TYPE_ORDER = ["편의점", "전문점", "직영점"]


def vending_kind(store: StoreRecord) -> str:
    blob = f"{store.major}{store.middle}{store.minor}{store.biz_detail}{store.name}"
    if store.is_locker or "보관함" in blob:
        return "보관함"
    if "사진" in blob:
        return "사진"
    if any(token in blob for token in ("장난감", "토이", "캡슐")):
        return "장난감"
    if "멀티" in blob:
        return "멀티"
    if "커피" in blob:
        return "커피"
    if any(token in blob.upper() for token in ("캔", "페트", "패트", "PAIR")):
        return "캔"
    return "기타"


def shop_place(store: StoreRecord) -> str:
    loc = (store.location or "").strip()
    name = store.name or ""
    blob = loc + name
    if "맞이방" in blob:
        return "맞이방"
    if "타는곳" in blob.replace(" ", "") or "타는 곳" in blob:
        return "타는곳"
    if loc:
        return store.location_display()
    for key in ("대합실", "상행", "하행"):
        if key in name:
            return key
    return "기타"


def shop_major_label(store: StoreRecord) -> str:
    if store.biz_type == "편의점":
        return "편의점"
    if store.biz_type:
        return store.biz_type
    return "전문점"


def _is_total_row_label(*parts: str) -> bool:
    blob = header_key("".join(parts))
    return "총계" in blob or blob in {"총", "합계", "총합"}


def _count_total_index(table) -> int | None:
    for i, row in enumerate(table.rows):
        if i == 0:
            continue
        labels = [get_cell_text(c) for c in list(row.cells)[:2]]
        if _is_total_row_label(*labels):
            return i
    return None


def _ensure_rows_before_total(table, needed: int, total_idx: int) -> int:
    available = total_idx - 1
    while available < needed:
        insert_table_row_before(table, total_idx, template_idx=1)
        total_idx += 1
        available += 1
    return total_idx


def _write_count_cell(
    row, text: str, font_name: str | None = None, size_pt: float | None = None, bold: bool = True
) -> None:
    cell = last_cell(row)
    if cell is not None:
        set_cell_text_preserve(cell, text, font_name, size_pt, bold=bold)


def _write_count_label(
    row, col: int, text: str, font_name: str, size_pt: float, bold: bool = True
) -> None:
    if col >= len(row.cells):
        return
    set_cell_text_preserve(row.cells[col], text, font_name, size_pt, bold=bold)


def shop_count_rows(shops: list[StoreRecord]) -> list[tuple[str, str, int]]:
    grouped: dict[tuple[str, str], int] = defaultdict(int)
    for store in shops:
        grouped[(shop_place(store), shop_major_label(store))] += 1
    places = sorted({place for place, _ in grouped}, key=lambda p: (p != "맞이방", p))
    rows = []
    for place in places:
        majors = [m for p, m in grouped if p == place]
        majors = sorted(set(majors), key=lambda m: (SHOP_TYPE_ORDER.index(m) if m in SHOP_TYPE_ORDER else 99, m))
        for major in majors:
            rows.append((place, major, grouped[(place, major)]))
    return rows


def vending_count_map(vendings: list[StoreRecord]) -> dict[str, int]:
    counts: dict[str, int] = defaultdict(int)
    for store in vendings:
        counts[vending_kind(store)] += 1
    return dict(counts)


def update_ops_counts(prs, shops: list[StoreRecord], vendings: list[StoreRecord]) -> None:
    conv = sum(1 for s in shops if s.biz_type == "편의점")
    spec = sum(1 for s in shops if s.biz_type != "편의점")
    locker = sum(1 for s in vendings if s.is_locker)
    machine = sum(1 for s in vendings if not s.is_locker)
    new_blob = f"편의점({conv}),전문점({spec}),자판기({machine}),보관함({locker})"
    for slide in prs.slides:
        for shape in iter_shapes(slide):
            try:
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            text = get_cell_text(cell)
                            if "운영현황" in text:
                                new_text = re.sub(
                                    r"편의점\([^)]*\),?\s*전문점\([^)]*\),?\s*자판기\([^)]*\),?\s*보관함\([^)]*\)",
                                    new_blob,
                                    text,
                                )
                                if new_text == text:
                                    new_text = re.sub(
                                        r"(운영현황\s*:\s*)(.*)",
                                        rf"\g<1>{new_blob}",
                                        text,
                                    )
                                set_cell_text_preserve(cell, new_text)
            except Exception:
                continue


def _unique_names(values: list[str]) -> list[str]:
    seen: list[str] = []
    for raw in values:
        name = (raw or "").strip()
        if name and name not in seen:
            seen.append(name)
    return seen


def _space_korean_name(name: str) -> str:
    compact = re.sub(r"\s+", "", name or "")
    if re.fullmatch(r"[가-힣]{2,4}", compact):
        return " ".join(compact)
    return name.strip()


def update_staff_line(prs, shops: list[StoreRecord], vendings: list[StoreRecord]) -> int:
    """담당 FC / SV를 마스터 담당FC의 고유 인원으로 맞춘다."""
    shop_fcs = _unique_names([s.fc for s in shops])
    vend_fcs = _unique_names([s.fc for s in vendings])
    fc_names = shop_fcs or vend_fcs
    sv_names = [n for n in vend_fcs if n not in fc_names]
    fc_text = ", ".join(_space_korean_name(n) for n in fc_names) if fc_names else ""
    sv_text = ", ".join(_space_korean_name(n) for n in sv_names) if sv_names else ""
    if not fc_text:
        return 0
    changed = 0
    staff_slash = re.compile(
        r"(담당\s*F\s*C\s*:\s*)(.*?)(\s*/\s*담당\s*S\s*V\s*:\s*)(.*)$",
        flags=re.I | re.S,
    )
    staff_both = re.compile(
        r"(담당\s*F\s*C\s*:\s*)(.*?)(\s*,\s*담당\s*S\s*V\s*:\s*)(.*)$",
        flags=re.I | re.S,
    )
    staff_fc = re.compile(r"(담당\s*F\s*C\s*:\s*)(.*)$", flags=re.I | re.S)
    for slide in prs.slides:
        for shape in iter_shapes(slide):
            cells = []
            try:
                if shape.has_table:
                    for row in shape.table.rows:
                        cells.extend(row.cells)
            except Exception:
                pass
            try:
                if shape.has_text_frame and not shape.has_table:
                    for para in shape.text_frame.paragraphs:
                        cells.append(para)
            except Exception:
                pass
            for item in cells:
                if hasattr(item, "text_frame"):
                    paras = item.text_frame.paragraphs
                    whole = get_cell_text(item)
                else:
                    paras = [item]
                    whole = item.text or ""
                if "담당" not in whole or "F" not in whole:
                    continue
                for para in paras:
                    text = para.text or ""
                    if "담당" not in text or not re.search(r"F\s*C", text, re.I):
                        continue
                    if sv_text and staff_slash.search(text):
                        new_text = staff_slash.sub(rf"\g<1>{fc_text}\g<3>{sv_text}", text)
                    elif sv_text and staff_both.search(text):
                        new_text = staff_both.sub(rf"\g<1>{fc_text}\g<3>{sv_text}", text)
                    elif sv_text:
                        new_text = staff_fc.sub(
                            rf"\g<1>{fc_text}/ 담당 S V: {sv_text}",
                            text,
                            count=1,
                        )
                    else:
                        new_text = staff_fc.sub(rf"\g<1>{fc_text}", text, count=1)
                    if new_text != text:
                        segs = labeled_review_segments(new_text, text)
                        if segs:
                            write_paragraph_segments(para, segs)
                        else:
                            set_run_text_preserve(para, new_text)
                        changed += 1
    return changed


VEND_LABEL_TO_KIND = {
    "커피": "커피",
    "캔": "캔",
    "멀티": "멀티",
    "사진": "사진",
    "장난감": "장난감",
    "토이캡슐": "장난감",
    "토이": "장난감",
    "캡슐": "장난감",
    "보관함": "보관함",
    "물품보관함": "보관함",
    "기타": "기타",
    "휴대폰충전기": "기타",
    "충전기": "기타",
}


def _count_label_kind(label: str) -> str | None:
    h = header_key(label)
    if h in VEND_LABEL_TO_KIND:
        return VEND_LABEL_TO_KIND[h]
    for ppt_label, kind in VEND_LABEL_TO_KIND.items():
        if header_key(ppt_label) in h or h in header_key(ppt_label):
            return kind
    if h in VEND_KIND_ORDER:
        return h
    return None


def update_count_summaries(prs, shops: list[StoreRecord], vendings: list[StoreRecord]) -> dict[str, str]:
    shop_rows = shop_count_rows(shops)
    vend_counts = vending_count_map(vendings)
    shop_summary = " · ".join(f"{place} {major} {n}개소" for place, major, n in shop_rows) + f" · 총 {len(shops)}개소"
    vend_parts = [f"{kind} {vend_counts[kind]}개소" for kind in VEND_KIND_ORDER if vend_counts.get(kind)]
    vend_summary = " · ".join(vend_parts) + f" · 총 {len(vendings)}개소"

    for found in find_tables(prs):
        if found.kind != "count_summary":
            continue
        table = found.shape.table
        font_name, size_pt = sample_table_data_font(table)
        header = [header_key(c) for c in table_rows_text(table)[0]]
        joined = "".join(header)
        total_idx = _count_total_index(table)
        if total_idx is None:
            continue
        if "업종" in joined:
            existing = []
            for i in range(1, total_idx):
                label = get_cell_text(table.rows[i].cells[0]).strip()
                if label:
                    existing.append(label)
            wanted: list[tuple[str, str]] = []
            covered: set[str] = set()
            for label in existing:
                kind = _count_label_kind(label) or header_key(label)
                if vend_counts.get(kind, 0) <= 0:
                    continue
                if kind in covered:
                    continue
                wanted.append((label, kind))
                covered.add(kind)
            for kind in VEND_KIND_ORDER:
                if vend_counts.get(kind, 0) and kind not in covered:
                    wanted.append((kind, kind))
                    covered.add(kind)
            total_idx = _ensure_rows_before_total(table, len(wanted), total_idx)
            for i, (label, kind) in enumerate(wanted):
                row = table.rows[i + 1]
                _write_count_label(row, 0, label, font_name, size_pt)
                _write_count_cell(row, str(vend_counts.get(kind, 0)), font_name, size_pt)
            for i in range(len(wanted) + 1, total_idx):
                row = table.rows[i]
                for cell in row.cells:
                    set_cell_text_preserve(cell, "", font_name, size_pt, review_red=False, bold=True)
            _write_count_cell(table.rows[total_idx], str(len(vendings)), font_name, size_pt)
            unify_table_data_style(table, font_name, size_pt, True)
        elif "대분류" in joined or "개소수" in joined:
            total_idx = _ensure_rows_before_total(table, len(shop_rows), total_idx)
            last_place = ""
            for i, (place, major, n) in enumerate(shop_rows):
                row = table.rows[i + 1]
                shown_place = place if place != last_place else ""
                last_place = place
                _write_count_label(row, 0, shown_place, font_name, size_pt)
                if len(row.cells) > 1:
                    _write_count_label(row, 1, major, font_name, size_pt)
                _write_count_cell(row, str(n), font_name, size_pt)
            for i in range(len(shop_rows) + 1, total_idx):
                row = table.rows[i]
                for cell in row.cells:
                    set_cell_text_preserve(cell, "", font_name, size_pt, review_red=False, bold=True)
            _write_count_cell(table.rows[total_idx], str(len(shops)), font_name, size_pt)
            unify_table_data_style(table, font_name, size_pt, True)
    return {"shop": shop_summary, "vending": vend_summary}


def _clip(text: str, limit: int = 70) -> str:
    text = re.sub(r"\s+", " ", (text or "")).strip()
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "…"


def _slide_title(prs, sidx: int) -> str:
    skip = {"contents", "1", "2", "3", "4", "5"}
    for shape in iter_shapes(prs.slides[sidx]):
        try:
            if shape.has_table:
                continue
            if not shape.has_text_frame:
                continue
        except Exception:
            continue
        text = (shape.text_frame.text or "").strip()
        first = (text.splitlines()[0] if text else "").strip()
        if not first or header_key(first).lower() in skip:
            continue
        if first.startswith("[삭제 메모]"):
            continue
        return first
    return ""


def _table_part_name(rows: list[list[str]]) -> str:
    kind = classify_table(rows)
    names = {
        "store_detail": "매장현황 세부내역",
        "vending_detail": "자판기현황 세부내역",
        "count_summary": "매장·자판기 개소 집계",
        "ridership": "승하차 인원 표",
        "ops_info": "담당 소속·운영현황",
        "photo_card": "매장 사진 카드",
    }
    if kind in names:
        return names[kind]
    blob = header_key("".join("".join(r) for r in rows[:4]))
    if "담당" in blob and ("FC" in blob.upper() or "F C" in blob):
        return "담당 소속·운영현황"
    if "운영현황" in blob or "담당소속" in blob:
        return "담당 소속·운영현황"
    return "표"


def _staff_pair(text: str) -> tuple[str, str]:
    text = text or ""
    fc_m = re.search(r"담당\s*F\s*C\s*:\s*", text, flags=re.I)
    if not fc_m:
        return "", ""
    sv_m = re.search(r"담당\s*S\s*V\s*:\s*", text, flags=re.I)
    if sv_m and sv_m.start() > fc_m.end():
        fc = text[fc_m.end() : sv_m.start()]
        sv = text[sv_m.end() :]
    else:
        fc = text[fc_m.end() :]
        sv = ""
    fc = re.sub(r"\s+", " ", fc).strip(" ,")
    sv = re.sub(r"\s+", " ", sv).strip(" ,")
    return fc, sv


def _line_change(old: str, new: str) -> str | None:
    old, new = (old or "").strip(), (new or "").strip()
    if norm_text(old) == norm_text(new):
        return None
    if "담당" in old + new and re.search(r"F\s*C", old + new, re.I):
        ofc, osv = _staff_pair(old)
        nfc, nsv = _staff_pair(new)
        bits = []
        if ofc != nfc:
            bits.append(f"담당 FC {ofc or '(없음)'} → {nfc or '(없음)'}")
        if osv != nsv:
            bits.append(f"담당 SV {osv or '(없음)'} → {nsv or '(없음)'}")
        return " · ".join(bits) if bits else f"{_clip(old, 40)} → {_clip(new, 40)}"
    if "운영현황" in old or "운영현황" in new:
        def blob(text: str) -> str:
            match = re.search(r"운영현황\s*:\s*(.*)", text)
            return (match.group(1) if match else text).strip()
        return f"운영현황 {blob(old) or '(없음)'} → {blob(new)}"
    if "승" in old and "하" in old:
        return None
    return f"{_clip(old, 36) or '(빈칸)'} → {_clip(new, 36) or '(빈칸)'}"


def _header_labels(header: list[str]) -> list[str]:
    labels = []
    for raw in header:
        h = header_key(raw)
        if "개소코드" in h:
            labels.append("개소코드")
        elif "개소명" in h:
            labels.append("개소명")
        elif "브랜드" in h:
            labels.append("브랜드")
        elif "대표" in h:
            labels.append("대표자")
        elif h == "위치":
            labels.append("위치")
        elif "면적" in h:
            labels.append("면적")
        elif "최초계약" in h:
            labels.append("최초계약일")
        elif "만료" in h:
            labels.append("만료일")
        elif "월평균" in h:
            labels.append("월평균매출")
        elif "용량" in h:
            labels.append("계약용량")
        else:
            labels.append(raw.strip() or "값")
    return labels


def _detail_row_changes(old_rows: list[list[str]], new_rows: list[list[str]]) -> list[str]:
    header = new_rows[0] if new_rows else (old_rows[0] if old_rows else [])
    cols = find_header_columns(header)
    labels = _header_labels(header)
    code_i = cols.get("code", 1 if len(header) > 1 else 0)
    name_i = cols.get("name", 2 if len(header) > 2 else 0)

    def keyed(rows: list[list[str]]) -> dict[str, list[str]]:
        out = {}
        for row in rows[1:]:
            if is_empty_row(row):
                continue
            code = re.sub(r"\D", "", row[code_i] if code_i < len(row) else "")
            key = code or (row[name_i] if name_i < len(row) else str(len(out)))
            out[key] = row
        return out

    old_map, new_map = keyed(old_rows), keyed(new_rows)
    notes = []
    for key in list(new_map)[:12]:
        nrow = new_map[key]
        name = nrow[name_i] if name_i < len(nrow) else key
        orow = old_map.get(key)
        if orow is None:
            notes.append(f"추가 {name}({key})")
            continue
        fields = []
        for i, label in enumerate(labels):
            if i == 0 or label in {"구분"}:
                continue
            ov = orow[i] if i < len(orow) else ""
            nv = nrow[i] if i < len(nrow) else ""
            if norm_text(ov) != norm_text(nv):
                fields.append(f"{label} {_clip(ov, 18) or '(빈칸)'}→{_clip(nv, 18)}")
        if fields:
            notes.append(f"{name}: " + ", ".join(fields[:4]))
    for key, orow in old_map.items():
        if key not in new_map:
            name = orow[name_i] if name_i < len(orow) else key
            notes.append(f"삭제 {name}({key})")
    extra = len(notes) - 8
    if extra > 0:
        notes = notes[:8] + [f"외 {extra}건"]
    return notes


def _count_table_changes(old_rows: list[list[str]], new_rows: list[list[str]]) -> list[str]:
    def pairs(rows: list[list[str]]) -> dict[str, str]:
        out = {}
        for row in rows[1:]:
            if is_empty_row(row):
                continue
            label = " ".join(c for c in row[:-1] if c.strip()) or row[0]
            out[header_key(label) or label] = (row[-1] if row else "").strip()
        return out

    old_p, new_p = pairs(old_rows), pairs(new_rows)
    notes = []
    for key, val in new_p.items():
        prev = old_p.get(key)
        if prev is None:
            notes.append(f"{key} {val} 추가")
        elif prev != val:
            notes.append(f"{key} {prev}→{val}")
    for key, val in old_p.items():
        if key not in new_p and val:
            notes.append(f"{key} {val} 삭제")
    return notes[:10]


def _ridership_changes(old_rows: list[list[str]], new_rows: list[list[str]]) -> list[str]:
    notes = []
    if old_rows and new_rows and old_rows[0] != new_rows[0]:
        notes.append("연도 " + " / ".join(c for c in old_rows[0][1:] if c) + " → " + " / ".join(c for c in new_rows[0][1:] if c))
    limit = min(len(old_rows), len(new_rows))
    for i in range(1, limit):
        if old_rows[i] != new_rows[i]:
            label = old_rows[i][0] if old_rows[i] else "행"
            notes.append(f"{label} {_clip(' / '.join(old_rows[i][1:]), 28)} → {_clip(' / '.join(new_rows[i][1:]), 28)}")
    return notes[:6]


def _ops_table_changes(old_rows: list[list[str]], new_rows: list[list[str]]) -> list[str]:
    notes = []
    old_cells = [c for row in old_rows for c in row]
    new_cells = [c for row in new_rows for c in row]
    for old, new in zip(old_cells, new_cells):
        line = _line_change(old, new)
        if line:
            notes.append(line)
    if len(new_cells) != len(old_cells):
        for extra in new_cells[len(old_cells) :]:
            line = _line_change("", extra)
            if line:
                notes.append(line)
    return notes


def _photo_card_changes(old_rows: list[list[str]], new_rows: list[list[str]]) -> list[str]:
    old_t = old_rows[0][0] if old_rows and old_rows[0] else ""
    new_t = new_rows[0][0] if new_rows and new_rows[0] else ""
    notes = []
    if norm_text(old_t) != norm_text(new_t):
        notes.append(f"제목 {old_t or '(빈칸)'} → {new_t or '(빈칸)'}")
    old_flat, new_flat = [c for r in old_rows for c in r], [c for r in new_rows for c in r]
    for old, new in zip(old_flat, new_flat):
        if "담당" in old + new:
            continue
        if norm_text(old) != norm_text(new) and old not in {old_t} and new not in {new_t}:
            if header_key(old) in {"개소코드", "위치", "면적"} or header_key(new) in {"개소코드", "위치", "면적"}:
                continue
            if re.fullmatch(r"\d{5,7}", new or "") or re.fullmatch(r"\d{5,7}", old or ""):
                if old != new:
                    notes.append(f"개소코드 {old or '(빈칸)'} → {new or '(빈칸)'}")
            elif ("㎡" in old + new or re.search(r"\d", new or "")) and old != new and len(notes) < 4:
                if any(k in "".join(old_flat) for k in ("위치", "면적")):
                    notes.append(f"{_clip(old, 20) or '(빈칸)'} → {_clip(new, 20)}")
    return notes[:5]


def _generic_table_changes(old_rows: list[list[str]], new_rows: list[list[str]]) -> list[str]:
    notes = []
    for old_row, new_row in zip(old_rows, new_rows):
        for old, new in zip(old_row, new_row):
            line = _line_change(old, new)
            if line:
                notes.append(line)
    extra = len(notes) - 8
    if extra > 0:
        notes = notes[:8] + [f"외 {extra}건"]
    return notes


def _chart_changes(old: dict, new: dict) -> list[str]:
    notes = []
    if (old or {}).get("cats") != (new or {}).get("cats"):
        notes.append("축 " + " / ".join((old or {}).get("cats") or []) + " → " + " / ".join((new or {}).get("cats") or []))
    old_s, new_s = (old or {}).get("series") or {}, (new or {}).get("series") or {}
    for name, vals in new_s.items():
        prev = old_s.get(name)
        if prev != vals:
            notes.append(f"{name} {_clip(', '.join(str(int(v) if v is not None else 0) for v in (prev or [])), 28)} → {_clip(', '.join(str(int(v) if v is not None else 0) for v in vals), 28)}")
    return notes[:4]


def build_change_summaries(
    snap: dict,
    prs,
    matches,
    leftover_ppt: list[dict],
    leftover_xls: list,
    shops: list[StoreRecord],
    vendings: list[StoreRecord],
    photo: dict,
    count_text: dict[str, str],
    deletion_notes: list,
) -> list[dict]:
    after = capture_snapshot(prs)
    old_tables, new_tables = snap.get("tables", {}), after.get("tables", {})
    old_texts, new_texts = snap.get("texts", {}), after.get("texts", {})
    old_charts, new_charts = snap.get("charts", {}), after.get("charts", {})
    rows: list[dict] = []

    def add(sidx: int, part: str, content: str) -> None:
        content = (content or "").strip()
        if not content:
            return
        title = _slide_title(prs, sidx)
        label = f"슬라이드 {sidx + 1}" + (f". {title}" if title else "")
        rows.append({"슬라이드": label, "슬라이드번호": sidx + 1, "부분": part, "내용": content})

    all_keys = set(old_tables) | set(new_tables)
    for key in sorted(all_keys, key=lambda k: (_snap_slide(after, snap, key=k), str(k))):
        sidx = _snap_slide(after, snap, key=key)
        old_rows = old_tables.get(key) or []
        new_rows = new_tables.get(key) or []
        if old_rows == new_rows:
            continue
        part = _table_part_name(new_rows or old_rows)
        kind = classify_table(new_rows or old_rows)
        if kind in {"store_detail", "vending_detail"}:
            notes = _detail_row_changes(old_rows, new_rows)
        elif kind == "count_summary":
            notes = _count_table_changes(old_rows, new_rows)
        elif kind == "ridership":
            notes = _ridership_changes(old_rows, new_rows)
        elif kind == "ops_info":
            notes = _ops_table_changes(old_rows, new_rows)
        elif kind == "photo_card":
            notes = _photo_card_changes(old_rows, new_rows)
        else:
            notes = _ops_table_changes(old_rows, new_rows) or _generic_table_changes(old_rows, new_rows)
        if notes:
            add(sidx, part, " · ".join(notes))

    text_keys = set(old_texts) | set(new_texts)
    for key in sorted(text_keys, key=lambda k: (_snap_slide(after, snap, key=k), str(k))):
        sidx = _snap_slide(after, snap, key=key)
        old_p = old_texts.get(key) or []
        new_p = new_texts.get(key) or []
        if old_p == new_p:
            continue
        old_join, new_join = "\n".join(old_p), "\n".join(new_p)
        if (new_join or "").startswith("[삭제 메모]"):
            continue
        line = _line_change(old_join, new_join)
        if not line:
            if not old_join and new_join:
                line = f"문구 추가 {_clip(new_join)}"
            elif old_join and not new_join:
                line = f"문구 삭제 {_clip(old_join)}"
            else:
                line = f"{_clip(old_join, 40)} → {_clip(new_join, 40)}"
        part = "담당 FC/SV" if re.search(r"F\s*C", old_join + new_join, re.I) else "문구"
        if "계약기간" in old_join + new_join or "월평균매출" in old_join + new_join:
            part = "특이사항"
        elif CODE_RE.search(old_join + new_join):
            part = "위치도·매장 라벨"
        add(sidx, part, line)

    chart_keys = set(old_charts) | set(new_charts)
    for key in sorted(chart_keys, key=lambda k: (_snap_slide(after, snap, key=k), str(k))):
        notes = _chart_changes(old_charts.get(key) or {}, new_charts.get(key) or {})
        if notes:
            add(_snap_slide(after, snap, key=key), "승하차 현황 그래프", " · ".join(notes))

    shop_slides = sorted({f.slide_index for f in find_tables(prs) if f.kind == "store_detail"})
    vend_slides = sorted({f.slide_index for f in find_tables(prs) if f.kind == "vending_detail"})
    for store in leftover_xls:
        if store.is_shop and shop_slides:
            sidx = shop_slides[0]
        elif store.is_vending and vend_slides:
            sidx = vend_slides[0]
        else:
            sidx = shop_slides[0] if shop_slides else 0
        add(sidx, "추가", f"엑셀 매장 {store.name}({store.code})를 세부내역에 추가")
    for item in leftover_ppt:
        name = item.get("name") or ""
        if "[삭제 메모]" in name:
            continue
        add(int(item.get("slide") or 0), "삭제", f"PPT만 있던 {name}({item.get('code') or '-'}) 비움")
    for name in photo.get("photo_updated") or []:
        card_slides = sorted({f.slide_index for f in find_tables(prs) if f.kind == "photo_card"})
        add(card_slides[0] if card_slides else 0, "매장 사진 카드", f"{name} 엑셀 기준으로 문구 갱신")
    for code in photo.get("photo_cleared") or []:
        card_slides = sorted({f.slide_index for f in find_tables(prs) if f.kind == "photo_card"})
        add(card_slides[0] if card_slides else 0, "매장 사진 카드", f"{code} 엑셀에 없어 제목·코드 비움(사진은 유지)")
    for name in photo.get("photo_added") or []:
        card_slides = sorted({f.slide_index for f in find_tables(prs) if f.kind == "photo_card"})
        add(card_slides[-1] if card_slides else 0, "매장 사진 카드", f"{name} 신규 매장 사진 폼 추가")
    for note in deletion_notes:
        name = note.name or "매장"
        if "[삭제 메모]" in name:
            name = "제외 매장"
        add(note.slide_index, "삭제 메모", f"{name}({note.code or '-'}): {note.reason}")

    # 같은 슬라이드·부분 중복 집계 문구는 표 diff와 겹치면 표 쪽만 남김
    seen = set()
    unique = []
    for row in rows:
        key = (row["슬라이드번호"], row["부분"], row["내용"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(row)
    unique.sort(key=lambda r: (r["슬라이드번호"], r["부분"]))
    return unique


def _daily_avg(ridership, months: list[int], year: int | None = None) -> tuple[float | None, float | None]:
    if ridership is None:
        return None, None
    year = year or getattr(ridership, "year", 0) or date.today().year
    on_vals = []
    off_vals = []
    days = 0
    for m in months:
        key = f"{m}월"
        mdays = calendar.monthrange(year, m)[1]
        if key in ridership.onboard:
            val = ridership.onboard[key]
            if val:
                on_vals.append(val)
                days += mdays
        if key in ridership.offboard:
            val = ridership.offboard[key]
            if val:
                off_vals.append(val)
    if on_vals and days:
        off_sum = sum(off_vals) if off_vals else 0
        return sum(on_vals) / days, off_sum / max(days, 1)
    days = sum(calendar.monthrange(year, m)[1] for m in months) or 365
    if ridership.total_on:
        return ridership.total_on / days, (ridership.total_off or 0) / days
    return None, None


def _header_year(text: str) -> int | None:
    compact = header_key(text or "")
    if "상반기" in compact or "하반기" in compact:
        return None
    match = re.search(r"(19|20)\d{2}", text or "")
    return int(match.group(0)) if match else None


def _target_years(as_of: date) -> list[int]:
    return [as_of.year - 3, as_of.year - 2, as_of.year - 1, as_of.year]


def _year_axis_label(year: int, as_of: date) -> str:
    if year == as_of.year and as_of.month <= 6:
        return f"{year}년 상반기"
    return f"{year}년"


def _avg_for_year(rec: RidershipYear, as_of: date) -> tuple[int, int, int] | None:
    if rec.year < as_of.year:
        months = list(range(1, 13))
    elif rec.year == as_of.year:
        months = list(range(1, min(max(as_of.month, 1), 12) + 1))
    else:
        return None
    on_avg, off_avg = _daily_avg(rec, months, rec.year)
    if on_avg is None:
        return None
    on_i = int(round(on_avg))
    off_i = int(round(off_avg or 0))
    return on_i, off_i, on_i + off_i


def _parse_ridership_int(text: str) -> int | None:
    cleaned = re.sub(r"[^\d\-]", "", text or "")
    if not cleaned or cleaned == "-":
        return None
    try:
        return int(cleaned)
    except ValueError:
        return None


def _complete_year_tuple(on, off, total) -> tuple[int, int, int] | None:
    if total is None and on is not None and off is not None:
        total = on + off
    if on is None and total is not None and off is not None:
        on = total - off
    if off is None and total is not None and on is not None:
        off = total - on
    if on is None and off is None and total is None:
        return None
    on = int(on or 0)
    off = int(off or 0)
    return on, off, int(total if total is not None else on + off)


def _excel_window_avgs(ridership: RidershipSet, as_of: date) -> dict[int, tuple[int, int, int]]:
    years = _target_years(as_of)
    lo, hi = years[0], years[-1]
    out: dict[int, tuple[int, int, int]] = {}
    for year, rec in ridership.by_year.items():
        if year < lo or year > hi:
            continue
        avg = _avg_for_year(rec, as_of)
        if avg is not None:
            out[year] = avg
    if out:
        return out
    if len(ridership.by_year) != 1:
        return out
    rec = next(iter(ridership.by_year.values()))
    avg = _avg_for_year(rec, as_of)
    if avg is not None:
        out[as_of.year] = avg
    return out


def _ridership_row_slot(kind: str) -> int | None:
    if "소계" in kind or "합계" in kind:
        return 2
    if "승" in kind:
        return 0
    if "하" in kind:
        return 1
    return None


def _extract_table_year_values(table) -> dict[int, tuple[int, int, int]]:
    header = [get_cell_text(c) for c in table.rows[0].cells]
    year_cols = {_header_year(h): idx for idx, h in enumerate(header) if idx and _header_year(h)}
    slots: dict[int, list] = {year: [None, None, None] for year in year_cols}
    for i in range(1, len(table.rows)):
        row = table.rows[i]
        if not row.cells:
            continue
        slot = _ridership_row_slot(header_key(get_cell_text(row.cells[0])))
        if slot is None:
            continue
        for year, col in year_cols.items():
            if col >= len(row.cells):
                continue
            slots[year][slot] = _parse_ridership_int(get_cell_text(row.cells[col]))
    result = {}
    for year, triple in slots.items():
        done = _complete_year_tuple(*triple)
        if done is not None:
            result[year] = done
    return result


def _extract_chart_year_values(chart) -> dict[int, tuple[int, int, int]]:
    plots = list(chart.plots)
    if not plots:
        return {}
    cats = [str(c) for c in plots[0].categories]
    series_map = {s.name: [v for v in s.values] for s in plots[0].series}
    if not any(_header_year(c) for c in cats):
        return {}
    result = {}
    for idx, cat in enumerate(cats):
        year = _header_year(cat)
        if year is None:
            continue
        on = off = None
        for name, vals in series_map.items():
            val = vals[idx] if idx < len(vals) else None
            if val is None:
                continue
            if "승" in (name or ""):
                on = int(round(val))
            elif "하" in (name or ""):
                off = int(round(val))
        done = _complete_year_tuple(on, off, None)
        if done is not None:
            result[year] = done
    return result


def sales_period_tag(as_of: date) -> str:
    yy = str(as_of.year)[2:]
    if as_of.month <= 6:
        return f"`{yy}년 상반기"
    return f"`{yy}년"


def update_cover_and_period_labels(prs, as_of: date | None) -> int:
    """표지 월과 매출 기간 표기(`26년 상반기 등)를 기준일에 맞춘다."""
    if as_of is None:
        return 0
    tag = sales_period_tag(as_of)
    cover_re = re.compile(r"([`''′]\d{2})\.\s*\d{1,2}\.")
    period_re = re.compile(r"[`''′]\d{2}년\s*(상반기|하반기)?")
    changed = 0

    def patch_para(para, allow_cover: bool, allow_period: bool) -> None:
        nonlocal changed
        text = para.text or ""
        if "구내영업료" in text:
            return
        new_text = text
        if allow_cover:
            new_text = cover_re.sub(rf"\g<1>. {as_of.month}.", new_text, count=1)
        if allow_period and period_re.search(new_text):
            new_text = period_re.sub(tag, new_text)
        if new_text != text:
            segs = labeled_review_segments(new_text, text)
            if segs:
                write_paragraph_segments(para, segs)
            else:
                set_run_text_preserve(para, new_text)
            changed += 1

    for sidx, slide in enumerate(prs.slides):
        for shape in iter_shapes(slide):
            try:
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            blob = get_cell_text(cell)
                            allow_period = "월평균매출액" in blob
                            if not allow_period and not (sidx == 0 and cover_re.search(blob)):
                                continue
                            for para in cell.text_frame.paragraphs:
                                patch_para(para, allow_cover=False, allow_period=allow_period)
            except Exception:
                pass
            try:
                if shape.has_text_frame:
                    blob = shape.text_frame.text or ""
                    allow_cover = sidx == 0 and cover_re.search(blob)
                    allow_period = "월평균매출액" in blob
                    if not allow_cover and not allow_period:
                        continue
                    for para in shape.text_frame.paragraphs:
                        patch_para(para, allow_cover=bool(allow_cover), allow_period=allow_period)
            except Exception:
                pass
    return changed


def _roll_four_year_values(
    target_years: list[int],
    excel_vals: dict[int, tuple[int, int, int]],
    ppt_vals: dict[int, tuple[int, int, int]],
) -> dict[int, tuple[int, int, int]]:
    """엑셀(기준일 최근 4개년) 우선, 같은 연도는 PPT 유지, 빈 칸은 가장 오래된 PPT 연도를 버리고 채운다."""
    filled: dict[int, tuple[int, int, int]] = {}
    used_ppt: set[int] = set()
    for year in target_years:
        if year in excel_vals:
            filled[year] = excel_vals[year]
        elif year in ppt_vals:
            filled[year] = ppt_vals[year]
            used_ppt.add(year)
    missing = [year for year in target_years if year not in filled]
    leftover = sorted(year for year in ppt_vals if year not in used_ppt and year not in excel_vals)
    if missing and leftover:
        take = leftover[-len(missing) :]
        for target, source in zip(missing, take):
            filled[target] = ppt_vals[source]
    return filled


def update_ridership(prs, ridership, as_of: date | None = None) -> None:
    if ridership is None:
        return
    as_of = as_of or date.today()
    if isinstance(ridership, RidershipYear):
        year = ridership.year or as_of.year
        ridership = RidershipSet(station=ridership.station, by_year={year: ridership})
    years = _target_years(as_of)
    labels = [_year_axis_label(y, as_of) for y in years]
    excel_vals = _excel_window_avgs(ridership, as_of)

    ppt_vals: dict[int, tuple[int, int, int]] = {}
    for found in find_tables(prs):
        if found.kind != "ridership":
            continue
        for year, vals in _extract_table_year_values(found.shape.table).items():
            ppt_vals.setdefault(year, vals)
    for slide in prs.slides:
        for shape in iter_shapes(slide):
            try:
                if shape.has_chart:
                    for year, vals in _extract_chart_year_values(shape.chart).items():
                        ppt_vals.setdefault(year, vals)
            except Exception:
                continue

    filled = _roll_four_year_values(years, excel_vals, ppt_vals)

    for found in find_tables(prs):
        if found.kind != "ridership":
            continue
        table = found.shape.table
        header = [get_cell_text(c) for c in table.rows[0].cells]
        data_cols = [idx for idx, h in enumerate(header) if idx > 0]
        if not data_cols:
            continue
        use_cols = data_cols[-4:] if len(data_cols) >= 4 else data_cols
        assign_years = years[-len(use_cols) :]
        assign_labels = labels[-len(use_cols) :]
        for year, lab, col in zip(assign_years, assign_labels, use_cols):
            set_cell_text_preserve(
                table.rows[0].cells[col],
                lab,
                review_red=norm_text(header[col]) != norm_text(lab),
            )
        for i in range(1, len(table.rows)):
            row = table.rows[i]
            if len(row.cells) == 0:
                continue
            slot = _ridership_row_slot(header_key(get_cell_text(row.cells[0])))
            if slot is None:
                continue
            for year, lab, col in zip(assign_years, assign_labels, use_cols):
                if col >= len(row.cells):
                    continue
                if year in filled:
                    set_cell_text_preserve(row.cells[col], f"{filled[year][slot]:,}")
                else:
                    set_cell_text_preserve(row.cells[col], "", review_red=False)

    for slide in prs.slides:
        for shape in iter_shapes(slide):
            try:
                if not shape.has_chart:
                    continue
                chart = shape.chart
                plots = list(chart.plots)
                if not plots:
                    continue
                cats = [str(c) for c in plots[0].categories]
                if not any(_header_year(c) for c in cats):
                    continue
                onboard = [filled[year][0] if year in filled else 0 for year in years]
                offboard = [filled[year][1] if year in filled else 0 for year in years]
                patch_category_chart(chart, labels, {"승": onboard, "하": offboard})
            except Exception:
                continue


REVIEW_SHAPE_NAME = "korail_review_note"
REVIEW_NOTE_PREFIX = "[삭제 메모]"
NOTE_RED = RGBColor(0xC0, 0x00, 0x00)


@dataclass
class DeletionNote:
    slide_index: int
    code: str
    name: str
    reason: str


def explain_removal(
    code: str,
    name: str,
    active: list[StoreRecord],
    inactive: list[StoreRecord],
    as_of: date,
) -> str:
    for store in active:
        if name and score_names(name, store.name) >= 90:
            return f"동일 매장은 현재 개소코드 {store.code}만 유지"
    for store in inactive:
        if code and store.code == code:
            return store.skip_reason or "현재 운영 계약 구간 아님"
    best = None
    best_score = 0.0
    for store in inactive:
        sc = score_names(name or "", store.name)
        if sc > best_score:
            best_score = sc
            best = store
    if best and best_score >= 85:
        return best.skip_reason or "현재 운영 계약 구간 아님"
    return f"기준일 {format_iso_date(as_of)} 현재 영업중(운영단계)이 아님"


def _current_slide_for_orig_code(prs, snap: dict | None, code: str) -> int | None:
    """슬라이드가 끼워진 뒤에도, 원본 개소코드가 있던 카드/표의 현재 번호를 찾는다."""
    if not code:
        return None
    table_snap = (snap or {}).get("tables", {})
    table_hit = None
    for found in find_tables(prs):
        orig = table_snap.get(_shape_key(found.slide_index, found.shape))
        if not orig:
            continue
        if found.kind == "photo_card":
            if _photo_card_code(orig) == code:
                return found.slide_index
            continue
        if table_hit is not None or found.kind not in {"store_detail", "vending_detail"}:
            continue
        cols = find_header_columns(orig[0])
        if "code" not in cols:
            continue
        for row in orig[1:]:
            row_code = re.sub(r"\D", "", row[cols["code"]]) if cols["code"] < len(row) else ""
            if row_code == code:
                table_hit = found.slide_index
                break
    return table_hit


def collect_deletion_notes(
    prs,
    snap: dict,
    leftover_ppt: list[dict],
    mapping: dict[str, str],
    active: list[StoreRecord],
    inactive: list[StoreRecord],
    as_of: date,
    map_cleared: list[dict],
) -> list[DeletionNote]:
    remaining = {s.code for s in active}
    notes: list[DeletionNote] = []
    seen: set[tuple[int, str, str]] = set()

    def add(slide_index: int, code: str, name: str, reason: str) -> None:
        key = (slide_index, code or "", name or "")
        if key in seen:
            return
        if not code and not name:
            return
        seen.add(key)
        notes.append(DeletionNote(slide_index, code or "", name or "", reason))

    for item in leftover_ppt:
        if "[삭제 메모]" in (item.get("name") or ""):
            continue
        code = item.get("code") or ""
        if code and code in remaining:
            continue
        slide = _current_slide_for_orig_code(prs, snap, code)
        if slide is None:
            slide = int(item.get("slide") or 0)
        add(
            slide,
            code,
            item.get("name") or "",
            explain_removal(code, item.get("name") or "", active, inactive, as_of),
        )

    table_snap = (snap or {}).get("tables", {})
    for found in find_tables(prs):
        if found.kind not in {"store_detail", "vending_detail", "photo_card"}:
            continue
        orig = table_snap.get(_shape_key(found.slide_index, found.shape))
        if not orig:
            continue
        if found.kind == "photo_card":
            code = _photo_card_code(orig)
            name = orig[0][0] if orig and orig[0] else ""
            mapped = mapping.get(code, code) if code else ""
            if code and mapped not in remaining and code not in remaining:
                add(
                    found.slide_index,
                    code,
                    name,
                    explain_removal(code, name, active, inactive, as_of),
                )
            continue
        cols = find_header_columns(orig[0])
        if "code" not in cols:
            continue
        for row in orig[1:]:
            if is_empty_row(row):
                continue
            code = re.sub(r"\D", "", row[cols["code"]]) if cols["code"] < len(row) else ""
            name = row[cols["name"]] if "name" in cols and cols["name"] < len(row) else ""
            mapped = mapping.get(code, code) if code else ""
            if mapped in remaining or code in remaining:
                continue
            add(
                found.slide_index,
                code,
                name,
                explain_removal(code, name, active, inactive, as_of),
            )

    leftover_by_code = {
        (item.get("code") or ""): (item.get("name") or "")
        for item in leftover_ppt
        if item.get("code")
    }
    for item in map_cleared:
        code = (item.get("code") or "").split(",")[0]
        name = leftover_by_code.get(code) or item.get("name") or ""
        if name == code:
            name = leftover_by_code.get(code, "")
        add(
            int(item.get("slide") or 0),
            code,
            name,
            explain_removal(code, name, active, inactive, as_of),
        )
    return notes


def _remove_review_shapes(slide) -> None:
    sp_tree = slide.shapes._spTree
    victims = []
    for shape in list(slide.shapes):
        name = getattr(shape, "name", "") or ""
        text = ""
        try:
            if shape.has_text_frame:
                text = shape.text_frame.text or ""
        except Exception:
            text = ""
        if name == REVIEW_SHAPE_NAME or text.startswith(REVIEW_NOTE_PREFIX):
            victims.append(shape._element)
    for el in victims:
        sp_tree.remove(el)


def _write_speaker_notes(slide, body: str) -> None:
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    existing = (tf.text or "").rstrip()
    existing = re.sub(rf"\n*{re.escape(REVIEW_NOTE_PREFIX)}[\s\S]*$", "", existing).rstrip()
    tf.text = f"{existing}\n\n{body}".strip() if existing else body


def apply_deletion_notes(prs, notes: list[DeletionNote], as_of: date) -> int:
    by_slide: dict[int, list[DeletionNote]] = defaultdict(list)
    for note in notes:
        by_slide[note.slide_index].append(note)
    marked = 0
    for sidx, slide in enumerate(prs.slides):
        _remove_review_shapes(slide)
        items = by_slide.get(sidx) or []
        if not items:
            continue
        lines = [
            f"{item.name or '매장'}({item.code or '-'}): {item.reason}" if item.code or item.name else item.reason
            for item in items
        ]
        header = f"{REVIEW_NOTE_PREFIX} 기준일 {format_iso_date(as_of)}"
        notes_body = header + "\n" + "\n".join(f"- {line}" for line in lines)
        _write_speaker_notes(slide, notes_body)

        shown = lines[:3]
        extra = len(lines) - len(shown)
        footer = header + " · " + " / ".join(shown)
        if extra > 0:
            footer += f" 외 {extra}건 (노트 참고)"
        if len(footer) > 180:
            footer = footer[:177] + "..."

        width = int(prs.slide_width)
        height = int(prs.slide_height)
        box_h = int(0.42 * 914400)
        margin = int(0.12 * 914400)
        shape = slide.shapes.add_textbox(Emu(margin), Emu(height - box_h - int(0.04 * 914400)), Emu(width - margin * 2), Emu(box_h))
        shape.name = REVIEW_SHAPE_NAME
        tf = shape.text_frame
        tf.word_wrap = True
        para = tf.paragraphs[0]
        set_run_text_preserve(para, footer)
        for run in para.runs:
            apply_run_font(run, DATA_FONT, 8)
            try:
                run.font.color.rgb = NOTE_RED
            except Exception:
                pass
        if not para.runs:
            run = para.add_run()
            run.text = footer
            apply_run_font(run, DATA_FONT, 8)
            try:
                run.font.color.rgb = NOTE_RED
            except Exception:
                pass
        marked += 1
    return marked
