from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path

from pptx import Presentation

from engine.excel_loader import ExcelBundle, StoreRecord, format_iso_date, load_excel_bundle
from engine.pptx_io import sanitize_chart_workbooks
from engine.updater import (
    apply_code_mapping,
    apply_deletion_notes,
    build_change_summaries,
    capture_snapshot,
    clear_unmatched_map_labels,
    collect_deletion_notes,
    detect_station,
    match_presentation,
    normalize_map_labels,
    update_count_summaries,
    update_detail_tables,
    update_feature_boxes,
    update_ops_counts,
    update_photo_cards,
    add_photo_forms_for_new_shops,
    update_ridership,
    update_staff_line,
    update_cover_and_period_labels,
)


@dataclass
class PipelineReport:
    station: str
    excel_store_count: int
    matches: list[dict] = field(default_factory=list)
    unmatched_ppt: list[dict] = field(default_factory=list)
    unmatched_excel: list[dict] = field(default_factory=list)
    code_replacements: int = 0
    shops_written: int = 0
    vendings_written: int = 0
    photo_updated: list[str] = field(default_factory=list)
    photo_cleared: list[str] = field(default_factory=list)
    photo_added: list[str] = field(default_factory=list)
    map_labels_cleared: list[str] = field(default_factory=list)
    output_path: str = ""
    source_name: str = ""
    error: str = ""
    as_of: str = ""
    inactive_count: int = 0
    notes_slides: int = 0
    changes: list[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)


def output_pptx_name(station: str, as_of: date | None) -> str:
    raw = re.sub(r'[<>:"/\\|?*]+', "_", (station or "역").strip()) or "역"
    raw = raw.strip(" .")
    day = as_of or date.today()
    return f"{raw}_{day.year}.{day.month}.{day.day}.pptx"


def _unique_path(dest: Path) -> Path:
    if not dest.exists():
        return dest
    stem, suffix = dest.stem, dest.suffix
    i = 2
    while True:
        candidate = dest.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def split_stores(stores: list[StoreRecord]) -> tuple[list[StoreRecord], list[StoreRecord]]:
    shops = [s for s in stores if s.is_shop]
    vendings = [s for s in stores if s.is_vending]
    shops.sort(key=lambda s: (s.biz_type, s.name))
    vendings.sort(key=lambda s: s.name)
    return shops, vendings


def run_pipeline(
    pptx_path: str | Path,
    sales_path: str | Path,
    master_path: str | Path,
    ridership_path: str | Path | None,
    output_path: str | Path,
    station: str | None = None,
    as_of: date | None = None,
) -> PipelineReport:
    prs = Presentation(str(pptx_path))
    station = station or detect_station(prs)
    bundle: ExcelBundle = load_excel_bundle(
        sales_path, master_path, ridership_path, station, as_of=as_of
    )

    matches, mapping, leftover_ppt, leftover_xls = match_presentation(prs, bundle.stores)
    snap = capture_snapshot(prs)
    replacements = apply_code_mapping(prs, mapping)

    shops, vendings = split_stores(bundle.stores)
    update_detail_tables(prs, shops, vendings, snap, mapping)

    by_code = {s.code: s for s in bundle.stores}
    photo = update_photo_cards(prs, by_code, snap)
    photo["photo_added"] = add_photo_forms_for_new_shops(prs, leftover_xls)
    update_feature_boxes(prs, by_code)
    map_cleared = clear_unmatched_map_labels(prs, set(by_code))
    normalize_map_labels(prs, shops, snap, mapping, map_cleared)
    update_ops_counts(prs, shops, vendings)
    update_staff_line(prs, shops, vendings)
    count_text = update_count_summaries(prs, shops, vendings)
    update_ridership(prs, bundle.ridership, bundle.as_of)
    update_cover_and_period_labels(prs, bundle.as_of)
    deletion_notes = collect_deletion_notes(
        prs,
        snap,
        leftover_ppt,
        mapping,
        bundle.stores,
        bundle.inactive,
        bundle.as_of or date.today(),
        map_cleared,
    )
    notes_slides = apply_deletion_notes(prs, deletion_notes, bundle.as_of or date.today())
    changes = build_change_summaries(
        snap,
        prs,
        matches,
        leftover_ppt,
        leftover_xls,
        shops,
        vendings,
        photo,
        count_text,
        deletion_notes,
    )
    sanitize_chart_workbooks(prs)
    as_of_value = bundle.as_of or date.today()
    try:
        prs.core_properties.title = f"{station}_{format_iso_date(as_of_value)}"
    except Exception:
        pass

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))

    return PipelineReport(
        station=station,
        excel_store_count=len(bundle.stores),
        matches=[
            {
                "old_code": m.old_code,
                "new_code": m.new_code,
                "ppt_name": m.ppt_name,
                "excel_name": m.excel_name,
                "score": m.score,
                "confidence": m.confidence,
                "aliases": m.aliases,
            }
            for m in matches
        ],
        unmatched_ppt=[{"code": p.get("code"), "name": p.get("name"), "brand": p.get("brand")} for p in leftover_ppt],
        unmatched_excel=[{"code": s.code, "name": s.name, "biz_type": s.biz_type} for s in leftover_xls],
        code_replacements=replacements,
        shops_written=len(shops),
        vendings_written=len(vendings),
        photo_updated=photo.get("photo_updated", []),
        photo_cleared=photo.get("photo_cleared", []),
        photo_added=photo.get("photo_added", []),
        map_labels_cleared=[item.get("code", "") for item in map_cleared],
        output_path=str(out),
        source_name=Path(pptx_path).name,
        as_of=bundle.as_of.isoformat() if bundle.as_of else "",
        inactive_count=len(bundle.inactive),
        notes_slides=notes_slides,
        changes=changes,
    )


def run_batch(
    pptx_paths: list[Path],
    sales_path: str | Path,
    master_path: str | Path,
    ridership_path: str | Path | None,
    output_dir: str | Path,
    station: str | None = None,
    as_of: date | None = None,
) -> list[PipelineReport]:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    reports: list[PipelineReport] = []
    for pptx_path in pptx_paths:
        out_path = output_dir / f"_work_{pptx_path.stem}.pptx"
        try:
            report = run_pipeline(
                pptx_path=pptx_path,
                sales_path=sales_path,
                master_path=master_path,
                ridership_path=ridership_path,
                output_path=out_path,
                station=station,
                as_of=as_of,
            )
            as_of_value = date.fromisoformat(report.as_of) if report.as_of else as_of
            final = _unique_path(output_dir / output_pptx_name(report.station, as_of_value))
            Path(report.output_path).replace(final)
            report.output_path = str(final)
            reports.append(report)
        except Exception as exc:
            reports.append(
                PipelineReport(
                    station=station or "",
                    excel_store_count=0,
                    source_name=pptx_path.name,
                    error=str(exc),
                )
            )
    return reports
