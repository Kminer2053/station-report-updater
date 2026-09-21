from __future__ import annotations

import re
import zipfile
from pathlib import Path

from engine.ppt_convert import convert_ppt_to_pptx, is_pptx_bytes

SAFE_SUFFIX = {".pptx", ".ppt"}
SKIP_DIR_PARTS = {"__macosx", ".ds_store"}


def _decode_zip_name(info: zipfile.ZipInfo) -> str:
    name = info.filename
    if info.flag_bits & 0x800:
        return name
    try:
        return name.encode("cp437").decode("euc-kr")
    except UnicodeError:
        try:
            return name.encode("cp437").decode("cp949")
        except UnicodeError:
            return name


def _is_safe_member(name: str) -> bool:
    path = Path(name.replace("\\", "/"))
    parts = [p.lower() for p in path.parts]
    if any(p in SKIP_DIR_PARTS or p.startswith(".") for p in parts):
        return False
    if path.is_absolute() or ".." in path.parts:
        return False
    return True


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


def extract_pptx_from_zip(zip_path: Path, dest_dir: Path, depth: int = 0) -> list[Path]:
    found: list[Path] = []
    dest_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = _decode_zip_name(info)
            if not _is_safe_member(name):
                continue
            suffix = Path(name).suffix.lower()
            raw = zf.read(info)
            if suffix in {".pptx", ".ppt"}:
                found.append(_write_presentation(dest_dir / Path(name).name, raw))
            elif suffix == ".zip" and depth < 1:
                nested = dest_dir / f"_nested_{Path(name).stem}.zip"
                nested.write_bytes(raw)
                found.extend(extract_pptx_from_zip(nested, dest_dir / Path(name).stem, depth + 1))
    return found


def _write_presentation(path: Path, data: bytes) -> Path:
    dest_dir = path.parent
    dest_dir.mkdir(parents=True, exist_ok=True)
    stem = path.stem
    if is_pptx_bytes(data):
        out = _unique_path(dest_dir / f"{stem}.pptx")
        out.write_bytes(data)
        return out
    src = _unique_path(dest_dir / f"{stem}.ppt")
    src.write_bytes(data)
    return convert_ppt_to_pptx(src, _unique_path(dest_dir / f"{stem}.pptx"))


def collect_pptx_uploads(
    uploads: list[tuple[str, bytes]],
    dest_dir: Path,
) -> list[Path]:
    """업로드된 ppt/pptx/zip 바이트에서 처리할 pptx 경로 목록을 만든다."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    found: list[Path] = []
    for filename, data in uploads:
        suffix = Path(filename).suffix.lower()
        if suffix in {".pptx", ".ppt"}:
            found.append(_write_presentation(dest_dir / Path(filename).name, data))
        elif suffix == ".zip":
            zip_path = dest_dir / f"_upload_{re.sub(r'[^0-9A-Za-z._-]+', '_', Path(filename).stem)}.zip"
            zip_path.write_bytes(data)
            found.extend(extract_pptx_from_zip(zip_path, dest_dir / zip_path.stem))
    return found


def pack_outputs_zip(pairs: list[tuple[str, bytes]], dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        used: set[str] = set()
        for name, data in pairs:
            arc = Path(name).name
            if not arc.lower().endswith(".pptx"):
                arc = f"{arc}.pptx"
            base = arc
            i = 2
            while arc.lower() in used:
                arc = f"{Path(base).stem}_{i}.pptx"
                i += 1
            used.add(arc.lower())
            zf.writestr(arc, data)
    return dest
