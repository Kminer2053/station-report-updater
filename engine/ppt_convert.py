from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
PPTX_MAGIC = b"PK"
PP_SAVE_AS_OPENXML = 24


class PptConvertError(RuntimeError):
    pass


def is_pptx_bytes(data: bytes) -> bool:
    return data[:2] == PPTX_MAGIC


def is_legacy_ppt_bytes(data: bytes) -> bool:
    return data[:8] == OLE_MAGIC


def convert_ppt_to_pptx(src: Path, dest: Path | None = None) -> Path:
    """구버전 .ppt를 .pptx로 변환. PowerPoint 또는 LibreOffice를 사용한다."""
    src = Path(src).resolve()
    dest = Path(dest).resolve() if dest is not None else src.with_suffix(".pptx")
    dest.parent.mkdir(parents=True, exist_ok=True)
    for converter in (_try_powerpoint, _try_libreoffice):
        try:
            converter(src, dest)
            if dest.exists() and dest.stat().st_size > 0:
                return dest
        except Exception:
            pass
        if dest.exists() and dest.stat().st_size == 0:
            dest.unlink()
    raise PptConvertError(
        f"{src.name}은 구버전(.ppt) 파일이라 이 PC에서 바로 열 수 없습니다. "
        "PowerPoint에서 ‘pptx로 저장’한 뒤 올려 주세요. "
        "PowerPoint나 LibreOffice가 설치된 PC에서는 .ppt도 자동으로 변환됩니다."
    )


def _try_powerpoint(src: Path, dest: Path) -> None:
    ps1 = dest.parent / f"_convert_{os.getpid()}.ps1"
    src_lit = str(src).replace("'", "''")
    dest_lit = str(dest).replace("'", "''")
    ps1.write_text(
        "\n".join(
            [
                "$ErrorActionPreference = 'Stop'",
                "$ppt = New-Object -ComObject PowerPoint.Application",
                "try {",
                f"  $src = '{src_lit}'",
                f"  $dest = '{dest_lit}'",
                "  if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }",
                "  $pres = $ppt.Presentations.Open($src, $true, $false, $false)",
                f"  $pres.SaveAs($dest, {PP_SAVE_AS_OPENXML})",
                "  $pres.Close()",
                "} finally {",
                "  $ppt.Quit()",
                "  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null",
                "}",
            ]
        ),
        encoding="utf-8",
    )
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(ps1),
            ],
            capture_output=True,
            text=True,
            timeout=180,
        )
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "").strip() or f"exit {result.returncode}"
            raise RuntimeError(err[:400])
    finally:
        try:
            ps1.unlink(missing_ok=True)
        except OSError:
            pass


def _find_soffice() -> Path | None:
    env = os.environ.get("LIBREOFFICE_PATH") or os.environ.get("SOFFICE_PATH")
    candidates = [
        Path(env) if env else None,
        Path(shutil.which("soffice") or ""),
        Path(shutil.which("soffice.exe") or ""),
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        Path("/usr/bin/soffice"),
        Path("/usr/bin/libreoffice"),
    ]
    for path in candidates:
        if path and path.is_file():
            return path
    return None


def _try_libreoffice(src: Path, dest: Path) -> None:
    soffice = _find_soffice()
    if soffice is None:
        raise RuntimeError("LibreOffice 없음")
    outdir = dest.parent
    result = subprocess.run(
        [
            str(soffice),
            "--headless",
            "--norestore",
            "--convert-to",
            "pptx",
            "--outdir",
            str(outdir),
            str(src),
        ],
        capture_output=True,
        text=True,
        timeout=180,
    )
    if result.returncode != 0:
        err = (result.stderr or result.stdout or "").strip() or f"exit {result.returncode}"
        raise RuntimeError(err[:400])
    produced = outdir / (src.stem + ".pptx")
    if produced.resolve() != dest.resolve() and produced.exists():
        if dest.exists():
            dest.unlink()
        produced.replace(dest)
