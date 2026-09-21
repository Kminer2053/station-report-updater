from __future__ import annotations

import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher

from engine.excel_loader import StoreRecord, _to_str


GENERIC_NAMES = {
    "편의점",
    "전문점",
    "자판기",
    "직영점",
    "매장",
    "위치도",
    "개소코드",
    "위치",
    "면적",
    "매장사진",
}


def is_usable_ppt_name(name: str, brand: str = "") -> bool:
    blob = re.sub(r"\s+", "", f"{name}{brand}")
    blob = re.sub(r"\d+", "", blob)
    blob = re.sub(r"[NO.]", "", blob, flags=re.I)
    if len(blob) < 2:
        return False
    if blob in GENERIC_NAMES:
        return False
    return True


BRAND_HINTS = [
    "다이소",
    "더베이크",
    "진김밥",
    "동산화원",
    "명품마루",
    "행복한광목",
    "스토리웨이",
    "셀프스토리웨이",
    "고향뜨락",
]

OPPOSITE_PAIRS = [("남", "북"), ("상행", "하행"), ("상홈남", "상홈북"), ("상홈중", "상홈북")]


def normalize_name(text: str) -> str:
    s = _to_str(text)
    s = s.replace(" ", "").replace("\n", "").replace("\t", "")
    s = s.replace("．", ".").replace("·", "")
    s = re.sub(r"[()\[\]{}_]", "", s)
    s = s.replace("자판기", "")
    return s.lower()


def tokens(text: str) -> set[str]:
    parts = re.split(r"[\s()/_-]+", _to_str(text))
    return {p for p in parts if len(p) >= 2}


def _has_opposite_location(a: str, b: str) -> bool:
    na, nb = normalize_name(a), normalize_name(b)
    for left, right in OPPOSITE_PAIRS:
        if (left in na and right in nb) or (right in na and left in nb):
            # "남" in both 주안 AND 남편의점 vs 북 - the pair 남/북
            return True
    return False


def score_names(ppt_name: str, excel_name: str, ppt_brand: str = "") -> float:
    query = f"{ppt_name} {ppt_brand}".strip()
    a, b = normalize_name(query), normalize_name(excel_name)
    if not a or not b:
        return 0.0

    if a == b:
        base = 100.0
    elif a in b or b in a:
        base = 92.0 + min(len(a), len(b)) / max(len(a), len(b)) * 8
    else:
        base = SequenceMatcher(None, a, b).ratio() * 100
        ta, tb = tokens(query), tokens(excel_name)
        if ta and tb:
            jacc = len(ta & tb) / len(ta | tb) * 100
            base = max(base, jacc)

    for brand in BRAND_HINTS:
        bn = brand.replace(" ", "")
        in_ppt = bn in normalize_name(ppt_name) or bn in normalize_name(ppt_brand)
        in_xls = bn in normalize_name(excel_name)
        if in_ppt and in_xls:
            base = max(base, 94.0)

    if "스마트" in a and "스마트" in b:
        base = max(base, 93.0)

    if _has_opposite_location(query, excel_name):
        base *= 0.35

    return base


@dataclass
class MatchResult:
    old_code: str
    new_code: str
    ppt_name: str
    excel_name: str
    score: float
    confidence: str
    aliases: list[str] = field(default_factory=list)


def greedy_match(
    ppt_rows: list[dict],
    stores: list[StoreRecord],
    threshold: float = 70.0,
) -> tuple[list[MatchResult], list[dict], list[StoreRecord]]:
    """ppt_rows: {code, name, brand}"""
    used_ppt: set[int] = set()
    used_xls: set[int] = set()
    used_old_codes: set[str] = set()
    matches: list[MatchResult] = []

    xls_by_code: dict[str, int] = {}
    for j, store in enumerate(stores):
        if store.code and store.code not in xls_by_code:
            xls_by_code[store.code] = j
    for i, ppt in enumerate(ppt_rows):
        old_code = _to_str(ppt.get("code"))
        j = xls_by_code.get(old_code)
        if not old_code or j is None or j in used_xls:
            continue
        store = stores[j]
        matches.append(
            MatchResult(
                old_code=old_code,
                new_code=store.code,
                ppt_name=_to_str(ppt.get("name")),
                excel_name=store.name,
                score=100.0,
                confidence="high",
            )
        )
        used_ppt.add(i)
        used_xls.add(j)
        used_old_codes.add(old_code)

    pairs: list[tuple[float, int, int]] = []
    for i, ppt in enumerate(ppt_rows):
        if not is_usable_ppt_name(ppt.get("name") or "", ppt.get("brand") or ""):
            continue
        for j, store in enumerate(stores):
            sc = score_names(ppt.get("name") or "", store.name, ppt.get("brand") or "")
            if sc >= threshold:
                pairs.append((sc, i, j))
    pairs.sort(key=lambda x: x[0], reverse=True)

    for sc, i, j in pairs:
        if i in used_ppt or j in used_xls:
            continue
        ppt = ppt_rows[i]
        if not is_usable_ppt_name(ppt.get("name") or "", ppt.get("brand") or ""):
            continue
        old_code = _to_str(ppt.get("code"))
        if old_code and old_code in used_old_codes:
            continue
        store = stores[j]
        matches.append(
            MatchResult(
                old_code=_to_str(ppt.get("code")),
                new_code=store.code,
                ppt_name=_to_str(ppt.get("name")),
                excel_name=store.name,
                score=round(sc, 1),
                confidence="high" if sc >= 90 else "medium",
            )
        )
        used_ppt.add(i)
        used_xls.add(j)
        if old_code:
            used_old_codes.add(old_code)

    # last-resort: leftover 편의점 1:1
    leftover_ppt = [ppt_rows[i] for i in range(len(ppt_rows)) if i not in used_ppt]
    leftover_ppt = [p for p in leftover_ppt if is_usable_ppt_name(p.get("name") or "", p.get("brand") or "")]
    leftover_xls = [stores[j] for j in range(len(stores)) if j not in used_xls]
    ppt_conv = [p for p in leftover_ppt if "편의점" in (p.get("name") or "") or p.get("brand") == "스토리웨이"]
    xls_conv = [s for s in leftover_xls if s.biz_type == "편의점"]
    if xls_conv and ppt_conv:
        # 남은 편의점은 점수 높은 쌍부터 1:1
        conv_pairs = []
        for p in ppt_conv:
            for s in xls_conv:
                conv_pairs.append((score_names(p.get("name") or "", s.name, p.get("brand") or ""), p, s))
        conv_pairs.sort(key=lambda x: x[0], reverse=True)
        used_p = set()
        used_s = set()
        for sc, p, s in conv_pairs:
            if sc < 45:
                continue
            pk, sk = id(p), id(s)
            if pk in used_p or sk in used_s:
                continue
            matches.append(
                MatchResult(
                    old_code=_to_str(p.get("code")),
                    new_code=s.code,
                    ppt_name=_to_str(p.get("name")),
                    excel_name=s.name,
                    score=round(sc, 1),
                    confidence="low",
                )
            )
            used_p.add(pk)
            used_s.add(sk)
            leftover_ppt = [x for x in leftover_ppt if x is not p]
            leftover_xls = [x for x in leftover_xls if x is not s]

    return matches, leftover_ppt, leftover_xls


def build_code_map(matches: list[MatchResult]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for m in matches:
        if m.old_code:
            mapping[m.old_code] = m.new_code
        for alias in m.aliases:
            if alias:
                mapping[alias] = m.new_code
    return mapping
