from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd


def _to_str(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def _to_float(value) -> float | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text or text.lower() in {"nan", "none"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalize_station(name: str) -> str:
    text = _to_str(name).replace(" ", "")
    if text.endswith("역"):
        text = text[:-1]
    return text


def format_iso_date(value: date | datetime | None, spaced: bool = False) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        value = value.date()
    if spaced:
        return f"{value.year}. {value.month}. {value.day}."
    return f"{value.year}.{value.month}.{value.day}."


def parse_date(value) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.date()


def format_ppt_date(value) -> str:
    parsed = parse_date(value)
    if parsed is not None:
        return format_iso_date(parsed, spaced=True)
    return _to_str(value)


def format_won(value: float | None, unit: str) -> str:
    if value is None:
        return ""
    if abs(value) < 0.5:
        return "-"
    if unit == "천원":
        n = int(round(value / 1000))
        if n == 0:
            return "-"
        return f"{n:,}"
    return f"{int(round(value)):,}"


@dataclass
class StoreRecord:
    code: str
    name: str
    station: str
    biz_type: str
    biz_detail: str
    major: str
    middle: str
    minor: str
    brand: str
    owner: str
    location: str
    area: str
    contract_start: str
    contract_end: str
    monthly_sales: float | None
    source: str = ""
    start_date: date | None = None
    end_date: date | None = None
    status: str = ""
    skip_reason: str = ""
    stage: str = ""
    open_date: date | None = None
    close_date: date | None = None
    fc: str = ""

    @property
    def is_locker(self) -> bool:
        blob = f"{self.biz_type}{self.biz_detail}{self.middle}{self.name}"
        return "보관함" in blob

    @property
    def is_vending(self) -> bool:
        return self.biz_type == "자판기" or self.is_locker

    @property
    def is_shop(self) -> bool:
        return not self.is_vending

    def short_brand(self) -> str:
        """표·사진 카드용 짧은 브랜드. 원본 PPT는 '스토리웨이', '다이소' 수준."""
        compact = f"{self.name} {self.brand}".replace(" ", "")
        if "카페스토리웨이" in compact and "라운지" in compact:
            return "카페스토리웨이"
        known = [
            "스토리웨이",
            "크리스피크림",
            "핫브레드",
            "마성떡볶이",
            "삼송빵집",
            "다이소",
            "더베이크",
            "진김밥",
            "동산화원",
            "행복한 광목",
            "행복한광목",
            "명품마루",
        ]
        blob = f"{self.name} {self.brand}"
        for brand in known:
            if brand.replace(" ", "") in blob.replace(" ", ""):
                if brand == "행복한광목":
                    return "행복한 광목"
                return brand
        paren = re.findall(r"\(([^)]+)\)", self.name or "")
        skip_paren = {"김미숙", "직영", "단기", "부경", "양오순"}
        for part in paren:
            compact = part.replace(" ", "")
            if part in skip_paren or compact in skip_paren:
                continue
            if re.fullmatch(r"\d+-\d+", compact):
                continue
            if compact.startswith("22년") or "주식회사" in part or part.startswith("("):
                continue
            if 1 < len(part) <= 10:
                return part
        brand = (self.brand or "").strip()
        if brand.startswith("스토리웨이"):
            return "스토리웨이"
        if brand in skip_paren or re.fullmatch(r"\d+-\d+", brand.replace(" ", "")):
            return ""
        if brand and len(brand) <= 8 and not brand.startswith("롯데"):
            return brand
        return ""

    def card_title(self) -> str:
        brand = self.short_brand()
        if "단기" in (self.name or "") and ("명품마루" in (self.name or "") or brand == "명품마루"):
            return "명품마루(단기)"
        if brand:
            return brand
        name = self.name or ""
        return name if len(name) <= 12 else name[:12]

    def display_name(self) -> str:
        name = self.name or ""
        compact = name.replace(" ", "")
        if "카페스토리웨이" in compact and "라운지" in compact:
            return "카페스토리웨이"
        return name

    def location_display(self) -> str:
        loc = (self.location or "").strip()
        loc = loc.replace("상행선타는곳", "상행 타는곳").replace("하행선타는곳", "하행 타는곳")
        loc = loc.replace("상행선 타는곳", "상행 타는곳").replace("하행선 타는곳", "하행 타는곳")
        return loc



@dataclass
class RidershipYear:
    station: str
    year: int = 0
    onboard: dict[str, float] = field(default_factory=dict)
    offboard: dict[str, float] = field(default_factory=dict)
    total_on: float | None = None
    total_off: float | None = None


@dataclass
class RidershipSet:
    station: str
    by_year: dict[int, RidershipYear] = field(default_factory=dict)


@dataclass
class ExcelBundle:
    stores: list[StoreRecord]
    ridership: RidershipSet | None
    station: str
    inactive: list[StoreRecord] = field(default_factory=list)
    as_of: date | None = None


def _flatten_columns(columns) -> list[str]:
    out = []
    for col in columns:
        if isinstance(col, tuple):
            parts = [_to_str(p) for p in col if _to_str(p) and not str(p).startswith("Unnamed")]
            # 하위 헤더가 더 구체적이면 그걸 우선
            name = parts[-1] if parts else ""
            if len(parts) >= 2 and parts[-1] in {"시작일", "종료일", "계약만료일"}:
                name = "_".join(parts[-2:])
            out.append(name.replace("\n", "").strip())
        else:
            out.append(_to_str(col).replace("\n", "").strip())
    # 중복 열 이름에 접미사
    seen: dict[str, int] = {}
    unique = []
    for name in out:
        n = seen.get(name, 0)
        seen[name] = n + 1
        unique.append(name if n == 0 else f"{name}_{n+1}")
    return unique


def _pick_col(df: pd.DataFrame, *candidates: str) -> str | None:
    cols = list(df.columns)
    for cand in candidates:
        for col in cols:
            if str(col).replace(" ", "") == cand.replace(" ", ""):
                return col
    for cand in candidates:
        for col in cols:
            if cand.replace(" ", "") in str(col).replace(" ", "").replace("\n", ""):
                return col
    return None


def load_sales(path: str | Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, header=1, engine="openpyxl")
    df.columns = [str(c).replace("\n", "").strip() for c in df.columns]
    if _to_str(df.iloc[0].get("매장코드")) == "매장코드":
        df = df.iloc[1:].reset_index(drop=True)
    return df


def load_master(path: str | Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, header=[5, 6], engine="openpyxl")
    df.columns = _flatten_columns(df.columns)
    if not df.empty and _to_str(df.iloc[0].get("매장코드")) == "매장코드":
        df = df.iloc[1:].reset_index(drop=True)
    return df


def _sheet_year(name: str) -> int | None:
    match = re.search(r"(19|20)\d{2}", name or "")
    return int(match.group(0)) if match else None


def _ridership_station_col(raw: pd.DataFrame):
    for col in raw.columns:
        parts = col if isinstance(col, tuple) else (col,)
        blob = "".join(_to_str(p) for p in parts).replace(" ", "")
        if "역명" in blob:
            return col
    if raw.shape[1] > 3:
        return raw.columns[3]
    return None


def _ridership_mode_key(top: str) -> str | None:
    """열 상단 헤더 → passenger / metro / total. SR·참고는 제외.

    '총합계(여객+광역)'처럼 총합에 광역·여객이 같이 적힌 경우는 총합으로 본다.
    """
    compact = (top or "").replace(" ", "").replace("\n", "")
    if not compact or "SR" in compact or compact.startswith("참고"):
        return None
    if "총합" in compact:
        return "total"
    if "광역" in compact:
        return "metro"
    if "여객" in compact:
        return "passenger"
    return None


def _has_ridership_amount(values: dict) -> bool:
    return any(v for v in values.values() if v)


def _merge_ridership_modes(
    passenger: dict[str, float],
    metro: dict[str, float],
    total: dict[str, float],
) -> dict[str, float]:
    """역에 실제 인원이 있는 수단만 합친다.

    여객만 → 여객, 광역만 → 광역, 둘 다 → 여객+광역.
    둘 다 없으면 총합. 총합을 여객·광역에 다시 더하지 않는다.
    """
    use_passenger = _has_ridership_amount(passenger)
    use_metro = _has_ridership_amount(metro)
    if not use_passenger and not use_metro:
        return {k: v for k, v in total.items() if v}
    out: dict[str, float] = {}
    for key in set(passenger) | set(metro):
        amount = 0.0
        if use_passenger:
            amount += passenger.get(key) or 0.0
        if use_metro:
            amount += metro.get(key) or 0.0
        if amount:
            out[key] = amount
    return out


def _merged_ridership_total(
    passenger_months: dict[str, float],
    metro_months: dict[str, float],
    passenger_sum: float,
    metro_sum: float,
    has_passenger_sum: bool,
    has_metro_sum: bool,
    total_sum: float,
    has_total_sum: bool,
) -> float | None:
    use_passenger = _has_ridership_amount(passenger_months) or (has_passenger_sum and passenger_sum)
    use_metro = _has_ridership_amount(metro_months) or (has_metro_sum and metro_sum)
    if use_passenger or use_metro:
        amount = 0.0
        if use_passenger:
            amount += passenger_sum if has_passenger_sum else sum(passenger_months.values())
        if use_metro:
            amount += metro_sum if has_metro_sum else sum(metro_months.values())
        return amount
    if has_total_sum:
        return total_sum
    return None


def _parse_ridership_sheet(raw: pd.DataFrame, station: str, year: int) -> RidershipYear | None:
    station_n = normalize_station(station)
    if raw is None or raw.empty:
        return None
    station_col = _ridership_station_col(raw)
    if station_col is None:
        return None
    matched = raw[raw[station_col].map(lambda v: normalize_station(_to_str(v)) == station_n)]
    if matched.empty:
        return None

    buckets = {
        "passenger": {"on": {}, "off": {}, "sum_on": 0.0, "sum_off": 0.0, "has_sum_on": False, "has_sum_off": False},
        "metro": {"on": {}, "off": {}, "sum_on": 0.0, "sum_off": 0.0, "has_sum_on": False, "has_sum_off": False},
        "total": {"on": {}, "off": {}, "sum_on": 0.0, "sum_off": 0.0, "has_sum_on": False, "has_sum_off": False},
    }

    for _, row in matched.iterrows():
        for col in raw.columns:
            if not isinstance(col, tuple) or len(col) < 3:
                continue
            a_s, b_s, c_s = _to_str(col[0]), _to_str(col[1]), _to_str(col[2])
            mode = _ridership_mode_key(a_s)
            if mode is None:
                continue
            val = _to_float(row[col])
            if val is None:
                continue
            bucket = buckets[mode]
            month = re.search(r"(\d+)월", b_s)
            if month and c_s == "승차":
                key = f"{int(month.group(1))}월"
                bucket["on"][key] = bucket["on"].get(key, 0.0) + val
            elif month and c_s == "하차":
                key = f"{int(month.group(1))}월"
                bucket["off"][key] = bucket["off"].get(key, 0.0) + val
            elif "합계" in b_s and c_s == "승차":
                bucket["sum_on"] += val
                bucket["has_sum_on"] = True
            elif "합계" in b_s and c_s == "하차":
                bucket["sum_off"] += val
                bucket["has_sum_off"] = True

    p_b, m_b, t_b = buckets["passenger"], buckets["metro"], buckets["total"]
    result = RidershipYear(station=station_n, year=year)
    result.onboard = _merge_ridership_modes(p_b["on"], m_b["on"], t_b["on"])
    result.offboard = _merge_ridership_modes(p_b["off"], m_b["off"], t_b["off"])
    result.total_on = _merged_ridership_total(
        p_b["on"], m_b["on"], p_b["sum_on"], m_b["sum_on"],
        p_b["has_sum_on"], m_b["has_sum_on"], t_b["sum_on"], t_b["has_sum_on"],
    )
    result.total_off = _merged_ridership_total(
        p_b["off"], m_b["off"], p_b["sum_off"], m_b["sum_off"],
        p_b["has_sum_off"], m_b["has_sum_off"], t_b["sum_off"], t_b["has_sum_off"],
    )
    if not result.onboard and result.total_on is None:
        return None
    return result


def load_ridership(path: str | Path, station: str, as_of: date | None = None) -> RidershipSet | None:
    as_of = as_of or date.today()
    by_year: dict[int, RidershipYear] = {}
    with pd.ExcelFile(path, engine="openpyxl") as xl:
        sheets = list(xl.sheet_names)
        for sheet in sheets:
            try:
                raw = pd.read_excel(xl, sheet_name=sheet, header=[0, 1, 2])
            except Exception:
                continue
            year = _sheet_year(sheet)
            if year is None:
                year = as_of.year if len(sheets) == 1 else _sheet_year(str(raw.columns[0][0] if len(raw.columns) else ""))
            if year is None:
                year = as_of.year
            parsed = _parse_ridership_sheet(raw, station, year)
            if parsed is not None:
                by_year[parsed.year] = parsed
    if not by_year:
        return None
    return RidershipSet(station=normalize_station(station), by_year=by_year)


def _last_complete_month(as_of: date | None) -> int | None:
    if as_of is None:
        return None
    last_day = calendar.monthrange(as_of.year, as_of.month)[1]
    if as_of.day < last_day:
        return as_of.month - 1 if as_of.month > 1 else None
    return as_of.month


def _sales_month_avg(row: pd.Series, df: pd.DataFrame, as_of: date | None = None) -> float | None:
    month_cols = []
    for c in df.columns:
        match = re.fullmatch(r"(\d+)월\.1", str(c))
        if match:
            month_cols.append((int(match.group(1)), c))
    last_m = _last_complete_month(as_of)
    vals = []
    for month, col in month_cols:
        if last_m is not None and month > last_m:
            continue
        val = _to_float(row[col])
        if val:
            vals.append(val)
    if vals:
        return sum(vals) / len(vals)
    total = _to_float(row.get("총매출"))
    if total is not None and last_m:
        return total / last_m
    if total is not None:
        return total / 12.0
    return None


def _pick_first_col(df: pd.DataFrame, *exact_names: str) -> str | None:
    for name in exact_names:
        for col in df.columns:
            if str(col).replace(" ", "").replace("\n", "") == name.replace(" ", ""):
                return col
    return None


def merge_stores(
    sales_df: pd.DataFrame,
    master_df: pd.DataFrame,
    station: str,
    as_of: date | None = None,
) -> list[StoreRecord]:
    station_n = normalize_station(station)
    sales = sales_df[sales_df["역명"].map(lambda v: normalize_station(_to_str(v)) == station_n)].copy()
    master = master_df[master_df["역명"].map(lambda v: normalize_station(_to_str(v)) == station_n)].copy()

    master_by_code: dict[str, pd.Series] = {}
    code_col = _pick_col(master, "매장코드") or "매장코드"
    for _, row in master.iterrows():
        code = _to_str(row.get(code_col))
        if code:
            master_by_code[code] = row

    brand_col = _pick_col(master, "브랜드매장", "브랜드")
    owner_col = _pick_col(master, "대표자")
    loc_col = _pick_col(master, "매장위치")
    area_col_m = _pick_col(master, "면적")
    # PPT 최초계약일·최대운영만료일 = 마스터 시작/종료. 기본계약일자는 당해 갱신분.
    start_col = _pick_first_col(master, "시작", "최초계약일", "최초계약일자")
    if start_col is None:
        start_col = _pick_first_col(master, "기본계약일자")
        if start_col is None:
            for col in master.columns:
                name = str(col).replace(" ", "").replace("\n", "")
                if "기본계약일자" in name and "만료" not in name and "종료" not in name:
                    start_col = col
                    break
    end_col = _pick_first_col(master, "종료", "최대운영만료일")
    if end_col is None:
        for col in master.columns:
            name = str(col).replace(" ", "")
            if "계약만료일" in name and "기본계약" not in name:
                end_col = col
                break
        if end_col is None:
            end_col = _pick_col(master, "계약만료일")
    status_col = _pick_col(master, "영업구분")
    stage_col = _pick_col(master, "운영단계")
    open_col = _pick_col(master, "영업개시일")
    close_col = _pick_col(master, "영업종료일")
    fc_col = _pick_col(master, "담당FC")

    stores: list[StoreRecord] = []
    for _, row in sales.iterrows():
        code = _to_str(row.get("매장코드"))
        if not code:
            continue
        mrow = master_by_code.get(code)
        brand = ""
        owner = ""
        location = ""
        area = _to_str(row.get("면적"))
        start = ""
        end = ""
        start_dt = None
        end_dt = None
        open_dt = None
        close_dt = None
        status = ""
        stage = ""
        fc = ""
        monthly = _sales_month_avg(row, sales_df, as_of=as_of)
        if mrow is not None:
            if brand_col:
                brand = _to_str(mrow.get(brand_col))
            if owner_col:
                owner = _to_str(mrow.get(owner_col))
            if loc_col:
                location = _to_str(mrow.get(loc_col))
            if area_col_m and _to_str(mrow.get(area_col_m)):
                area = _to_str(mrow.get(area_col_m))
            if start_col:
                start_dt = parse_date(mrow.get(start_col))
                start = format_ppt_date(mrow.get(start_col))
            if end_col:
                end_dt = parse_date(mrow.get(end_col))
                end = format_ppt_date(mrow.get(end_col))
            if open_col:
                open_dt = parse_date(mrow.get(open_col))
            if close_col:
                close_dt = parse_date(mrow.get(close_col))
            if status_col:
                status = _to_str(mrow.get(status_col))
            if stage_col:
                stage = _to_str(mrow.get(stage_col))
            if fc_col:
                fc = _to_str(mrow.get(fc_col))
        name = _to_str(row.get("매장명"))
        if not brand:
            paren = re.search(r"\(([^)]+)\)", name)
            if paren and paren.group(1) not in {"김미숙", "직영", "단기", "양오순"}:
                brand = paren.group(1)
        stores.append(
            StoreRecord(
                code=code,
                name=name,
                station=station_n,
                biz_type=_to_str(row.get("사업구분")),
                biz_detail=_to_str(row.get("사업상세구분")),
                major=_to_str(row.get("대분류")),
                middle=_to_str(row.get("중분류")),
                minor=_to_str(row.get("소분류")),
                brand=brand,
                owner=owner,
                location=location,
                area=area,
                contract_start=start,
                contract_end=end,
                monthly_sales=monthly,
                source="sales",
                start_date=start_dt,
                end_date=end_dt,
                status=status,
                stage=stage,
                open_date=open_dt,
                close_date=close_dt,
                fc=fc,
            )
        )
    return stores


CLOSED_STATUS_MARKERS = ("폐점", "종료", "폐쇄", "계약종료", "영업종료")


def inactive_reason(store: StoreRecord, as_of: date) -> str | None:
    """운영 중이면 None, 아니면 제외 사유.

    자판기는 기본계약일·계약만료일이 비어 있는 경우가 많다.
    마스터의 운영단계/영업구분(영업중), 영업개시일, 영업종료일을 기준으로 본다.
    """
    name = store.name or ""
    if "(폐쇄)" in name or name.rstrip().endswith("폐쇄"):
        return "매장명 폐쇄 표시"
    status_blob = f"{store.status}{store.stage}"
    if any(mark in status_blob for mark in CLOSED_STATUS_MARKERS):
        return f"영업구분 {store.status or store.stage}"
    if store.close_date and as_of > store.close_date:
        return f"영업종료일 {format_iso_date(store.close_date)} 경과"
    if store.open_date and as_of < store.open_date:
        return f"영업개시일 {format_iso_date(store.open_date)} 미도래"
    if "영업중" in (store.status or "") or "영업중" in (store.stage or ""):
        return None
    if not store.status and not store.stage:
        return "마스터 운영단계(영업중) 확인 안 됨"
    return "현재 영업중이 아님"


def filter_active_stores(
    stores: list[StoreRecord],
    as_of: date,
) -> tuple[list[StoreRecord], list[StoreRecord]]:
    active: list[StoreRecord] = []
    inactive: list[StoreRecord] = []
    for store in stores:
        reason = inactive_reason(store, as_of)
        if reason:
            store.skip_reason = reason
            inactive.append(store)
        else:
            active.append(store)
    return active, inactive


def load_excel_bundle(
    sales_path: str | Path,
    master_path: str | Path,
    ridership_path: str | Path | None,
    station: str,
    as_of: date | None = None,
) -> ExcelBundle:
    as_of = as_of or date.today()
    sales_df = load_sales(sales_path)
    master_df = load_master(master_path)
    stores = merge_stores(sales_df, master_df, station, as_of=as_of)
    active, inactive = filter_active_stores(stores, as_of)
    ridership = None
    if ridership_path:
        ridership = load_ridership(ridership_path, station, as_of=as_of)
    return ExcelBundle(
        stores=active,
        ridership=ridership,
        station=normalize_station(station),
        inactive=inactive,
        as_of=as_of,
    )
