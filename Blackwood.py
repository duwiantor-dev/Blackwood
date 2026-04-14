
import io
import re
from typing import List

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dashboard Analisa Produk", layout="wide")

PERIODS = ["7DAY", "14DAY", "30DAY"]
DIVISIONS = ["DIV03", "DIV04", "DIV05"]

MPLSSR_DIV_COLS = {
    "7DAY": {"DIV03": "03 OLP", "DIV04": "04 MOD", "DIV05": "05 OLR"},
    "14DAY": {"DIV03": "03 OLP.1", "DIV04": "04 MOD.1", "DIV05": "05 OLR.1"},
    "30DAY": {"DIV03": "03 OLP.2", "DIV04": "04 MOD.2", "DIV05": "05 OLR.2"},
}

VALID_PRICELIST_SHEETS = [
    "LAPTOP",
    "TELCO",
    "PC HOM ELE",
    "SOF COM SUP",
    "ACC",
    "SER OTH CON",
]

PRICE_SEGMENTS = [
    (0, 1_000_000, "< 1 JUTA"),
    (1_000_000, 1_500_000, "1 - 1.5 JUTA"),
    (1_500_000, 2_000_000, "1.5 - 2 JUTA"),
    (2_000_000, 2_500_000, "2 - 2.5 JUTA"),
    (2_500_000, 3_000_000, "2.5 - 3 JUTA"),
    (3_000_000, 4_000_000, "3 - 4 JUTA"),
    (4_000_000, 5_000_000, "4 - 5 JUTA"),
    (5_000_000, 7_000_000, "5 - 7 JUTA"),
    (7_000_000, 10_000_000, "7 - 10 JUTA"),
    (10_000_000, 12_500_000, "10 - 12.5 JUTA"),
    (12_500_000, 15_000_000, "12.5 - 15 JUTA"),
    (15_000_000, 20_000_000, "15 - 20 JUTA"),
    (20_000_000, 25_000_000, "20 - 25 JUTA"),
    (25_000_000, 30_000_000, "25 - 30 JUTA"),
    (30_000_000, 40_000_000, "30 - 40 JUTA"),
    (40_000_000, np.inf, "40 JUTA - UP"),
]

SEGMENT_ORDER = [label for _, _, label in PRICE_SEGMENTS]

st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 1rem;}
    .upload-card-wrap {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        background: #ffffff;
        padding: 14px;
        margin-bottom: 16px;
    }
    .table-card {
        border: 1px solid #d9d9d9;
        border-radius: 8px;
        background: #fff;
        padding: 10px;
        margin-bottom: 14px;
    }
    .alert-wrap {
        border: 1px solid #d9d9d9;
        border-radius: 8px;
        background: #fff;
        overflow: auto;
        max-height: 520px;
    }
    table.alert-table {
        border-collapse: collapse;
        width: max-content;
        min-width: 100%;
        table-layout: fixed;
        font-size: 12px;
    }
    table.alert-table th, table.alert-table td {
        border: 1px solid #e5e7eb;
        padding: 6px 8px;
        text-align: left;
        white-space: nowrap;
    }
    table.alert-table thead th {
        position: sticky;
        top: 0;
        background: #f8fafc;
        z-index: 2;
    }
    .main-fixed-wrap {
        border: 1px solid #d9d9d9;
        border-radius: 8px;
        background: #fff;
        overflow: auto;
        max-height: 520px;
    }
    table.main-fixed {
        border-collapse: collapse;
        width: max-content;
        min-width: 100%;
        table-layout: fixed;
        font-size: 12px;
    }
    table.main-fixed th, table.main-fixed td {
        border: 1px solid #e5e7eb;
        padding: 6px 8px;
        text-align: left;
        white-space: nowrap;
    }
    table.main-fixed thead th {
        position: sticky;
        top: 0;
        background: #f8fafc;
        z-index: 2;
    }
    table.main-fixed th:nth-child(1),
    table.main-fixed td:nth-child(1) { min-width: 180px; }
    table.main-fixed th:nth-child(2),
    table.main-fixed td:nth-child(2) { min-width: 90px; }
    table.main-fixed th:nth-child(3),
    table.main-fixed td:nth-child(3) { min-width: 80px; }
    table.main-fixed th:nth-child(4),
    table.main-fixed td:nth-child(4) {
        min-width: 420px;
        max-width: 420px;
        white-space: normal;
        word-break: break-word;
        line-height: 1.25;
    }
    table.main-fixed th:nth-child(5),
    table.main-fixed td:nth-child(5) { min-width: 80px; }
    table.main-fixed th:nth-child(6),
    table.main-fixed td:nth-child(6),
    table.main-fixed th:nth-child(7),
    table.main-fixed td:nth-child(7),
    table.main-fixed th:nth-child(8),
    table.main-fixed td:nth-child(8),
    table.main-fixed th:nth-child(9),
    table.main-fixed td:nth-child(9) { min-width: 70px; }
    .bg-red {
        background: #ffebee;
        color: #c62828;
        font-weight: 700;
    }
    .status-refill {
        background: #ffebee;
        color: #c62828;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

def normalize_text(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.strip()
        .str.upper()
        .replace({"NAN": np.nan, "NONE": np.nan, "": np.nan})
    )

def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")

def price_segment(price: float) -> str:
    if pd.isna(price):
        return "UNKNOWN"
    for low, high, label in PRICE_SEGMENTS:
        if low <= float(price) < high:
            return label
    return "UNKNOWN"

def segment_sort_key(label: str) -> int:
    try:
        return SEGMENT_ORDER.index(label)
    except ValueError:
        return len(SEGMENT_ORDER)

def first_row_contains_text(df: pd.DataFrame, text: str):
    target = str(text).strip().upper()
    for idx in range(len(df)):
        row_text = df.iloc[idx].astype(str).str.upper()
        if row_text.str.contains(target, na=False).any():
            return idx
    return None

def _ffill_header(values: List) -> List:
    out = []
    last = None
    for v in values:
        if pd.notna(v) and str(v).strip() != "":
            last = str(v).strip().upper()
            out.append(last)
        else:
            out.append(last)
    return out


def _norm_header_cell(value) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return ""
    return str(value).strip().upper()

def find_div05_stock_columns(columns: List[str], excel_row3_raw: List, excel_row4_raw: List) -> List[str]:
    # Excel row 3 dan row 4 bisa berupa merged cell.
    # Supaya semua kolom di dalam merge range ikut terbaca, keduanya di-forward fill ke kanan.
    row3_vals = [_norm_header_cell(v) for v in _ffill_header(excel_row3_raw)]
    row4_vals = [_norm_header_cell(v).replace(" ", "") for v in _ffill_header(excel_row4_raw)]

    # Titik awal range area ditandai kata RAM di row 3.
    ram_start = next((i for i, v in enumerate(row3_vals) if "RAM" in v), None)
    if ram_start is None:
        return []

    # Dari kolom RAM sampai kolom paling kanan file.
    idx_ram_to_right = set(range(ram_start, len(columns)))

    # Ambil semua kolom yang pada row 4 bernilai 5B
    idx_5b = {i for i, v in enumerate(row4_vals) if v == "5B"}

    final_idx = sorted(idx_ram_to_right & idx_5b)
    return [columns[i] for i in final_idx if 0 <= i < len(columns)]

def area_code_matches(value, prefixes: List[str]) -> bool:
    if pd.isna(value):
        return False
    txt = str(value).strip().upper().replace(" ", "")
    return any(txt.startswith(p) for p in prefixes)

def normalize_warehouse_code(value) -> str:
    if pd.isna(value):
        return np.nan
    txt = str(value).strip().upper()
    if "-" in txt:
        txt = txt.split("-", 1)[1]
    txt = txt.replace(" ", "")
    txt = re.sub(r"0+(?=\d)", "", txt)
    txt = txt.replace("0A", "A").replace("0B", "B").replace("0C", "C")
    return txt


def normalize_sales_pivot_gudang(value) -> str:
    if pd.isna(value):
        return np.nan
    txt = str(value).strip().upper()
    if "-" in txt:
        txt = txt.split("-", 1)[1]
    txt = txt.replace(" ", "")
    m = re.match(r"([A-Z]+)0*(\d+[A-Z]?)$", txt)
    if m:
        prefix = m.group(1)
        suffix = m.group(2)
        return f"{prefix} {suffix}"
    return txt


def normalize_team_code(value) -> str:
    if pd.isna(value):
        return np.nan
    txt = str(value).strip().upper()
    if "-" in txt:
        txt = txt.split("-", 1)[1]
    txt = txt.replace(" ", "")
    m = re.match(r"([A-Z]+)0*(\d+)$", txt)
    if m:
        return f"{m.group(1)} {int(m.group(2))}A"
    m = re.match(r"([A-Z]+)0*(\d+)([A-Z])$", txt)
    if m:
        return f"{m.group(1)} {int(m.group(2))}{m.group(3)}"
    return txt

def ensure_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


# =========================================================
# MPLSSR
# =========================================================
def load_mplssr(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="ALL", header=1)
    df = df.iloc[4:].copy().reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

    base_cols = ["PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI"]
    for c in base_cols:
        if c not in df.columns:
            df[c] = np.nan

    div_cols = []
    for p in PERIODS:
        for col in MPLSSR_DIV_COLS[p].values():
            if col in df.columns:
                div_cols.append(col)

    df = df[base_cols + div_cols].copy()

    for c in base_cols:
        df[c] = normalize_text(df[c])

    df = df[df["KODE BARANG"].notna()].copy()
    df = df[~df["KODE BARANG"].isin(["TOTAL", "SHARE%"])]

    rows = []
    for period in PERIODS:
        for div, col in MPLSSR_DIV_COLS[period].items():
            if col not in df.columns:
                continue
            tmp = df[["PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI"]].copy()
            tmp["PERIOD"] = period
            tmp["DIVISION"] = div
            tmp["QTY"] = to_num(df[col]).fillna(0)
            rows.append(tmp)

    out = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(
        columns=["PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI", "PERIOD", "DIVISION", "QTY"]
    )
    out["MERGE_KEY"] = normalize_text(out["KODE BARANG"])
    out["SKU NO"] = out["MERGE_KEY"]
    return out

# =========================================================
# PRICELIST
# =========================================================

def parse_pricelist_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    raw = xls.parse(sheet_name=sheet_name, header=None).copy()

    if sheet_name.upper() == "LAPTOP":
        coming_idx = first_row_contains_text(raw, "COMING")
        end_coming_idx = first_row_contains_text(raw, "END COMING")
        if coming_idx is not None and end_coming_idx is not None and end_coming_idx >= coming_idx:
            raw = raw.drop(index=range(coming_idx, end_coming_idx + 1)).reset_index(drop=True)

    row1 = raw.iloc[1].tolist()
    row2_raw = raw.iloc[2].tolist()
    row2 = _ffill_header(row2_raw)
    row3_raw = raw.iloc[3].tolist()
    row3 = [str(x).strip().upper() if pd.notna(x) and str(x).strip() != "" else None for x in row3_raw]

    columns = []
    for i, v in enumerate(row1):
        v1 = str(v).strip().upper() if pd.notna(v) and str(v).strip() != "" else None
        if v1 is not None:
            columns.append(v1)
        elif row2[i] is not None and row3[i] is not None:
            columns.append(f"{row2[i]}__{row3[i]}")
        elif row3[i] is not None:
            columns.append(row3[i])
        else:
            columns.append(f"COL_{i}")

    df = raw.iloc[5:].copy().reset_index(drop=True)
    df.columns = columns

    for c in ["SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "TOT", "M3"]:
        if c not in df.columns:
            df[c] = np.nan

    df["SKU NO"] = normalize_text(df["SKU NO"])
    df["PRODUCT"] = normalize_text(df["PRODUCT"])
    df["KODEBARANG"] = normalize_text(df["KODEBARANG"])
    df["SPESIFIKASI"] = normalize_text(df["SPESIFIKASI"])

    df = df[df["KODEBARANG"].notna()].copy()
    df = df[~df["KODEBARANG"].isin(["TOTAL"])]

    stock03_cols = [columns[i] for i, area in enumerate(row3) if area_code_matches(area, ["3", "03"])]
    stock04_cols = [columns[i] for i, area in enumerate(row3) if area_code_matches(area, ["4", "04"])]
    stock05_cols = find_div05_stock_columns(columns, row2_raw, row3_raw)

    df["PRICE"] = to_num(df["M3"]) * 1000
    df["STOK_DIV03"] = df[stock03_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock03_cols else 0
    df["STOK_DIV04"] = df[stock04_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock04_cols else 0
    df["STOK_DIV05"] = df[stock05_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock05_cols else 0
    df["CATEGORY"] = sheet_name.upper()
    df["PRICE_SEGMENT"] = df["PRICE"].apply(price_segment)
    df["MERGE_KEY"] = df["KODEBARANG"]

    return df[[
        "SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "PRICE",
        "STOK_DIV03", "STOK_DIV04", "STOK_DIV05",
        "CATEGORY", "PRICE_SEGMENT", "MERGE_KEY"
    ]]

def load_pricelist(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    sheets = [s for s in xls.sheet_names if s.upper() in VALID_PRICELIST_SHEETS]
    frames = [parse_pricelist_sheet(xls, s) for s in sheets]
    if not frames:
        return pd.DataFrame(columns=[
            "SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "PRICE",
            "STOK_DIV03", "STOK_DIV04", "STOK_DIV05",
            "CATEGORY", "PRICE_SEGMENT", "MERGE_KEY"
        ])
    out = pd.concat(frames, ignore_index=True)
    out = out.loc[:, ~out.columns.duplicated()]
    out = out.drop_duplicates(subset=["MERGE_KEY"], keep="first")
    return out

# =========================================================
# PRICELIST WITH WAREHOUSES
# =========================================================
def parse_pricelist_sheet_with_warehouses(xls: pd.ExcelFile, sheet_name: str):
    raw = xls.parse(sheet_name=sheet_name, header=None).copy()

    if sheet_name.upper() == "LAPTOP":
        coming_idx = first_row_contains_text(raw, "COMING")
        end_coming_idx = first_row_contains_text(raw, "END COMING")
        if coming_idx is not None and end_coming_idx is not None and end_coming_idx >= coming_idx:
            raw = raw.drop(index=range(coming_idx, end_coming_idx + 1)).reset_index(drop=True)

    row1 = raw.iloc[1].tolist()
    row2 = _ffill_header(raw.iloc[2].tolist())
    row3 = [str(x).strip().upper() if pd.notna(x) and str(x).strip() != "" else None for x in raw.iloc[3].tolist()]
    row4 = [str(x).strip().upper() if pd.notna(x) and str(x).strip() != "" else None for x in raw.iloc[4].tolist()] if len(raw) > 4 else [None] * len(row1)

    columns = []
    warehouse_meta = []

    for i, v in enumerate(row1):
        v1 = str(v).strip().upper() if pd.notna(v) and str(v).strip() != "" else None
        group = row3[i]
        wh = row4[i]

        if v1 is not None:
            columns.append(v1)
            warehouse_meta.append((None, None))
        elif group is not None and wh is not None:
            columns.append(f"{group}__{wh}")
            warehouse_meta.append((group, wh))
        elif group is not None:
            columns.append(group)
            warehouse_meta.append((group, None))
        else:
            columns.append(f"COL_{i}")
            warehouse_meta.append((None, None))

    df = raw.iloc[5:].copy().reset_index(drop=True)
    df.columns = columns

    for c in ["SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "M3"]:
        if c not in df.columns:
            df[c] = np.nan

    df["SKU NO"] = normalize_text(df["SKU NO"])
    df["PRODUCT"] = normalize_text(df["PRODUCT"])
    df["KODEBARANG"] = normalize_text(df["KODEBARANG"])
    df["SPESIFIKASI"] = normalize_text(df["SPESIFIKASI"])

    df = df[df["KODEBARANG"].notna()].copy()
    df = df[~df["KODEBARANG"].isin(["TOTAL"])].copy()
    df["PRICE"] = to_num(df["M3"]) * 1000

    warehouse_stock_cols = {}
    default_stock_cols = []
    jkt_stock_cols = []

    for i, col in enumerate(columns):
        group, wh = warehouse_meta[i]
        if group is None:
            continue
        group_clean = str(group).strip().upper().replace(" ", "")
        wh_clean = normalize_warehouse_code(wh) if wh is not None else None

        if group_clean == "DEFAULT":
            default_stock_cols.append(col)
        if group_clean == "JKT":
            jkt_stock_cols.append(col)
        if group_clean and wh_clean:
            warehouse_stock_cols[f"{group_clean} {wh_clean}"] = col

    if "BRAND" not in df.columns:
        df["BRAND"] = np.nan
    df["BRAND"] = normalize_text(df["BRAND"])
    df["PRICE_SEGMENT"] = df["PRICE"].apply(price_segment)
    keep_cols = ["SKU NO", "PRODUCT", "BRAND", "KODEBARANG", "SPESIFIKASI", "PRICE", "PRICE_SEGMENT"] + list(set(default_stock_cols + jkt_stock_cols + list(warehouse_stock_cols.values())))
    out = df[keep_cols].copy()
    out = out.loc[:, ~out.columns.duplicated()].copy()
    out["DEFAULT_STOCK_TOTAL"] = df.loc[:, ~df.columns.duplicated()][default_stock_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if default_stock_cols else 0
    return out, warehouse_stock_cols

def load_pricelist_with_warehouses(file):
    xls = pd.ExcelFile(file)
    sheets = [s for s in xls.sheet_names if s.upper() in VALID_PRICELIST_SHEETS]
    frames = []
    merged_map = {}
    for s in sheets:
        part, part_map = parse_pricelist_sheet_with_warehouses(xls, s)
        frames.append(part)
        merged_map.update(part_map)

    if not frames:
        return pd.DataFrame(), {}

    out = pd.concat(frames, ignore_index=True)
    out = out.loc[:, ~out.columns.duplicated()]
    out = out.drop_duplicates(subset=["SKU NO", "KODEBARANG"], keep="first").reset_index(drop=True)
    return out, merged_map

# =========================================================
# SALES PIVOT
# =========================================================
def load_sales_pivot(file) -> pd.DataFrame:
    raw = pd.read_excel(file, header=1).copy()
    raw.columns = [str(c).strip().upper() for c in raw.columns]
    raw = raw.loc[:, ~pd.Index(raw.columns).duplicated()].copy()

    team_col = next((c for c in raw.columns if c == "TEAM" or "TEAM" in c), None)
    kode_barang_col = next((c for c in raw.columns if "KODE BARANG" in c or "KODEBARANG" in c), None)
    qty_col = next((c for c in raw.columns if c == "QTY" or "QTY" in c or "PCS" in c or "TERJUAL" in c), None)
    tgl_col = next((c for c in raw.columns if c == "TGL" or "TGL" in c), None)

    if team_col is None or kode_barang_col is None or qty_col is None or tgl_col is None:
        raise ValueError(
            f"Format SALES PIVOT tidak cocok. Kolom terbaca: {list(raw.columns)}. "
            "Pastikan header row 2 berisi TEAM, KODE BARANG, QTY, dan TGL."
        )

    df = raw[[team_col, kode_barang_col, qty_col, tgl_col]].copy()
    df.columns = ["TEAM", "KODE BARANG", "QTY", "TGL"]
    df["TEAM"] = df["TEAM"].apply(normalize_team_code)
    df["KODE BARANG"] = normalize_text(df["KODE BARANG"])
    df["QTY"] = to_num(df["QTY"]).fillna(0)
    df["TGL"] = ensure_datetime(df["TGL"])

    df = df[df["TEAM"].notna()].copy()
    df = df[df["KODE BARANG"].notna()].copy()
    df = df[df["QTY"] > 0].copy()
    df = df[df["TGL"].notna()].copy()

    if df.empty:
        return pd.DataFrame(columns=["TEAM", "KODE BARANG", "QTY", "TGL", "PERIOD"])

    max_date = df["TGL"].max().normalize()

    period_frames = []
    for days, label in [(7, "7DAY"), (14, "14DAY"), (30, "30DAY")]:
        start_date = max_date - pd.Timedelta(days=days - 1)
        tmp = df[df["TGL"].dt.normalize() >= start_date].copy()
        if tmp.empty:
            continue
        agg = (
            tmp.groupby(["TEAM", "KODE BARANG"], as_index=False)["QTY"]
            .sum()
        )
        agg["PERIOD"] = label
        period_frames.append(agg)

    if not period_frames:
        return pd.DataFrame(columns=["TEAM", "KODE BARANG", "QTY", "PERIOD"])

    return pd.concat(period_frames, ignore_index=True).sort_values(
        ["PERIOD", "QTY", "TEAM", "KODE BARANG"], ascending=[True, False, True, True]
    ).reset_index(drop=True)

def build_sales_pivot_alerts(
    sales_pivot: pd.DataFrame,
    pricelist_wh: pd.DataFrame,
    warehouse_stock_cols: dict,
    period: str,
    selected_products=None,
    selected_brands=None,
    selected_segments=None,
) -> pd.DataFrame:
    empty_cols = ["TEAM", "KODE BARANG", "SPESIFIKASI", "QTY", "STOK", "KET", "GUDANG READY"]
    if sales_pivot.empty or pricelist_wh.empty:
        return pd.DataFrame(columns=empty_cols)

    base = sales_pivot[sales_pivot["PERIOD"] == period].copy()
    if base.empty:
        return pd.DataFrame(columns=empty_cols)

    base = (
        base.groupby(["TEAM", "KODE BARANG"], as_index=False)["QTY"]
        .sum()
    )

    pl = pricelist_wh.copy()
    if "KODEBARANG" not in pl.columns:
        return pd.DataFrame(columns=empty_cols)

    if selected_products:
        pl = pl[pl["PRODUCT"].isin(selected_products)]
    if selected_brands:
        pl = pl[pl["BRAND"].isin(selected_brands)]
    if selected_segments:
        pl = pl[pl["PRICE_SEGMENT"].isin(selected_segments)]

    merged = base.merge(pl, how="left", left_on="KODE BARANG", right_on="KODEBARANG")

    def find_stock_col_by_team(team_code):
        if pd.isna(team_code):
            return None
        team_code = str(team_code).strip().upper()
        exact_col = warehouse_stock_cols.get(team_code)
        if exact_col and exact_col in merged.columns:
            return exact_col
        compact = team_code.replace(" ", "")
        for col in merged.columns:
            col_txt = str(col).strip().upper().replace(" ", "")
            if compact in col_txt:
                return col
        return None

    all_ready_codes = sorted([str(k).strip().upper() for k in warehouse_stock_cols.keys() if str(k).strip().upper() != "DEFAULT"])

    def get_current_stock(row):
        stock_col = find_stock_col_by_team(row.get("TEAM"))
        if stock_col and stock_col in row.index:
            val = pd.to_numeric(row[stock_col], errors="coerce")
            return 0 if pd.isna(val) else float(val)
        return 0.0

    def get_ready_warehouses(row):
        current_team = str(row.get("TEAM")).strip().upper()
        ready_list = []
        for team_code in all_ready_codes:
            if team_code == current_team:
                continue
            stock_col = find_stock_col_by_team(team_code)
            if not stock_col:
                continue
            val = pd.to_numeric(row.get(stock_col), errors="coerce")
            val = 0 if pd.isna(val) else float(val)
            if val > 0:
                ready_list.append(team_code)
        return ", ".join(ready_list)

    merged["SPESIFIKASI"] = merged.get("SPESIFIKASI", "").fillna("")
    merged["STOK"] = merged.apply(get_current_stock, axis=1)
    merged["GUDANG READY"] = merged.apply(get_ready_warehouses, axis=1)
    merged["KET"] = np.where(
        (to_num(merged["QTY"]).fillna(0) > to_num(merged["STOK"]).fillna(0)) &
        (to_num(merged["STOK"]).fillna(0) <= 0) &
        (merged["GUDANG READY"].astype(str).str.strip() != ""),
        "REFILL",
        np.where(
            (to_num(merged["QTY"]).fillna(0) > to_num(merged["STOK"]).fillna(0)),
            "CEK",
            ""
        )
    )

    work_df = merged[merged["KET"] != ""].copy()
    if work_df.empty:
        return pd.DataFrame(columns=empty_cols)

    work_df["KEBUTUHAN_STOK"] = to_num(work_df["QTY"]).fillna(0) - to_num(work_df["STOK"]).fillna(0)
    out = work_df[["TEAM", "KODE BARANG", "SPESIFIKASI", "QTY", "STOK", "KET", "GUDANG READY", "KEBUTUHAN_STOK"]].copy()
    out["QTY"] = pd.to_numeric(out["QTY"], errors="coerce").fillna(0).round(0).astype(int)
    out["STOK"] = pd.to_numeric(out["STOK"], errors="coerce").fillna(0).round(0).astype(int)
    out = out.sort_values(["KEBUTUHAN_STOK", "QTY", "TEAM", "KODE BARANG"], ascending=[False, False, True, True]).reset_index(drop=True)
    return out[["TEAM", "KODE BARANG", "SPESIFIKASI", "QTY", "STOK", "KET", "GUDANG READY"]]

def render_sales_pivot_alert_table(df: pd.DataFrame):
    if df.empty:
        st.info("Analisa Stok belum menemukan data.")
        return

    show_df = df.copy()
    for col in ["QTY", "STOK"]:
        show_df[col] = pd.to_numeric(show_df[col], errors="coerce").fillna(0).round(0).astype(int)

    html = []
    html.append('<div class="alert-wrap"><table class="alert-table"><thead><tr>')
    for col in show_df.columns:
        html.append(f"<th>{col}</th>")
    html.append("</tr></thead><tbody>")

    for _, row in show_df.iterrows():
        html.append("<tr>")
        for col in show_df.columns:
            cls = ""
            if col in ["STOK", "KET"] and str(row.get("KET", "")).strip().upper() in ["REFILL", "CEK"]:
                cls = ' class="bg-red"'
            html.append(f"<td{cls}>{row[col]}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    st.markdown("".join(html), unsafe_allow_html=True)

# =========================================================
# BUILD TABLES
# =========================================================
def build_master(sales: pd.DataFrame, stock: pd.DataFrame) -> pd.DataFrame:
    df = sales.merge(stock, how="left", on="MERGE_KEY", suffixes=("_sales", "_stock"))

    for col in ["PRICE", "STOK_DIV03", "STOK_DIV04", "STOK_DIV05", "CATEGORY", "PRICE_SEGMENT"]:
        if col not in df.columns:
            df[col] = np.nan

    df["KODEBARANG"] = normalize_text(df["KODE BARANG"])
    df["SPESIFIKASI_FINAL"] = df["SPESIFIKASI_sales"].fillna(df.get("SPESIFIKASI_stock"))
    df["PRODUCT_FINAL"] = df["PRODUCT_sales"].fillna(df.get("PRODUCT_stock"))
    df["BRAND"] = normalize_text(df["BRAND"])
    df["QTY"] = to_num(df["QTY"]).fillna(0)

    stock_map = {"DIV03": "STOK_DIV03", "DIV04": "STOK_DIV04", "DIV05": "STOK_DIV05"}
    df["STOK_DIVISI"] = df.apply(lambda r: pd.to_numeric(r.get(stock_map[r["DIVISION"]]), errors="coerce"), axis=1).fillna(0)
    return df

def build_segment_table(df, period, comparison_division_label="03 OLP"):
    tmp = df[df["PERIOD"] == period].copy()
    tmp["SEGMENT"] = tmp["PRICE"].apply(price_segment)
    seg = tmp.groupby(["SEGMENT", "DIVISION"])["QTY"].sum().unstack().fillna(0).reset_index()
    for div in DIVISIONS:
        if div not in seg.columns:
            seg[div] = 0
    seg = seg[["SEGMENT", "DIV03", "DIV04", "DIV05"]].copy()
    seg = seg.sort_values("SEGMENT", key=lambda s: s.map(segment_sort_key)).reset_index(drop=True)
    seg.columns = ["SEGMENT", "03 OLP", "04 MOD", "05 OLR"]

    compare_col = comparison_division_label if comparison_division_label in ["03 OLP", "04 MOD"] else "03 OLP"
    seg["DELTA"] = to_num(seg["05 OLR"]).fillna(0) - to_num(seg[compare_col]).fillna(0)
    return seg

def build_brand_table(df, period, comparison_division_label="03 OLP"):
    brand = df[df["PERIOD"] == period].copy()
    brand = brand.groupby(["BRAND", "DIVISION"])["QTY"].sum().unstack().fillna(0).reset_index()
    for div in DIVISIONS:
        if div not in brand.columns:
            brand[div] = 0
    brand = brand[["BRAND", "DIV03", "DIV04", "DIV05"]].copy()
    brand["TOTAL"] = brand[["DIV03", "DIV04", "DIV05"]].sum(axis=1)
    brand = brand.sort_values(["TOTAL", "BRAND"], ascending=[False, True]).head(10).drop(columns=["TOTAL"])
    brand.columns = ["BRAND", "03 OLP", "04 MOD", "05 OLR"]

    compare_col = comparison_division_label if comparison_division_label in ["03 OLP", "04 MOD"] else "03 OLP"
    brand["DELTA"] = to_num(brand["05 OLR"]).fillna(0) - to_num(brand[compare_col]).fillna(0)
    return brand

def render_left_table(df, title, selected_division="05 OLR"):
    def is_number(v):
        return isinstance(v, (int, float, np.integer, np.floating)) and not pd.isna(v)

    def fmt_number(v):
        return f"{int(round(float(v))):,}".replace(",", ".")

    html = []
    html.append("""
    <div class="table-card">
      <div style="font-weight:700;font-size:16px;margin-bottom:8px;">""" + title + """</div>
      <div style="overflow-x:auto;">
        <table style="border-collapse:collapse;width:100%;font-size:12px;">
    """)
    html.append("<thead><tr>")
    for col in df.columns:
        html.append(f'<th style="border:1px solid #2b2b2b;background:#f3f4f6;padding:6px;text-align:left;">{col}</th>')
    html.append("</tr></thead><tbody>")

    compare_cols = ["03 OLP", "04 MOD", "05 OLR"]
    has_compare_cols = all(c in df.columns for c in compare_cols)

    for _, row in df.iterrows():
        html.append("<tr>")

        losing_selected = False
        if has_compare_cols and selected_division in compare_cols:
            current_val = row[selected_division]
            other_vals = [row[c] for c in compare_cols if c != selected_division]
            if is_number(current_val):
                numeric_others = [v for v in other_vals if is_number(v)]
                if numeric_others:
                    losing_selected = any(float(current_val) < float(v) for v in numeric_others)

        for col in df.columns:
            val = row[col]
            try:
                if pd.notna(val) and isinstance(val, (int, float, np.integer, np.floating)):
                    display = fmt_number(val)
                else:
                    display = "" if pd.isna(val) else str(val)
            except Exception:
                display = str(val)

            style = 'border:1px solid #2b2b2b;padding:6px;text-align:left;'
            if col == selected_division and losing_selected:
                style += 'color:#c62828;font-weight:700;background:#ffebee;'
            if col == "DELTA" and is_number(val) and float(val) < 0:
                style += 'color:#c62828;font-weight:700;background:#ffebee;'

            html.append(f'<td style="{style}">{display}</td>')
        html.append("</tr>")

    html.append("</tbody></table></div></div>")
    st.markdown("".join(html), unsafe_allow_html=True)

def build_main_table_filtered(
    df: pd.DataFrame,
    period: str,
    comparison_division_label: str,
    selected_segments=None,
    selected_brands=None,
    selected_products=None,
) -> pd.DataFrame:
    base = df[df["PERIOD"] == period].copy()

    if selected_segments:
        base = base[base["PRICE"].apply(price_segment).isin(selected_segments)]
    if selected_brands:
        base = base[base["BRAND"].isin(selected_brands)]
    if selected_products:
        base = base[base["PRODUCT_FINAL"].isin(selected_products)]

    qty = (
        base.groupby(["KODEBARANG", "SPESIFIKASI_FINAL", "PRICE", "DIVISION"], as_index=False)["QTY"]
        .sum()
        .pivot(index=["KODEBARANG", "SPESIFIKASI_FINAL", "PRICE"], columns="DIVISION", values="QTY")
        .fillna(0)
        .reset_index()
    )

    for div in DIVISIONS:
        if div not in qty.columns:
            qty[div] = 0

    product_brand = (
        base[["KODEBARANG", "PRODUCT_FINAL", "BRAND"]]
        .dropna(subset=["KODEBARANG"])
        .drop_duplicates(subset=["KODEBARANG"], keep="first")
        .copy()
    )

    stock_label_map = {"03 OLP": "STOK_DIV03", "04 MOD": "STOK_DIV04", "05 OLR": "STOK_DIV05"}
    stok_col = "STOK_DIV05"

    stock_df = (
        base[["KODEBARANG", "STOK_DIV03", "STOK_DIV04", "STOK_DIV05"]]
        .dropna(subset=["KODEBARANG"])
        .drop_duplicates(subset=["KODEBARANG"], keep="first")
        .copy()
    )
    stock_df["STOK"] = to_num(stock_df[stok_col]).fillna(0)

    out = qty.merge(product_brand, how="left", on="KODEBARANG")
    out = out.merge(stock_df, how="left", on="KODEBARANG")
    out["STOK"] = to_num(out["STOK"]).fillna(0)
    out["PRICE"] = to_num(out["PRICE"]).fillna(0)
    out["STOK_DIV03"] = to_num(out["STOK_DIV03"]).fillna(0)
    out["STOK_DIV04"] = to_num(out["STOK_DIV04"]).fillna(0)
    out["STOK_DIV05"] = to_num(out["STOK_DIV05"]).fillna(0)

    out = out.rename(columns={
        "PRODUCT_FINAL": "PRODUCT",
        "SPESIFIKASI_FINAL": "SPESIFIKASI",
        "PRICE": "M3",
        "DIV03": "03 OLP",
        "DIV04": "04 MOD",
        "DIV05": "05 OLR",
    })

    ordered_cols = [
        "KODEBARANG", "PRODUCT", "BRAND", "SPESIFIKASI", "M3",
        "03 OLP", "04 MOD", "05 OLR", "STOK",
        "STOK_DIV03", "STOK_DIV04", "STOK_DIV05"
    ]
    for col in ordered_cols:
        if col not in out.columns:
            out[col] = 0 if col not in ["KODEBARANG", "PRODUCT", "BRAND", "SPESIFIKASI"] else ""

    out["M3"] = (to_num(out["M3"]).fillna(0) / 1000).round(0)
    compare_map = {"03 OLP": "03 OLP", "04 MOD": "04 MOD"}
    compare_col = compare_map.get(comparison_division_label, "03 OLP")
    out["DELTA"] = to_num(out["05 OLR"]).fillna(0) - to_num(out[compare_col]).fillna(0)
    return out[ordered_cols + ["DELTA"]].sort_values(["DELTA", compare_col, "05 OLR"], ascending=[True, False, True]).reset_index(drop=True)

def render_main_table_dynamic(df: pd.DataFrame, comparison_division_label: str):
    display_df = df.copy()

    compare_cols = ["03 OLP", "04 MOD", "05 OLR"]
    stock_hidden_map = {"03 OLP": "STOK_DIV03", "04 MOD": "STOK_DIV04", "05 OLR": "STOK_DIV05"}
    selected_stock_hidden = "STOK_DIV05"

    def losing_division(row):
        current_val = row.get("05 OLR", 0)
        compare_val = row.get(comparison_division_label, 0)
        try:
            return float(current_val) < float(compare_val)
        except Exception:
            return False

    def stok_problem(row):
        try:
            stok_selected = float(row.get("STOK", 0))
            qty_selected = float(row.get("05 OLR", 0))
            other_stock_cols = [stock_hidden_map[c] for c in compare_cols if c != "05 OLR"]
            other_stock_values = [float(row.get(c, 0)) for c in other_stock_cols]
            cond_a = stok_selected < qty_selected
            cond_b = stok_selected == 0 and qty_selected > 0 and any(v > 0 for v in other_stock_values)
            return cond_a or cond_b
        except Exception:
            return False

    display_df["_LOSS_DIVISION_FLAG"] = display_df.apply(losing_division, axis=1)
    display_df["_STOK_ALERT_FLAG"] = display_df.apply(stok_problem, axis=1)

    visible_df = display_df[["KODEBARANG", "PRODUCT", "BRAND", "SPESIFIKASI", "M3", "03 OLP", "04 MOD", "05 OLR", "DELTA", "STOK"]].copy()
    visible_df["M3"] = pd.to_numeric(visible_df["M3"], errors="coerce").fillna(0).round(0).astype(int)
    for col in ["03 OLP", "04 MOD", "05 OLR", "DELTA", "STOK"]:
        numeric_col = pd.to_numeric(visible_df[col], errors="coerce").fillna(0).round(0)
        visible_df[col] = numeric_col.astype(int)

    def fmt_value(val, col_name):
        if col_name in ["M3", "03 OLP", "04 MOD", "05 OLR", "DELTA", "STOK"]:
            try:
                return f"{int(val)}"
            except Exception:
                return "0"
        return "" if pd.isna(val) else str(val)

    html = []
    html.append('<div class="main-fixed-wrap"><table class="main-fixed"><thead><tr>')
    for col in visible_df.columns:
        html.append(f"<th>{col}</th>")
    html.append("</tr></thead><tbody>")

    for idx, row in visible_df.iterrows():
        html.append("<tr>")
        original = display_df.loc[idx]
        for col in visible_df.columns:
            cls = ""
            if col == "05 OLR" and bool(original["_LOSS_DIVISION_FLAG"]):
                cls = ' class="bg-red"'
            if col == "DELTA":
                try:
                    if float(original.get("DELTA", 0)) < 0:
                        cls = ' class="bg-red"'
                except Exception:
                    pass
            if col == "STOK" and bool(original["_STOK_ALERT_FLAG"]):
                cls = ' class="bg-red"'
            html.append(f"<td{cls}>{fmt_value(row[col], col)}</td>")
        html.append("</tr>")

    html.append("</tbody></table></div>")
    st.markdown("".join(html), unsafe_allow_html=True)
    return visible_df

# =========================================================
# UI
# =========================================================
st.title("Dashboard Analisa Produk")

st.markdown('<div class="upload-card-wrap">', unsafe_allow_html=True)
st.markdown("### Upload File")

upload_col1, upload_col2, upload_col3 = st.columns(3)
with upload_col1:
    st.markdown("**Upload MPLSSR**")
    mplssr_file = st.file_uploader("", type=["xlsx", "xls"], key="upload_mplssr_main", label_visibility="collapsed")
    st.caption("200MB per file • XLSX, XLS")

with upload_col2:
    st.markdown("**Upload Pricelist**")
    pricelist_file = st.file_uploader("", type=["xlsx", "xls"], key="upload_pricelist_main", label_visibility="collapsed")
    st.caption("200MB per file • XLSX, XLS")

with upload_col3:
    st.markdown("**Upload SALES PIVOT**")
    sales_pivot_file = st.file_uploader("", type=["xlsx", "xls"], key="upload_sales_pivot_main", label_visibility="collapsed")
    st.caption("200MB per file • XLSX, XLS")

all_required_uploaded = all([mplssr_file is not None, pricelist_file is not None, sales_pivot_file is not None])

process_upload = st.button(
    "PROSES FILE",
    type="primary",
    
    disabled=not all_required_uploaded,
)

if not all_required_uploaded:
    st.info("Silakan upload MPLSSR, Pricelist, dan SALES PIVOT dulu, lalu klik PROSES FILE.")
elif not process_upload and "processed_data" not in st.session_state:
    st.info("Semua file sudah di-upload. Klik PROSES FILE untuk generate dashboard.")
st.markdown('</div>', unsafe_allow_html=True)

if process_upload:
    try:
        sales = load_mplssr(mplssr_file)
        stock = load_pricelist(pricelist_file)
        master = build_master(sales, stock)
        pricelist_wh, warehouse_stock_cols = load_pricelist_with_warehouses(pricelist_file)
        sales_pivot = load_sales_pivot(sales_pivot_file)
        st.session_state["processed_data"] = {
            "sales": sales,
            "stock": stock,
            "master": master,
            "pricelist_wh": pricelist_wh,
            "warehouse_stock_cols": warehouse_stock_cols,
            "sales_pivot": sales_pivot,
        }
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

if "processed_data" not in st.session_state:
    st.stop()

sales = st.session_state["processed_data"]["sales"]
stock = st.session_state["processed_data"]["stock"]
master = st.session_state["processed_data"]["master"]
pricelist_wh = st.session_state["processed_data"]["pricelist_wh"]
warehouse_stock_cols = st.session_state["processed_data"]["warehouse_stock_cols"]
sales_pivot = st.session_state["processed_data"]["sales_pivot"]

product_options = sorted(master["PRODUCT_FINAL"].dropna().unique().tolist())
default_product = ["LAPTOP R"] if "LAPTOP R" in product_options else []

st.markdown('<div class="upload-card-wrap">', unsafe_allow_html=True)
st.markdown("### Filter Dashboard")
with st.form("unified_filter_form"):
    fcol1, fcol2, fcol3, fcol4, fcol5, fcol6 = st.columns([1.2, 1.2, 1, 1.2, 1.1, 0.6])
    with fcol1:
        selected_products = st.multiselect("Product", product_options, default=default_product)
    with fcol2:
        selected_brands = st.multiselect("Brand", sorted(master["BRAND"].dropna().unique().tolist()))
    with fcol3:
        selected_period = st.selectbox("Period", PERIODS, index=0)
    with fcol4:
        selected_segments = st.multiselect("Range Harga", [s[2] for s in PRICE_SEGMENTS] + ["UNKNOWN"])
    with fcol5:
        comparison_division = st.selectbox("Perbandingan", ["03 OLP", "04 MOD", "05 OLR"], index=1)
    with fcol6:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        apply_filter = st.form_submit_button("PROSES", use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

filtered = master.copy()
if selected_products:
    filtered = filtered[filtered["PRODUCT_FINAL"].isin(selected_products)]
if selected_brands:
    filtered = filtered[filtered["BRAND"].isin(selected_brands)]
if selected_segments:
    filtered = filtered[filtered["PRICE"].apply(price_segment).isin(selected_segments)]

if filtered.empty:
    st.warning("Data kosong setelah filter diterapkan.")
    st.stop()

left, right = st.columns(2)
with left:
    render_left_table(build_segment_table(filtered, selected_period, comparison_division), f"Segmentasi Harga - {selected_period}", selected_division=comparison_division)
with right:
    render_left_table(build_brand_table(filtered, selected_period, comparison_division), f"Segmentasi Brand - {selected_period}", selected_division=comparison_division)

sales_pivot_alerts = build_sales_pivot_alerts(
    sales_pivot,
    pricelist_wh,
    warehouse_stock_cols,
    period=selected_period,
    selected_products=selected_products,
    selected_brands=selected_brands,
    selected_segments=selected_segments,
)

st.markdown("### Tabel Utama Analisa")
main_table_export = build_main_table_filtered(
    filtered,
    selected_period,
    comparison_division,
    selected_segments=selected_segments,
    selected_brands=selected_brands,
    selected_products=selected_products,
)
main_table_export = render_main_table_dynamic(main_table_export, comparison_division)

st.markdown("### Analisa Stok")
render_sales_pivot_alert_table(sales_pivot_alerts)

st.markdown("<div style='height:120px;'></div>", unsafe_allow_html=True)
