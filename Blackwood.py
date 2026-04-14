
import io
from typing import List

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dashboard Analisa Sales vs Stock", layout="wide")

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
    (1_000_000, 1_500_000, "1 – 1.5 JUTA"),
    (1_500_000, 2_000_000, "1.5 – 2 JUTA"),
    (2_000_000, 2_500_000, "2 – 2.5 JUTA"),
    (2_500_000, 3_000_000, "2.5 – 3 JUTA"),
    (3_000_000, 4_000_000, "3 – 4 JUTA"),
    (4_000_000, 5_000_000, "4 – 5 JUTA"),
    (5_000_000, 7_000_000, "5 – 7 JUTA"),
    (7_000_000, 10_000_000, "7 – 10 JUTA"),
    (10_000_000, 12_500_000, "10 – 12.5 JUTA"),
    (12_500_000, 15_000_000, "12.5 – 15 JUTA"),
    (15_000_000, 20_000_000, "15 – 20 JUTA"),
    (20_000_000, 25_000_000, "20 – 25 JUTA"),
    (25_000_000, 30_000_000, "25 – 30 JUTA"),
    (30_000_000, 40_000_000, "30 – 40 JUTA"),
    (40_000_000, float("inf"), "40 JUTA – UP"),
]


st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 1rem;}
    .main-table-wrap {overflow-x:auto; border:1px solid #b7b7b7; background:#fff;}
    table.report {border-collapse:collapse; width:max-content; min-width:100%;}
    table.report th, table.report td {
        border:1px solid #2b2b2b;
        padding:4px 6px;
        font-size:12px;
        white-space:nowrap;
    }
    table.report th {
        background:#9fc5e8;
        text-align:center;
        font-weight:700;
    }
    table.report td:nth-child(1),
    table.report td:nth-child(2),
    table.report td:nth-child(3) {
        text-align:left;
    }
    table.report td:not(:nth-child(1)):not(:nth-child(2)):not(:nth-child(3)) {
        text-align:right;
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

SEGMENT_ORDER = [label for _, _, label in PRICE_SEGMENTS]

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

def area_code_matches(value, prefixes: List[str]) -> bool:
    if pd.isna(value):
        return False
    txt = str(value).strip().upper().replace(" ", "")
    return any(txt.startswith(p) for p in prefixes)

def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    cols = []
    for col in out.columns:
        if isinstance(col, tuple):
            parts = [str(x) for x in col if str(x) not in ["", "nan", "None"]]
            cols.append("|".join(parts))
        else:
            cols.append(str(col))
    out.columns = cols
    return out

# =========================================================
# MPLSSR
# =========================================================
def load_mplssr(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="ALL", header=1)
    df = df.iloc[4:].copy().reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

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
    row2 = _ffill_header(raw.iloc[2].tolist())
    row3 = [str(x).strip().upper() if pd.notna(x) and str(x).strip() != "" else None for x in raw.iloc[3].tolist()]

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

    for c in ["PRODUCT", "KODEBARANG", "SPESIFIKASI", "TOT", "M3"]:
        if c not in df.columns:
            df[c] = np.nan

    df["PRODUCT"] = normalize_text(df["PRODUCT"])
    df["KODEBARANG"] = normalize_text(df["KODEBARANG"])
    df["SPESIFIKASI"] = normalize_text(df["SPESIFIKASI"])

    df = df[df["KODEBARANG"].notna()].copy()
    df = df[~df["KODEBARANG"].isin(["TOTAL"])]

    stock03_cols = [columns[i] for i, area in enumerate(row3) if area_code_matches(area, ["3", "03"])]
    stock04_cols = [columns[i] for i, area in enumerate(row3) if area_code_matches(area, ["4", "04"])]
    stock05_cols = [columns[i] for i, area in enumerate(row3) if area_code_matches(area, ["5", "05"])]

    df["PRICE"] = to_num(df["M3"]) * 1000
    df["STOK_DIV03"] = df[stock03_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock03_cols else 0
    df["STOK_DIV04"] = df[stock04_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock04_cols else 0
    df["STOK_DIV05"] = df[stock05_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock05_cols else 0
    df["CATEGORY"] = sheet_name.upper()
    df["PRICE_SEGMENT"] = df["PRICE"].apply(price_segment)
    df["MERGE_KEY"] = df["KODEBARANG"]

    return df[[
        "PRODUCT", "KODEBARANG", "SPESIFIKASI", "PRICE",
        "STOK_DIV03", "STOK_DIV04", "STOK_DIV05",
        "CATEGORY", "PRICE_SEGMENT", "MERGE_KEY"
    ]]

def load_pricelist(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    sheets = [s for s in xls.sheet_names if s.upper() in VALID_PRICELIST_SHEETS]
    frames = [parse_pricelist_sheet(xls, s) for s in sheets]
    if not frames:
        return pd.DataFrame(columns=[
            "PRODUCT", "KODEBARANG", "SPESIFIKASI", "PRICE",
            "STOK_DIV03", "STOK_DIV04", "STOK_DIV05",
            "CATEGORY", "PRICE_SEGMENT", "MERGE_KEY"
        ])
    out = pd.concat(frames, ignore_index=True)
    out = out.drop_duplicates(subset=["MERGE_KEY"], keep="first")
    return out

# =========================================================
# BUILD TABLE
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

def build_main_table(filtered: pd.DataFrame) -> pd.DataFrame:
    qty = (
        filtered.pivot_table(
            index=["KODEBARANG", "SPESIFIKASI_FINAL", "PRICE"],
            columns=["PERIOD", "DIVISION"],
            values="QTY",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )
    qty = flatten_columns(qty)

    stock_src = (
        filtered[["KODEBARANG", "SPESIFIKASI_FINAL", "PRICE", "DIVISION", "STOK_DIVISI"]]
        .drop_duplicates()
        .copy()
    )
    stock_piv = (
        stock_src.pivot_table(
            index=["KODEBARANG", "SPESIFIKASI_FINAL", "PRICE"],
            columns=["DIVISION"],
            values="STOK_DIVISI",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )
    stock_piv = flatten_columns(stock_piv)

    out = qty.merge(stock_piv, how="outer", on=["KODEBARANG", "SPESIFIKASI_FINAL", "PRICE"])
    out = out.fillna(0)

    final = pd.DataFrame()
    final["KODEBARANG"] = out["KODEBARANG"]
    final["SPESIFIKASI"] = out["SPESIFIKASI_FINAL"]
    final["M3"] = (to_num(out["PRICE"]).fillna(0) / 1000).round(0)

    div_name = {"DIV03": "03 OLP", "DIV04": "04 MOD", "DIV05": "05 OLR"}

    for period in PERIODS:
        period_label = {"7DAY": "7 DAY ANALISA", "14DAY": "14 DAY ANALISA", "30DAY": "30 DAY ANALISA"}[period]
        for div in DIVISIONS:
            qty_col = f"{period}|{div}"
            if qty_col in out.columns:
                final[f"{period_label}|{div_name[div]}|QTY"] = to_num(out[qty_col]).fillna(0)
            else:
                final[f"{period_label}|{div_name[div]}|QTY"] = 0

            if div in out.columns:
                final[f"{period_label}|{div_name[div]}|STOK"] = to_num(out[div]).fillna(0)
            else:
                final[f"{period_label}|{div_name[div]}|STOK"] = 0

    final = final.sort_values(["KODEBARANG", "SPESIFIKASI"], ascending=[True, True]).reset_index(drop=True)
    return final

def render_exact_header_table(df: pd.DataFrame):
    def fmt_num(v):
        try:
            v = float(v)
            if np.isnan(v):
                return ""
            return f"{int(round(v)):,}".replace(",", ".")
        except Exception:
            return str(v)

    periods = ["7 DAY ANALISA", "14 DAY ANALISA", "30 DAY ANALISA"]
    divs = ["03 OLP", "04 MOD", "05 OLR"]

    html = []
    html.append('<div class="main-table-wrap"><table class="report">')
    html.append("<thead>")
    html.append(
        "<tr>"
        '<th rowspan="3">KODEBARANG</th>'
        '<th rowspan="3">SPESIFIKASI</th>'
        '<th rowspan="3">M3</th>'
        '<th colspan="6"></th>'
        '<th colspan="6"></th>'
        '<th colspan="6"></th>'
        "</tr>"
    )
    html.append("<tr>")
    for _ in periods:
        for div in divs:
            html.append(f'<th colspan="2">{div}</th>')
    html.append("</tr>")

    html.append("<tr>")
    for _ in periods:
        for _ in divs:
            html.append("<th>QTY</th><th>STOK</th>")
    html.append("</tr>")

    html.append("<tr>")
    html.append("<th></th><th></th><th></th>")
    for period in periods:
        html.append(f'<th colspan="6">{period}</th>')
    html.append("</tr>")
    html.append("</thead>")

    html.append("<tbody>")
    for _, row in df.iterrows():
        html.append("<tr>")
        html.append(f"<td>{row['KODEBARANG'] if pd.notna(row['KODEBARANG']) else ''}</td>")
        html.append(f"<td>{row['SPESIFIKASI'] if pd.notna(row['SPESIFIKASI']) else ''}</td>")
        html.append(f"<td>{fmt_num(row['M3'])}</td>")
        for period in periods:
            for div in divs:
                html.append(f"<td>{fmt_num(row.get(f'{period}|{div}|QTY', 0))}</td>")
                html.append(f"<td>{fmt_num(row.get(f'{period}|{div}|STOK', 0))}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def render_main_table_simple(df: pd.DataFrame, master_df: pd.DataFrame):
    display_df = df.copy()

    product_brand = (
        master_df[["KODEBARANG", "PRODUCT_FINAL", "BRAND"]]
        .dropna(subset=["KODEBARANG"])
        .drop_duplicates(subset=["KODEBARANG"], keep="first")
        .copy()
    )
    display_df = display_df.merge(product_brand, how="left", on="KODEBARANG")
    display_df = display_df.rename(columns={
        "PRODUCT_FINAL": "PRODUCT",
        "BRAND": "BRAND",
    })

    ordered_front = ["KODEBARANG", "PRODUCT", "BRAND", "SPESIFIKASI", "M3"]
    other_cols = [c for c in display_df.columns if c not in ordered_front]
    display_df = display_df[ordered_front + other_cols]

    st.dataframe(
        display_df,
        use_container_width=True,
        height=520,
    )

    return display_df


def build_main_table_filtered(
    df: pd.DataFrame,
    period: str,
    stock_division_label: str,
    selected_segments=None,
    selected_brands=None,
) -> pd.DataFrame:
    base = df[df["PERIOD"] == period].copy()

    if selected_segments:
        base = base[base["PRICE"].apply(price_segment).isin(selected_segments)]
    if selected_brands:
        base = base[base["BRAND"].isin(selected_brands)]

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
    stok_col = stock_label_map.get(stock_division_label, "STOK_DIV05")

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
        "BRAND": "BRAND",
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
    out = out[ordered_cols].sort_values(["KODEBARANG", "SPESIFIKASI"], ascending=[True, True]).reset_index(drop=True)
    return out

def render_main_table_dynamic(df: pd.DataFrame, selected_division_label: str, selected_stock_division_label: str):
    display_df = df.copy()

    compare_cols = ["03 OLP", "04 MOD", "05 OLR"]
    stock_hidden_map = {"03 OLP": "STOK_DIV03", "04 MOD": "STOK_DIV04", "05 OLR": "STOK_DIV05"}
    selected_stock_hidden = stock_hidden_map.get(selected_stock_division_label, "STOK_DIV05")

    def losing_division(row):
        current_val = row.get(selected_division_label, 0)
        other_vals = [row.get(c, 0) for c in compare_cols if c != selected_division_label]
        try:
            return any(float(current_val) < float(v) for v in other_vals)
        except Exception:
            return False

    def stok_problem(row):
        try:
            stok_selected = float(row.get("STOK", 0))
            qty_selected = float(row.get(selected_division_label, 0))
            other_stock_cols = [stock_hidden_map[c] for c in compare_cols if c != selected_stock_division_label]
            other_stock_values = [float(row.get(c, 0)) for c in other_stock_cols]
            cond_a = stok_selected < qty_selected
            cond_b = stok_selected == 0 and qty_selected > 0 and any(v > 0 for v in other_stock_values)
            return cond_a or cond_b
        except Exception:
            return False

    display_df["_LOSS_DIVISION_FLAG"] = display_df.apply(losing_division, axis=1)
    display_df["_STOK_ALERT_FLAG"] = display_df.apply(stok_problem, axis=1)

    visible_df = display_df[["KODEBARANG", "PRODUCT", "BRAND", "SPESIFIKASI", "M3", "03 OLP", "04 MOD", "05 OLR", "STOK"]].copy()

    visible_df["M3"] = pd.to_numeric(visible_df["M3"], errors="coerce").fillna(0).round(0).astype(int)
    for col in ["03 OLP", "04 MOD", "05 OLR", "STOK"]:
        numeric_col = pd.to_numeric(visible_df[col], errors="coerce").fillna(0).round(0)
        visible_df[col] = numeric_col.astype(int)

    def fmt_value(val, col_name):
        if col_name in ["M3", "03 OLP", "04 MOD", "05 OLR", "STOK"]:
            try:
                return f"{int(val)}"
            except Exception:
                return "0"
        return "" if pd.isna(val) else str(val)

    html = []
    html.append("""
    <style>
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
    table.main-fixed td:nth-child(5) { min-width: 70px; }
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
    </style>
    """)
    html.append('<div class="main-fixed-wrap"><table class="main-fixed"><thead><tr>')
    for col in visible_df.columns:
        html.append(f"<th>{col}</th>")
    html.append("</tr></thead><tbody>")

    for idx, row in visible_df.iterrows():
        html.append("<tr>")
        original = display_df.loc[idx]
        for col in visible_df.columns:
            cls = ""
            if col == selected_division_label and bool(original["_LOSS_DIVISION_FLAG"]):
                cls = ' class="bg-red"'
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


st.sidebar.header("Upload File")
mplssr_file = st.sidebar.file_uploader("Upload MPLSSR", type=["xlsx", "xls"])
pricelist_file = st.sidebar.file_uploader("Upload Pricelist", type=["xlsx", "xls"])


if not mplssr_file or not pricelist_file:
    st.info("Silakan upload file MPLSSR dan Pricelist.")
    st.stop()

try:
    sales = load_mplssr(mplssr_file)
    stock = load_pricelist(pricelist_file)
    master = build_master(sales, stock)
except Exception as e:
    st.error(f"Gagal membaca file: {e}")
    st.stop()

product_options = sorted(master["PRODUCT_FINAL"].dropna().unique().tolist())
default_product = ["LAPTOP R"] if "LAPTOP R" in product_options else []

with st.sidebar:
    st.markdown("---")
    with st.form("filter_form"):
        selected_products = st.multiselect("Product", product_options, default=default_product)
        selected_brands = st.multiselect("Brand", sorted(master["BRAND"].dropna().unique().tolist()))
        process_clicked = st.form_submit_button("PROSES", use_container_width=True)

if "filter_submitted" not in st.session_state:
    st.session_state["filter_submitted"] = True

if process_clicked:
    st.session_state["filter_submitted"] = True

filtered = master.copy()
if selected_products:
    filtered = filtered[filtered["PRODUCT_FINAL"].isin(selected_products)]
if selected_brands:
    filtered = filtered[filtered["BRAND"].isin(selected_brands)]

if filtered.empty:
    st.warning("Data kosong setelah filter diterapkan.")
    st.stop()


# =========================================================
# SEGMENTATION CARDS
# =========================================================

def build_segment_table(df, period):
    tmp = df[df["PERIOD"] == period].copy()
    tmp["SEGMENT"] = tmp["PRICE"].apply(price_segment)
    seg = tmp.groupby(["SEGMENT", "DIVISION"])["QTY"].sum().unstack().fillna(0).reset_index()
    for div in DIVISIONS:
        if div not in seg.columns:
            seg[div] = 0
    seg = seg[["SEGMENT", "DIV03", "DIV04", "DIV05"]].copy()
    seg = seg.sort_values("SEGMENT", key=lambda s: s.map(segment_sort_key)).reset_index(drop=True)
    seg.columns = ["SEGMENT", "03 OLP", "04 MOD", "05 OLR"]
    return seg

def build_brand_table(df, period):
    brand = df[df["PERIOD"] == period].copy()
    brand = brand.groupby(["BRAND", "DIVISION"])["QTY"].sum().unstack().fillna(0).reset_index()
    for div in DIVISIONS:
        if div not in brand.columns:
            brand[div] = 0
    brand = brand[["BRAND", "DIV03", "DIV04", "DIV05"]].copy()
    brand["TOTAL"] = brand[["DIV03", "DIV04", "DIV05"]].sum(axis=1)
    brand = brand.sort_values(["TOTAL", "BRAND"], ascending=[False, True]).head(10).drop(columns=["TOTAL"])
    brand.columns = ["BRAND", "03 OLP", "04 MOD", "05 OLR"]
    return brand

def render_left_table(df, title, selected_division="05 OLR"):
    def is_number(v):
        return isinstance(v, (int, float, np.integer, np.floating)) and not pd.isna(v)

    def fmt_number(v):
        return f"{int(round(float(v))):,}".replace(",", ".")

    html = []
    html.append("""
    <div style="border:1px solid #d9d9d9;border-radius:8px;background:#fff;padding:8px;">
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

            html.append(f'<td style="{style}">{display}</td>')
        html.append("</tr>")

    html.append("</tbody></table></div></div>")
    st.markdown("".join(html), unsafe_allow_html=True)

with st.form("segmentasi_form"):
    seg_filter_col1, seg_filter_col2, seg_filter_col3 = st.columns([1, 1, 0.5])
    with seg_filter_col1:
        segmentasi_period = st.selectbox(
            "Filter Segmentasi",
            PERIODS,
            index=0,
            key="segmentasi_period_top",
        )
    with seg_filter_col2:
        selected_division_segment = st.selectbox(
            "Filter Divisi",
            ["03 OLP", "04 MOD", "05 OLR"],
            index=2,
            key="segmentasi_division_top",
        )
    with seg_filter_col3:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        segmentasi_process = st.form_submit_button("PROSES")

if "segmentasi_submitted" not in st.session_state:
    st.session_state["segmentasi_submitted"] = True

if segmentasi_process:
    st.session_state["segmentasi_submitted"] = True

left, right = st.columns(2)

with left:
    render_left_table(
        build_segment_table(filtered, segmentasi_period),
        f"Segmentasi Harga - {segmentasi_period}",
        selected_division=selected_division_segment,
    )

with right:
    render_left_table(
        build_brand_table(filtered, segmentasi_period),
        f"Segmentasi Brand - {segmentasi_period}",
        selected_division=selected_division_segment,
    )


st.markdown("### Tabel Utama Analisa")

with st.form("main_table_form"):
    main_filter_col1, main_filter_col2, main_filter_col3, main_filter_col4, main_filter_col5, main_filter_col6 = st.columns([1, 1, 1, 1.4, 1.4, 0.6])
    with main_filter_col1:
        main_period = st.selectbox("Filter Period", PERIODS, index=0, key="main_period_filter")
    with main_filter_col2:
        main_division_label = st.selectbox("Filter Divisi", ["03 OLP", "04 MOD", "05 OLR"], index=2, key="main_division_filter")
    with main_filter_col3:
        main_stock_division = st.selectbox("Filter Stok Divisi", ["03 OLP", "04 MOD", "05 OLR"], index=2, key="main_stock_division_filter")
    with main_filter_col4:
        main_segment_filter = st.multiselect(
            "Filter Segmentasi",
            [s[2] for s in PRICE_SEGMENTS] + ["UNKNOWN"],
            key="main_segment_filter",
        )
    with main_filter_col5:
        main_brand_filter = st.multiselect(
            "Filter Brand",
            sorted(filtered["BRAND"].dropna().unique().tolist()),
            key="main_brand_filter",
        )
    with main_filter_col6:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        main_process = st.form_submit_button("PROSES")

if "main_table_submitted" not in st.session_state:
    st.session_state["main_table_submitted"] = True

if main_process:
    st.session_state["main_table_submitted"] = True

main_table_export = build_main_table_filtered(
    filtered,
    main_period,
    main_stock_division,
    selected_segments=main_segment_filter,
    selected_brands=main_brand_filter,
)
main_table_export = render_main_table_dynamic(main_table_export, main_division_label, main_stock_division)

out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    main_table_export.to_excel(writer, index=False, sheet_name="main_table")

st.download_button(
    "Download hasil analisa",
    data=out.getvalue(),
    file_name="dashboard_sales_stock_main_table.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Debug hasil parsing"):
    st.write("Sample MPLSSR:", sales.head(20))
    st.write("Sample Pricelist:", stock.head(20))
    st.write("Sample Master:", master.head(20))
    st.write("Sample Main Table:", main_table_export.head(20))
