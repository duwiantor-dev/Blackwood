
import io
from typing import Dict, List

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dashboard Analisa Sales vs Stock", layout="wide")

PRICE_SEGMENTS = [
    (0, 1_000_000, "< 1 JUTA"),
    (1_000_000, 1_500_000, "1 - 1.5 JUTA"),
    (1_500_000, 2_000_000, "1.5 - 2 JUTA"),
    (2_000_000, 2_500_000, "2 - 2.5 JUTA"),
    (2_500_000, 3_000_000, "2.5 - 3 JUTA"),
    (3_000_000, 4_000_000, "3 - 4 JUTA"),
    (4_000_000, 5_000_000, "4 - 5 JUTA"),
    (5_000_000, 10_000_000, "5 - 10 JUTA"),
    (10_000_000, np.inf, "10 JUTA - UP"),
]
DIVISIONS = ["DIV03", "DIV04", "DIV05"]
PERIODS = ["7DAY", "14DAY", "30DAY"]
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

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 1rem;}
    .stDataFrame {border-radius: 10px;}
    </style>
    """,
    unsafe_allow_html=True,
)

def normalize_text(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.upper().replace({"NAN": np.nan, "NONE": np.nan, "": np.nan})

def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")

def price_segment(price: float) -> str:
    if pd.isna(price):
        return "UNKNOWN"
    for low, high, label in PRICE_SEGMENTS:
        if low <= float(price) < high:
            return label
    return "UNKNOWN"

def first_row_contains_text(df: pd.DataFrame, text: str):
    target = str(text).strip().upper()
    for idx in range(len(df)):
        row = df.iloc[idx].astype(str).str.upper()
        if row.str.contains(target, na=False).any():
            return idx
    return None

def build_merge_key(*series_list: pd.Series) -> pd.Series:
    normalized = [normalize_text(s) for s in series_list]
    out = normalized[0].copy()
    for s in normalized[1:]:
        out = out.fillna(s)
    return out

def area_code_matches(value, prefixes: List[str]) -> bool:
    if pd.isna(value):
        return False
    txt = str(value).strip().upper().replace(" ", "")
    return any(txt.startswith(p) for p in prefixes)

def load_mplssr(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="ALL", header=1)
    df = df.iloc[4:].copy().reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    base_cols = ["PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI"]
    wanted_div_cols = [c for p in PERIODS for c in MPLSSR_DIV_COLS[p].values() if c in df.columns]

    for c in base_cols:
        if c not in df.columns:
            df[c] = np.nan

    df = df[base_cols + wanted_div_cols].copy()

    for c in base_cols:
        df[c] = normalize_text(df[c])

    df = df[df["KODE BARANG"].notna()].copy()
    df = df[~df["KODE BARANG"].isin(["TOTAL", "SHARE%"])]

    rows = []
    for period in PERIODS:
        for div, col in MPLSSR_DIV_COLS[period].items():
            if col not in df.columns:
                continue
            tmp = df[["PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI", col]].copy()
            tmp["QTY"] = to_num(tmp[col]).fillna(0)
            tmp["PERIOD"] = period
            tmp["DIVISION"] = div
            tmp["SKU NO"] = np.nan
            rows.append(tmp[["SKU NO", "PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI", "PERIOD", "DIVISION", "QTY"]])

    out = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(
        columns=["SKU NO", "PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI", "PERIOD", "DIVISION", "QTY"]
    )
    out["_MERGE_KEY"] = build_merge_key(out["KODE BARANG"], out["SKU NO"])
    return out

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

def parse_pricelist_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    raw = xls.parse(sheet_name=sheet_name, header=None)
    raw = raw.iloc[:, : raw.shape[1]].copy()

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

    for c in ["SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "TOT", "M3", "SRP"]:
        if c not in df.columns:
            df[c] = np.nan

    stock03_cols = [columns[i] for i, area in enumerate(row3) if area_code_matches(area, ["3"])]
    stock04_cols = [columns[i] for i, area in enumerate(row3) if area_code_matches(area, ["4"])]
    stock05_cols = [columns[i] for i, area in enumerate(row3) if area_code_matches(area, ["5", "05"])]

    df["SKU NO"] = normalize_text(df["SKU NO"])
    df["PRODUCT"] = normalize_text(df["PRODUCT"])
    df["KODEBARANG"] = normalize_text(df["KODEBARANG"])
    df["SPESIFIKASI"] = normalize_text(df["SPESIFIKASI"])
    df = df[df["KODEBARANG"].notna()].copy()
    df = df[~df["KODEBARANG"].isin(["TOTAL"])]

    df["PRICE"] = to_num(df["M3"]) * 1000
    df["STOK_TOTAL"] = to_num(df["TOT"]).fillna(0)
    df["STOK_DIV03"] = df[stock03_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock03_cols else 0
    df["STOK_DIV04"] = df[stock04_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock04_cols else 0
    df["STOK_DIV05"] = df[stock05_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock05_cols else 0
    df["CATEGORY"] = sheet_name.upper()
    df["PRICE_SEGMENT"] = df["PRICE"].apply(price_segment)
    df["_MERGE_KEY"] = build_merge_key(df["KODEBARANG"], df["SKU NO"])

    return df[[
        "SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "PRICE", "PRICE_SEGMENT",
        "STOK_TOTAL", "STOK_DIV03", "STOK_DIV04", "STOK_DIV05", "CATEGORY", "_MERGE_KEY"
    ]]

def load_pricelist(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    sheets = [s for s in xls.sheet_names if s.upper() in VALID_PRICELIST_SHEETS]
    frames = [parse_pricelist_sheet(xls, s) for s in sheets]
    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=[
            "SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "PRICE", "PRICE_SEGMENT",
            "STOK_TOTAL", "STOK_DIV03", "STOK_DIV04", "STOK_DIV05", "CATEGORY", "_MERGE_KEY"
        ]
    )
    out = out.drop_duplicates(subset=["_MERGE_KEY"], keep="first")
    return out

def build_master(sales: pd.DataFrame, stock: pd.DataFrame) -> pd.DataFrame:
    df = sales.merge(stock, how="left", on="_MERGE_KEY")
    for col in [
        "PRODUCT_y", "SPESIFIKASI_y", "CATEGORY", "PRICE", "PRICE_SEGMENT",
        "STOK_TOTAL", "STOK_DIV03", "STOK_DIV04", "STOK_DIV05", "KODEBARANG"
    ]:
        if col not in df.columns:
            df[col] = np.nan

    empty_series = pd.Series(index=df.index, dtype=object)
    df["SKU NO"] = build_merge_key(df.get("SKU NO_x", empty_series), df.get("SKU NO_y", empty_series))
    df["KODEBARANG_FINAL"] = build_merge_key(df.get("KODE BARANG", empty_series), df.get("KODEBARANG", empty_series), df["SKU NO"])
    df["PRODUCT_FINAL"] = df["PRODUCT_x"].fillna(df.get("PRODUCT_y"))
    df["SPEC_FINAL"] = df["SPESIFIKASI_x"].fillna(df.get("SPESIFIKASI_y"))
    df["CATEGORY"] = df["CATEGORY"].fillna(df["PRODUCT_FINAL"])
    df["PRICE_SEGMENT"] = df["PRICE_SEGMENT"].fillna("UNKNOWN")
    df["PRICE"] = to_num(df["PRICE"])
    df["QTY"] = to_num(df["QTY"]).fillna(0)

    stock_map = {"DIV03": "STOK_DIV03", "DIV04": "STOK_DIV04", "DIV05": "STOK_DIV05"}
    df["STOK_DIVISI"] = df.apply(lambda r: r.get(stock_map.get(r["DIVISION"]), 0), axis=1)
    return df

def build_main_table(df: pd.DataFrame) -> pd.DataFrame:
    base = df.copy()

    qty_piv = (
        base.pivot_table(
            index=["KODEBARANG_FINAL", "SPEC_FINAL", "PRICE"],
            columns=["PERIOD", "DIVISION"],
            values="QTY",
            aggfunc="sum",
            fill_value=0,
        )
    )

    stock_base = (
        base[["KODEBARANG_FINAL", "SPEC_FINAL", "PRICE", "DIVISION", "STOK_DIVISI"]]
        .drop_duplicates()
        .copy()
    )
    stock_base["dummy_period"] = "ALL"

    stock_piv = (
        stock_base.pivot_table(
            index=["KODEBARANG_FINAL", "SPEC_FINAL", "PRICE"],
            columns=["dummy_period", "DIVISION"],
            values="STOK_DIVISI",
            aggfunc="sum",
            fill_value=0,
        )
    )

    out = qty_piv.copy() if not qty_piv.empty else pd.DataFrame(index=stock_piv.index)
    if not stock_piv.empty:
        out = out.join(stock_piv, how="outer")

    out = out.fillna(0).reset_index()

    rename_map = {
        "KODEBARANG_FINAL": "KODEBARANG",
        "SPEC_FINAL": "SPESIFIKASI",
        "PRICE": "M3",
    }

    if isinstance(out.columns, pd.MultiIndex):
        new_cols = []
        for col in out.columns:
            if not isinstance(col, tuple):
                new_cols.append(rename_map.get(col, col))
                continue

            if len(col) == 3:
                metric, period, div = col
                if metric == "QTY":
                    period_label = {"7DAY": "7 DAY ANALISA", "14DAY": "14 DAY ANALISA", "30DAY": "30 DAY ANALISA"}.get(period, period)
                    div_label = {"DIV03": "03 OLP", "DIV04": "04 MOD", "DIV05": "05 OLR"}.get(div, div)
                    new_cols.append(f"{period_label}|{div_label}|QTY")
                elif metric == "STOK_DIVISI":
                    div_label = {"DIV03": "03 OLP", "DIV04": "04 MOD", "DIV05": "05 OLR"}.get(div, div)
                    new_cols.append(f"ALL|{div_label}|STOK")
                else:
                    new_cols.append("|".join([str(x) for x in col if str(x) != ""]))
            elif len(col) == 2:
                a, b = col
                if a == "STOK_DIVISI":
                    div_label = {"DIV03": "03 OLP", "DIV04": "04 MOD", "DIV05": "05 OLR"}.get(b, b)
                    new_cols.append(f"ALL|{div_label}|STOK")
                elif a in rename_map:
                    new_cols.append(rename_map.get(a, a))
                else:
                    new_cols.append("|".join([str(x) for x in col if str(x) != ""]))
            else:
                flat = "|".join([str(x) for x in col if str(x) != ""])
                new_cols.append(rename_map.get(flat, flat))
        out.columns = new_cols

    for period_label in ["7 DAY ANALISA", "14 DAY ANALISA", "30 DAY ANALISA"]:
        for div_label in ["03 OLP", "04 MOD", "05 OLR"]:
            qty_col = f"{period_label}|{div_label}|QTY"
            stok_all_col = f"ALL|{div_label}|STOK"
            final_stok_col = f"{period_label}|{div_label}|STOK"

            if qty_col not in out.columns:
                out[qty_col] = 0
            if stok_all_col not in out.columns:
                out[stok_all_col] = 0
            out[final_stok_col] = out[stok_all_col]

    drop_cols = [c for c in out.columns if str(c).startswith("ALL|")]
    if drop_cols:
        out = out.drop(columns=drop_cols)

    ordered_cols = ["KODEBARANG", "SPESIFIKASI", "M3"]
    for period_label in ["7 DAY ANALISA", "14 DAY ANALISA", "30 DAY ANALISA"]:
        for div_label in ["03 OLP", "04 MOD", "05 OLR"]:
            ordered_cols.extend([
                f"{period_label}|{div_label}|QTY",
                f"{period_label}|{div_label}|STOK",
            ])

    for col in ordered_cols:
        if col not in out.columns:
            out[col] = 0 if col != "KODEBARANG" and col != "SPESIFIKASI" else ""

    out = out[ordered_cols].copy()
    out["M3"] = to_num(out["M3"]).fillna(0) / 1000
    out = out.sort_values(["KODEBARANG", "SPESIFIKASI"], ascending=[True, True]).reset_index(drop=True)
    return out

st.title("Dashboard Analisa Sales vs Stock")
st.caption("QTY diambil dari MPLSSR. STOK dan harga diambil dari Pricelist. Fokus awal dibuat seperti dashboard pada contoh.")

st.sidebar.header("Upload File")
mplssr_file = st.sidebar.file_uploader("Upload MPLSSR", type=["xlsx", "xls"])
pricelist_file = st.sidebar.file_uploader("Upload Pricelist", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.write("**Rules:**")
st.sidebar.caption("""- QTY: MPLSSR
- STOK: Pricelist
- Harga: kolom M3
- STOK DIV05: berdasarkan kode area row 3
- Sheet LAPTOP: hapus blok COMING sampai END COMING
- Setelah 2 file ter-upload, tabel langsung tampil otomatis""")

if not mplssr_file or not pricelist_file:
    st.info("Silakan upload file MPLSSR dan Pricelist untuk menampilkan dashboard.")
    st.stop()

try:
    sales = load_mplssr(mplssr_file)
    stock = load_pricelist(pricelist_file)
    master = build_master(sales, stock)
except Exception as e:
    st.error(f"Gagal membaca file: {e}")
    st.stop()

default_product = ["LAPTOP R"] if "LAPTOP R" in master["PRODUCT_FINAL"].dropna().unique().tolist() else []

with st.sidebar:
    st.markdown("---")
    metric_type = st.selectbox("Metric utama", ["QTY", "STOK"], index=0)
    selected_products = st.multiselect(
        "Product",
        sorted(master["PRODUCT_FINAL"].dropna().unique().tolist()),
        default=default_product,
    )
    selected_brands = st.multiselect("Brand", sorted(master["BRAND"].dropna().unique().tolist()))
    selected_periods = st.multiselect("Periode", PERIODS, default=PERIODS)
    selected_segments = st.multiselect("Segment Harga", [s[2] for s in PRICE_SEGMENTS] + ["UNKNOWN"])

filtered = master.copy()
if selected_products:
    filtered = filtered[filtered["PRODUCT_FINAL"].isin(selected_products)]
if selected_brands:
    filtered = filtered[filtered["BRAND"].isin(selected_brands)]
if selected_periods:
    filtered = filtered[filtered["PERIOD"].isin(selected_periods)]
if selected_segments:
    filtered = filtered[filtered["PRICE_SEGMENT"].isin(selected_segments)]

if filtered.empty:
    st.warning("Data kosong setelah filter diterapkan.")
    st.stop()

st.markdown("### Tabel Utama Analisa")

main_table = build_main_table(filtered)

period_order = ["7 DAY ANALISA", "14 DAY ANALISA", "30 DAY ANALISA"]
div_order = ["03 OLP", "04 MOD", "05 OLR"]

column_config = {
    "KODEBARANG": st.column_config.TextColumn("KODEBARANG", width="medium"),
    "SPESIFIKASI": st.column_config.TextColumn("SPESIFIKASI", width="large"),
    "M3": st.column_config.NumberColumn("M3", format="%.0f", width="small"),
}

for period_label in period_order:
    for div_label in div_order:
        for metric_label in ["QTY", "STOK"]:
            col_name = f"{period_label}|{div_label}|{metric_label}"
            column_config[col_name] = st.column_config.NumberColumn(
                f"{div_label}\n{metric_label}",
                format="%.0f",
                width="small",
            )

st.dataframe(
    main_table,
    use_container_width=True,
    height=520,
    column_config=column_config,
)

out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    main_table.to_excel(writer, index=False, sheet_name="main_table")

st.download_button(
    "Download hasil analisa",
    data=out.getvalue(),
    file_name="dashboard_sales_stock_main_table.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Debug hasil parsing"):
    st.write("Sales:", sales.head(20))
    st.write("Stock:", stock.head(20))
    st.write("Master:", master.head(20))
    st.write("Main Table:", main_table.head(20))
