import io
from dataclasses import dataclass
from typing import Optional, List

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Analisa Penjualan Divisi", layout="wide")

# =========================================================
# Helpers
# =========================================================
PRICE_SEGMENTS = [
    (0, 1_000_000, "< 1 JUTA"),
    (1_000_000, 1_500_000, "1 – 1.5 JUTA"),
    (1_500_000, 2_000_000, "1.5 – 2 JUTA"),
    (2_000_000, 2_500_000, "2 – 2.5 JUTA"),
    (2_500_000, 3_000_000, "2.5 – 3 JUTA"),
    (3_000_000, 4_000_000, "3 – 4 JUTA"),
    (4_000_000, 5_000_000, "4 – 5 JUTA"),
    (5_000_000, 10_000_000, "5 – 10 JUTA"),
    (10_000_000, np.inf, "10 JUTA – UP"),
]
DIVISIONS = ["DIV03", "DIV04", "DIV05"]
PERIODS = ["7DAY", "14DAY", "30DAY"]


def find_segment(price: float) -> str:
    if pd.isna(price):
        return "UNKNOWN"
    for low, high, label in PRICE_SEGMENTS:
        if low <= price < high:
            return label
    return "UNKNOWN"


def safe_read_excel(uploaded_file) -> dict:
    xl = pd.ExcelFile(uploaded_file)
    return {sheet: xl.parse(sheet) for sheet in xl.sheet_names}


def normalize_text(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.strip()
        .str.upper()
        .replace({"NAN": np.nan, "NONE": np.nan})
    )


def find_sheet_by_keyword(sheets: dict, keyword: str) -> Optional[pd.DataFrame]:
    keyword = keyword.upper()
    for name, df in sheets.items():
        if keyword in name.upper():
            out = df.copy()
            out["_sheet_name"] = name
            return out
    return None


def money_fmt(v: float) -> str:
    if pd.isna(v):
        return "-"
    return f"Rp {v:,.0f}".replace(",", ".")


def qty_fmt(v: float) -> str:
    if pd.isna(v):
        return "0"
    return f"{v:,.0f}".replace(",", ".")


def growth_fmt(cur: float, base: float) -> str:
    if pd.isna(base) or base == 0:
        return "-"
    return f"{((cur - base) / base) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


@dataclass
class Mapping:
    product: str
    brand: str
    specification: str
    division: str
    qty: str
    price: str
    sku: Optional[str] = None


# =========================================================
# UI Header
# =========================================================
st.title("📊 Web App Analisa Penjualan Div03 / Div04 / Div05")
st.caption(
    "Upload file MPLSSR dan Pricelist. App ini membandingkan QTY penjualan per divisi, per segmen harga, brand, dan spesifikasi."
)

# =========================================================
# Sidebar Upload
# =========================================================
st.sidebar.header("Upload Data")
mplssr_file = st.sidebar.file_uploader("Upload MPLSSR", type=["xlsx", "xls"])
pricelist_file = st.sidebar.file_uploader("Upload Pricelist", type=["xlsx", "xls"])

with st.sidebar.expander("Aturan Analisa 7/14/30 Day", expanded=True):
    st.markdown(
        """
        - **7DAY** = alert cepat  
        - **30DAY** = baseline  
        - **14DAY** = validasi tren  
        
        **Interpretasi:**
        - Jika **7-day naik** tapi **30-day stagnan** → kemungkinan efek promo
        - Jika **7-day & 30-day sama-sama naik** → growth lebih real
        """
    )

if not mplssr_file or not pricelist_file:
    st.info("Silakan upload 2 file Excel terlebih dahulu: MPLSSR dan Pricelist.")
    st.stop()

# =========================================================
# Read Files
# =========================================================
try:
    mplssr_sheets = safe_read_excel(mplssr_file)
    pricelist_sheets = safe_read_excel(pricelist_file)
except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")
    st.stop()

st.sidebar.success(f"MPLSSR sheets: {', '.join(mplssr_sheets.keys())}")
st.sidebar.success(f"Pricelist sheets: {', '.join(pricelist_sheets.keys())}")

# =========================================================
# Auto detect sheets by period
# =========================================================
period_sources = {}
for p in PERIODS:
    period_sources[p] = find_sheet_by_keyword(mplssr_sheets, p)

st.subheader("1) Mapping Kolom Data")

sample_period = next((df for df in period_sources.values() if df is not None), None)
if sample_period is None:
    st.error("Tidak ditemukan sheet 7DAY / 14DAY / 30DAY pada file MPLSSR. Pastikan nama sheet mengandung kata tersebut.")
    st.stop()

sample_cols = list(sample_period.columns)
price_sample_df = next(iter(pricelist_sheets.values())).copy()
price_cols = list(price_sample_df.columns)

c1, c2 = st.columns(2)
with c1:
    st.markdown("**Mapping MPLSSR**")
    map_product = st.selectbox("Kolom Product", sample_cols, index=sample_cols.index(sample_cols[0]))
    map_brand = st.selectbox("Kolom Brand", sample_cols, index=min(1, len(sample_cols) - 1))
    map_spec = st.selectbox("Kolom Spesifikasi", sample_cols, index=min(2, len(sample_cols) - 1))
    map_division = st.selectbox("Kolom Divisi", sample_cols, index=min(3, len(sample_cols) - 1))
    map_qty = st.selectbox("Kolom Qty", sample_cols, index=min(4, len(sample_cols) - 1))
    map_sku = st.selectbox("Kolom SKU/Kode (opsional)", ["<none>"] + sample_cols, index=0)

with c2:
    st.markdown("**Mapping Pricelist**")
    price_product = st.selectbox("Kolom Product (Pricelist)", price_cols, index=price_cols.index(price_cols[0]))
    price_brand = st.selectbox("Kolom Brand (Pricelist)", price_cols, index=min(1, len(price_cols) - 1))
    price_spec = st.selectbox("Kolom Spesifikasi (Pricelist)", price_cols, index=min(2, len(price_cols) - 1))
    price_price = st.selectbox("Kolom Harga", price_cols, index=min(3, len(price_cols) - 1))
    price_sku = st.selectbox("Kolom SKU/Kode (Pricelist, opsional)", ["<none>"] + price_cols, index=0)

mapping = Mapping(
    product=map_product,
    brand=map_brand,
    specification=map_spec,
    division=map_division,
    qty=map_qty,
    sku=None if map_sku == "<none>" else map_sku,
)

# =========================================================
# Prepare Pricelist Master
# =========================================================
def prepare_pricelist(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out = out.rename(
        columns={
            price_product: "PRODUCT",
            price_brand: "BRAND",
            price_spec: "SPECIFICATION",
            price_price: "PRICE",
        }
    )
    if price_sku != "<none>":
        out = out.rename(columns={price_sku: "SKU"})
    else:
        out["SKU"] = np.nan

    out["PRODUCT"] = normalize_text(out["PRODUCT"])
    out["BRAND"] = normalize_text(out["BRAND"])
    out["SPECIFICATION"] = normalize_text(out["SPECIFICATION"])
    out["PRICE"] = pd.to_numeric(out["PRICE"], errors="coerce")
    out["PRICE_SEGMENT"] = out["PRICE"].apply(find_segment)

    key_cols = ["SKU"] if out["SKU"].notna().any() else ["PRODUCT", "BRAND", "SPECIFICATION"]
    out = out.drop_duplicates(subset=key_cols)
    return out[["PRODUCT", "BRAND", "SPECIFICATION", "SKU", "PRICE", "PRICE_SEGMENT"]]

pricelist_master = prepare_pricelist(price_sample_df)

# =========================================================
# Prepare MPLSSR Period Data
# =========================================================
def prepare_sales(df: pd.DataFrame, period_name: str) -> pd.DataFrame:
    out = df.copy()
    out = out.rename(
        columns={
            mapping.product: "PRODUCT",
            mapping.brand: "BRAND",
            mapping.specification: "SPECIFICATION",
            mapping.division: "DIVISION",
            mapping.qty: "QTY",
        }
    )
    if mapping.sku:
        out = out.rename(columns={mapping.sku: "SKU"})
    else:
        out["SKU"] = np.nan

    out["PRODUCT"] = normalize_text(out["PRODUCT"])
    out["BRAND"] = normalize_text(out["BRAND"])
    out["SPECIFICATION"] = normalize_text(out["SPECIFICATION"])
    out["DIVISION"] = normalize_text(out["DIVISION"])
    out["QTY"] = pd.to_numeric(out["QTY"], errors="coerce").fillna(0)
    out["PERIOD"] = period_name

    if out["SKU"].notna().any() and pricelist_master["SKU"].notna().any():
        merged = out.merge(
            pricelist_master,
            how="left",
            on="SKU",
            suffixes=("", "_P"),
        )
        for col in ["PRODUCT", "BRAND", "SPECIFICATION"]:
            merged[col] = merged[col].fillna(merged.get(f"{col}_P"))
    else:
        merged = out.merge(
            pricelist_master,
            how="left",
            on=["PRODUCT", "BRAND", "SPECIFICATION"],
            suffixes=("", "_P"),
        )

    if "PRICE_SEGMENT" not in merged.columns:
        merged["PRICE_SEGMENT"] = "UNKNOWN"
    return merged[["PRODUCT", "BRAND", "SPECIFICATION", "SKU", "DIVISION", "QTY", "PERIOD", "PRICE", "PRICE_SEGMENT"]]

all_period_frames = []
for period_name, df in period_sources.items():
    if df is not None:
        try:
            all_period_frames.append(prepare_sales(df, period_name))
        except Exception as e:
            st.error(f"Gagal memproses sheet {period_name}: {e}")
            st.stop()

if not all_period_frames:
    st.error("Tidak ada data period yang berhasil diproses.")
    st.stop()

sales = pd.concat(all_period_frames, ignore_index=True)
sales = sales[sales["DIVISION"].isin(DIVISIONS)].copy()

# =========================================================
# Filters
# =========================================================
st.subheader("2) Filter Analisa")
f1, f2, f3 = st.columns(3)
with f1:
    selected_products = st.multiselect(
        "Filter by Product",
        sorted([x for x in sales["PRODUCT"].dropna().unique().tolist()]),
        default=[],
    )
with f2:
    selected_periods = st.multiselect("Periode", PERIODS, default=PERIODS)
with f3:
    selected_divisions = st.multiselect("Divisi", DIVISIONS, default=DIVISIONS)

filtered = sales.copy()
if selected_products:
    filtered = filtered[filtered["PRODUCT"].isin(selected_products)]
if selected_periods:
    filtered = filtered[filtered["PERIOD"].isin(selected_periods)]
if selected_divisions:
    filtered = filtered[filtered["DIVISION"].isin(selected_divisions)]

if filtered.empty:
    st.warning("Data kosong setelah filter diterapkan.")
    st.stop()

# =========================================================
# Trend Insight
# =========================================================
def aggregate_by_period(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby("PERIOD", as_index=False)["QTY"].sum()

period_summary = aggregate_by_period(filtered)
qty_7 = period_summary.loc[period_summary["PERIOD"] == "7DAY", "QTY"].sum()
qty_14 = period_summary.loc[period_summary["PERIOD"] == "14DAY", "QTY"].sum()
qty_30 = period_summary.loc[period_summary["PERIOD"] == "30DAY", "QTY"].sum()

trend_note = "Belum cukup data untuk membaca tren."
if qty_7 > qty_30 and qty_30 > 0:
    trend_note = "7-day naik di atas 30-day. Cek apakah ada promo atau push jangka pendek."
elif qty_7 > 0 and qty_30 > 0 and qty_7 >= qty_30 * 0.95:
    trend_note = "7-day dan 30-day sama-sama kuat. Potensi growth lebih real."
elif qty_7 < qty_30:
    trend_note = "7-day di bawah 30-day. Momentum jangka pendek melemah."

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total QTY 7DAY", qty_fmt(qty_7))
k2.metric("Total QTY 14DAY", qty_fmt(qty_14))
k3.metric("Total QTY 30DAY", qty_fmt(qty_30))
k4.metric("Growth 7 vs 30", growth_fmt(qty_7, qty_30))
st.info(trend_note)

# =========================================================
# Core comparison pivot
# =========================================================
def comparison_pivot(df: pd.DataFrame, dimension: str) -> pd.DataFrame:
    piv = (
        df.groupby([dimension, "DIVISION"], as_index=False)["QTY"]
        .sum()
        .pivot(index=dimension, columns="DIVISION", values="QTY")
        .fillna(0)
        .reset_index()
    )
    for div in DIVISIONS:
        if div not in piv.columns:
            piv[div] = 0
    piv["TOTAL"] = piv[DIVISIONS].sum(axis=1)
    return piv[[dimension] + DIVISIONS + ["TOTAL"]].sort_values("TOTAL", ascending=False)


def render_card_grid(df: pd.DataFrame, dimension: str, title: str, az_sort: bool = False):
    st.subheader(title)
    table = comparison_pivot(df, dimension)
    if az_sort:
        sort_choice = st.radio(
            f"Urutkan {dimension}", ["A-Z", "Z-A", "By Total Desc"], horizontal=True, key=f"sort_{dimension}"
        )
        if sort_choice == "A-Z":
            table = table.sort_values(dimension, ascending=True)
        elif sort_choice == "Z-A":
            table = table.sort_values(dimension, ascending=False)
        else:
            table = table.sort_values("TOTAL", ascending=False)

    cols = st.columns(3)
    for i, row in table.iterrows():
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**{row[dimension]}**")
                c1, c2 = st.columns(2)
                c1.metric("DIV03", qty_fmt(row["DIV03"]))
                c2.metric("DIV04", qty_fmt(row["DIV04"]))
                st.metric("DIV05", qty_fmt(row["DIV05"]))
                st.caption(f"Total: {qty_fmt(row['TOTAL'])}")

    with st.expander(f"Lihat tabel detail {title.lower()}"):
        st.dataframe(table, use_container_width=True)

# =========================================================
# Analysis Cards
# =========================================================
segment_df = filtered.copy()
brand_df = filtered.copy()
spec_df = filtered.copy()

render_card_grid(segment_df, "PRICE_SEGMENT", "3) Card Analisa per Segment Harga")
render_card_grid(brand_df, "BRAND", "4) Card Analisa per Brand")
render_card_grid(spec_df, "SPECIFICATION", "5) Card Analisa per Spesifikasi", az_sort=True)

# =========================================================
# Detailed summary tables by period
# =========================================================
st.subheader("6) Ringkasan Periode per Divisi")
period_div = (
    filtered.groupby(["PERIOD", "DIVISION"], as_index=False)["QTY"].sum()
    .pivot(index="PERIOD", columns="DIVISION", values="QTY")
    .fillna(0)
    .reset_index()
)
for div in DIVISIONS:
    if div not in period_div.columns:
        period_div[div] = 0
period_div["TOTAL"] = period_div[DIVISIONS].sum(axis=1)
st.dataframe(period_div, use_container_width=True)

# =========================================================
# Alert table
# =========================================================
st.subheader("7) Alert Cepat 7DAY vs 30DAY")
alert_dim = st.selectbox("Basis alert", ["PRODUCT", "BRAND", "SPECIFICATION", "PRICE_SEGMENT"])
base = (
    filtered.groupby([alert_dim, "PERIOD"], as_index=False)["QTY"].sum()
    .pivot(index=alert_dim, columns="PERIOD", values="QTY")
    .fillna(0)
    .reset_index()
)
for p in PERIODS:
    if p not in base.columns:
        base[p] = 0
base["DELTA_7_VS_30"] = base["7DAY"] - base["30DAY"]
base["GROWTH_7_VS_30_%"] = np.where(base["30DAY"] == 0, np.nan, (base["7DAY"] - base["30DAY"]) / base["30DAY"] * 100)
base["INSIGHT"] = np.select(
    [
        (base["7DAY"] > base["30DAY"]) & (base["30DAY"] > 0),
        (base["7DAY"] > 0) & (base["30DAY"] > 0) & (base["7DAY"] >= base["30DAY"] * 0.95),
        (base["7DAY"] < base["30DAY"]),
    ],
    [
        "Naik cepat, cek efek promo",
        "Growth cenderung real",
        "Momentum melemah",
    ],
    default="Perlu validasi",
)
st.dataframe(base.sort_values("DELTA_7_VS_30", ascending=False), use_container_width=True)

# =========================================================
# Download cleaned data
# =========================================================
st.subheader("8) Download Hasil Olahan")
out_buffer = io.BytesIO()
with pd.ExcelWriter(out_buffer, engine="openpyxl") as writer:
    filtered.to_excel(writer, index=False, sheet_name="filtered_data")
    period_div.to_excel(writer, index=False, sheet_name="period_summary")
    comparison_pivot(segment_df, "PRICE_SEGMENT").to_excel(writer, index=False, sheet_name="segment_cards")
    comparison_pivot(brand_df, "BRAND").to_excel(writer, index=False, sheet_name="brand_cards")
    comparison_pivot(spec_df, "SPECIFICATION").to_excel(writer, index=False, sheet_name="spec_cards")
    base.to_excel(writer, index=False, sheet_name="alerts")

st.download_button(
    "Download hasil analisa (.xlsx)",
    data=out_buffer.getvalue(),
    file_name="hasil_analisa_penjualan.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(
    "Catatan: karena struktur Excel Anda bisa berbeda-beda, aplikasi ini dibuat fleksibel dengan mapping kolom di awal. Setelah struktur final Anda pasti, mapping ini bisa saya ubah menjadi otomatis penuh."
)
