import io
from typing import List

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Analisa Penjualan Div03 / Div04 / Div05", layout="wide")

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
DIVISION_MAP = {
    "DIV03": {"7DAY": "03 OLP", "14DAY": "03 OLP", "30DAY": "03 OLP"},
    "DIV04": {"7DAY": "04 MOD", "14DAY": "04 MOD", "30DAY": "04 MOD"},
    "DIV05": {"7DAY": "05 OLR", "14DAY": "05 OLR", "30DAY": "05 OLR"},
}
PERIOD_KEYS = {
    "7DAY": ["7D SSR", "7D AVG", "T7D", "00 AGR", "01 DIS", "02 COM", "03 OLP", "04 MOD", "05 OLR", "06 COM"],
    "14DAY": ["14D SSR", "14D AVG", "T14D", "00 AGR", "01 DIS", "02 COM", "03 OLP", "04 MOD", "05 OLR", "06 COM"],
    "30DAY": ["30D SSR", "30D AVG", "T30D", "00 AGR", "01 DIS", "02 COM", "03 OLP", "04 MOD", "05 OLR", "06 COM"],
}


# =========================================================
# Helpers
# =========================================================
def normalize_text(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.upper().replace({"NAN": np.nan, "NONE": np.nan, "": np.nan})


def clean_money(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def find_price_segment(price: float) -> str:
    if pd.isna(price):
        return "UNKNOWN"
    price = float(price)
    for low, high, label in PRICE_SEGMENTS:
        if low <= price < high:
            return label
    return "UNKNOWN"


def qty_fmt(v: float) -> str:
    if pd.isna(v):
        return "0"
    return f"{v:,.0f}".replace(",", ".")


def pct_fmt(v: float) -> str:
    if pd.isna(v):
        return "-"
    s = f"{v:,.2f}%"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def growth_pct(current: float, baseline: float) -> float:
    if pd.isna(baseline) or baseline == 0:
        return np.nan
    return (current - baseline) / baseline * 100


def infer_brand_from_spec(spec: str) -> str:
    if pd.isna(spec):
        return np.nan
    text = str(spec).strip().upper()
    if not text:
        return np.nan
    return text.split()[0]


# =========================================================
# Parsers for uploaded files
# =========================================================
def load_mplssr(uploaded_file) -> pd.DataFrame:
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    raw.columns = [str(c).strip() for c in raw.columns]

    rename_map = {
        raw.columns[0]: "SKU NO",
        raw.columns[1]: "PRODUCT",
        raw.columns[2]: "BRAND",
        raw.columns[3]: "KODE BARANG",
        raw.columns[4]: "SPESIFIKASI",
    }
    raw = raw.rename(columns=rename_map)

    raw = raw.iloc[5:].copy()
    raw["SKU NO"] = normalize_text(raw["SKU NO"])
    raw["PRODUCT"] = normalize_text(raw["PRODUCT"])
    raw["BRAND"] = normalize_text(raw["BRAND"])
    raw["SPESIFIKASI"] = normalize_text(raw["SPESIFIKASI"])
    raw = raw[raw["SKU NO"].notna()].copy()
    raw = raw[~raw["SKU NO"].isin(["TOTAL", "SHARE%"])]

    needed_cols = ["SKU NO", "PRODUCT", "BRAND", "SPESIFIKASI"]
    for _, cols in PERIOD_KEYS.items():
        needed_cols.extend(cols)
    needed_cols = [c for c in needed_cols if c in raw.columns]
    raw = raw[needed_cols].copy()

    for col in raw.columns:
        if col not in ["SKU NO", "PRODUCT", "BRAND", "SPESIFIKASI"]:
            raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    rows = []
    for period in ["7DAY", "14DAY", "30DAY"]:
        for div, mapping in DIVISION_MAP.items():
            measure_col = mapping[period]
            if measure_col not in raw.columns:
                continue
            tmp = raw[["SKU NO", "PRODUCT", "BRAND", "SPESIFIKASI", measure_col]].copy()
            tmp = tmp.rename(columns={measure_col: "QTY"})
            tmp["PERIOD"] = period
            tmp["DIVISION"] = div
            rows.append(tmp)

    sales = pd.concat(rows, ignore_index=True)
    sales["QTY"] = pd.to_numeric(sales["QTY"], errors="coerce").fillna(0)
    sales = sales[sales["QTY"] != 0].copy()
    return sales



def load_pricelist(uploaded_file) -> pd.DataFrame:
    xls = pd.ExcelFile(uploaded_file)
    frames: List[pd.DataFrame] = []

    for sheet_name in xls.sheet_names:
        if sheet_name.upper() == "AGS10":
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            expected = ["PRODUCT", "SKU NO", "KODEBARANG", "SPESIFIKASI", "NOTES", "TOT"]
            for col in expected:
                if col not in df.columns:
                    df[col] = np.nan
            df["SRP"] = np.nan
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            if "SKU NO" not in df.columns or "PRODUCT" not in df.columns or "SPESIFIKASI" not in df.columns:
                continue
            if "SRP" not in df.columns:
                df["SRP"] = np.nan
            if "TOT" not in df.columns:
                df["TOT"] = np.nan

        df = df[[c for c in ["SKU NO", "PRODUCT", "SPESIFIKASI", "SRP", "TOT"] if c in df.columns]].copy()
        df["SKU NO"] = normalize_text(df["SKU NO"])
        df["PRODUCT"] = normalize_text(df["PRODUCT"])
        df["SPESIFIKASI"] = normalize_text(df["SPESIFIKASI"])
        df["PRICE"] = clean_money(df.get("SRP")) * 1000
        df["STOCK"] = clean_money(df.get("TOT"))
        df["SOURCE_SHEET"] = sheet_name.upper()
        df = df[df["SKU NO"].notna()].copy()
        df = df[~df["SKU NO"].isin(["TOTAL"])]
        df = df[df["SPESIFIKASI"].notna()].copy()
        df["BRAND_PRICELIST"] = df["SPESIFIKASI"].apply(infer_brand_from_spec)
        frames.append(df[["SKU NO", "PRODUCT", "SPESIFIKASI", "PRICE", "STOCK", "SOURCE_SHEET", "BRAND_PRICELIST"]])

    pricelist = pd.concat(frames, ignore_index=True)
    pricelist = pricelist.sort_values(["SKU NO", "PRICE"], ascending=[True, False])
    pricelist = pricelist.drop_duplicates(subset=["SKU NO"], keep="first")
    pricelist["PRICE_SEGMENT"] = pricelist["PRICE"].apply(find_price_segment)
    return pricelist


# =========================================================
# UI
# =========================================================
st.title("📊 Analisa Penjualan Div03 / Div04 / Div05")
st.caption("Khusus disesuaikan untuk format file MPLSSR dan Pricelist yang Anda upload.")

st.sidebar.header("Upload File")
mplssr_file = st.sidebar.file_uploader("Upload MPLSSR", type=["xlsx", "xls"])
pricelist_file = st.sidebar.file_uploader("Upload Pricelist", type=["xlsx", "xls"])

with st.sidebar.expander("Logika 7 / 14 / 30 Day", expanded=True):
    st.markdown(
        """
- **7DAY** = alert cepat  
- **14DAY** = validasi tren  
- **30DAY** = baseline  

**Interpretasi:**
- Jika **7DAY naik** tapi **30DAY stagnan** → kemungkinan efek promo
- Jika **7DAY dan 30DAY sama-sama naik** → growth lebih real
        """
    )

if not mplssr_file or not pricelist_file:
    st.info("Silakan upload file MPLSSR dan Pricelist terlebih dahulu.")
    st.stop()

try:
    sales = load_mplssr(mplssr_file)
    pricelist = load_pricelist(pricelist_file)
except Exception as e:
    st.error(f"Gagal membaca file: {e}")
    st.stop()

master = sales.merge(pricelist, how="left", on=["SKU NO", "PRODUCT", "SPESIFIKASI"])
master["BRAND_FINAL"] = master["BRAND"].fillna(master["BRAND_PRICELIST"])
master["PRICE_SEGMENT"] = master["PRICE_SEGMENT"].fillna("UNKNOWN")
master["SOURCE_SHEET"] = master["SOURCE_SHEET"].fillna("UNKNOWN")

# =========================================================
# Filters
# =========================================================
st.subheader("Filter")
c1, c2, c3, c4 = st.columns(4)
with c1:
    product_options = sorted([x for x in master["PRODUCT"].dropna().unique().tolist()])
    selected_products = st.multiselect("Product", product_options)
with c2:
    selected_periods = st.multiselect("Periode", ["7DAY", "14DAY", "30DAY"], default=["7DAY", "14DAY", "30DAY"])
with c3:
    brand_options = sorted([x for x in master["BRAND_FINAL"].dropna().unique().tolist()])
    selected_brands = st.multiselect("Brand", brand_options)
with c4:
    source_options = sorted([x for x in master["SOURCE_SHEET"].dropna().unique().tolist()])
    selected_sources = st.multiselect("Kategori Pricelist", source_options)

filtered = master.copy()
if selected_products:
    filtered = filtered[filtered["PRODUCT"].isin(selected_products)]
if selected_periods:
    filtered = filtered[filtered["PERIOD"].isin(selected_periods)]
if selected_brands:
    filtered = filtered[filtered["BRAND_FINAL"].isin(selected_brands)]
if selected_sources:
    filtered = filtered[filtered["SOURCE_SHEET"].isin(selected_sources)]

if filtered.empty:
    st.warning("Tidak ada data setelah filter diterapkan.")
    st.stop()

# =========================================================
# KPI + trend insight
# =========================================================
def total_period_qty(df: pd.DataFrame, period: str) -> float:
    return df.loc[df["PERIOD"] == period, "QTY"].sum()

qty7 = total_period_qty(filtered, "7DAY")
qty14 = total_period_qty(filtered, "14DAY")
qty30 = total_period_qty(filtered, "30DAY")
g7030 = growth_pct(qty7, qty30)

trend_note = "Perlu validasi tambahan."
if qty7 > qty30 and qty30 > 0:
    trend_note = "7DAY lebih tinggi dari 30DAY. Ada indikasi alert cepat atau dorongan promo."
elif qty7 > 0 and qty30 > 0 and qty7 >= qty30 * 0.95:
    trend_note = "7DAY dan 30DAY sama-sama kuat. Growth cenderung real."
elif qty7 < qty30:
    trend_note = "7DAY di bawah 30DAY. Momentum jangka pendek melemah."

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total QTY 7DAY", qty_fmt(qty7))
k2.metric("Total QTY 14DAY", qty_fmt(qty14))
k3.metric("Total QTY 30DAY", qty_fmt(qty30))
k4.metric("Growth 7DAY vs 30DAY", pct_fmt(g7030))
st.info(trend_note)


# =========================================================
# Table builders
# =========================================================
def make_compare_table(df: pd.DataFrame, dimension: str) -> pd.DataFrame:
    piv = (
        df.groupby([dimension, "DIVISION"], as_index=False)["QTY"]
        .sum()
        .pivot(index=dimension, columns="DIVISION", values="QTY")
        .fillna(0)
        .reset_index()
    )
    for div in ["DIV03", "DIV04", "DIV05"]:
        if div not in piv.columns:
            piv[div] = 0
    piv["TOTAL"] = piv[["DIV03", "DIV04", "DIV05"]].sum(axis=1)
    return piv[[dimension, "DIV03", "DIV04", "DIV05", "TOTAL"]].sort_values("TOTAL", ascending=False)


def render_cards(table: pd.DataFrame, key_col: str, title: str, az_sort: bool = False):
    st.subheader(title)
    if az_sort:
        sort_mode = st.radio(
            f"Urutan {key_col}",
            ["A-Z", "Z-A", "TOTAL DESC"],
            horizontal=True,
            key=f"sort_{key_col}",
        )
        if sort_mode == "A-Z":
            table = table.sort_values(key_col, ascending=True)
        elif sort_mode == "Z-A":
            table = table.sort_values(key_col, ascending=False)
        else:
            table = table.sort_values("TOTAL", ascending=False)

    cols = st.columns(3)
    for idx, row in table.iterrows():
        with cols[idx % 3]:
            with st.container(border=True):
                st.markdown(f"**{row[key_col]}**")
                a, b = st.columns(2)
                a.metric("DIV03", qty_fmt(row["DIV03"]))
                b.metric("DIV04", qty_fmt(row["DIV04"]))
                st.metric("DIV05", qty_fmt(row["DIV05"]))
                st.caption(f"Total: {qty_fmt(row['TOTAL'])}")

    with st.expander(f"Lihat tabel detail - {title}"):
        st.dataframe(table, use_container_width=True)


# =========================================================
# Main analysis cards
# =========================================================
segment_table = make_compare_table(filtered, "PRICE_SEGMENT")
brand_table = make_compare_table(filtered, "BRAND_FINAL")
spec_table = make_compare_table(filtered, "SPESIFIKASI")

render_cards(segment_table, "PRICE_SEGMENT", "Card Analisa per Segment Harga")
render_cards(brand_table, "BRAND_FINAL", "Card Analisa per Brand")
render_cards(spec_table, "SPESIFIKASI", "Card Analisa per Spesifikasi", az_sort=True)

# =========================================================
# Top tables
# =========================================================
left, right = st.columns(2)
with left:
    st.subheader("Top 10 Brand")
    st.dataframe(brand_table.head(10), use_container_width=True)
with right:
    st.subheader("Top 10 Segment Harga")
    st.dataframe(segment_table.head(10), use_container_width=True)

# =========================================================
# Alert cepat 7d vs 30d
# =========================================================
st.subheader("Alert Cepat 7DAY vs 30DAY")
alert_basis = st.selectbox("Basis alert", ["PRODUCT", "BRAND_FINAL", "SPESIFIKASI", "PRICE_SEGMENT"])
alert = (
    filtered.groupby([alert_basis, "PERIOD"], as_index=False)["QTY"]
    .sum()
    .pivot(index=alert_basis, columns="PERIOD", values="QTY")
    .fillna(0)
    .reset_index()
)
for p in ["7DAY", "14DAY", "30DAY"]:
    if p not in alert.columns:
        alert[p] = 0
alert["DELTA_7_vs_30"] = alert["7DAY"] - alert["30DAY"]
alert["GROWTH_7_vs_30_%"] = np.where(alert["30DAY"] == 0, np.nan, (alert["7DAY"] - alert["30DAY"]) / alert["30DAY"] * 100)
alert["INSIGHT"] = np.select(
    [
        (alert["7DAY"] > alert["30DAY"]) & (alert["30DAY"] > 0),
        (alert["7DAY"] > 0) & (alert["30DAY"] > 0) & (alert["7DAY"] >= alert["30DAY"] * 0.95),
        (alert["7DAY"] < alert["30DAY"]),
    ],
    [
        "Naik cepat, cek promo",
        "Growth cenderung real",
        "Momentum melemah",
    ],
    default="Perlu validasi",
)
st.dataframe(alert.sort_values("DELTA_7_vs_30", ascending=False), use_container_width=True)

# =========================================================
# Debug info and download
# =========================================================
with st.expander("Cek hasil join data"):
    st.write("Sample data gabungan")
    st.dataframe(filtered.head(50), use_container_width=True)
    st.caption(f"Jumlah baris sales: {len(sales):,} | jumlah baris pricelist master: {len(pricelist):,} | jumlah baris final: {len(master):,}".replace(",", "."))

output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    filtered.to_excel(writer, index=False, sheet_name="filtered_data")
    segment_table.to_excel(writer, index=False, sheet_name="segment_harga")
    brand_table.to_excel(writer, index=False, sheet_name="brand")
    spec_table.to_excel(writer, index=False, sheet_name="spesifikasi")
    alert.to_excel(writer, index=False, sheet_name="alert_7_vs_30")

st.download_button(
    "Download hasil analisa",
    data=output.getvalue(),
    file_name="hasil_analisa_div03_div04_div05.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption("Versi ini sudah disesuaikan ke struktur file MPLSSR ALL dan Pricelist multi-sheet yang Anda upload.")
