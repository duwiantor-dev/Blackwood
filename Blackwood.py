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

# =========================================================
# Styling
# =========================================================
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 1rem;}
    .metric-card {
        background: #ffffff;
        border: 1px solid #e6e9ef;
        border-radius: 14px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        min-height: 108px;
    }
    .metric-label {font-size: 0.82rem; color: #6b7280; margin-bottom: 6px;}
    .metric-value {font-size: 1.55rem; font-weight: 700; color: #111827;}
    .metric-sub {font-size: 0.82rem; color: #6b7280; margin-top: 6px;}
    .section-card {
        background: #ffffff;
        border: 1px solid #e6e9ef;
        border-radius: 16px;
        padding: 14px 14px 10px 14px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        margin-bottom: 14px;
    }
    .up-chip {background:#e8f6ee; color:#15803d; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:600;}
    .down-chip {background:#fdecec; color:#b91c1c; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:600;}
    .neutral-chip {background:#eef2ff; color:#4338ca; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:600;}
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# Helpers
# =========================================================
def normalize_text(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.upper().replace({"NAN": np.nan, "NONE": np.nan, "": np.nan})


def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def fmt_int(v: float) -> str:
    if pd.isna(v):
        return "0"
    return f"{v:,.0f}".replace(",", ".")


def fmt_pct(v: float) -> str:
    if pd.isna(v):
        return "-"
    return f"{v:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_price(v: float) -> str:
    if pd.isna(v):
        return "-"
    return f"Rp {v:,.0f}".replace(",", ".")


def growth_pct(cur: float, base: float) -> float:
    if pd.isna(base) or base == 0:
        return np.nan
    return (cur - base) / base * 100


def price_segment(price: float) -> str:
    if pd.isna(price):
        return "UNKNOWN"
    for low, high, label in PRICE_SEGMENTS:
        if low <= float(price) < high:
            return label
    return "UNKNOWN"


def chip_html(text: str, kind: str) -> str:
    cls = {"up": "up-chip", "down": "down-chip", "neutral": "neutral-chip"}.get(kind, "neutral-chip")
    return f'<span class="{cls}">{text}</span>'


def metric_box(label: str, value: str, sub: str = ""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# =========================================================
# MPLSSR parser
# =========================================================
def load_mplssr(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="ALL")
    df.columns = [str(c).strip() for c in df.columns]
    base_cols = ["SKU NO", "PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI"]
    df = df[base_cols + [c for p in PERIODS for c in MPLSSR_DIV_COLS[p].values() if c in df.columns]].copy()
    for c in base_cols:
        df[c] = normalize_text(df[c])
    df = df[df["SKU NO"].notna()].copy()
    df = df[~df["SKU NO"].isin(["TOTAL", "SHARE%"])]

    rows = []
    for period in PERIODS:
        for div, col in MPLSSR_DIV_COLS[period].items():
            if col not in df.columns:
                continue
            tmp = df[["SKU NO", "PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI", col]].copy()
            tmp["QTY"] = to_num(tmp[col]).fillna(0)
            tmp["PERIOD"] = period
            tmp["DIVISION"] = div
            rows.append(tmp[["SKU NO", "PRODUCT", "BRAND", "KODE BARANG", "SPESIFIKASI", "PERIOD", "DIVISION", "QTY"]])
    out = pd.concat(rows, ignore_index=True)
    return out

# =========================================================
# Pricelist parser
# =========================================================
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


def parse_pricelist_sheet(file, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    raw = raw.iloc[:, : raw.shape[1]].copy()

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

    df = raw.iloc[4:].copy().reset_index(drop=True)
    df.columns = columns

    for c in ["SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "TOT", "M3", "SRP"]:
        if c not in df.columns:
            df[c] = np.nan

    ram_start_idx = None
    for i, grp in enumerate(row2):
        if grp == "RAM":
            ram_start_idx = i
            break

    stock05_cols = []
    if ram_start_idx is not None:
        stock05_cols = [columns[i] for i in range(ram_start_idx, len(columns))]

    stock03_cols = [c for c in columns if c in ["JKT__3A", "JKT__3B"]]
    stock04_cols = [c for c in columns if c in ["JKT__4A", "JKT__4B"]]

    df["SKU NO"] = normalize_text(df["SKU NO"])
    df["PRODUCT"] = normalize_text(df["PRODUCT"])
    df["KODEBARANG"] = normalize_text(df["KODEBARANG"])
    df["SPESIFIKASI"] = normalize_text(df["SPESIFIKASI"])
    df = df[df["SKU NO"].notna()].copy()
    df = df[~df["SKU NO"].isin(["TOTAL"])]

    df["PRICE"] = to_num(df["M3"]) * 1000
    df["STOK_TOTAL"] = to_num(df["TOT"]).fillna(0)
    df["STOK_DIV03"] = df[stock03_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock03_cols else 0
    df["STOK_DIV04"] = df[stock04_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock04_cols else 0
    df["STOK_DIV05"] = df[stock05_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if stock05_cols else 0
    df["CATEGORY"] = sheet_name.upper()
    df["PRICE_SEGMENT"] = df["PRICE"].apply(price_segment)

    return df[[
        "SKU NO", "PRODUCT", "KODEBARANG", "SPESIFIKASI", "PRICE", "PRICE_SEGMENT",
        "STOK_TOTAL", "STOK_DIV03", "STOK_DIV04", "STOK_DIV05", "CATEGORY"
    ]]


def load_pricelist(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    sheets = [s for s in xls.sheet_names if s.upper() in VALID_PRICELIST_SHEETS]
    frames = [parse_pricelist_sheet(file, s) for s in sheets]
    out = pd.concat(frames, ignore_index=True)
    out = out.drop_duplicates(subset=["SKU NO"], keep="first")
    return out

# =========================================================
# Transform for dashboard
# =========================================================
def build_master(sales: pd.DataFrame, stock: pd.DataFrame) -> pd.DataFrame:
    df = sales.merge(stock, how="left", on="SKU NO")
    for col in ["PRODUCT_y", "SPESIFIKASI_y", "KODEBARANG", "CATEGORY", "PRICE", "PRICE_SEGMENT", "STOK_TOTAL", "STOK_DIV03", "STOK_DIV04", "STOK_DIV05"]:
        if col not in df.columns:
            df[col] = np.nan

    df["PRODUCT_FINAL"] = df["PRODUCT_x"].fillna(df.get("PRODUCT_y"))
    df["SPEC_FINAL"] = df["SPESIFIKASI_x"].fillna(df.get("SPESIFIKASI_y"))
    df["CATEGORY"] = df["CATEGORY"].fillna(df["PRODUCT_FINAL"])
    df["PRICE_SEGMENT"] = df["PRICE_SEGMENT"].fillna("UNKNOWN")
    df["PRICE"] = to_num(df["PRICE"])
    df["QTY"] = to_num(df["QTY"]).fillna(0)

    stock_map = {"DIV03": "STOK_DIV03", "DIV04": "STOK_DIV04", "DIV05": "STOK_DIV05"}
    df["STOK_DIVISI"] = df.apply(lambda r: r.get(stock_map.get(r["DIVISION"]), 0), axis=1)
    df["SELL_THRU_%"] = np.where(df["STOK_DIVISI"] > 0, df["QTY"] / df["STOK_DIVISI"] * 100, np.nan)
    return df


def compare_table(df: pd.DataFrame, dim: str) -> pd.DataFrame:
    qty = df.groupby([dim, "DIVISION"], as_index=False)["QTY"].sum().pivot(index=dim, columns="DIVISION", values="QTY").fillna(0)
    stock = df.groupby([dim, "DIVISION"], as_index=False)["STOK_DIVISI"].sum().pivot(index=dim, columns="DIVISION", values="STOK_DIVISI").fillna(0)
    qty.columns = [f"QTY_{c}" for c in qty.columns]
    stock.columns = [f"STOK_{c}" for c in stock.columns]
    out = qty.join(stock, how="outer").fillna(0).reset_index()
    for d in DIVISIONS:
        if f"QTY_{d}" not in out.columns:
            out[f"QTY_{d}"] = 0
        if f"STOK_{d}" not in out.columns:
            out[f"STOK_{d}"] = 0
    out["QTY_TOTAL"] = out[[f"QTY_{d}" for d in DIVISIONS]].sum(axis=1)
    out["STOK_TOTAL"] = out[[f"STOK_{d}" for d in DIVISIONS]].sum(axis=1)
    out["SELL_THRU_%"] = np.where(out["STOK_TOTAL"] > 0, out["QTY_TOTAL"] / out["STOK_TOTAL"] * 100, np.nan)
    return out.sort_values(["QTY_TOTAL", "STOK_TOTAL"], ascending=[False, False])


def top_table(df: pd.DataFrame, dim: str, top_n: int = 10) -> pd.DataFrame:
    t = compare_table(df, dim).head(top_n).copy()
    return t


def trend_table(df: pd.DataFrame, dim: str) -> pd.DataFrame:
    piv = df.groupby([dim, "PERIOD"], as_index=False)["QTY"].sum().pivot(index=dim, columns="PERIOD", values="QTY").fillna(0).reset_index()
    for p in PERIODS:
        if p not in piv.columns:
            piv[p] = 0
    piv["DELTA_7_vs_30"] = piv["7DAY"] - piv["30DAY"]
    piv["GROWTH_7_vs_30_%"] = np.where(piv["30DAY"] > 0, (piv["7DAY"] - piv["30DAY"]) / piv["30DAY"] * 100, np.nan)
    piv["INSIGHT"] = np.select(
        [
            (piv["7DAY"] > piv["30DAY"]) & (piv["30DAY"] > 0),
            (piv["7DAY"] < piv["30DAY"]) & (piv["30DAY"] > 0),
        ],
        ["Alert cepat / cek promo", "Momentum melemah"],
        default="Perlu validasi"
    )
    return piv.sort_values("DELTA_7_vs_30", ascending=False)


def render_analysis_cards(table: pd.DataFrame, dim: str, title: str, sort_az: bool = False):
    with st.container(border=False):
        st.markdown(f"### {title}")
        if sort_az:
            sort_mode = st.radio(f"Urutan {dim}", ["TOTAL", "A-Z", "Z-A"], horizontal=True, key=f"sort_{dim}")
            if sort_mode == "A-Z":
                table = table.sort_values(dim, ascending=True)
            elif sort_mode == "Z-A":
                table = table.sort_values(dim, ascending=False)
        cols = st.columns(3)
        for i, (_, r) in enumerate(table.iterrows()):
            with cols[i % 3]:
                with st.container(border=True):
                    st.markdown(f"**{r[dim]}**")
                    q1, q2, q3 = st.columns(3)
                    q1.metric("QTY D03", fmt_int(r["QTY_DIV03"]))
                    q2.metric("QTY D04", fmt_int(r["QTY_DIV04"]))
                    q3.metric("QTY D05", fmt_int(r["QTY_DIV05"]))
                    s1, s2, s3 = st.columns(3)
                    s1.metric("STK D03", fmt_int(r["STOK_DIV03"]))
                    s2.metric("STK D04", fmt_int(r["STOK_DIV04"]))
                    s3.metric("STK D05", fmt_int(r["STOK_DIV05"]))
                    st.caption(f"Total QTY: {fmt_int(r['QTY_TOTAL'])} | Total STOK: {fmt_int(r['STOK_TOTAL'])} | Sell Thru: {fmt_pct(r['SELL_THRU_%'])}")
        with st.expander(f"Lihat tabel detail - {title}"):
            st.dataframe(table, use_container_width=True)

# =========================================================
# Sidebar
# =========================================================
st.title("Dashboard Analisa Sales vs Stock")
st.caption("QTY diambil dari MPLSSR. STOK dan harga diambil dari Pricelist. Fokus awal dibuat seperti dashboard pada contoh.")

st.sidebar.header("Upload File")
mplssr_file = st.sidebar.file_uploader("Upload MPLSSR", type=["xlsx", "xls"])
pricelist_file = st.sidebar.file_uploader("Upload Pricelist", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.write("**Rules:**")
st.sidebar.caption("- QTY: MPLSSR
- STOK: Pricelist
- Harga: kolom M3
- STOK DIV05: dari RAM sampai kolom paling belakang")

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

# =========================================================
# Filters
# =========================================================
with st.sidebar:
    st.markdown("---")
    metric_type = st.selectbox("Metric utama", ["QTY", "STOK", "SELL THRU"])
    selected_products = st.multiselect("Product", sorted(master["PRODUCT_FINAL"].dropna().unique().tolist()))
    selected_categories = st.multiselect("Category", sorted(master["CATEGORY"].dropna().unique().tolist()))
    selected_brands = st.multiselect("Brand", sorted(master["BRAND"].dropna().unique().tolist()))
    selected_periods = st.multiselect("Periode", PERIODS, default=PERIODS)
    selected_segments = st.multiselect("Segment Harga", [s[2] for s in PRICE_SEGMENTS] + ["UNKNOWN"])

filtered = master.copy()
if selected_products:
    filtered = filtered[filtered["PRODUCT_FINAL"].isin(selected_products)]
if selected_categories:
    filtered = filtered[filtered["CATEGORY"].isin(selected_categories)]
if selected_brands:
    filtered = filtered[filtered["BRAND"].isin(selected_brands)]
if selected_periods:
    filtered = filtered[filtered["PERIOD"].isin(selected_periods)]
if selected_segments:
    filtered = filtered[filtered["PRICE_SEGMENT"].isin(selected_segments)]

if filtered.empty:
    st.warning("Data kosong setelah filter diterapkan.")
    st.stop()

# =========================================================
# KPI row
# =========================================================
qty7 = filtered.loc[filtered["PERIOD"] == "7DAY", "QTY"].sum()
qty14 = filtered.loc[filtered["PERIOD"] == "14DAY", "QTY"].sum()
qty30 = filtered.loc[filtered["PERIOD"] == "30DAY", "QTY"].sum()
stok_total = filtered[["SKU NO", "DIVISION", "STOK_DIVISI"]].drop_duplicates()["STOK_DIVISI"].sum()
sell_thru = growth_pct(qty7 + qty14 + qty30, stok_total)
trend = growth_pct(qty7, qty30)

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    metric_box("Total QTY 7DAY", fmt_int(qty7), "Alert cepat")
with c2:
    metric_box("Total QTY 14DAY", fmt_int(qty14), "Validasi tren")
with c3:
    metric_box("Total QTY 30DAY", fmt_int(qty30), "Baseline")
with c4:
    metric_box("Total Stock", fmt_int(stok_total), "Dari pricelist")
with c5:
    metric_box("Growth 7 vs 30", fmt_pct(trend), "Naik/turun momentum")

if pd.notna(trend) and trend > 0:
    st.markdown(chip_html("7DAY di atas 30DAY - cek apakah real growth atau efek promo", "up"), unsafe_allow_html=True)
elif pd.notna(trend) and trend < 0:
    st.markdown(chip_html("7DAY di bawah 30DAY - momentum melemah", "down"), unsafe_allow_html=True)
else:
    st.markdown(chip_html("Belum cukup sinyal untuk membaca tren", "neutral"), unsafe_allow_html=True)

# =========================================================
# Main table like screenshot
# =========================================================
st.markdown("### Tabel Utama Analisa")
base_period = st.selectbox("Periode tabel utama", PERIODS, index=2)
main = filtered[filtered["PERIOD"] == base_period].copy()
main["DELTA_QTY_7_vs_30"] = main.groupby("SKU NO")["QTY"].transform("sum")
main_table = (
    main.groupby(["SKU NO", "PRODUCT_FINAL", "BRAND", "CATEGORY", "SPEC_FINAL", "PRICE_SEGMENT", "PRICE", "DIVISION"], as_index=False)[["QTY", "STOK_DIVISI"]]
    .sum()
    .pivot(index=["SKU NO", "PRODUCT_FINAL", "BRAND", "CATEGORY", "SPEC_FINAL", "PRICE_SEGMENT", "PRICE"], columns="DIVISION", values=["QTY", "STOK_DIVISI"])
)
main_table.columns = [f"{a}_{b}" for a, b in main_table.columns]
main_table = main_table.fillna(0).reset_index()
for d in DIVISIONS:
    if f"QTY_{d}" not in main_table.columns:
        main_table[f"QTY_{d}"] = 0
    if f"STOK_DIVISI_{d}" not in main_table.columns:
        main_table[f"STOK_DIVISI_{d}"] = 0
main_table["QTY_TOTAL"] = main_table[[f"QTY_{d}" for d in DIVISIONS]].sum(axis=1)
main_table["STOK_TOTAL"] = main_table[[f"STOK_DIVISI_{d}" for d in DIVISIONS]].sum(axis=1)
main_table["SELL_THRU_%"] = np.where(main_table["STOK_TOTAL"] > 0, main_table["QTY_TOTAL"] / main_table["STOK_TOTAL"] * 100, np.nan)
main_table = main_table.sort_values(["QTY_TOTAL", "SELL_THRU_%"], ascending=[False, False])
st.dataframe(main_table, use_container_width=True, height=420)

# =========================================================
# Top summary blocks
# =========================================================
left, right = st.columns(2)
with left:
    st.markdown("### Top 10 Product")
    t = top_table(filtered, "PRODUCT_FINAL", 10)
    st.dataframe(t[["PRODUCT_FINAL", "QTY_DIV03", "QTY_DIV04", "QTY_DIV05", "QTY_TOTAL", "STOK_TOTAL", "SELL_THRU_%"]], use_container_width=True, height=340)
with right:
    st.markdown("### Top 10 Brand")
    t = top_table(filtered, "BRAND", 10)
    st.dataframe(t[["BRAND", "QTY_DIV03", "QTY_DIV04", "QTY_DIV05", "QTY_TOTAL", "STOK_TOTAL", "SELL_THRU_%"]], use_container_width=True, height=340)

# =========================================================
# Lower analysis cards
# =========================================================
segment_table = compare_table(filtered, "PRICE_SEGMENT")
brand_table = compare_table(filtered, "BRAND")
spec_table = compare_table(filtered, "SPEC_FINAL")

render_analysis_cards(segment_table, "PRICE_SEGMENT", "Card Analisa per Segment Harga")
render_analysis_cards(brand_table.head(18), "BRAND", "Card Analisa per Brand")
render_analysis_cards(spec_table.head(18), "SPEC_FINAL", "Card Analisa per Spesifikasi", sort_az=True)

# =========================================================
# Alert table
# =========================================================
st.markdown("### Alert 7DAY vs 30DAY")
alert_basis = st.selectbox("Basis alert", ["PRODUCT_FINAL", "BRAND", "SPEC_FINAL", "PRICE_SEGMENT", "CATEGORY"])
alert = trend_table(filtered, alert_basis)
st.dataframe(alert, use_container_width=True, height=340)

# =========================================================
# Download outputs
# =========================================================
out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    master.to_excel(writer, index=False, sheet_name="master")
    main_table.to_excel(writer, index=False, sheet_name="main_table")
    segment_table.to_excel(writer, index=False, sheet_name="segment")
    brand_table.to_excel(writer, index=False, sheet_name="brand")
    spec_table.to_excel(writer, index=False, sheet_name="spesifikasi")
    alert.to_excel(writer, index=False, sheet_name="alert")

st.download_button(
    "Download hasil analisa",
    data=out.getvalue(),
    file_name="dashboard_sales_stock.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Debug hasil parsing"):
    st.write("Sales:", sales.head(20))
    st.write("Stock:", stock.head(20))
    st.write("Master:", master.head(20))
