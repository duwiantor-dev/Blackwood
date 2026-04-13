
import io
from typing import List

import numpy as np
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, ColumnsAutoSizeMode

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


def render_main_table_aggrid(df: pd.DataFrame):
    display_df = df.copy()

    # Tambahkan kolom PRODUCT dan BRAND di depan tabel utama
    product_brand = (
        filtered[["KODEBARANG", "PRODUCT_FINAL", "BRAND"]]
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

    gb = GridOptionsBuilder.from_dataframe(display_df)
    gb.configure_default_column(
        filter=True,
        sortable=True,
        resizable=True,
        floatingFilter=True,
        editable=False,
    )

    # Kolom utama dibuat lebih ramping supaya enak dibaca cepat
    gb.configure_column("KODEBARANG", header_name="KODEBARANG", minWidth=120, width=140, pinned="left")
    gb.configure_column("PRODUCT", header_name="PRODUCT", minWidth=90, width=110, pinned="left")
    gb.configure_column("BRAND", header_name="BRAND", minWidth=90, width=100, pinned="left")
    gb.configure_column("SPESIFIKASI", header_name="SPESIFIKASI", minWidth=220, width=280)
    gb.configure_column("M3", header_name="M3", type=["numericColumn"], minWidth=80, width=90)

    # Kolom period/division dibuat kecil dan dinamis
    for col in display_df.columns:
        if col not in ["KODEBARANG", "PRODUCT", "BRAND", "SPESIFIKASI", "M3"]:
            gb.configure_column(
                col,
                minWidth=78,
                width=82,
                type=["numericColumn"],
            )

    gb.configure_grid_options(
        animateRows=False,
        suppressColumnVirtualisation=False,
        suppressRowVirtualisation=False,
        enableCellTextSelection=True,
        ensureDomOrder=True,
        rowHeight=30,
        headerHeight=34,
    )

    grid_options = gb.build()

    AgGrid(
        display_df,
        gridOptions=grid_options,
        height=520,
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=False,
        columns_auto_size_mode=ColumnsAutoSizeMode.NO_AUTOSIZE,
        theme="streamlit",
        enable_enterprise_modules=False,
        reload_data=False,
    )

    return display_df

# =========================================================
# UI
# =========================================================
st.title("Dashboard Analisa Sales vs Stock")
st.caption("QTY diambil dari MPLSSR. STOK dan harga diambil dari Pricelist.")

st.sidebar.header("Upload File")
mplssr_file = st.sidebar.file_uploader("Upload MPLSSR", type=["xlsx", "xls"])
pricelist_file = st.sidebar.file_uploader("Upload Pricelist", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.write("**Rules:**")
st.sidebar.caption("""- QTY: MPLSSR
- STOK: Pricelist
- Harga: kolom M3
- MPLSSR: header row 2, data row 7
- Pricelist: header row 2, data row 6
- 05 OLR cek kode area row 3
- Sheet LAPTOP hapus COMING s/d END COMING
- Gunakan tombol PROSES setelah pilih filter
- Filter range harga di sidebar dihapus""")

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
    st.session_state["filter_submitted"] = False

if process_clicked:
    st.session_state["filter_submitted"] = True

if not st.session_state["filter_submitted"]:
    st.info("Silakan pilih filter terlebih dulu, lalu klik PROSES.")
    st.stop()

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
    seg.columns = ["SEGMENT", "DIV 03", "DIV 04", "DIV 05"]
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
    brand.columns = ["BRAND", "DIV 03", "DIV 04", "DIV 05"]
    return brand

def render_left_table(df, title):
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
    for _, row in df.iterrows():
        html.append("<tr>")
        for col in df.columns:
            val = row[col]
            try:
                if pd.notna(val) and isinstance(val, (int, float, np.integer, np.floating)):
                    display = f"{int(round(float(val))):,}".replace(",", ".")
                else:
                    display = "" if pd.isna(val) else str(val)
            except Exception:
                display = str(val)
            html.append(f'<td style="border:1px solid #2b2b2b;padding:6px;text-align:left;">{display}</td>')
        html.append("</tr>")
    html.append("</tbody></table></div></div>")
    st.markdown("".join(html), unsafe_allow_html=True)

seg_filter_col1, seg_filter_col2 = st.columns([1, 3])
with seg_filter_col1:
    segmentasi_period = st.selectbox(
        "Filter Segmentasi",
        PERIODS,
        index=0,
        key="segmentasi_period_top",
    )

left, right = st.columns(2)

with left:
    render_left_table(build_segment_table(filtered, segmentasi_period), f"Segmentasi Harga - {segmentasi_period}")

with right:
    render_left_table(build_brand_table(filtered, segmentasi_period), f"Segmentasi Brand - {segmentasi_period}")


st.markdown("### Tabel Utama Analisa")
st.caption("Tabel utama sekarang memakai AgGrid, jadi bisa filter, sort, resize kolom, dan tampil lebih mirip Excel.")

main_table = build_main_table(filtered)
main_table_export = render_main_table_aggrid(main_table)

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
