# ==================================================================
# BANK XYZ — CUSTOMER EXPERIENCE DASHBOARD (v2)
# ------------------------------------------------------------------
# Jalankan dari root project:   streamlit run dashboard.py
#
# Struktur file yang dibutuhkan (relatif terhadap file ini):
#   1. data/Deka_project_dataset_BankXYZ.csv
#   2. metadata/metadata_dashboard.xlsx  (atau .csv)  -> hasil buat_metadata.py
#
# Fitur:
#   - Navigasi halaman di SIDEBAR (Ringkasan, Brand Image, Branch
#     Facilities, Service Experience, ATM Experience)
#   - Filter tersimpan saat pindah halaman (tidak ke-reset) + tombol
#     Reset Filter + expander "Filter Tambahan"
#   - Setiap halaman: skor keseluruhan di atas, lalu sub-kategori
#     (mis. Toilet, Parkir, Ruang Tunggu) masing-masing dengan
#     overall + grafik atributnya
# ==================================================================

from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==================================================================
# KONFIGURASI
# ==================================================================

BASE_DIR = Path(__file__).parent

DATA_PATH = BASE_DIR / "data" / "Deka_project_dataset_BankXYZ.csv"
META_XLSX = BASE_DIR / "metadata" / "metadata_dashboard.xlsx"
META_CSV = BASE_DIR / "metadata" / "metadata_dashboard.csv"

# Palet biru tua – biru muda – putih
C_DARK = "#0F4C81"
C_MID = "#2E77AE"
C_LIGHT = "#7FB3D5"
C_PALE = "#D6E9F8"
C_BG = "#F4F9FF"
C_TEXT = "#0B2540"
C_RED = "#E8B4B8"
BLUE_SEQ = ["#D6E9F8", "#A9CCE3", "#7FB3D5", "#5499C7", "#2E77AE", "#0F4C81"]

PAGE_RINGKASAN = "Ringkasan"
PAGE_BI = "Brand Image"
PAGE_BF = "Branch Facilities"
PAGE_SE = "Service Experience"
PAGE_ATM = "ATM Experience"
PAGES = {
    PAGE_RINGKASAN: "📊",
    PAGE_BI: "⭐",
    PAGE_BF: "🏢",
    PAGE_SE: "🤝",
    PAGE_ATM: "🏧",
}

TP_SERVICE = ["Customer Service", "Teller", "Security",
              "Customer Advisor", "Service Electronics"]
TP_LABEL = {"Security": "Sekuriti",
            "Service Electronics": "Sarana Elektronik"}

ORDER_USIA = [
    "17 -19 tahun", "20 - 25 tahun", "26 - 30 tahun", "31 - 35 tahun",
    "36 - 40 tahun", "41 - 45 tahun", "46 - 50 tahun",
    "50 tahun dan ke atas",
]
ORDER_LAMA = [
    "1 bulan s/d 3 bulan", "3 bulan s/d 11 bulan",
    "1 tahun s/d 2 tahun 11 bulan", "3 tahun s/d 4 tahun 11 bulan",
    "5 tahun atau lebih",
]
ORDER_BY_VAR = {"S2_2": ORDER_USIA, "S4": ORDER_LAMA}

st.set_page_config(
    page_title="Bank XYZ Customer Experience Dashboard",
    page_icon="🏦",
    layout="wide",
)

# ==================================================================
# TEMA / CSS  (kontras diperbaiki: semua teks sidebar gelap di atas
# latar biru muda, kotak pilihan putih, menu dropdown putih)
# ==================================================================

CSS = """
<style>
.stApp { background-color: __BG__; }
.block-container { padding-top: 1.4rem; }

h1, h2, h3 { color: __DARK__ !important; font-weight: 800; }

/* ---------- SIDEBAR ---------- */
section[data-testid="stSidebar"] {
    background-color: #E7F0FA;
    border-right: 1px solid #C7DAEE;
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] small {
    color: __TEXT__ !important;
}
section[data-testid="stSidebar"] label { font-weight: 700 !important; }

/* kotak selectbox / multiselect: PUTIH dengan teks gelap */
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    border: 1px solid #9DBFE0 !important;
    border-radius: 8px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] span,
section[data-testid="stSidebar"] [data-baseweb="select"] div {
    color: __TEXT__ !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] svg {
    fill: __DARK__ !important;
}
/* tag pada multiselect */
section[data-testid="stSidebar"] [data-baseweb="tag"] {
    background-color: __DARK__ !important;
}
section[data-testid="stSidebar"] [data-baseweb="tag"] span,
section[data-testid="stSidebar"] [data-baseweb="tag"] svg {
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
}

/* menu dropdown (muncul di luar sidebar) */
div[data-baseweb="popover"] ul[role="listbox"] {
    background-color: #FFFFFF !important;
    border: 1px solid #9DBFE0 !important;
}
div[data-baseweb="popover"] li[role="option"],
div[data-baseweb="popover"] li[role="option"] div,
div[data-baseweb="popover"] li[role="option"] span {
    color: __TEXT__ !important;
    background-color: #FFFFFF;
}
div[data-baseweb="popover"] li[role="option"]:hover,
div[data-baseweb="popover"] li[aria-selected="true"] {
    background-color: __PALE__ !important;
}

/* navigasi halaman (radio) tampil seperti menu */
section[data-testid="stSidebar"] div[role="radiogroup"] > label {
    display: block;
    background: #FFFFFF;
    border: 1px solid #C7DAEE;
    border-radius: 10px;
    padding: 9px 12px;
    margin-bottom: 6px;
    cursor: pointer;
}
section[data-testid="stSidebar"] div[role="radiogroup"]
    > label:has(input:checked) {
    background: __DARK__;
    border-color: __DARK__;
}
section[data-testid="stSidebar"] div[role="radiogroup"]
    > label:has(input:checked) * {
    color: #FFFFFF !important;
}

/* tombol reset */
section[data-testid="stSidebar"] button {
    background-color: __DARK__ !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 8px !important;
    width: 100%;
}
section[data-testid="stSidebar"] button p { color:#FFFFFF !important; }

/* expander filter tambahan */
section[data-testid="stSidebar"] details {
    background: #FFFFFF;
    border: 1px solid #C7DAEE;
    border-radius: 10px;
}

/* ---------- KARTU METRIK ---------- */
div[data-testid="stMetric"] {
    background-color: #FFFFFF;
    border: 1px solid #D7E6F5;
    border-radius: 14px;
    padding: 14px 16px;
    box-shadow: 0 2px 10px rgba(15,76,129,0.08);
}
div[data-testid="stMetric"] label { color: __DARK__ !important;
    font-weight: 700; }
div[data-testid="stMetricValue"] { color: #0B1F33 !important;
    font-weight: 800; }
div[data-testid="stMetricDelta"] { color: #4B5563 !important; }
/* allow delta subtext and label to wrap instead of truncating */
div[data-testid="stMetricDelta"],
div[data-testid="stMetricDelta"] *,
div[data-testid="stMetricLabel"],
div[data-testid="stMetricLabel"] *,
[data-testid="stMetric"] [data-testid="stMetricDelta"],
[data-testid="stMetric"] [data-testid="stMetricDelta"] * {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    overflow-wrap: break-word !important;
    word-wrap: break-word !important;
    max-width: 100% !important;
}

/* ---------- TAB SUB-KATEGORI ---------- */
button[data-baseweb="tab"] { color: #33506B !important; font-weight: 600; }
button[data-baseweb="tab"][aria-selected="true"] {
    color: __DARK__ !important;
    border-bottom: 3px solid __DARK__;
}

div[data-testid="stExpander"] { background:#FFFFFF; border-radius:12px; }
hr { border-color: __PALE__; }

/* ---------- EXPANDER — header and body text always dark on white ---------- */
/* covers summary tag directly (Streamlit uses <details><summary>) */
details summary,
details > summary,
details summary p,
details summary span,
details summary svg,
div[data-testid="stExpander"] summary,
div[data-testid="stExpander"] summary p,
div[data-testid="stExpander"] summary span,
div[data-testid="stExpander"] summary svg {
    color: __TEXT__ !important;
    fill: __TEXT__ !important;
}
/* body text inside expander — all common inline elements */
div[data-testid="stExpander"] p,
div[data-testid="stExpander"] strong,
div[data-testid="stExpander"] em,
div[data-testid="stExpander"] span,
div[data-testid="stExpander"] li,
div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] p,
div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] strong,
div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] em,
div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] span {
    color: __TEXT__ !important;
}

/* ---------- TOGGLE / CHECKBOX WIDGET — always visible ---------- */
/* Streamlit toggle uses data-baseweb="checkbox" internally */
div[data-testid="stToggle"] label,
div[data-testid="stToggle"] label p,
div[data-testid="stToggle"] label span,
div[data-testid="stToggle"] > div > label > div:last-child,
div[data-testid="stToggle"] > div > label > div:last-child *,
label[data-baseweb="checkbox"] > div:last-child,
label[data-baseweb="checkbox"] > div:last-child p,
label[data-baseweb="checkbox"] > div:last-child span,
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] span,
[data-testid="stWidgetLabel"] * {
    color: __TEXT__ !important;
    opacity: 1 !important;
    visibility: visible !important;
}
</style>
"""
st.markdown(
    CSS.replace("__BG__", C_BG).replace("__DARK__", C_DARK)
       .replace("__PALE__", C_PALE).replace("__TEXT__", C_TEXT),
    unsafe_allow_html=True,
)

# ==================================================================
# LOAD DATA & METADATA
# ==================================================================

def to_score(series: pd.Series, max_valid: int = 6) -> pd.Series:
    """'6  SANGAT PUAS' -> 6 ; '99 TIDAK RELEVAN' / kosong -> NaN."""
    s = series.astype(str).str.extract(r"^\s*(\d+)")[0]
    s = pd.to_numeric(s, errors="coerce")
    s = s.where(s <= max_valid)
    return s.astype("Float64")


@st.cache_data(show_spinner="Memuat data…")
def load_all():
    df = pd.read_csv(DATA_PATH, sep=";", header=1,
                     low_memory=False, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]

    if META_XLSX.exists():
        meta = pd.read_excel(META_XLSX, sheet_name="Metadata")
    else:
        meta = pd.read_csv(META_CSV, encoding="utf-8-sig")
    meta.columns = [str(c).strip() for c in meta.columns]
    for c in ("variable", "question", "label", "bank", "section",
              "touchpoint", "subgroup", "role"):
        meta[c] = meta[c].astype(str).str.strip()
    meta["include"] = pd.to_numeric(meta["include"],
                                    errors="coerce").fillna(0).astype(int)
    meta["scale_max"] = pd.to_numeric(meta["scale_max"],
                                      errors="coerce").fillna(6).astype(int)
    # hanya variabel yang benar-benar ada di data
    meta = meta[meta["variable"].isin(df.columns)].reset_index(drop=True)
    return df, meta


if not DATA_PATH.exists():
    st.error("File **data/Deka_project_dataset_BankXYZ.csv** tidak ditemukan. "
             "Pastikan file data ada di folder `data/` pada root project.")
    st.stop()
if not META_XLSX.exists() and not META_CSV.exists():
    st.error("Metadata dashboard belum ada. Jalankan dulu "
             "`python preprocessing/buat_metadata.py` untuk membuat "
             "**metadata/metadata_dashboard.xlsx / .csv**.")
    st.stop()

df, META = load_all()

ATTR = META[(META["role"] == "Atribut") & (META["include"] == 1)
            & (META["bank"] == "XYZ")]
OVERALL = META[(META["role"] == "Overall") & (META["include"] == 1)
               & (META["bank"] == "XYZ")]
EXTRA_FILTERS = META[(META["role"].str.startswith("Filter"))
                     & (META["subgroup"] == "Tambahan")]
# Rev 2 — Loyalty Drivers; Rev 3 — Brand Image Importance
LOYALTY_DRIVERS = META[(META["role"] == "Loyalty Driver") & (META["include"] == 1)]
IMPORTANCE = META[(META["role"] == "Importance") & (META["include"] == 1)]


def attrs_of(section=None, touchpoint=None) -> pd.DataFrame:
    sub = ATTR
    if section:
        sub = sub[sub["section"] == section]
    if touchpoint:
        sub = sub[sub["touchpoint"] == touchpoint]
    return sub


def overall_of(section=None, touchpoint=None) -> pd.DataFrame:
    sub = OVERALL
    if section:
        sub = sub[sub["section"] == section]
    if touchpoint:
        sub = sub[sub["touchpoint"] == touchpoint]
    return sub


# ==================================================================
# SIDEBAR — NAVIGASI + FILTER (persisten antar halaman)
# ==================================================================

MAIN_FILTERS = [
    ("f_prov", "PROV", "Provinsi"),
    ("f_kab", "KABKOTA", "Kabupaten/Kota"),
    ("f_cab", "CABANG", "Cabang"),
    ("f_usia", "S2_2", "Kelompok Usia"),
    ("f_lama", "S4", "Lama Menjadi Nasabah"),
]


def ordered_options(values, var):
    vals = pd.Series(values).dropna().unique().tolist()
    order = ORDER_BY_VAR.get(var)
    if order:
        return [v for v in order if v in vals] + sorted(
            v for v in vals if v not in order)
    return sorted(vals)


def reset_filters():
    for key, _, _ in MAIN_FILTERS:
        st.session_state[key] = "Semua"
    for var in EXTRA_FILTERS["variable"]:
        st.session_state[f"fx_{var}"] = []


def select_persist(key, label, options):
    """Selectbox ber-key: nilai tersimpan di session_state sehingga
    tidak ke-reset saat pindah halaman; nilai tak valid dikembalikan
    ke 'Semua' (mis. saat pilihan kab/kota berubah karena provinsi)."""
    opts = ["Semua"] + options
    if key not in st.session_state or st.session_state[key] not in opts:
        st.session_state[key] = "Semua"
    return st.sidebar.selectbox(label, opts, key=key)


with st.sidebar:
    st.title("🏦 Bank XYZ")
    page = st.radio("Halaman", list(PAGES),
                    format_func=lambda p: f"{PAGES[p]}  {p}",
                    key="nav_page", label_visibility="collapsed")
    st.markdown("---")
    st.subheader("📌 Filter Data")

# Filter utama (berjenjang: Provinsi -> Kab/Kota -> Cabang)
prov = select_persist("f_prov", "Provinsi",
                      ordered_options(df["PROV"], "PROV"))
_t = df if prov == "Semua" else df[df["PROV"] == prov]
kab = select_persist("f_kab", "Kabupaten/Kota",
                     ordered_options(_t["KABKOTA"], "KABKOTA"))
if kab != "Semua":
    _t = _t[_t["KABKOTA"] == kab]
cab = select_persist("f_cab", "Cabang",
                     ordered_options(_t["CABANG"], "CABANG"))
usia = select_persist("f_usia", "Kelompok Usia",
                      ordered_options(df["S2_2"], "S2_2"))
lama = select_persist("f_lama", "Lama Menjadi Nasabah",
                      ordered_options(df["S4"], "S4"))

# Filter tambahan (multiselect, kosong = semua)
with st.sidebar.expander("⚙️ Filter Tambahan"):
    for _, r in EXTRA_FILTERS.iterrows():
        var = r["variable"]
        st.multiselect(r["label"], ordered_options(df[var], var),
                       key=f"fx_{var}", placeholder="Semua")

st.sidebar.button("🔄 Reset Semua Filter", on_click=reset_filters)

# Terapkan filter
fdf = df
for key, col, _ in MAIN_FILTERS:
    val = st.session_state.get(key, "Semua")
    if val != "Semua":
        fdf = fdf[fdf[col] == val]
for _, r in EXTRA_FILTERS.iterrows():
    sel = st.session_state.get(f"fx_{r['variable']}", [])
    if sel:
        fdf = fdf[fdf[r["variable"]].isin(sel)]

st.sidebar.markdown("---")
st.sidebar.metric("Jumlah Responden Terfilter", f"{len(fdf):,}")
if len(fdf) == 0:
    st.warning("Tidak ada responden yang sesuai dengan kombinasi filter ini. "
               "Silakan longgarkan filter atau tekan Reset.")
    st.stop()
if len(fdf) < 30:
    st.sidebar.caption("⚠️ Responden < 30 — hasil agregat kurang "
                       "representatif.")

# ==================================================================
# FUNGSI PERHITUNGAN & GRAFIK
# ==================================================================

def mean_score(data, var, scale=6):
    if var not in data.columns:
        return None
    s = to_score(data[var], scale)
    return None if s.notna().sum() == 0 else float(s.mean())


def t2b(data, var, scale=6):
    s = to_score(data[var], scale).dropna()
    if len(s) == 0:
        return None
    return float(s.isin([scale - 1, scale]).mean() * 100)


def nps(data, var="G1A"):
    s = to_score(data[var], 10).dropna()
    if len(s) == 0:
        return None, None, None, None
    p = float((s >= 9).mean() * 100)
    d = float((s <= 6).mean() * 100)
    return p - d, p, 100 - p - d, d


def attribute_table(data, attr_rows: pd.DataFrame) -> pd.DataFrame:
    """Skor & T2B per atribut (1 baris per variabel XYZ)."""
    rows = []
    for _, r in attr_rows.iterrows():
        m = mean_score(data, r["variable"])
        if m is None:
            continue
        rows.append({"variable": r["variable"], "Atribut": r["label"],
                     "Subkategori": r["subgroup"], "Skor": round(m, 2),
                     "T2B": t2b(data, r["variable"])})
    return pd.DataFrame(rows)


def style_fig(fig, height=420):
    fig.update_layout(
        height=height, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(color="#0B1F33", size=13),
        legend=dict(font=dict(size=13, color="#0B1F33")),
        title_font=dict(color=C_DARK, size=16),
        margin=dict(l=10, r=10, t=48, b=10),
    )
    fig.update_xaxes(tickfont=dict(color="#0B1F33", size=12),
                     title_font=dict(color="#0B1F33", size=12))
    fig.update_yaxes(tickfont=dict(color="#0B1F33", size=12),
                     title_font=dict(color="#0B1F33", size=12))
    return fig


PLOTLY_CFG = {"displaylogo": False,
              "modeBarButtonsToRemove": ["lasso2d", "select2d"]}


def bar_attributes(am: pd.DataFrame, title: str):
    am = am.sort_values("Skor")
    fig = px.bar(am, x="Skor", y="Atribut", orientation="h",
                 color="Skor", color_continuous_scale=BLUE_SEQ,
                 range_color=(max(1, am["Skor"].min() - 0.05),
                              min(6, am["Skor"].max() + 0.05)),
                 text="Skor", title=title)
    fig.update_traces(textposition="outside", cliponaxis=False,
                      texttemplate="%{text:.2f}")
    fig.update_coloraxes(showscale=False)
    fig.update_layout(xaxis_range=[1, 6.6], yaxis_title=None,
                      xaxis_title="Skor (1–6)")
    st.plotly_chart(style_fig(fig, max(230, 30 * len(am) + 110)),
                    width="stretch", config=PLOTLY_CFG)


def bar_subgroups(am: pd.DataFrame, title: str):
    g = (am.groupby("Subkategori", as_index=False)
           .agg(Skor=("Skor", "mean"), Atribut=("Atribut", "size"))
           .sort_values("Skor"))
    g["Skor"] = g["Skor"].round(2)
    fig = px.bar(g, x="Skor", y="Subkategori", orientation="h",
                 color="Skor", color_continuous_scale=BLUE_SEQ,
                 text="Skor", title=title,
                 hover_data={"Atribut": True})
    fig.update_traces(textposition="outside", cliponaxis=False,
                      texttemplate="%{text:.2f}")
    fig.update_coloraxes(showscale=False)
    fig.update_layout(xaxis_range=[1, 6.6], yaxis_title=None,
                      xaxis_title="Rata-rata skor sub-kategori")
    st.plotly_chart(style_fig(fig, max(200, 36 * len(g) + 100)),
                    width="stretch", config=PLOTLY_CFG)


def heatmap_usia(data, attr_rows, title, max_attrs=8):
    base = attribute_table(data, attr_rows)
    if base.empty:
        return
    pick = base.sort_values("Skor").head(max_attrs)
    usia_groups = [u for u in ORDER_USIA if u in data["S2_2"].unique()]
    if not usia_groups:
        return
    matrix, labels = [], []
    for _, r in pick.iterrows():
        row = [mean_score(data[data["S2_2"] == u], r["variable"])
               for u in usia_groups]
        matrix.append([None if v is None else round(v, 2) for v in row])
        labels.append(r["Atribut"])
    fig = px.imshow(matrix,
                    x=[u.replace(" tahun", "") for u in usia_groups],
                    y=labels, color_continuous_scale=BLUE_SEQ,
                    text_auto=".2f", aspect="auto", title=title,
                    labels=dict(color="Skor"))
    st.plotly_chart(style_fig(fig, max(330, 34 * len(labels) + 140)),
                    width="stretch", config=PLOTLY_CFG)
    st.caption("Menampilkan atribut dengan skor terendah (prioritas "
               "perbaikan). Semakin gelap = semakin puas.")


def fmt(v, suffix=""):
    return "-" if v is None else f"{v:.2f}{suffix}"


# Rev 1 — actual respondent count for a touchpoint in the current filtered data
def tp_n_actual(data: pd.DataFrame, touchpoint: str) -> int:
    rows = ATTR[ATTR["touchpoint"] == touchpoint]
    if rows.empty:
        return 0
    col = rows.iloc[0]["variable"]
    if col not in data.columns:
        return 0
    return int(to_score(data[col], 6).notna().sum())


# Rev 3 — Importance-Performance Analysis scatter for Brand Image
def ipa_chart(data: pd.DataFrame, perf_rows: pd.DataFrame,
              imp_rows: pd.DataFrame, title: str):
    perf = []
    for _, r in perf_rows.iterrows():
        m = mean_score(data, r["variable"])
        if m is not None:
            perf.append({"label": r["label"], "Performance": round(m, 2),
                         "Subkategori": r["subgroup"]})
    imp = []
    for _, r in imp_rows.iterrows():
        m = mean_score(data, r["variable"])
        if m is not None:
            imp.append({"label": r["label"], "Importance": round(m, 2)})
    if not perf or not imp:
        return
    merged = (pd.DataFrame(perf)
              .merge(pd.DataFrame(imp), on="label", how="inner"))
    if merged.empty:
        return
    avg_p = merged["Performance"].mean()
    avg_i = merged["Importance"].mean()
    fig = px.scatter(
        merged, x="Performance", y="Importance",
        hover_name="label", color="Subkategori",
        title=title,
        color_discrete_sequence=BLUE_SEQ,
    )
    fig.add_vline(x=avg_p, line_dash="dash", line_color="#B0C8E0", line_width=1)
    fig.add_hline(y=avg_i, line_dash="dash", line_color="#B0C8E0", line_width=1)
    # quadrant labels (placed at corners, offset from the mean lines)
    pad = 0.02
    x_lo = merged["Performance"].min() - pad
    x_hi = merged["Performance"].max() + pad
    y_lo = merged["Importance"].min() - pad
    y_hi = merged["Importance"].max() + pad
    for txt, ax, ay, color in [
        ("⚠️ Prioritas Perbaikan", x_lo, y_hi, "#C0392B"),
        ("✅ Pertahankan",          x_hi, y_hi, "#0F4C81"),
        ("Monitor",                 x_lo, y_lo, "#9DBFE0"),
        ("Efisiensi Lebih",         x_hi, y_lo, "#9DBFE0"),
    ]:
        fig.add_annotation(x=ax, y=ay, text=txt, showarrow=False,
                           font=dict(color=color, size=11),
                           xanchor="left" if ax == x_lo else "right")
    fig.update_traces(marker=dict(size=11, line=dict(width=1, color="#FFFFFF")))
    fig.update_coloraxes(showscale=False)
    fig.update_layout(
        xaxis=dict(title="Skor Kepuasan / Performance (1–6)",
                   range=[x_lo - 0.05, x_hi + 0.05]),
        yaxis=dict(title="Skor Kepentingan / Importance (1–6)",
                   range=[y_lo - 0.05, y_hi + 0.05]),
        legend_title="Sub-Kategori",
    )
    st.plotly_chart(style_fig(fig, 520), width="stretch", config=PLOTLY_CFG)
    st.caption(
        "Kuadran dibagi oleh rata-rata skor kepuasan (vertikal) dan kepentingan (horizontal). "
        "**Kiri atas** = penting tapi kurang memuaskan → prioritas perbaikan. "
        "**Kanan atas** = penting dan sudah memuaskan → pertahankan. "
        "**Kanan bawah** = memuaskan tapi kurang penting → efisiensi lebih. "
        "**Kiri bawah** = kurang penting dan kurang memuaskan → monitor."
    )


def overall_metric_cards(data, ov_rows: pd.DataFrame):
    """Single aggregated KPI: mean of all holistic overall questions for this
    section/touchpoint.  Avoids showing one arbitrary sub-area and hides the
    survey boilerplate ('Dengan mempertimbangkan berbagai hal…')."""
    if ov_rows.empty:
        return
    scores = [mean_score(data, r["variable"]) for _, r in ov_rows.iterrows()]
    scores = [s for s in scores if s is not None]
    if not scores:
        return
    avg = round(sum(scores) / len(scores), 2)
    n = len(scores)
    delta = f"rata-rata dari {n} penilaian holistik" if n > 1 else "penilaian holistik responden"
    st.metric("Penilaian Keseluruhan",
              f"{avg:.2f} / 6",
              delta,
              delta_color="off")


def render_section_page(title, caption, attr_rows, ov_rows,
                        heat_key, custom_card=None):
    """Render a full section page.

    custom_card: optional tuple (label, value, delta) to fill the 3rd KPI
    slot when the section has no Overall-role variable in metadata
    (e.g. Brand Image).  If None, falls back to overall_metric_cards().
    """
    st.header(f"{PAGES.get(title, '')} {title}")
    st.caption(caption)

    am = attribute_table(fdf, attr_rows)
    if am.empty:
        st.info("Tidak ada data atribut untuk bagian ini "
                "(periksa kolom include pada metadata).")
        return

    t2b_avg = am["T2B"].dropna().mean() if am["T2B"].notna().any() else None
    cols = st.columns(4)
    cols[0].metric("Skor Keseluruhan", f"{am['Skor'].mean():.2f} / 6",
                   f"{len(am)} atribut", delta_color="off")
    cols[1].metric("Rata-rata % Puas (T2B)",
                   "-" if t2b_avg is None else f"{t2b_avg:.0f}%",
                   "jawaban 5–6", delta_color="off")
    with cols[2]:
        if custom_card is not None:
            st.metric(custom_card[0], custom_card[1],
                      custom_card[2] if len(custom_card) > 2 else None,
                      delta_color="off")
        else:
            overall_metric_cards(fdf, ov_rows)
    cols[3].metric("Responden", f"{len(fdf):,}")

    st.markdown("---")

    n_sub = am["Subkategori"].nunique()
    if n_sub > 1:
        bar_subgroups(am, f"Skor per Sub-Kategori — {title}")

    # detail per sub-kategori (overall sub-kategori di atas + atribut)
    order = (am.groupby("Subkategori")["Skor"].mean()
               .sort_values().index.tolist())
    tabs = st.tabs([f"{s}" for s in order]) if n_sub > 1 else [st.container()]
    for tab, sub in zip(tabs, order):
        with tab:
            sam = am[am["Subkategori"] == sub]
            st2b = (sam["T2B"].dropna().mean()
                    if sam["T2B"].notna().any() else None)
            c1, c2, c3 = st.columns(3)
            c1.metric(f"Overall {sub}", f"{sam['Skor'].mean():.2f} / 6",
                      delta_color="off")
            c2.metric("% Puas (T2B)",
                      "-" if st2b is None else f"{st2b:.0f}%",
                      delta_color="off")
            c3.metric("Jumlah Atribut", f"{len(sam)}")
            bar_attributes(sam, f"Kepuasan per Atribut — {sub}")

    with st.expander("🎯 Prioritas Perbaikan & Kekuatan Utama"):
        l, r = st.columns(2)
        show_cols = {"Atribut": "Atribut", "Subkategori": "Sub-Kategori",
                     "Skor": "Skor", "T2B": "% T2B"}
        low = am.sort_values("Skor").head(5)[list(show_cols)]
        high = am.sort_values("Skor", ascending=False).head(5)[list(show_cols)]
        with l:
            st.markdown("**🔻 5 skor terendah (prioritas)**")
            st.dataframe(low.rename(columns=show_cols), hide_index=True,
                         width="stretch")
        with r:
            st.markdown("**🔺 5 skor tertinggi (kekuatan)**")
            st.dataframe(high.rename(columns=show_cols), hide_index=True,
                         width="stretch")

    if st.toggle("Tampilkan perbandingan antar kelompok usia (heatmap)",
                 key=heat_key):
        heatmap_usia(fdf, attr_rows,
                     f"Atribut Prioritas {title} per Kelompok Usia")


# ==================================================================
# HALAMAN 1 — RINGKASAN
# ==================================================================

def page_ringkasan():
    st.title("🏦 Bank XYZ Customer Experience Dashboard")
    st.caption("Monitoring **CSAT**, **Loyalty**, **NPS**, dan **Customer "
               "Experience** Bank XYZ berdasarkan survei nasabah. Gunakan "
               "filter di sidebar — pilihan filter tetap tersimpan saat "
               "berpindah halaman.")

    csat_m, csat_p = mean_score(fdf, "E1A"), t2b(fdf, "E1A")
    loy_m, loy_p = mean_score(fdf, "F1A"), t2b(fdf, "F1A")
    nps_val, prom, pas, det = nps(fdf)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CSAT (Kepuasan)", fmt(csat_m, " / 6"),
              None if csat_p is None else f"{csat_p:.0f}% puas (T2B)",
              delta_color="off")
    c2.metric("Loyalty", fmt(loy_m, " / 6"),
              None if loy_p is None else f"{loy_p:.0f}% setuju (T2B)",
              delta_color="off")
    c3.metric("NPS", "-" if nps_val is None else f"{nps_val:.0f}",
              None if prom is None else f"{prom:.0f}% promoter",
              delta_color="off")
    c4.metric("Responden", f"{len(fdf):,}")

    st.markdown("---")
    col_l, col_r = st.columns(2)

    # Radar skor per touchpoint
    with col_l:
        tp_scores = []
        for tp in ATTR["touchpoint"].unique():
            am = attribute_table(fdf, attrs_of(touchpoint=tp))
            if not am.empty:
                tp_scores.append((TP_LABEL.get(tp, tp),
                                  round(am["Skor"].mean(), 2)))
        if tp_scores:
            cats = [t[0] for t in tp_scores]
            vals = [t[1] for t in tp_scores]
            rmin = min(4.0, min(vals) - 0.3)
            fig = go.Figure(go.Scatterpolar(
                r=vals + vals[:1], theta=cats + cats[:1],
                fill="toself", line_color=C_MID,
                fillcolor="rgba(46,119,174,0.25)"))
            fig.update_layout(
                title="Skor Customer Experience per Touchpoint",
                polar=dict(radialaxis=dict(range=[rmin, 6], visible=True)),
                showlegend=False)
            st.plotly_chart(style_fig(fig, 430),
                            width="stretch", config=PLOTLY_CFG)

    with col_r:
        # Komposisi NPS
        if nps_val is not None:
            nps_df = pd.DataFrame({
                "Kategori": ["Promoter (9-10)", "Passive (7-8)",
                             "Detractor (0-6)"],
                "Persen": [prom, pas, det]})
            fig = px.bar(nps_df, x="Persen", y=["NPS"] * 3,
                         color="Kategori", orientation="h", text="Persen",
                         color_discrete_sequence=[C_DARK, C_LIGHT, C_RED],
                         title="Komposisi NPS Bank XYZ")
            fig.update_traces(texttemplate="%{x:.0f}%",
                              textposition="inside")
            fig.update_layout(barmode="stack", yaxis_title=None,
                              xaxis_title="%", yaxis_visible=False,
                              legend_title=None,
                              legend=dict(orientation="h", y=-0.35))
            st.plotly_chart(style_fig(fig, 200),
                            width="stretch", config=PLOTLY_CFG)

        # XYZ vs Kompetitor
        comp = []
        for label, vx, vk, sc in [("CSAT", "E1A", "E1B", 6),
                                  ("Loyalty", "F1A", "F1B", 6),
                                  ("NPS (rata-rata 0-10)", "G1A", "G1C", 10)]:
            mx, mk = mean_score(fdf, vx, sc), mean_score(fdf, vk, sc)
            if mx is not None:
                comp.append({"Indikator": label, "Bank": "XYZ", "Skor": mx})
            if mk is not None:
                comp.append({"Indikator": label, "Bank": "Kompetitor",
                             "Skor": mk})
        if comp:
            fig = px.bar(pd.DataFrame(comp), x="Indikator", y="Skor",
                         color="Bank", barmode="group", text_auto=".2f",
                         color_discrete_map={"XYZ": C_DARK,
                                             "Kompetitor": "#94A3B8"},
                         title="Bank XYZ vs Kompetitor")
            fig.update_layout(legend_title=None)
            st.plotly_chart(style_fig(fig, 320),
                            width="stretch", config=PLOTLY_CFG)

    # Prioritas perbaikan lintas touchpoint
    st.markdown("---")
    all_am = attribute_table(fdf, ATTR)
    if not all_am.empty:
        tp_of = dict(zip(ATTR["variable"], ATTR["touchpoint"]))
        low = all_am.sort_values("Skor").head(8).copy()
        low["Atribut"] = [
            f"{TP_LABEL.get(tp_of.get(v, ''), tp_of.get(v, ''))} — {a}"
            for v, a in zip(low["variable"], low["Atribut"])]
        bar_attributes(low, "🎯 8 Atribut Prioritas Perbaikan "
                            "(skor terendah, seluruh touchpoint)")

    # Rev 2 — Loyalty Drivers: why do customers stay? (T_H* battery)
    if not LOYALTY_DRIVERS.empty:
        st.markdown("---")
        st.subheader("🔗 Driver Loyalitas — Mengapa Nasabah Tetap Setia?")
        st.caption(
            "15 dimensi yang menjelaskan *mengapa* nasabah berniat tetap menggunakan Bank XYZ "
            "(bukan hanya seberapa tinggi niatnya). Diurutkan dari skor tertinggi ke terendah."
        )
        ld_rows = []
        for _, r in LOYALTY_DRIVERS.iterrows():
            m = mean_score(fdf, r["variable"])
            if m is not None:
                ld_rows.append({"Driver": r["label"], "Skor": round(m, 2)})
        if ld_rows:
            ld_df = pd.DataFrame(ld_rows).sort_values("Skor")
            fig = px.bar(
                ld_df, x="Skor", y="Driver", orientation="h",
                color="Skor", color_continuous_scale=BLUE_SEQ,
                range_color=(max(1, ld_df["Skor"].min() - 0.1),
                             min(6, ld_df["Skor"].max() + 0.1)),
                text="Skor",
                title="Skor per Driver Loyalitas (skala 1–6)",
            )
            fig.update_traces(textposition="outside", cliponaxis=False,
                              texttemplate="%{text:.2f}")
            fig.update_coloraxes(showscale=False)
            fig.update_layout(xaxis_range=[1, 6.6], yaxis_title=None,
                              xaxis_title="Skor (1–6)")
            st.plotly_chart(
                style_fig(fig, max(300, 32 * len(ld_df) + 120)),
                width="stretch", config=PLOTLY_CFG,
            )

    # Profil responden
    st.markdown("---")
    st.subheader("👤 Profil Responden")
    p1, p2, p3 = st.columns(3)
    if "S1" in fdf.columns:
        with p1:
            g = fdf["S1"].value_counts().reset_index()
            g.columns = ["Gender", "Jumlah"]
            fig = px.pie(g, names="Gender", values="Jumlah",
                         title="Gender", hole=0.45,
                         color_discrete_sequence=[C_DARK, C_LIGHT, C_PALE])
            st.plotly_chart(style_fig(fig, 320),
                            width="stretch", config=PLOTLY_CFG)
    if "S2_2" in fdf.columns:
        with p2:
            u = (fdf["S2_2"].value_counts()
                 .reindex(ORDER_USIA).dropna().reset_index())
            u.columns = ["Usia", "Jumlah"]
            fig = px.bar(u, x="Usia", y="Jumlah", title="Kelompok Usia",
                         text="Jumlah", color_discrete_sequence=[C_MID])
            fig.update_layout(xaxis_title=None)
            st.plotly_chart(style_fig(fig, 320),
                            width="stretch", config=PLOTLY_CFG)
    if "S4" in fdf.columns:
        with p3:
            l = (fdf["S4"].value_counts()
                 .reindex(ORDER_LAMA).dropna().reset_index())
            l.columns = ["Lama Menjadi Nasabah", "Jumlah"]
            fig = px.bar(l, x="Jumlah", y="Lama Menjadi Nasabah",
                         orientation="h", title="Lama Menjadi Nasabah",
                         text="Jumlah", color_discrete_sequence=[C_MID])
            fig.update_layout(yaxis_title=None)
            st.plotly_chart(style_fig(fig, 320),
                            width="stretch", config=PLOTLY_CFG)


# ==================================================================
# HALAMAN SERVICE EXPERIENCE (5 touchpoint petugas)
# ==================================================================

def page_service():
    st.header(f"{PAGES[PAGE_SE]} Service Experience")
    st.caption("Pengalaman layanan dari petugas cabang: Customer Service, "
               "Teller, Sekuriti, Customer Advisor, dan Sarana Elektronik "
               "Layanan.")

    avail = [tp for tp in TP_SERVICE
             if not attrs_of(touchpoint=tp).empty]
    if not avail:
        st.info("Tidak ada atribut Service Experience pada metadata.")
        return

    cols = st.columns(len(avail))
    for i, tp in enumerate(avail):
        am = attribute_table(fdf, attrs_of(touchpoint=tp))
        n_tp = tp_n_actual(fdf, tp)
        # Rev 1: surface actual n; flag when sample is small
        n_label = f"n={n_tp:,}"
        if n_tp < 100:
            n_label = f"⚠️ n={n_tp} (sampel kecil)"
        cols[i].metric(TP_LABEL.get(tp, tp),
                       "-" if am.empty else f"{am['Skor'].mean():.2f}",
                       f"{len(am)} atribut · {n_label}", delta_color="off")

    st.markdown("---")
    # pilihan touchpoint juga dibuat persisten antar halaman
    cur = st.session_state.get("se_tp_keep", avail[0])
    if cur not in avail:
        cur = avail[0]
    sel = st.selectbox("Pilih area layanan untuk detail:", avail,
                       index=avail.index(cur),
                       format_func=lambda t: TP_LABEL.get(t, t))
    st.session_state["se_tp_keep"] = sel

    render_section_page(
        TP_LABEL.get(sel, sel),
        f"Detail kepuasan layanan {TP_LABEL.get(sel, sel)} per "
        f"sub-kategori dan atribut.",
        attrs_of(touchpoint=sel), overall_of(touchpoint=sel),
        heat_key=f"hm_se_{sel}")


# ==================================================================
# HALAMAN BRAND IMAGE (Rev 3 — IPA scatter + existing bar charts)
# ==================================================================

def page_brand_image():
    # Importance mean (T_C1A* battery)
    imp_bi = IMPORTANCE[IMPORTANCE["section"] == PAGE_BI]
    imp_scores = [mean_score(fdf, r["variable"]) for _, r in imp_bi.iterrows()]
    imp_scores = [s for s in imp_scores if s is not None]
    imp_avg = round(sum(imp_scores) / len(imp_scores), 2) if imp_scores else None

    # Performance mean (T_C1B* attributes) — for gap calculation
    perf_bi_rows = attrs_of(section=PAGE_BI)
    perf_scores = [mean_score(fdf, r["variable"]) for _, r in perf_bi_rows.iterrows()]
    perf_scores = [s for s in perf_scores if s is not None]
    perf_avg = round(sum(perf_scores) / len(perf_scores), 2) if perf_scores else None

    # Gap: importance − performance.
    # Positive → customers expect more than the bank delivers (underdelivering).
    # Negative → bank delivers beyond what customers expect (safe zone).
    if imp_avg is not None and perf_avg is not None:
        gap = round(imp_avg - perf_avg, 2)
        gap_label = (
            f"gap: {gap:+.2f} — nasabah mengharapkan lebih"
            if gap > 0 else
            f"gap: {gap:+.2f} — bank melampaui ekspektasi"
        )
    else:
        gap_label = f"{len(imp_scores)} atribut kepentingan"

    importance_card = (
        "Rata-rata Kepentingan",
        "-" if imp_avg is None else f"{imp_avg:.2f} / 6",
        gap_label,
    )

    render_section_page(
        PAGE_BI,
        "Persepsi nasabah terhadap citra Bank XYZ (skala 1\u20136, jawaban "
        "'tidak relevan' dikeluarkan dari perhitungan).",
        attrs_of(section=PAGE_BI), overall_of(section=PAGE_BI),
        heat_key="hm_bi", custom_card=importance_card,
    )
    # Rev 3 — IPA: only show when importance data is available
    imp_bi = IMPORTANCE[IMPORTANCE["section"] == PAGE_BI]
    perf_bi = attrs_of(section=PAGE_BI)
    if not imp_bi.empty and not perf_bi.empty:
        st.markdown("---")
        st.subheader("\U0001f4cc Importance-Performance Analysis (IPA)")
        st.caption(
            "Setiap titik = satu atribut Brand Image. "
            "Sumbu X = skor kepuasan/performance (T_C1B). "
            "Sumbu Y = skor kepentingan/importance (T_C1A). "
            "Garis putus-putus = rata-rata masing-masing."
        )
        ipa_chart(fdf, perf_bi, imp_bi, "IPA — Brand Image Bank XYZ")


# ==================================================================
# HALAMAN BRANCH FACILITIES
# ==================================================================

def page_branch_facilities():
    render_section_page(
        PAGE_BF,
        "Penilaian nasabah terhadap fasilitas fisik cabang Bank XYZ "
        "(skala 1\u20136).",
        attrs_of(section=PAGE_BF), overall_of(section=PAGE_BF),
        heat_key="hm_bf",
    )


# ==================================================================
# HALAMAN ATM EXPERIENCE
# ==================================================================

def page_atm():
    render_section_page(
        PAGE_ATM,
        "Pengalaman nasabah menggunakan ATM Bank XYZ (skala 1\u20136).",
        attrs_of(section=PAGE_ATM), overall_of(section=PAGE_ATM),
        heat_key="hm_atm",
    )


# ==================================================================
# ROUTING UTAMA
# ==================================================================

if page == PAGE_RINGKASAN:
    page_ringkasan()
elif page == PAGE_BI:
    page_brand_image()
elif page == PAGE_BF:
    page_branch_facilities()
elif page == PAGE_SE:
    page_service()
elif page == PAGE_ATM:
    page_atm()
