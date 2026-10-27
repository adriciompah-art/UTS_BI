# ============================================================
# UTS BUSINESS INTELLIGENCE & DATA MINING
# Dashboard Interaktif: EDA → Preprocessing → Apriori → K-Means → BI → Rekomendasi
# Jalankan: streamlit run app.py
# ============================================================
import streamlit as st
import pandas as pd
import numpy as np
import os
import base64
import io
import datetime
import matplotlib.pyplot as plt
import seaborn as sns
import networkx as nx
import plotly.express as px
import plotly.graph_objects as go
import warnings
warnings.filterwarnings('ignore')

from mlxtend.frequent_patterns import apriori, association_rules
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score

# ============================================================
# 0. KONFIGURASI HALAMAN + TEMA
# ============================================================
st.set_page_config(
    page_title="UTS Business Intelligence Kelas B — UNIMOR",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (10, 6)
plt.rcParams['font.size'] = 10

FILE_PATH = 'P8. UTS_Business Intelligence.xlsx'
SHEET_RAW = 'Data_Raw'

# Default parameter (bisa diubah interaktif di Tab Apriori / K-Means)
DEFAULT_SUPPORT = 0.02
DEFAULT_CONF = 0.30
DEFAULT_LIFT = 1.0
K_CANDIDATES = list(range(2, 9))
DEFAULT_K = 4

# ---------- Helper format Indonesia (titik ribuan) ----------
def fmt_rp(x):
    """Format Rupiah Indonesia: Rp 1.234.567 (titik, tanpa desimal)."""
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return "-"
        return "Rp " + f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "-"

def fmt_id(x, desimal=0):
    """Format angka Indonesia dengan titik ribuan. desimal=0/1/2."""
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return "-"
        if desimal == 0:
            return f"{float(x):,.0f}".replace(",", ".")
        if desimal == 1:
            s = f"{float(x):,.1f}"          # 1,234.5 -> 1.234,5
            return s.replace(",", "X").replace(".", ",").replace("X", ".")
        s = f"{float(x):,.2f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "-"

def fmt_pct(x, desimal=2):
    try:
        return fmt_id(float(x) * 100, desimal) + "%"
    except Exception:
        return "-"

def ticks_rp(vals, n=5):
    """Buat (tickvals, ticktext) Rupiah bertitik untuk sumbu Plotly."""
    try:
        import numpy as np
        vals = [float(v) for v in vals if v is not None]
        vmax = max(vals) if vals else 1.0
        if vmax <= 0:
            vmax = 1.0
        ticks = np.linspace(0, vmax, n).tolist()
        labels = ["Rp " + fmt_id(v) for v in ticks]
        return ticks, labels
    except Exception:
        return None, None


# ============================================================
# 0b. IDENTITAS UTS KELAS B + ASET GAMBAR (UNIMOR & MAHASISWA)
# ------------------------------------------------------------
# PENTING — kejujuran data:
# - Proyek ini TIDAK berisi data identitas/nim/foto di file Excel/Word,
#   jadi semua kolom identitas default KOSONG (placeholder "—").
# - Isi identitas lewat panel "Lengkapi identitas" di Dashboard
#   atau letakkan file gambar di folder assets/ (lihat assets/README.md).
# - Jangan mengarang NIM/nama/foto: hanya tampilkan data yang diisi pemilik.
# ============================================================
IDENTITAS_DEFAULT = {
    "nama": "Adrianus S. Ciompah",  # Nama lengkap mahasiswa
    "nim": "51230033",              # NIM/NPM
    "prodi": "Teknologi Informasi",  # Program studi
    "fakultas": "",                 # Fakultas (belum diketahui — tampil "—")
    "matkul": "Business Intelligence",
    "kelas": "B",
    "keterangan": "Ujian Tengah Semester (UTS)",
    "semester": "7",                # Semester
    "tahun_akademik": "2026/2027",  # Tahun akademik
    "dosen": "",                    # Dosen pengampu (belum diketahui — tampil "—")
}
ASSET_DIR = "UTS_BI"
ASSET_FOTO = [os.path.join(ASSET_DIR, f) for f in
              ("foto_mahasiswa.png", "foto_mahasiswa.jpg", "foto_mahasiswa.jpeg")]
ASSET_LOGO = [os.path.join(ASSET_DIR, f) for f in
              ("logo_unimor.png", "logo_unimor.jpg", "logo_unimor.jpeg")]
ASSET_KAMPUS = [os.path.join(ASSET_DIR, f) for f in
                ("gedung_unimor.jpg", "gedung_unimor.jpeg", "gedung_unimor.png",
                 "kampus_unimor.jpg", "kampus_unimor.png")]

for _k, _v in (("id_nama", IDENTITAS_DEFAULT["nama"]), ("id_nim", IDENTITAS_DEFAULT["nim"]),
               ("id_prodi", IDENTITAS_DEFAULT["prodi"]), ("id_fakultas", IDENTITAS_DEFAULT["fakultas"]),
               ("id_semester", IDENTITAS_DEFAULT["semester"]), ("id_tahun", IDENTITAS_DEFAULT["tahun_akademik"]),
               ("id_dosen", IDENTITAS_DEFAULT["dosen"])):
    if _k not in st.session_state:
        st.session_state[_k] = _v
if "foto_bytes" not in st.session_state:
    st.session_state["foto_bytes"] = None
if "masuk" not in st.session_state:
    st.session_state["masuk"] = False  # False = halaman depan (tanpa sidebar), True = masuk sistem


def find_asset(candidates):
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def img_to_data_uri(path_or_bytes, ext="png"):
    try:
        if isinstance(path_or_bytes, (bytes, bytearray)):
            raw = bytes(path_or_bytes)
        else:
            with open(path_or_bytes, "rb") as f:
                raw = f.read()
        b64 = base64.b64encode(raw).decode("ascii")
        return f"data:image/{ext};base64,{b64}"
    except Exception:
        return None


def get_identitas():
    return {
        "nama": st.session_state.get("id_nama", "").strip(),
        "nim": st.session_state.get("id_nim", "").strip(),
        "prodi": st.session_state.get("id_prodi", "").strip(),
        "fakultas": st.session_state.get("id_fakultas", "").strip(),
        "matkul": IDENTITAS_DEFAULT["matkul"],
        "kelas": IDENTITAS_DEFAULT["kelas"],
        "keterangan": IDENTITAS_DEFAULT["keterangan"],
        "semester": st.session_state.get("id_semester", "").strip(),
        "tahun_akademik": st.session_state.get("id_tahun", "").strip(),
        "dosen": st.session_state.get("id_dosen", "").strip(),
    }


def tampil(val):
    return val if val else "—"


def inisial(nama):
    parts = [p for p in (nama or "").split() if p]
    if not parts:
        return "UTS"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def foto_data_uri():
    if st.session_state.get("foto_bytes"):
        return img_to_data_uri(st.session_state["foto_bytes"], "png")
    p = find_asset(ASSET_FOTO)
    if p:
        return img_to_data_uri(p, os.path.splitext(p)[1].lstrip(".").lower() or "png")
    return None


def goto(label):
    st.session_state["masuk"] = True
    st.session_state["nav"] = label
    st.rerun()


def masuk_sistem(label="🏠 Beranda Sistem"):
    st.session_state["masuk"] = True
    st.session_state["nav"] = label
    st.rerun()


def _pilih_nav(label):
    # Callback on_click: berjalan SEBELUM script run, jadi boleh mengubah state widget radio.
    st.session_state["nav"] = label


def sapaan():
    h = datetime.datetime.now().hour
    if 4 <= h < 11:
        return "Selamat pagi"
    if 11 <= h < 15:
        return "Selamat siang"
    if 15 <= h < 19:
        return "Selamat sore"
    return "Selamat malam"



# ============================================================
# 1. CSS KUSTOM — tata letak rapi & menarik
# ============================================================
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Poppins:wght@500;600;700;800&display=swap');
  :root {
    --blue: #2563EB; --cyan: #06B6D4; --purple: #8B5CF6; --violet: #6366F1;
    --teal: #14B8A6; --orange: #F97316; --yellow: #FACC15;
    --ink: #0F172A; --bg-soft: #EFF6FF;
  }
  .stApp, .stApp p, .stApp li, .stApp label, div[data-testid="stMetricLabel"] { font-family: 'Inter', system-ui, sans-serif; }
  .hero h1, h1, h2 { font-family: 'Poppins', 'Inter', system-ui, sans-serif; }
  /* ===== BASE: terang, akademik, berirama warna ===== */
  .stApp { background: linear-gradient(180deg, #FFFFFF 0%, #EFF6FF 60%, #F5F0FF 100%); }
  .stApp > header { background: transparent; }
  .block-container {
    padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1280px;
    background-image:
      linear-gradient(rgba(37,99,235,.05) 1px, transparent 1px),
      linear-gradient(90deg, rgba(37,99,235,.05) 1px, transparent 1px),
      radial-gradient(rgba(139,92,246,.07) 1px, transparent 1.5px);
    background-size: 30px 30px, 30px 30px, 26px 26px;
    border-radius: 18px;
  }
  /* ---------- HERO COLORFUL ---------- */
  .hero {
    position: relative; overflow: hidden; isolation: isolate;
    color: #F2F6FF; border-radius: 20px; padding: 32px 34px 26px 34px; margin-bottom: 18px;
    border: 1px solid rgba(147,197,253,.4);
    box-shadow: 0 16px 40px rgba(15,23,42,.30);
    animation: heroIn .45s ease both;
  }
  @keyframes heroIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: none; } }
  .theme-home { background: linear-gradient(135deg, #0F172A 0%, #2563EB 48%, #06B6D4 78%, #8B5CF6 115%); }
  .theme-eda  { background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 50%, #06B6D4 100%); }
  .theme-pre  { background: linear-gradient(135deg, #0F172A 0%, #3730A3 45%, #2563EB 75%, #06B6D4 110%); }
  .theme-apriori { background: linear-gradient(135deg, #4C1D95 0%, #6366F1 48%, #8B5CF6 80%, #F97316 135%); }
  .theme-kmeans  { background: linear-gradient(135deg, #083344 0%, #0891B2 40%, #14B8A6 70%, #2563EB 115%); }
  .theme-dash { background: linear-gradient(135deg, #1E1B4B 0%, #6366F1 45%, #2563EB 75%, #06B6D4 110%); }
  .theme-reco { background: linear-gradient(135deg, #1D4ED8 0%, #7C3AED 55%, #2563EB 85%, #F97316 130%); }
  .hero-grid {
    position: absolute; inset: 0; z-index: 0; pointer-events: none;
    background-image:
      linear-gradient(rgba(255,255,255,.10) 1px, transparent 1px),
      linear-gradient(90deg, rgba(255,255,255,.10) 1px, transparent 1px);
    background-size: 30px 30px;
    mask-image: radial-gradient(ellipse 90% 100% at 20% 0%, black 40%, transparent 100%);
  }
  .hero-circuit {
    position: absolute; inset: 0; z-index: 0; pointer-events: none; opacity: .55;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='220' height='160' viewBox='0 0 220 160'%3E%3Cg fill='none' stroke='%23A5F3FC' stroke-opacity='.5' stroke-width='1.2'%3E%3Cpath d='M10 30 H70 V70 H120'/%3E%3Cpath d='M10 120 H60 V95'/%3E%3Cpath d='M150 20 V60 H210'/%3E%3Cpath d='M160 130 H210 V100'/%3E%3Ccircle cx='120' cy='70' r='4' fill='%23FACC15' fill-opacity='.9' stroke='none'/%3E%3Ccircle cx='60' cy='95' r='3.4' fill='%23BFDBFE' stroke='none'/%3E%3Ccircle cx='150' cy='60' r='3.4' fill='%23F0ABFC' stroke='none'/%3E%3C/g%3E%3C/svg%3E");
    background-repeat: repeat;
    mask-image: linear-gradient(90deg, black 30%, transparent 85%);
  }
  .hero-orb { position: absolute; z-index: 0; border-radius: 50%; filter: blur(48px); pointer-events: none; animation: floatY 9s ease-in-out infinite; }
  .orb-a { width: 340px; height: 340px; right: -90px; top: -130px;
    background: radial-gradient(circle, rgba(6,182,212,.55), transparent 65%); }
  .orb-b { width: 300px; height: 300px; right: 200px; bottom: -170px;
    background: radial-gradient(circle, rgba(139,92,246,.5), transparent 65%); animation-delay: -3s; }
  .orb-c { width: 190px; height: 190px; left: 38%; top: -110px;
    background: radial-gradient(circle, rgba(250,204,21,.30), transparent 65%); animation-delay: -6s; }
  @keyframes floatY { 0%,100% { transform: translateY(0); } 50% { transform: translateY(16px); } }
  .hero-binary {
    position: absolute; right: 18px; bottom: 8px; z-index: 0;
    font-family: ui-monospace, 'Cascadia Code', Consolas, monospace;
    font-size: .68rem; letter-spacing: .28em; color: rgba(191,219,254,.55);
    user-select: none; pointer-events: none;
  }
  .hero-eyebrow {
    position: relative; z-index: 1;
    display: inline-flex; align-items: center; gap: 8px;
    font-family: ui-monospace, Consolas, monospace; font-size: .7rem; letter-spacing: .24em;
    color: #EAF6FF; background: rgba(255,255,255,.12);
    border: 1px solid rgba(255,255,255,.35); border-radius: 999px; padding: 5px 14px; margin-bottom: 12px;
    backdrop-filter: blur(8px);
  }
  .hero-eyebrow::before { content: ""; width: 34px; height: 2px; border-radius: 2px;
    background: linear-gradient(90deg, #FACC15, #F97316, #22D3EE); display: inline-block; }
  .pulse-dot { width: 7px; height: 7px; border-radius: 50%; background: #67E8F9;
    box-shadow: 0 0 10px rgba(103,232,249,.95); animation: pulseSoft 3.2s ease-in-out infinite; }
  @keyframes pulseSoft { 0%,100% { opacity: .55; } 50% { opacity: 1; } }
  .hero h1 { position: relative; z-index: 1; margin: 0 0 8px 0; font-size: 2rem; color: #FFFFFF; letter-spacing: -.015em; line-height: 1.2; }
  .hero p { position: relative; z-index: 1; margin: 4px 0; color: #E4EDFF; font-size: 1rem; }
  .hero .small { color: #D3E5FF !important; }
  .grad-text-warm { background: linear-gradient(90deg, #FDE68A, #FACC15 40%, #F97316); -webkit-background-clip: text; background-clip: text; color: transparent; font-weight: 800; }
  .grad-text-cool { background: linear-gradient(90deg, #A5F3FC, #67E8F9 45%, #C4B5FD); -webkit-background-clip: text; background-clip: text; color: transparent; font-weight: 800; }
  .grad-text-violet { background: linear-gradient(90deg, #DDD6FE, #C4B5FD 45%, #F0ABFC); -webkit-background-clip: text; background-clip: text; color: transparent; font-weight: 800; }
  .step-pills { position: relative; z-index: 1; margin-top: 12px; }
  .step-pills span {
    display: inline-block; padding: 5px 13px; border-radius: 999px; font-size: .78rem; margin: 3px 4px 0 0;
    font-family: ui-monospace, Consolas, monospace; letter-spacing: .02em; color: #fff;
    border: 1px solid rgba(255,255,255,.35); backdrop-filter: blur(8px);
    background: rgba(255,255,255,.12); transition: transform .25s ease, box-shadow .25s ease;
  }
  .step-pills span:hover { transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0,0,0,.25); }
  .step-pills span:nth-child(1) { background: linear-gradient(135deg, rgba(37,99,235,.85), rgba(6,182,212,.85)); }
  .step-pills span:nth-child(2) { background: linear-gradient(135deg, rgba(6,182,212,.85), rgba(20,184,166,.85)); }
  .step-pills span:nth-child(3) { background: linear-gradient(135deg, rgba(99,102,241,.85), rgba(139,92,246,.85)); }
  .step-pills span:nth-child(4) { background: linear-gradient(135deg, rgba(20,184,166,.85), rgba(37,99,235,.85)); }
  .step-pills span:nth-child(5) { background: linear-gradient(135deg, rgba(37,99,235,.85), rgba(139,92,246,.85)); }
  .step-pills span:nth-child(6) { background: linear-gradient(135deg, rgba(249,115,22,.9), rgba(250,204,21,.85)); color: #1F2937; border-color: rgba(255,255,255,.5); }
  .hero-chips { position: relative; z-index: 1; display: flex; gap: 10px; margin-top: 14px; flex-wrap: wrap; }
  .hero-chips .chip {
    display: inline-flex; align-items: center; gap: 8px; font-size: .8rem; color: #F2F7FF;
    background: rgba(255,255,255,.12); border: 1px solid rgba(255,255,255,.30);
    border-radius: 12px; padding: 7px 13px; backdrop-filter: blur(10px);
    box-shadow: 0 6px 18px rgba(0,0,0,.18); animation: floatY 7s ease-in-out infinite;
  }
  .hero-chips .chip b { font-family: ui-monospace, Consolas, monospace; font-size: .72rem; }
  .chip .dot { width: 9px; height: 9px; border-radius: 50%; }
  .chip-cyan .dot { background: #22D3EE; box-shadow: 0 0 8px #22D3EE; }
  .chip-violet .dot { background: #C4B5FD; box-shadow: 0 0 8px #A78BFA; animation-delay: -2s; }
  .chip-teal .dot { background: #2DD4BF; box-shadow: 0 0 8px #2DD4BF; }
  .chip-orange .dot { background: #FB923C; box-shadow: 0 0 8px #FB923C; }
  .hero-glowline {
    position: absolute; left: 0; right: 0; top: 0; height: 4px; z-index: 2;
    background: linear-gradient(90deg, #FACC15 0%, #F97316 18%, #38BDF8 45%, #818CF8 70%, #22D3EE 100%);
    box-shadow: 0 1px 12px rgba(56,189,248,.6);
    animation: glowDrift 7s ease-in-out infinite alternate;
  }
  @keyframes glowDrift { from { opacity: .8; } to { opacity: 1; } }
  .card {
    background: white; border: 1px solid #e5e7eb; border-radius: 14px;
    padding: 16px 18px; box-shadow: 0 2px 10px rgba(0,0,0,.04); margin-bottom: 14px;
  }
  .eyebrow { font-family: ui-monospace, Consolas, monospace; font-size: .7rem; letter-spacing: .22em; color: #1D4ED8; margin-bottom: 2px; }
  .eyebrow::before { content: "// "; color: #F97316; }
  /* icon warna per kategori */
  .ico-ai { color: #8B5CF6; } .ico-data { color: #06B6D4; } .ico-code { color: #2563EB; }
  .ico-db { color: #14B8A6; } .ico-net { color: #38BDF8; } .ico-inno { color: #F97316; } .ico-win { color: #CA8A04; }
  .note { background: #f0f9ff; border-left: 4px solid #0284c7; border-radius: 0 10px 10px 0;
    padding: 10px 14px; font-size: .88rem; color: #0c4a6e; margin: 10px 0; }
  .note-green { background:#f0fdf4; border-left-color:#16a34a; color:#14532d; }
  .note-amber { background:#fffbeb; border-left-color:#d97706; color:#78350f; }
  .note-purple { background:#F5F0FF; border-left-color:#8B5CF6; color:#4C1D95; }
  .note-teal { background:#EFFDF9; border-left-color:#14B8A6; color:#134E4A; }
  .kpi { text-align:center; }
  div[data-testid="stMetric"] {
    background: linear-gradient(180deg, #FFFFFF 0%, #F2F8FF 100%);
    border: 1px solid #D8E3F5; border-radius: 12px; padding: 12px;
    border-top: 4px solid #2563EB;
    box-shadow: 0 4px 14px rgba(37,99,235,.10);
    transition: transform .25s ease, box-shadow .25s ease;
  }
  div[data-testid="stMetric"]:hover { transform: translateY(-3px); box-shadow: 0 10px 24px rgba(37,99,235,.18); }
  div[data-testid="column"]:nth-child(2) div[data-testid="stMetric"] { border-top-color: #06B6D4; }
  div[data-testid="column"]:nth-child(3) div[data-testid="stMetric"] { border-top-color: #8B5CF6; }
  div[data-testid="column"]:nth-child(4) div[data-testid="stMetric"] { border-top-color: #14B8A6; }
  div[data-testid="column"]:nth-child(5) div[data-testid="stMetric"] { border-top-color: #F97316; }
  /* sidebar */
  section[data-testid="stSidebar"] { background: linear-gradient(180deg, #E3ECFD 0%, #F2F6FF 35%, #FAF5FF 100%); border-right: 1px solid #CBD9F2; }
  section[data-testid="stSidebar"]::before {
    content: ""; display: block; height: 4px; margin: 0 -1rem;
    background: linear-gradient(90deg, #2563EB, #06B6D4 35%, #8B5CF6 70%, #F97316);
    opacity: .9;
  }
  .side-badge {
    background: linear-gradient(120deg, #0F172A 0%, #2563EB 60%, #8B5CF6 110%);
    border: 1px solid rgba(147,197,253,.45); border-radius: 14px; padding: 13px 15px; color: #EAF2FF;
    box-shadow: 0 8px 20px rgba(15,23,42,.28); margin-bottom: 6px; position: relative; overflow: hidden;
  }
  .side-badge::after {
    content: ""; position: absolute; inset: 0; pointer-events: none;
    background-image: linear-gradient(rgba(165,243,252,.20) 1px, transparent 1px),
      linear-gradient(90deg, rgba(196,181,253,.20) 1px, transparent 1px);
    background-size: 18px 18px;
    mask-image: linear-gradient(90deg, black, transparent 78%);
  }
  .side-badge .mono { font-family: ui-monospace, Consolas, monospace; font-size: .66rem; letter-spacing: .26em; color: #A5F3FC; }
  .side-badge h2 { margin: 2px 0 0 0 !important; font-size: 1.12rem !important; color: #fff !important; }
  h1, h2, h3 { letter-spacing: -.01em; color: #0F172A; }
  .hero h1 { color: #FFFFFF; }
  .small { font-size: .82rem; color: #64748b; }
  hr { margin: 1.2em 0; border-color: #D8E3F5; }
  div[data-testid="stContainer"] { border-radius: 14px; }
  div[data-testid="stMetricValue"] > div { white-space: nowrap; overflow: visible; text-overflow: clip; font-size: 1.2rem; font-variant-numeric: tabular-nums; }
  div[data-testid="stMetricLabel"] { white-space: normal; }
  div[data-testid="stContainer"][data-testid="stVerticalBlockBorderWrapper"] {
    box-shadow: 0 5px 18px rgba(15,23,42,.07);
    border-color: #D8E3F5 !important;
    background: linear-gradient(180deg, rgba(255,255,255,.99), rgba(240,246,255,.99)) !important;
    transition: transform .25s ease, box-shadow .25s ease;
  }
  div[data-testid="stContainer"][data-testid="stVerticalBlockBorderWrapper"]:hover {
    box-shadow: 0 10px 26px rgba(37,99,235,.14);
  }
  div[data-testid="stMetricValue"], .stDataFrame { font-variant-numeric: tabular-nums; }
  /* tombol unduh/aksi: biru-cyan + orange untuk prioritas */
  div[data-testid="stDownloadButton"] > button, div[data-testid="stButton"] > button {
    background: linear-gradient(135deg, #2563EB, #06B6D4) !important; color: #fff !important;
    border: none !important; border-radius: 10px !important; font-weight: 600;
    box-shadow: 0 4px 14px rgba(37,99,235,.35); transition: all .25s ease;
  }
  div[data-testid="stDownloadButton"] > button:hover, div[data-testid="stButton"] > button:hover {
    transform: translateY(-2px); box-shadow: 0 8px 22px rgba(6,182,212,.45);
    background: linear-gradient(135deg, #1D4ED8, #0891B2) !important;
  }
  /* tombol utama MULAI SEKARANG: oranye menonjol */
  div[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(135deg, #EA580C, #F97316 55%, #FACC15 130%) !important;
    color: #fff !important; font-size: 1.1rem !important; font-weight: 800 !important;
    padding: .85rem 1.4rem !important; border-radius: 14px !important;
    box-shadow: 0 10px 28px rgba(249,115,22,.5); transition: all .25s ease;
    animation: ctaGlow 2.6s ease-in-out infinite;
  }
  div[data-testid="stButton"] > button[kind="primary"]:hover {
    transform: translateY(-3px) scale(1.01); box-shadow: 0 14px 34px rgba(249,115,22,.6);
  }
  @keyframes ctaGlow { 0%,100% { filter: brightness(1); } 50% { filter: brightness(1.12); } }
  /* tabs berwarna */
  button[data-baseweb="tab"] { font-weight: 600; }
  button[data-baseweb="tab"][aria-selected="true"] {
    background: linear-gradient(135deg, rgba(37,99,235,.12), rgba(139,92,246,.12)) !important;
    border-radius: 10px 10px 0 0;
  }
  input[type="range"] { accent-color: #2563EB; }
  details[data-testid="stExpander"] { border: 1px solid #D8E3F5; border-radius: 12px; background: #F7FBFF; }
  .stRadio > div { gap: .5rem; }
  ::selection { background: rgba(139,92,246,.25); }
  /* ---------- COVER IDENTITAS (halaman depan, ala contoh) ---------- */
  .cover { position: relative; border-radius: 22px; overflow: hidden; isolation: isolate;
    padding: 46px 34px 38px 34px; text-align: center; color: #fff;
    background-color: #151238; background-size: cover; background-position: center;
    border: 1px solid rgba(139,92,246,.45); box-shadow: 0 22px 55px rgba(10,8,35,.5);
    animation: heroIn .45s ease both; margin-bottom: 16px; }
  .cover::before { content: ""; position: absolute; inset: 0; z-index: 0;
    background: linear-gradient(135deg, rgba(12,10,42,.90) 0%, rgba(38,28,92,.84) 45%, rgba(40,70,180,.80) 100%); }
  .cover-grid { position: absolute; inset: 0; z-index: 0; pointer-events: none;
    background-image: linear-gradient(rgba(255,255,255,.07) 1px, transparent 1px),
      linear-gradient(90deg, rgba(255,255,255,.07) 1px, transparent 1px);
    background-size: 30px 30px;
    mask-image: radial-gradient(ellipse 80% 70% at 50% 0%, black, transparent 85%); }
  .cover-inner { position: relative; z-index: 1; max-width: 860px; margin: 0 auto; }
  .cv-badge { display: inline-flex; align-items: center; gap: 8px;
    font-family: ui-monospace, Consolas, monospace; font-size: .72rem; letter-spacing: .24em;
    color: #C9D6FF; background: rgba(255,255,255,.10); border: 1px solid rgba(255,255,255,.35);
    border-radius: 999px; padding: 8px 22px; backdrop-filter: blur(8px); }
  .cv-ring { width: 202px; height: 202px; border-radius: 50%; margin: 22px auto 0 auto; padding: 6px;
    background: conic-gradient(from 200deg, #0EA5A4, #2563EB, #38BDF8, #0E7490, #0EA5A4);
    box-shadow: 0 0 44px rgba(6,182,212,.55), 0 12px 30px rgba(0,0,0,.4); line-height: 0; }
  .cv-crop { width: 100%; height: 100%; border-radius: 50%; overflow: hidden;
    border: 5px solid #fff; margin: 0; padding: 0; background: #1E293B; line-height: 0; }
  .cv-photo { width: 100%; height: 100%; display: block; border: none; margin: 0; padding: 0;
    object-fit: cover; object-position: 50% 18%;
    transform: scale(1.45); transform-origin: 50% 28%; background: #1E293B; }
  .cv-initials { width: 100%; height: 100%; border-radius: 50%; display: flex; align-items: center; justify-content: center;
    border: 5px solid #fff; margin: 0;
    font-family: 'Poppins', sans-serif; font-weight: 800; font-size: 2.6rem; color: #fff;
    background: linear-gradient(135deg, #2563EB, #06B6D4 55%, #8B5CF6); line-height: 1; }
  .cv-name { font-family: 'Poppins', sans-serif; font-weight: 800; font-size: 2.1rem; color: #fff; margin: 18px 0 10px 0; }
  .cv-npm { display: inline-block; background: #fff; color: #1E1B4B; font-weight: 800; font-size: 1rem;
    border-radius: 999px; padding: 9px 30px; box-shadow: 0 8px 22px rgba(0,0,0,.35); }
  .cv-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 28px; text-align: left; }
  .cv-card { background: rgba(255,255,255,.09); border: 1px solid rgba(255,255,255,.24);
    border-radius: 14px; padding: 14px 18px; backdrop-filter: blur(10px);
    transition: transform .25s ease, box-shadow .25s ease; }
  .cv-card:hover { transform: translateY(-3px); box-shadow: 0 10px 26px rgba(6,182,212,.25); }
  .cv-card small { display: block; color: #9FB6FF; letter-spacing: .2em; font-size: .66rem; font-weight: 700;
    font-family: ui-monospace, Consolas, monospace; margin-bottom: 5px; }
  .cv-card b { color: #fff; font-size: 1.04rem; font-weight: 700; }
  .cv-note { margin-top: 16px; font-size: .8rem; color: #C4CFF5; }
  /* responsif */
  @media (max-width: 768px) {
    .cv-grid { grid-template-columns: 1fr; }
    .cv-name { font-size: 1.5rem; }
    .cv-ring { width: 168px; height: 168px; }
    .cover { padding: 34px 18px 28px 18px; }
    .hero { padding: 22px 20px 18px 20px; }
    .hero h1 { font-size: 1.4rem; }
    .hero-binary, .hero-chips { display: none; }
    .orb-c { display: none; }
    .block-container { background-size: 24px 24px, 24px 24px, 22px 22px; }
    .pf-photo, .pf-initials { width: 128px !important; height: 128px !important; }
    .campus-img { height: 200px !important; }
  }
  /* ---------- DASHBOARD KELAS B: profil, kampus, kartu nav ---------- */
  @keyframes fadeUp { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: none; } }
  div[data-testid="stVerticalBlockBorderWrapper"] { animation: fadeUp .4s ease both; }
  .pf-wrap { display: flex; gap: 16px; align-items: center; }
  .pf-center { text-align: center; }
  .pf-center .pf-ring { display: inline-block; }
  .pf-center .pf-name { margin: 10px 0 2px 0; }
  .pf-ring {
    padding: 4px; border-radius: 50%;
    background: conic-gradient(from 210deg, #2563EB, #06B6D4, #8B5CF6, #6366F1, #2563EB);
    box-shadow: 0 8px 22px rgba(37,99,235,.35), 0 0 0 4px rgba(6,182,212,.12);
    flex-shrink: 0; line-height: 0;
  }
  .pf-photo, .pf-initials {
    width: 168px; height: 168px; border-radius: 50%; display: block;
    border: none; margin: 0;
  }
  .pf-photo { object-fit: cover; object-position: center top; background: #DBEAFE; }
  .pf-initials {
    display: flex; align-items: center; justify-content: center;
    font-family: 'Poppins', sans-serif; font-weight: 800; font-size: 2rem; color: #fff;
    background: linear-gradient(135deg, #2563EB, #06B6D4 55%, #8B5CF6);
  }
  .pf-name { font-family: 'Poppins', sans-serif; font-weight: 700; font-size: 1.25rem; color: #0F172A; margin: 0; }
  .pf-nim { font-family: ui-monospace, Consolas, monospace; color: #1D4ED8; font-size: .95rem; }
  .pf-badge {
    display: inline-block; font-size: .72rem; font-weight: 700; letter-spacing: .08em;
    color: #fff; background: linear-gradient(135deg, #2563EB, #8B5CF6);
    border-radius: 999px; padding: 3px 12px; margin: 2px 4px 2px 0;
  }
  .pf-badge.cyan { background: linear-gradient(135deg, #0891B2, #14B8A6); }
  .pf-badge.orange { background: linear-gradient(135deg, #F97316, #FACC15); color: #431407; }
  .campus-banner { position: relative; border-radius: 16px; overflow: hidden; border: 1px solid #D8E3F5;
    box-shadow: 0 10px 26px rgba(15,23,42,.15); line-height: 0; background: #0F172A; }
  .campus-img { width: 100%; height: 300px; object-fit: cover; object-position: center; display: block; margin: 0; }
  .campus-overlay { position: absolute; inset: 0;
    background: linear-gradient(180deg, rgba(15,23,42,0) 30%, rgba(15,23,42,.55) 75%, rgba(37,99,235,.55) 100%); }
  .campus-caption { position: absolute; left: 16px; right: 16px; bottom: 12px; color: #fff; }
  .campus-caption b { font-family: 'Poppins', sans-serif; font-size: 1.05rem; }
  .campus-caption span { font-size: .82rem; color: #D8E8FF; }
  .campus-placeholder { border-radius: 16px; border: 1.5px dashed #93B4E8;
    min-height: 300px; display: flex; gap: 14px; align-items: center; justify-content: center;
    text-align: center; padding: 26px 30px;
    background-image:
      linear-gradient(rgba(37,99,235,.07) 1px, transparent 1px),
      linear-gradient(90deg, rgba(37,99,235,.07) 1px, transparent 1px),
      linear-gradient(135deg, #DBEAFE, #EFF6FF 55%, #EDE9FE);
    background-size: 26px 26px, 26px 26px, cover;
    color: #334155; font-size: .88rem; }
  .nav-ico { font-size: 1.5rem; }
  .greet-pill { display: inline-block; font-weight: 600; font-size: .85rem; color: #0F172A;
    background: linear-gradient(135deg, #FEF9C3, #FDE68A); border: 1px solid #FACC15;
    padding: 5px 14px; border-radius: 999px; margin-bottom: 8px; }
  @media (prefers-reduced-motion: reduce) {
    .pulse-dot, .hero-glowline, .hero-orb, .hero-chips .chip, .hero,
    div[data-testid="stButton"] > button[kind="primary"] { animation: none !important; }
    div[data-testid="stMetric"], div[data-testid="stContainer"][data-testid="stVerticalBlockBorderWrapper"] { transition: none; }
  }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 2. FUNGSI DATA (CACHED)
# ============================================================
@st.cache_data(show_spinner=False)
def load_raw(path, sheet):
    """Baca mentah tanpa dibersihkan (untuk audit sebelum vs sesudah)."""
    try:
        return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except ImportError as exc:
        raise ImportError(
            "Library openpyxl belum terpasang pada Python yang menjalankan aplikasi. "
            "Jalankan: python -m pip install openpyxl. "
            "Jika deploy di Streamlit Cloud, tambahkan openpyxl ke requirements.txt."
        ) from exc

@st.cache_data(show_spinner=False)
def load_and_clean(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except ImportError as exc:
        raise ImportError(
            "Library openpyxl belum terpasang pada Python yang menjalankan aplikasi. "
            "Jalankan: python -m pip install openpyxl. "
            "Jika deploy di Streamlit Cloud, tambahkan openpyxl ke requirements.txt."
        ) from exc
    rows_before = len(df)
    audit = {}
    audit['missing_awal'] = int(df.isnull().sum().sum())
    audit['duplikat_awal'] = int(df.duplicated().sum())

    # 1. Drop duplikat
    df = df.drop_duplicates()

    # 2. Rapikan teks kategori
    for col in ['Region', 'Customer_Type', 'Category', 'Payment_Method', 'Product_Name']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.title()
            df[col] = df[col].replace({'Nan': np.nan, 'None': np.nan, '': np.nan})
    if 'Payment_Method' in df.columns:
        df['Payment_Method'] = df['Payment_Method'].replace({'Qris': 'QRIS'})

    # 3. Drop missing kolom kunci
    before = len(df)
    df = df.dropna(subset=['Customer_ID', 'Product_ID', 'Product_Name'])
    audit['hapus_missing_kunci'] = before - len(df)

    # 4. Tanggal & numerik
    df['Transaction_Date'] = pd.to_datetime(df['Transaction_Date'], errors='coerce')
    before = len(df)
    df = df.dropna(subset=['Transaction_Date'])
    audit['hapus_tanggal_invalid'] = before - len(df)

    for col in ['Quantity', 'Unit_Price', 'Discount_Rate', 'Line_Total']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # 5. Filter nilai wajar
    before = len(df)
    df = df[(df['Quantity'] > 0) & (df['Quantity'] < 1000)
            & (df['Unit_Price'] > 0) & (df['Unit_Price'] < 1_000_000)
            & (df['Discount_Rate'].between(0, 1))]
    audit['hapus_tidak_wajar'] = before - len(df)

    # 6. Hitung ulang Line_Total agar konsisten
    df['Line_Total'] = df['Quantity'] * df['Unit_Price'] * (1 - df['Discount_Rate'])

    # 7. Capping outlier IQR (tetap pertahankan baris, batasi nilai ekstrem)
    Q1, Q3 = df['Line_Total'].quantile(0.25), df['Line_Total'].quantile(0.75)
    IQR = Q3 - Q1
    lower, upper = Q1 - 1.5 * IQR, Q3 + 1.5 * IQR
    n_capped_low = int((df['Line_Total'] < lower).sum())
    n_capped_high = int((df['Line_Total'] > upper).sum())
    df['Line_Total'] = np.clip(df['Line_Total'], lower, upper)
    audit['capped_bawah'] = n_capped_low
    audit['capped_atas'] = n_capped_high
    audit['batas_iqr'] = (float(lower), float(upper))

    # 8. basket_size = jumlah produk unik per transaksi
    basket_size = df.groupby('Transaction_ID')['Product_ID'].nunique().reset_index()
    basket_size.columns = ['Transaction_ID', 'basket_size']
    df = df.merge(basket_size, on='Transaction_ID', how='left')

    audit['missing_akhir'] = int(df.isnull().sum().sum())
    audit['duplikat_akhir'] = int(df.duplicated().sum())
    return df, rows_before, audit

@st.cache_data(show_spinner=False)
def build_rfm(df):
    snapshot_date = df['Transaction_Date'].max() + pd.Timedelta(days=1)
    rfm = df.groupby('Customer_ID').agg(
        Recency=('Transaction_Date', lambda x: (snapshot_date - x.max()).days),
        Frequency=('Transaction_ID', 'nunique'),
        Monetary=('Line_Total', 'sum')
    ).reset_index()
    return rfm, snapshot_date

@st.cache_data(show_spinner=False)
def build_basket(df):
    basket = df.groupby(['Transaction_ID', 'Product_Name'])['Quantity'].sum().unstack().fillna(0)
    basket = (basket > 0).astype(int)
    return basket

@st.cache_data(show_spinner=False)
def run_apriori(_basket, min_support, min_confidence, min_lift):
    frequent = apriori(_basket, min_support=min_support, use_colnames=True)
    if len(frequent) == 0:
        return pd.DataFrame(), frequent
    rules = association_rules(frequent, metric="confidence", min_threshold=min_confidence)
    rules = rules[rules['lift'] >= min_lift]
    rules = rules.sort_values(['lift', 'confidence'], ascending=False).reset_index(drop=True)
    return rules, frequent

@st.cache_data(show_spinner=False)
def evaluate_k(_rfm, k_list):
    X = _rfm[['Recency', 'Frequency', 'Monetary']]
    scaler = StandardScaler()
    Xs = scaler.fit_transform(X)
    inertia, sil = [], []
    for k in k_list:
        km = KMeans(n_clusters=k, random_state=42, n_init=10)
        labels = km.fit_predict(Xs)
        inertia.append(float(km.inertia_))
        sil.append(float(silhouette_score(Xs, labels)) if len(set(labels)) > 1 else 0.0)
    return inertia, sil, Xs

@st.cache_data(show_spinner=False)
def run_kmeans_final(_rfm, _Xs, k_opt):
    km = KMeans(n_clusters=k_opt, random_state=42, n_init=10)
    out = _rfm.copy()
    out['Cluster'] = km.fit_predict(_Xs)
    profile = out.groupby('Cluster').agg(
        R_mean=('Recency', 'mean'), F_mean=('Frequency', 'mean'),
        M_mean=('Monetary', 'mean'), Jumlah=('Customer_ID', 'count'),
        R_med=('Recency', 'median'), F_med=('Frequency', 'median'), M_med=('Monetary', 'median')
    ).round(2).reset_index()
    # Pelabelan otomatis berbasis median profil
    f_med, m_med, r_med = profile['F_mean'].median(), profile['M_mean'].median(), profile['R_mean'].median()
    labels = []
    for _, r in profile.iterrows():
        if r['F_mean'] >= f_med and r['M_mean'] >= m_med:
            labels.append('Loyal Customers')
        elif r['R_mean'] >= r_med and r['F_mean'] < f_med:
            labels.append('Churn Risk')
        elif r['F_mean'] < f_med and r['M_mean'] < m_med:
            labels.append('New / Potential')
        else:
            labels.append('Regular')
    profile['Label'] = labels
    out['Segment'] = out['Cluster'].map(dict(zip(profile['Cluster'], profile['Label'])))
    return out, profile

# ============================================================
# 3. LOAD DATA
# ============================================================
try:
    df_raw_audit = load_raw(FILE_PATH, SHEET_RAW)
except Exception as e:
    st.error(f"❌ Gagal membaca file `{FILE_PATH}` sheet `{SHEET_RAW}`: {e}")
    if "openpyxl" in str(e).lower() or isinstance(e, ImportError):
        st.info(
            "Perbaikan: instal dependensi dengan `python -m pip install openpyxl`. "
            "Untuk Streamlit Cloud, pastikan `openpyxl` tercantum di `requirements.txt`, "
            "lalu lakukan redeploy/reboot aplikasi."
        )
    st.stop()

with st.spinner("🔄 Membersihkan data & menyiapkan fitur..."):
    df, rows_before, audit = load_and_clean(FILE_PATH, SHEET_RAW)
    rfm, snapshot_date = build_rfm(df)
    basket = build_basket(df)
    inertia_all, sil_all, X_scaled = evaluate_k(rfm, K_CANDIDATES)

# ============================================================
# 4. GERBANG HALAMAN: depan (tanpa sidebar) vs sistem (sidebar)
# ============================================================
MASUK = bool(st.session_state.get("masuk", False))
if not MASUK:
    st.markdown("""<style>
      section[data-testid="stSidebar"] { display: none !important; }
      button[data-testid="stSidebarCollapsedControl"] { display: none !important; }
    </style>""", unsafe_allow_html=True)

if MASUK:
    st.sidebar.markdown("""<div class="side-badge"><div class="mono">UNIMOR // BI KELAS B • UTS</div>
<h2>🎓 BI Kelas B — UNIMOR</h2><div style="font-size:.76rem;color:#BFDBFE;">Future Tech • Smart Analytics • Data Intelligence</div></div>""", unsafe_allow_html=True)
    if st.sidebar.button("← Halaman Depan", use_container_width=True):
        st.session_state["masuk"] = False
        st.rerun()
    st.sidebar.caption("Beranda Sistem → EDA → Preprocessing → Apriori → K-Means → Dashboard → Rekomendasi")

    MENU = st.sidebar.radio(
        "🧭 Navigasi Sistem:",
        ["🏠 Beranda Sistem",
         "1️⃣ Business Understanding & EDA",
         "2️⃣ Preprocessing & Fitur",
         "3️⃣ Apriori — Pola Belanja",
         "4️⃣ K-Means — Segmentasi",
         "5️⃣ Dashboard BI",
         "6️⃣ Rekomendasi & Laporan"],
        index=0,
        key="nav"
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Filter Global")
    st.sidebar.caption("📝 *Catatan: filter ini memengaruhi KPI, EDA, dan Dashboard. Apriori & K-Means tetap memakai data bersih penuh agar pola tidak bias.*")

    min_date = df['Transaction_Date'].min().date()
    max_date = df['Transaction_Date'].max().date()
    date_range = st.sidebar.date_input("📅 Rentang tanggal", [min_date, max_date], min_value=min_date, max_value=max_date)
    if isinstance(date_range, (list, tuple)):
        d0, d1 = (date_range + [date_range[0], date_range[0]])[:2] if len(date_range) == 1 else date_range[:2]
    else:
        d0 = d1 = date_range

    categories = ['Semua'] + sorted(df['Category'].dropna().unique().tolist())
    selected_cat = st.sidebar.selectbox("🏷️ Kategori produk", categories)

    st.sidebar.markdown("---")
    st.sidebar.markdown(f"📄 **Data awal:** {rows_before:,} baris  \n✅ **Data bersih:** {len(df):,} baris  \n🗓️ **Snapshot RFM:** {snapshot_date.date()}  \n🧺 **Transaksi unik:** {df['Transaction_ID'].nunique():,}  \n👥 **Pelanggan:** {df['Customer_ID'].nunique():,}")
    st.sidebar.caption("📝 Snapshot = tanggal acuan Recency (tgl terakhir + 1 hari).")
else:
    MENU = "🛬 Halaman Depan"
    min_date = df['Transaction_Date'].min().date()
    max_date = df['Transaction_Date'].max().date()
    d0, d1, selected_cat = min_date, max_date, 'Semua'

# Terapkan filter global → df_f
df_f = df[(df['Transaction_Date'].dt.date >= d0) & (df['Transaction_Date'].dt.date <= d1)].copy()
if selected_cat != 'Semua':
    df_f = df_f[df_f['Category'] == selected_cat].copy()
if len(df_f) == 0:
    st.warning("⚠️ Filter menghasilkan data kosong. Longgarkan rentang tanggal / kategori.")
    df_f = df.copy()
df_f['Bulan'] = df_f['Transaction_Date'].dt.to_period('M').astype(str)

def hero(title, desc, pills, eyebrow="FUTURE TECH • AKADEMIK • PROFESIONAL", theme="theme-home", chips=None):
    if chips is None:
        chips = [("chip-cyan", "◈ Apriori"), ("chip-violet", "⬡ K-Means"), ("chip-teal", "⬢ RFM"), ("chip-orange", "⬣ Dashboard")]
    _chips = "".join([f'<span class="chip {c}"><span class="dot"></span><b>{t}</b></span>' for c, t in chips])
    st.markdown(f"""<div class="hero {theme}"><div class="hero-glowline"></div><div class="hero-grid"></div>
    <div class="hero-circuit"></div><div class="hero-orb orb-a"></div><div class="hero-orb orb-b"></div><div class="hero-orb orb-c"></div>
    <div class="hero-binary">01100010 01101001 01010100 01101111 01101011 01101111 // AI • DATA • BI</div>
    <div class="hero-eyebrow"><span class="pulse-dot"></span>{eyebrow}</div><h1>{title}</h1><p>{desc}</p>
    <div class="step-pills">{''.join([f'<span>{p}</span>' for p in pills])}</div>
    <div class="hero-chips">{_chips}</div></div>""",
    unsafe_allow_html=True)

def note(text, kind=""):
    st.markdown(f'<div class="note {kind}">{text}</div>', unsafe_allow_html=True)

if not MASUK:
    idt = get_identitas()
    _bg = find_asset(ASSET_KAMPUS)
    _bg_uri = img_to_data_uri(_bg, __import__("os").path.splitext(_bg)[1].lstrip(".").lower() or "jpg") if _bg else None
    _bg_style = f"background-image: url('{_bg_uri}');" if _bg_uri else ""
    _furi = foto_data_uri()
    _fimg = (f'<img class="cv-photo" src="{_furi}" alt="Foto {tampil(idt["nama"])}">'
             if _furi else f'<div class="cv-initials">{inisial(idt["nama"])}</div>')
    _cards = [
        ("\U0001F3DB\ufe0f", "UNIVERSITAS", "Universitas Timor (UNIMOR)"),
        ("\U0001F4BB", "PROGRAM STUDI", tampil(idt["prodi"])),
        ("\U0001F5D3\ufe0f", "SEMESTER", f"Semester {tampil(idt['semester'])}"),
        ("\U0001F3EB", "KELAS", tampil(idt["kelas"])),
        ("\U0001F4D8", "MATA KULIAH", tampil(idt["matkul"])),
        ("\U0001F4C5", "TAHUN AKADEMIK", tampil(idt["tahun_akademik"])),
        ("\U0001F468\u200d\U0001F3EB", "DOSEN PENGAMPU", tampil(idt["dosen"])),
        ("\U0001F4DD", "KETERANGAN", tampil(idt["keterangan"])),
    ]
    _cards_html = "".join(
        [f'<div class="cv-card"><small>{ic} &nbsp;{lb}</small><b>{vl}</b></div>' for ic, lb, vl in _cards])
    st.markdown(
        f'<div class="cover" style="{_bg_style}"><div class="cover-grid"></div>'
        f'<div class="cover-inner"><div class="cv-badge">\u2728 COVER \u2022 IDENTITAS MAHASISWA</div><br>'
        f'<div class="cv-ring"><div class="cv-crop">{_fimg}</div></div>'
        f'<div class="cv-name">{tampil(idt["nama"])}</div>'
        f'<div><span class="cv-npm">\U0001F393 NPM {tampil(idt["nim"])}</span></div>'
        f'<div class="cv-grid">{_cards_html}</div>'
        f'<div class="cv-note">Foto gedung kampus UNIMOR sebagai latar cover \u2022 UTS Business Intelligence Kelas B</div>'
        f'</div></div>',
        unsafe_allow_html=True)
    with st.expander("\u270f\ufe0f Ubah identitas & foto (tanpa edit kode)"):
        st.text_input("Nama lengkap", key="id_nama")
        st.text_input("NPM", key="id_nim")
        cc1, cc2 = st.columns(2)
        with cc1:
            st.text_input("Program Studi", key="id_prodi")
            st.text_input("Semester", key="id_semester")
            st.text_input("Dosen pengampu", key="id_dosen")
        with cc2:
            st.text_input("Fakultas", key="id_fakultas")
            st.text_input("Tahun akademik", key="id_tahun")
        up = st.file_uploader("Foto mahasiswa (PNG/JPG, tampil bulat otomatis)", type=["png", "jpg", "jpeg"])
        if up is not None:
            st.session_state["foto_bytes"] = up.getvalue()
            st.success("\u2705 Foto dimuat — tampil bulat di cover.")
        st.caption("\U0001F4DD Kolom kosong tampil sebagai \u201c—\u201d (tidak dikarang).")

    _sp1, _cta, _sp2 = st.columns([1, 2, 1])
    with _cta:
        if st.button("\U0001F680 MULAI SEKARANG \u2014 Masuk ke Sistem", use_container_width=True, type="primary"):
            masuk_sistem("\U0001F3E0 Beranda Sistem")
    st.caption("\U0001F4DD *Klik tombol di atas untuk masuk ke halaman sistem (sidebar + 6 modul analisis). Kembali kapan saja lewat tombol \u201c\u2190 Halaman Depan\u201d di sidebar.*")

elif MENU == "🏠 Beranda Sistem":
    hero('🏠 Beranda <span class="grad-text-cool">Sistem Analitik</span>',
         "Ringkasan data & peta pengerjaan — pilih modul analisis di sidebar untuk mulai bekerja.",
         ["EDA", "Preprocessing", "Apriori", "K-Means", "Dashboard BI", "Rekomendasi"],
         theme="theme-home", eyebrow="SISTEM • ANALITIK • DECISION SUPPORT")

    st.markdown("### 📊 Ringkasan Data")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("💰 Total Penjualan (bersih)", f"{fmt_rp(df['Line_Total'].sum())}")
    c2.metric("🧾 Transaksi Unik", f"{df['Transaction_ID'].nunique():,}")
    c3.metric("👥 Pelanggan Unik", f"{df['Customer_ID'].nunique():,}")
    c4.metric("🧺 Rata-rata Basket", f"{df['basket_size'].mean():.2f} produk/nota")
    st.caption("📝 Total = Σ *Line_Total* bersih • Transaksi = ID unik (1 nota = beberapa baris) • Mengikuti filter sidebar.")

    with st.expander("🎯 Permasalahan Bisnis (studi kasus)"):
        st.markdown(
            "- Data selama ini hanya untuk **laporan rutin** — belum dipakai mengambil keputusan.\n"
            "- Kombinasi produk laris, segmen pelanggan bernilai/berisiko, dan strategi promosi **belum diketahui**.\n"
            "- Data mentah bermasalah: missing, duplikat, teks tak konsisten, qty/harga tak wajar, outlier.")

    with st.expander("❓ Pertanyaan Bisnis yang Dijawab Sistem"):
        st.markdown(
            "1. Produk apa yang sering dibeli bersama? → **Apriori**\n"
            "2. Segmen mana paling bernilai / berisiko churn? → **K-Means**\n"
            "3. Kapan penjualan naik & apa pendorongnya? → **Dashboard BI**\n"
            "4. Tindakan apa yang prioritas? → **Rekomendasi**")

    with st.expander("🗺️ Peta Alur Pengerjaan (6 tahap)"):
        st.markdown(
            "**1. EDA** → **2. Preprocessing** (RFM + basket) → **3. Apriori** → **4. K-Means** → "
            "**5. Dashboard** → **6. Rekomendasi**. Selesaikan tahap 2 sebelum menjalankan algoritma.")

    with st.expander("📦 Struktur dataset & contoh data mentah"):
        st.markdown(
            f"**Data_Raw:** {df_raw_audit.shape[0]:,} baris × {df_raw_audit.shape[1]} kolom (1 baris = 1 item nota). "
            "Kolom kunci: `Transaction_ID`, `Customer_ID`, `Product_ID/Name`, `Transaction_Date`, `Quantity`, "
            "`Unit_Price`, `Discount_Rate`, `Line_Total`. Sheet pendukung: `Data_Dictionary`, `Data_Produk`, "
            "`Data_Pelanggan`, `Info_Dataset`. Data mentah sengaja bermasalah untuk menguji preprocessing.")
        st.dataframe(df_raw_audit.head(5), use_container_width=True)

    st.button("▶️ Mulai dari EDA →", use_container_width=True,
              on_click=_pilih_nav, args=("1️⃣ Business Understanding & EDA",))

# ============================================================
# HALAMAN 1: EDA
# ============================================================
elif MENU == "1️⃣ Business Understanding & EDA":
    hero('1️⃣ Business Understanding & <span class="grad-text-cool">EDA</span> <span class="small" style="color:#dbeafe">(Bobot 10%)</span>',
        "Tujuan: <b>memahami masalah bisnis → memeriksa kualitas data → menemukan sinyal awal</b> sebelum algoritma dijalankan.",
         ["Soal 1 · Masalah & pertanyaan bisnis", "Soal 2 · Profil data + missing + visual + insight"],
        theme="theme-eda", eyebrow="TAHAP 01 // BUSINESS UNDERSTANDING & EDA")
    note("📝 <b>Alur halaman ini:</b> ① Ringkasan angka pokok → ② Tipe data & statistik → ③ Audit kualitas (missing/duplikat/outlier) → ④ Dua visual utama → ⑤ Tiga temuan awal. Semua angka memakai <b>filter global</b> di sidebar.")

    with st.container(border=True):
        st.markdown("### 📌 Ringkasan Dataset (memakai filter aktif)")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Baris (item)", f"{len(df_f):,}")
        k2.metric("Transaksi unik", f"{df_f['Transaction_ID'].nunique():,}")
        k3.metric("Pelanggan unik", f"{df_f['Customer_ID'].nunique():,}")
        k4.metric("Produk unik", f"{df_f['Product_Name'].nunique():,}")
        k5.metric("Total penjualan", f"{fmt_rp(df_f['Line_Total'].sum())}")
        st.caption(f"📝 Rentang tanggal data penuh: {min_date} s/d {max_date} • Filter aktif: {d0} s/d {d1} • Kategori: {selected_cat} • Rata-rata belanja/nota: {fmt_rp(df_f.groupby('Transaction_ID')['Line_Total'].sum().mean())}")

    c1, c2 = st.columns(2)
    with c1:
        with st.container(border=True):
            st.markdown("### 📊 Statistik Deskriptif")
            st.caption("📝 Catatan: `describe()` meringkas kolom numerik (rata-rata, min, max, kuartil). Gunakan untuk mendeteksi nilai aneh — mis. Quantity negatif atau diskon > 1.")
            st.dataframe(df_f[['Quantity', 'Unit_Price', 'Discount_Rate', 'Line_Total', 'basket_size']].describe().T.style.format("{:,.2f}"), use_container_width=True)
    with c2:
        with st.container(border=True):
            st.markdown("### 🧾 Tipe Data, Missing & Nilai Unik")
            st.caption("📝 Catatan: kolom ID/tanggal/kategori adalah teks & tanggal (bukan angka). Kolom `Customer_ID` boleh kosong di sebagian baris — penanganannya dibahas di Tahap 2.")
            info = pd.DataFrame({'Tipe': df_f.dtypes.astype(str), 'Missing (filter aktif)': df_f.isnull().sum(), 'Nilai unik': df_f.nunique()})
            st.dataframe(info, use_container_width=True)
            with st.expander("🔎 Audit kualitas DATA MENTAH (sebelum preprocessing) — bukti Soal 2c"):
                miss = df_raw_audit.isnull().sum().sort_values(ascending=False)
                st.markdown(f"**Duplikat mentah:** {int(df_raw_audit.duplicated().sum()):,} baris &nbsp;•&nbsp; **Total sel kosong:** {int(miss.sum()):,}")
                st.dataframe(miss[miss > 0].rename("Missing").to_frame(), use_container_width=True)
                st.caption("📝 Nilai tidak wajar yang sengaja disisipkan: Quantity ≤ 0 / ≥ 1000, Unit_Price ≤ 0 / sangat besar, Discount di luar [0,1], teks tidak konsisten, tanggal invalid, dan outlier Line_Total. Semua ditangani di halaman 2️⃣.")

    with st.container(border=True):
        st.markdown("### 📈 Visualisasi EDA (minimal 2 sesuai soal)")
        st.caption("📝 Catatan: visual kiri menjawab 'kategori apa yang paling ramai', visual kanan menjawab 'kapan transaksi naik/turun'. Arahkan kursor ke grafik untuk detail (interaktif).")
        v1, v2 = st.columns(2)
        with v1:
            cat_counts = df_f['Category'].value_counts().reset_index()
            cat_counts.columns = ['Category', 'Jumlah']
            fig = px.bar(cat_counts, x='Category', y='Jumlah', color='Jumlah', color_continuous_scale='Blues',
                         title='Distribusi Baris Transaksi per Kategori', text_auto=True)
            fig.update_layout(xaxis_tickangle=-20, height=420)
            st.plotly_chart(fig, use_container_width=True)
            st.caption("📝 Cara baca: batang tertinggi = kategori dengan baris item terbanyak pada filter aktif.")
        with v2:
            monthly_n = df_f.groupby('Bulan').size().reset_index(name='Jumlah')
            fig = px.line(monthly_n, x='Bulan', y='Jumlah', markers=True, title='Tren Jumlah Baris Transaksi per Bulan')
            fig.update_layout(height=420)
            st.plotly_chart(fig, use_container_width=True)
            st.caption("📝 Cara baca: titik naik = bulan ramai, titik turun = bulan sepi — bahan tren penjualan di Dashboard.")
        with st.expander("➕ Visual tambahan: Top 10 produk & metode pembayaran"):
            e1, e2 = st.columns(2)
            with e1:
                top_p = df_f['Product_Name'].value_counts().head(10).reset_index()
                top_p.columns = ['Produk', 'Jumlah']
                st.plotly_chart(px.bar(top_p, x='Jumlah', y='Produk', orientation='h', color='Jumlah',
                                       color_continuous_scale='Greens', title='Top 10 Produk (frekuensi baris)'), use_container_width=True)
            with e2:
                pay = df_f['Payment_Method'].value_counts().reset_index()
                pay.columns = ['Metode', 'Jumlah']
                st.plotly_chart(px.pie(pay, names='Metode', values='Jumlah', hole=.4, title='Komposisi Metode Pembayaran'), use_container_width=True)
            st.caption("📝 Catatan: frekuensi ≠ revenue. Produk sering dibeli belum tentu penyumbang Rp terbesar — bandingkan dengan Top-10 Revenue di Dashboard.")

    with st.container(border=True):
        st.markdown("### 💡 Tiga Temuan Awal (Soal 2e — otomatis dari data)")
        top_cat = df_f['Category'].value_counts()
        avg_nota = df_f.groupby('Transaction_ID')['Line_Total'].sum().mean()
        st.markdown(f"""
    1. **Kategori dominan:** `{top_cat.idxmax()}` ({int(top_cat.max()):,} baris) — kandidat fokus bundling & stok.
    2. **Skala bisnis (filter aktif):** {df_f['Transaction_ID'].nunique():,} nota dari {df_f['Customer_ID'].nunique():,} pelanggan, rata-rata **{fmt_rp(avg_nota)}/nota** dan **{df_f['basket_size'].mean():.2f} produk unik/nota**.
    3. **Kualitas data butuh preprocessing:** data mentah mengandung {int(df_raw_audit.duplicated().sum()):,} duplikat + {int(df_raw_audit.isnull().sum().sum()):,} sel kosong + nilai tidak wajar — lihat detail di halaman **2️⃣ Preprocessing**.
        """)
        note("📝 <b>Implikasi ke bisnis:</b> (1) Kategori ramai = prioritas rak & promo; (2) Rata-rata nota & basket = baseline KPI sebelum intervensi; (3) Tanpa pembersihan, semua metrik di atas bias. Lanjut ke halaman 2️⃣ untuk melihat perbaikannya.", "note-green")

# ============================================================
# HALAMAN 2: PREPROCESSING
# ============================================================
elif MENU == "2️⃣ Preprocessing & Fitur":
    hero('2️⃣ Preprocessing & <span class="grad-text-cool">Feature Engineering</span> <span class="small" style="color:#dbeafe">(Bobot 25%)</span>',
        "Tujuan: <b>mengubah data kotor → data siap algoritma</b> (bersih + fitur RFM + basket_size + matriks siap Apriori/K-Means).",
         ["Soal 3 · Cleaning + alasan + before/after", "Soal 4 · RFM + basket + scaling + strategi big data"],
        theme="theme-pre", eyebrow="TAHAP 02 // PREPROCESSING & FEATURE ENGINEERING")
    note("📝 <b>Alur halaman ini:</b> ① Peta masalah→penanganan→alasan ② Angka before/after ③ Distribusi Line_Total ④ Fitur RFM & basket_size ⑤ Scaling & strategi data besar ⑥ Download hasil. Semua keputusan harus beralasan (syarat UTS).")

    with st.container(border=True):
        st.markdown("### 🧮 Ringkasan Before → After")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Baris awal", f"{rows_before:,}")
        m2.metric("Baris bersih", f"{len(df):,}", delta=f"-{rows_before - len(df):,} dibuang")
        m3.metric("Duplikat sisa", f"{audit['duplikat_akhir']:,}", delta="0 (tuntas)", delta_color="normal")
        m4.metric("Sel kosong sisa*", f"{audit['missing_akhir']:,}")
        st.caption("📝 *Sisa missing hanya di kolom non-kunci yang memang boleh kosong. Kolom kunci (Customer/Product/Date) sudah 0 missing. Lihat tabel keputusan di bawah.")

    with st.container(border=True):
        st.markdown("### 🗺️ Keputusan Preprocessing (Soal 3a–3b: masalah → penanganan → alasan)")
        st.caption("📝 Catatan: setiap baris = satu jenis masalah di Data_Raw. Kolom 'Bukti/Jumlah' menunjukkan skala masalahnya — transparan & dapat direproduksi.")
        dec = pd.DataFrame([
            ["Duplikat persis", "Drop duplikat", f"{audit['duplikat_awal']:,} baris", "Baris identik menggandakan agregat & basket"],
            ["Missing kolom kunci (Customer/Product)", "Drop baris", f"{audit['hapus_missing_kunci']:,} baris", "Kunci tidak bisa diimputasi; RFM & basket butuh ID valid"],
            ["Teks tidak konsisten (kapital/spasi/Qris)", "strip().title() + QRIS fix", "seluruh kolom kategori", "Standarisasi agar 'minuman' & 'Minuman' tidak dihitung beda"],
            ["Tanggal invalid / gagal parse", "to_datetime(coerce) → drop NaT", f"{audit['hapus_tanggal_invalid']:,} baris", "Tren & Recency butuh tanggal valid"],
            ["Quantity ≤0 / ≥1000", "Drop baris", f"termasuk dalam {audit['hapus_tidak_wajar']:,} baris", "Qty 0/negatif & borongan tak wajar merusak total"],
            ["Unit_Price ≤0 / ≥1 jt", "Drop baris", "—", "Harga 0/negatif/ekstrem bukan transaksi normal"],
            ["Discount di luar [0,1]", "Drop baris", "—", "Diskon proporsi valid hanya 0–100%"],
            ["Line_Total tidak konsisten", "Hitung ulang Qty×Price×(1−Disc)", "seluruh baris", "Menjamin konsistensi aritmetika"],
            ["Outlier Line_Total", f"Capping IQR [{fmt_id(audit['batas_iqr'][0])} – {fmt_id(audit['batas_iqr'][1])}]", f"{audit['capped_atas']+audit['capped_bawah']:,} tercap", "Ekstrem dipertahankan tapi dibatasi (tidak dibuang) — adil secara bisnis"],
        ], columns=["Masalah", "Penanganan", "Bukti / Jumlah", "Alasan (logis)"])
        st.dataframe(dec, use_container_width=True, hide_index=True)
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**📊 Distribusi Line_Total sesudah cleaning + capping**")
            from matplotlib.ticker import FuncFormatter as _FF
            fig, ax = plt.subplots(figsize=(7, 3.6))
            sns.histplot(df['Line_Total'], bins=50, kde=True, ax=ax, color="#2563eb")
            ax.xaxis.set_major_formatter(_FF(lambda v, _: fmt_id(v)))
            ax.set_xlabel("Line_Total (Rp, titik ribuan)"); ax.set_title("Distribusi nilai per baris item")
            st.pyplot(fig, use_container_width=True)
            st.caption("📝 Sumbu-x memakai titik ribuan Indonesia. Ekor kanan panjang = outlier yang sudah dibatasi (capping).")
        with g2:
            st.markdown("**🧺 Distribusi basket_size (produk unik / nota)**")
            fig2, ax2 = plt.subplots(figsize=(7, 3.6))
            sns.histplot(df['basket_size'], bins=range(1, int(df['basket_size'].max()) + 2), ax=ax2, color="#06b6d4", discrete=True)
            ax2.set_xlabel("basket_size"); ax2.set_title("Berapa macam produk per nota?")
            st.pyplot(fig2, use_container_width=True)
            st.caption(f"📝 Rata-rata basket = {df['basket_size'].mean():.2f}. Nota dengan basket ≥ 2 adalah bahan utama Apriori (kombinasi produk).")

    with st.container(border=True):
        st.markdown("### 🧬 Feature Engineering (Soal 4)")
        t1, t2 = st.tabs(["👥 RFM per Pelanggan", "🧺 Basket & Matriks Siap Algoritma"])
        with t1:
            note("📝 <b>Apa itu RFM?</b> <b>R</b>ecency = hari sejak belanja terakhir (tgl acuan: snapshot). <b>F</b>requency = jumlah nota unik. <b>M</b>onetary = total Rp bersih. Satu baris = satu pelanggan — syarat K-Means.")
            st.caption(f"📝 Snapshot date (acuan Recency) = {snapshot_date.date()} = tanggal transaksi terakhir + 1 hari. Didokumentasikan agar reproducible.")
            _rfm_show = rfm.head(10).copy()
            _rfm_show['Monetary'] = _rfm_show['Monetary'].apply(fmt_id)
            st.dataframe(_rfm_show, use_container_width=True)
            st.caption("📝 Kolom Monetary ditampilkan bertitik ribuan Indonesia agar mudah dibaca.")
            c1, c2, c3 = st.columns(3)
            c1.metric("Rata-rata Recency", f"{rfm['Recency'].mean():.1f} hari")
            c2.metric("Rata-rata Frequency", f"{rfm['Frequency'].mean():.2f} nota")
            c3.metric("Rata-rata Monetary", f"{fmt_rp(rfm['Monetary'].mean())}")
            with st.expander("❓ Kenapa RFM cocok untuk segmentasi?"):
                st.markdown("- **Perilaku, bukan demografi:** R/F/M langsung mencerminkan nilai & keterlibatan pelanggan.  \n- **Kompatibel K-Means:** 3 fitur numerik → mudah di-scaling & di-cluster.  \n- **Actionable:** tiap dimensi punya aksi (R tinggi → reaktivasi; F/M tinggi → reward loyal).")
        with t2:
            note("📝 <b>Alur fitur:</b> `basket_size` = n unik Product_ID per Transaction_ID → matriks transaksi×produk (1 = dibeli, 0 = tidak). Matriks inilah input Apriori. Untuk K-Means, matriks numeriknya adalah tabel RFM yang di-standardisasi.")
            st.markdown(f"**Matriks basket:** {basket.shape[0]:,} transaksi × {basket.shape[1]} produk &nbsp;•&nbsp; **Kepadatan:** {(basket.values.sum() / basket.size * 100):.2f}% sel = 1")
            st.dataframe(basket.iloc[:8, :8], use_container_width=True)
            st.caption("📝 Catatan: JANGAN menghitung baris item sebagai transaksi terpisah — satu Transaction_ID = satu keranjang (sesuai Soal Tugas 9).")

    with st.container(border=True):
        st.markdown("### ⚖️ Standardisasi & Strategi Data Besar (Soal 4e–4f)")
        c1, c2 = st.columns(2)
        with c1:
            note("📝 <b>Kenapa StandardScaler (z-score)?</b> R (hari), F (nota), M (rupiah) satuannya beda jauh. Tanpa scaling, Monetary mendominasi jarak Euclidean dan klaster rusak. Z-score membuat tiap fitur mean=0, std=1 sehingga setara. Dipilih atas MinMax karena tahan terhadap sisa outlier.")
            scaler_demo = StandardScaler().fit_transform(rfm[['Recency', 'Frequency', 'Monetary']])
            demo = pd.DataFrame(scaler_demo, columns=['R_z', 'F_z', 'M_z']).describe().T[['mean', 'std', 'min', 'max']].round(2)
            st.dataframe(demo, use_container_width=True)
            st.caption("📝 Bukti scaling benar: mean ≈ 0 dan std ≈ 1 untuk ketiga fitur.")
        with c2:
            note("📝 <b>Strategi jika data membesar (jutaan baris):</b> ① <b>Sampling</b> stratified per bulan/kategori untuk eksplorasi cepat; ② <b>Chunking</b> — baca & agregasi per potongan (pandas chunksize) agar hemat RAM; ③ <b>Agregasi dini</b> — langsung ke level nota/pelanggan sebelum mining. Di proyek ini agregasi RFM & basket sudah menjadi bentuk agregasinya.", "note-green")
            st.markdown("""
    - **Sampling:** cocok untuk tuning support/K awal.
    - **Chunking:** cocok untuk cleaning file raksasa.
    - **Agregasi:** wajib sebelum Apriori/K-Means (hemat memori, makna bisnis jelas).
            """)
        st.markdown("### ⬇️ Unduh Hasil Preprocessing (Luaran UTS)")
        d1, d2, d3 = st.columns(3)
        d1.download_button("⬇️ Data_Clean.csv", df.to_csv(index=False).encode('utf-8'), "Data_Clean.csv", "text/csv", use_container_width=True)
        d2.download_button("⬇️ RFM.csv", rfm.to_csv(index=False).encode('utf-8'), "RFM.csv", "text/csv", use_container_width=True)
        d3.download_button("⬇️ Dokumentasi perubahan (txt)",
            f"Baris awal: {rows_before}\nBaris bersih: {len(df)}\nDihapus: {rows_before-len(df)}\nDuplikat awal: {audit['duplikat_awal']}\nMissing awal: {audit['missing_awal']}\nSnapshot: {snapshot_date.date()}\nBatas IQR: {audit['batas_iqr']}\n".encode('utf-8'),
            "dokumentasi_preprocessing.txt", "text/plain", use_container_width=True)
        st.caption("📝 Catatan: file-file ini adalah 'Dataset hasil preprocessing + dokumentasi perubahan' yang wajib dikumpulkan (Luaran UTS).")

# ============================================================
# HALAMAN 3: APRIORI
# ============================================================
elif MENU == "3️⃣ Apriori — Pola Belanja":
    hero('3️⃣ Market Basket Analysis — <span class="grad-text-violet">Apriori</span> <span class="small" style="color:#dbeafe">(Bobot 20%)</span>',
        "Tujuan: <b>menemukan kombinasi produk yang sering dibeli bersama</b> untuk bundling, cross-selling, dan tata rak.",
         ["Soal 5 · Matriks + support/confidence + top-10 + visual", "Soal 6 · Interpretasi + 3 rekomendasi"],
        theme="theme-apriori", eyebrow="TAHAP 03 // APRIORI • MARKET BASKET ANALYSIS")
    note("📝 <b>Alur halaman ini:</b> ① Pahami S/C/L → ② Atur threshold (geser slider) → ③ Baca Top-10 aturan → ④ Pilih visual → ⑤ Baca interpretasi & rekomendasi. Menggeser slider = mengulang mining secara live.")

    with st.container(border=True):
        st.markdown("### 📖 Memahami Support, Confidence & Lift (wajib paham sebelum baca tabel)")
        m1, m2, m3 = st.columns(3)
        m1.markdown("**Support**  \n`P(A & B)` — seberapa *sering* kombinasi muncul dari seluruh transaksi.  \n📝 *Contoh: 0.03 = 3% nota berisi A+B.*")
        m2.markdown("**Confidence**  \n`P(B|A)` — jika beli A, seberapa *yakin* ikut beli B.  \n📝 *Contoh: 0.5 = separuh pembeli A juga ambil B.*")
        m3.markdown("**Lift**  \n`Conf / P(B)` — apakah hubungan *lebih kuat dari kebetulan* (>1 = menarik).  \n📝 *Contoh: lift 2 = pembeli A 2× lebih mungkin ambil B.*")
        st.caption("📝 Peringkat 'terbaik' di dashboard ini = urut **lift tertinggi → confidence** (aturan paling tidak-kebetulan & paling dapat diandalkan). Kriteria ini dinyatakan eksplisit sesuai tuntutan soal.")

    with st.container(border=True):
        st.markdown("### 🎚️ Parameter Interaktif (Soal 5b — geser & lihat hasilnya berubah)")
        p1, p2, p3 = st.columns(3)
        with p1:
            min_sup = st.slider("Minimum Support", 0.005, 0.10, float(DEFAULT_SUPPORT), 0.005, format="%.3f")
            st.caption("📝 Support kecil → banyak aturan (termasuk yang jarang). Support besar → hanya pola sangat umum.")
        with p2:
            min_conf = st.slider("Minimum Confidence", 0.10, 0.80, float(DEFAULT_CONF), 0.05, format="%.2f")
            st.caption("📝 Confidence kecil → aturan lemah ikut lolos. Naikkan untuk aturan yang lebih 'pasti'.")
        with p3:
            min_lift = st.slider("Minimum Lift", 1.00, 3.00, float(DEFAULT_LIFT), 0.10, format="%.2f")
            st.caption("📝 Lift = 1 artinya independen (tidak menarik). Syarat ≥ 1 menyaring hubungan kebetulan.")
        with st.expander("❓ Alasan pemilihan nilai default (support 0.02, confidence 0.30, lift 1.0)"):
            st.markdown(f"""
    - Dataset punya **{basket.shape[0]:,} transaksi × {basket.shape[1]} produk** dengan kepadatan {(basket.values.sum()/basket.size*100):.2f}% — pola belanja **menyebar**, bukan didominasi 1-2 paket.
    - **Support 0.02** ≈ minimal {int(0.02*basket.shape[0]):,} nota — cukup rendah untuk menangkap bundling nyata, cukup tinggi agar tidak banjir ribuan aturan noise.
    - **Confidence 0.30** = aturan minimal benar 3 dari 10 kali — ambang praktis untuk promo (tidak terlalu longgar, tidak terlalu ketat).
    - **Lift ≥ 1.0** = hanya hubungan di atas kebetulan. Silakan geser slider untuk uji sensitivitas dan laporkan perubahannya.
            """)
        with st.spinner("⛏️ Menjalankan Apriori..."):
            rules, frequent = run_apriori(basket, float(min_sup), float(min_conf), float(min_lift))
        k1, k2, k3 = st.columns(3)
        k1.metric("📦 Frequent Itemsets", f"{len(frequent):,}")
        k2.metric("🔗 Association Rules", f"{len(rules):,}")
        k3.metric("🚀 Lift tertinggi", f"{rules['lift'].max():.2f}" if len(rules) else "—")
        note(f"📝 <b>Cara baca metrik:</b> dengan support={min_sup:.3f} & confidence={min_conf:.2f} ditemukan <b>{len(rules):,}</b> aturan. Jika 0 aturan → turunkan support/confidence (pola terlalu ketat). Jika ribuan → naikkan (terlalu longgar).", "note-amber" if len(rules) == 0 else "note-green")

    if len(rules) == 0:
        st.warning("⚠️ Tidak ada aturan memenuhi threshold. Turunkan Minimum Support / Confidence pada slider di atas, lalu tabel akan terisi otomatis.")
    else:
        with st.container(border=True):
            st.markdown("### 🏆 Top 10 Association Rules (Soal 5f)")
            st.caption("📝 Catatan: 'Antecedent → Consequent' dibaca 'jika beli kiri, maka cenderung beli kanan'. Tabel bisa dicari & diunduh.")
            q = st.text_input("🔎 Cari produk dalam aturan (mis. 'kopi', 'susu', 'sabun')", "")
            disp = rules.copy()
            disp['antecedents'] = disp['antecedents'].apply(lambda x: ', '.join(sorted(list(x))))
            disp['consequents'] = disp['consequents'].apply(lambda x: ', '.join(sorted(list(x))))
            if q.strip():
                mask = disp['antecedents'].str.contains(q, case=False) | disp['consequents'].str.contains(q, case=False)
                disp = disp[mask]
                st.caption(f"📝 Ditemukan {len(disp)} aturan mengandung '{q}'.")
            st.dataframe(disp[['antecedents', 'consequents', 'support', 'confidence', 'lift']].head(10).style.format({'support': '{:.4f}', 'confidence': '{:.4f}', 'lift': '{:.2f}'}), use_container_width=True)
            st.download_button("⬇️ Unduh semua rules (CSV)", disp.to_csv(index=False).encode('utf-8'), "association_rules.csv", "text/csv")

        with st.container(border=True):
            st.markdown("### 📊 Visualisasi Apriori (Soal 5g — pilih salah satu)")
            viz = st.radio("Pilih visualisasi:", ["Scatter (Support vs Confidence)", "Network Graph", "Heatmap Co-occurrence", "Bar Metrik Top-10"], horizontal=True)
            if viz == "Scatter (Support vs Confidence)":
                note("📝 <b>Cara baca scatter:</b> tiap titik = 1 aturan. Sumbu X = support (frekuensi), Y = confidence (kepastian), warna/ukuran = lift (kekuatan). Titik kanan-atas besar = aturan terbaik.")
                # Siapkan kolom string agar Plotly tidak error (frozenset tidak bisa di-serialize)
                _plot = rules.copy()
                _plot['antecedent_str'] = _plot['antecedents'].apply(lambda x: ', '.join(sorted(list(x))))
                _plot['consequent_str'] = _plot['consequents'].apply(lambda x: ', '.join(sorted(list(x))))
                _plot['lift_plot'] = _plot['lift'].replace([float('inf'), float('-inf')], float('nan')).fillna(_plot['lift'].median() if _plot['lift'].notna().any() else 1.0)
                _plot['label'] = _plot['antecedent_str'] + ' \u2192 ' + _plot['consequent_str']
                fig = px.scatter(
                    _plot, x='support', y='confidence', color='lift_plot', size='lift_plot',
                    hover_data={'label': True, 'support': ':.4f', 'confidence': ':.3f', 'lift_plot': ':.2f',
                                'antecedent_str': False, 'consequent_str': False},
                    labels={'lift_plot': 'lift', 'support': 'support', 'confidence': 'confidence', 'label': 'aturan'},
                    color_continuous_scale='Plasma', size_max=22, opacity=0.8,
                    title='Sebaran Association Rules (warna & ukuran = lift)')
                fig.update_layout(height=480)
                st.plotly_chart(fig, use_container_width=True)
                st.caption("📝 Catatan: jika titik menumpuk di kiri-bawah, artinya threshold longgar — naikkan Minimum Support/Confidence agar hanya aturan kuat yang tampil.")
            elif viz == "Network Graph":
                note("📝 <b>Cara baca network:</b> panah A → B = 'pembeli A cenderung ambil B'. Angka di garis = lift. Node berderajat tinggi = produk 'penghubung' — kandidat taruh di tengah rak.")
                top = rules.head(15)
                G = nx.DiGraph()
                for _, r in top.iterrows():
                    for a in r['antecedents']:
                        for c in r['consequents']:
                            G.add_edge(a, c, weight=float(r['lift']))
                fig_nx, ax = plt.subplots(figsize=(13, 8))
                pos = nx.spring_layout(G, k=1.6, seed=42)
                nx.draw_networkx_nodes(G, pos, node_size=2600, node_color='#bfdbfe', edgecolors='#1e3a8a', ax=ax)
                nx.draw_networkx_labels(G, pos, font_size=8, font_weight='bold', ax=ax)
                nx.draw_networkx_edges(G, pos, width=[G[u][v]['weight'] * 1.4 for u, v in G.edges()],
                                       edge_color='gray', arrows=True, arrowsize=18, alpha=.65, ax=ax)
                nx.draw_networkx_edge_labels(G, pos, edge_labels={(u, v): f"{G[u][v]['weight']:.2f}" for u, v in G.edges()}, font_size=7, ax=ax)
                ax.set_title('Network Graph Top-15 Rules (angka = lift)'); ax.axis('off')
                st.pyplot(fig_nx, use_container_width=True)
            elif viz == "Heatmap Co-occurrence":
                note("📝 <b>Cara baca heatmap:</b> angka = berapa nota berisi kedua produk sekaligus. Warna makin merah = makin sering bersama — kandidat bundling & rak berdekatan.")
                top_products = df['Product_Name'].value_counts().head(15).index.tolist()
                cols = [c for c in top_products if c in basket.columns]
                co = basket[cols].T.dot(basket[cols])
                fig_hm, ax = plt.subplots(figsize=(11, 9))
                sns.heatmap(co, annot=True, fmt='d', cmap='YlOrRd', square=True, cbar_kws={'label': 'Nota bersama'}, ax=ax)
                ax.set_title('Co-occurrence Top-15 Produk'); plt.xticks(rotation=45, ha='right'); plt.yticks(rotation=0)
                st.pyplot(fig_hm, use_container_width=True)
            else:
                note("📝 <b>Cara baca bar chart:</b> tiga panel membandingkan support / confidence / lift untuk 10 aturan teratas. Aturan ideal = tinggi di ketiganya.")
                top = rules.head(10).copy()
                top['rule'] = top['antecedents'].apply(lambda x: ', '.join(sorted(list(x)))) + ' → ' + top['consequents'].apply(lambda x: ', '.join(sorted(list(x))))
                fig, axes = plt.subplots(1, 3, figsize=(18, 6))
                for ax, m in zip(axes, ['support', 'confidence', 'lift']):
                    sns.barplot(data=top, x=m, y='rule', ax=ax, palette='viridis')
                    ax.set_title(f'Top-10 · {m.capitalize()}'); ax.set_ylabel('')
                plt.tight_layout(); st.pyplot(fig, use_container_width=True)

        with st.container(border=True):
            st.markdown("### 🔎 Interpretasi & Rekomendasi (Soal 6)")
            t1 = rules.iloc[0]
            a1, c1_ = ', '.join(sorted(list(t1['antecedents']))), ', '.join(sorted(list(t1['consequents'])))
            st.success(f"**Aturan terkuat:** pembeli **{a1}** → cenderung **{c1_}**  \nSupport `{t1['support']:.4f}` (muncul di {t1['support']*100:.2f}% nota) • Confidence `{t1['confidence']:.3f}` ({t1['confidence']*100:.1f}% pembeli kiri ikut ambil kanan) • Lift `{t1['lift']:.2f}` (peluang {t1['lift']:.1f}× lipat vs acak).")
            note("📝 <b>Contoh menerjemahkan angka:</b> support = 'seberapa umum', confidence = 'seberapa yakin', lift = 'seberapa istimewa'. Lift 2+ dengan confidence 40%+ adalah bahan bundling yang kuat — bukan sekadar produk laris yang memang dibeli semua orang.")
            st.markdown(f"""
    **3 rekomendasi cross-selling / bundling (berbasis bukti, bukan perkiraan):**
    1. 🎁 **Paket hemat `{a1} + {c1_}`** (diskon bundling) — didukung aturan lift tertinggi di atas.
    2. 🛒 **Cross-selling di kasir & struk:** setiap nota berisi `{a1}` tawarkan `{c1_}` (cek Top-10 untuk varian lain).
    3. 🗄️ **Rak berdekatan:** tempatkan pasangan co-occurrence tertinggi (lihat Heatmap/Network) dalam 1 lorong agar impulse-buying naik.
            """)
            st.caption("📝 Setiap rekomendasi merujuk ke aturan/heatmap di atas — inilah yang dimaksud soal 'alasan mengacu hasil analisis'.")

# ============================================================
# HALAMAN 4: K-MEANS
# ============================================================
elif MENU == "4️⃣ K-Means — Segmentasi":
    hero('4️⃣ Segmentasi Pelanggan — <span class="grad-text-cool">K-Means</span> <span class="small" style="color:#dbeafe">(Bobot 20%)</span>',
        "Tujuan: <b>mengelompokkan pelanggan berdasar perilaku belanja (RFM)</b> agar promosi tepat sasaran.",
         ["Soal 7 · Fitur + Elbow/Silhouette + k optimal + scatter", "Soal 8 · Profil + label + strategi + prioritas"],
        theme="theme-kmeans", eyebrow="TAHAP 04 // K-MEANS • SEGMENTASI RFM")
    note("📝 <b>Alur halaman ini:</b> ① Pahami fitur & scaling → ② Baca Elbow + Silhouette → ③ Geser k pilihan → ④ Baca profil & scatter → ⑤ Terapkan strategi per segmen.")

    with st.container(border=True):
        st.markdown("### 🧬 Fitur & Persiapan (Soal 7a)")
        st.markdown(f"Satu baris = **satu pelanggan** ({len(rfm):,} pelanggan). Fitur: **Recency** (hari), **Frequency** (nota), **Monetary** (Rp). Preprocessing: **StandardScaler** (z-score).")
        note("📝 <b>Kenapa harus scaling?</b> Tanpa scaling, Monetary (puluhan–ratusan ribu) menenggelamkan Recency/Frequency. Z-score menyetarakan ketiganya (mean 0, std 1) sehingga jarak K-Means adil. Buktinya bisa dicek di halaman 2️⃣ (mean≈0, std≈1).")

    with st.container(border=True):
        st.markdown("### 📉 Elbow Method & Silhouette Score (Soal 7b–7c)")
        st.caption("📝 Catatan: Elbow = cari 'siku' di mana penurunan inertia melambat. Silhouette = makin dekat 1 makin bagus (klaster rapat & terpisah). K optimal = siku + silhouette tinggi + masuk akal bisnis.")
        e1, e2 = st.columns(2)
        with e1:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=K_CANDIDATES, y=inertia_all, mode='lines+markers', name='Inertia', line=dict(color='#2563eb', width=3)))
            fig.update_layout(title='Elbow Method (makin kecil makin rapat)', xaxis_title='k', yaxis_title='Inertia', height=380)
            st.plotly_chart(fig, use_container_width=True)
        with e2:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=K_CANDIDATES, y=sil_all, mode='lines+markers', name='Silhouette', line=dict(color='#f97316', width=3)))
            fig.update_layout(title='Silhouette Score (makin tinggi makin baik)', xaxis_title='k', yaxis_title='Score', height=380)
            st.plotly_chart(fig, use_container_width=True)
        sil_tbl = pd.DataFrame({'k': K_CANDIDATES, 'Inertia': np.round(inertia_all, 1), 'Silhouette': np.round(sil_all, 4)})
        st.dataframe(sil_tbl.style.highlight_max(subset=['Silhouette'], color='#dcfce7'), use_container_width=True, hide_index=True)
        best_sil_k = int(sil_tbl.loc[sil_tbl['Silhouette'].idxmax(), 'k'])
        note(f"📝 <b>Pembacaan otomatis:</b> silhouette tertinggi ada di <b>k={best_sil_k}</b> (skor {max(sil_all):.4f}). Nilai default dashboard = <b>k={DEFAULT_K}</b>. Silakan uji k lain di bawah dan bandingkan profilnya — pilihan akhir harus menggabungkan metrik + interpretasi bisnis (Soal 7d).", "note-green")

    with st.container(border=True):
        k_opt = st.slider("🎯 Pilih jumlah klaster final (k)", 2, 8, int(DEFAULT_K))
        st.caption("📝 Catatan: menggeser k akan menghitung ulang segmentasi secara live. Perhatikan bagaimana profil & scatter berubah — k yang terlalu kecil mencampur segmen, terlalu besar memecah segmen alami.")
        rfm_seg, profile = run_kmeans_final(rfm, X_scaled, int(k_opt))
        s1, s2 = st.columns([1, 1.6])
        with s1:
            st.markdown("**Proporsi segmen**")
            seg_counts = rfm_seg['Segment'].value_counts().reset_index()
            seg_counts.columns = ['Segment', 'Jumlah']
            st.plotly_chart(px.pie(seg_counts, names='Segment', values='Jumlah', hole=.4, title=f'Pembagian {len(rfm_seg):,} pelanggan (k={k_opt})'), use_container_width=True)
            st.caption("📝 Cara baca: iris terbesar = segmen terbanyak — tapi prioritas bisnis belum tentu yang terbesar (lihat halaman 6️⃣).")
        with s2:
            st.markdown("**Scatter RFM per segmen (ukuran = Frequency)**")
            st.plotly_chart(px.scatter(rfm_seg, x='Recency', y='Monetary', color='Segment', size='Frequency',
                                       hover_data=['Customer_ID', 'Cluster'], title='Recency vs Monetary'), use_container_width=True)
            st.caption("📝 Cara baca: kanan = baru belanja (Recency kecil biasanya di kiri — cek sumbu), atas = belanja besar, lingkaran besar = sering belanja. Loyal = atas-besar, Churn = kanan-kecil.")
        st.markdown("### 📋 Profil Setiap Klaster (Soal 8a — R/F rata-rata + median + jumlah)")
        st.caption("📝 Catatan: R_mean kecil = baru belanja (bagus). F/M besar = sering & besar (bagus). Label otomatis di bawah DIVERIFIKASI dari angka profil — bukan asal tempel.")
        st.dataframe(profile.style.format({'R_mean': '{:.1f}', 'F_mean': '{:.2f}', 'M_mean': lambda x: fmt_id(x), 'R_med': '{:.1f}', 'F_med': '{:.2f}', 'M_med': lambda x: fmt_id(x)}), use_container_width=True)
        st.download_button("⬇️ Unduh profil klaster (CSV)", profile.to_csv(index=False).encode('utf-8'), "cluster_profile.csv", "text/csv")
        st.markdown("### 🏷️ Interpretasi + Strategi per Segmen (Soal 8b–8d)")
        strat = {'Loyal Customers': '👑 Membership premium + poin + early-access produk baru. Tujuan: pertahankan & upsell.',
                 'Churn Risk': '📲 Reaktivasi WA/email + voucher “kami rindu Anda” + survei alasan pergi. Tujuan: tarik kembali.',
                 'New / Potential': '🌱 Onboarding + diskon pembelian ke-2 + edukasi produk. Tujuan: ubah jadi Regular/Loyal.',
                 'Regular': '🛒 Upsell produk premium + cross-sell berbasis Apriori. Tujuan: naikkan F & M.'}
        for _, r in profile.iterrows():
            with st.container(border=True):
                st.markdown(f"**Klaster {int(r['Cluster'])} → {r['Label']}** ({int(r['Jumlah'])} pelanggan)")
                st.markdown(f"Rata-rata: Recency **{r['R_mean']:.1f} hari** • Frequency **{r['F_mean']:.2f} nota** • Monetary **{fmt_rp(r['M_mean'])}**")
                st.markdown(f"📝 Strategi: {strat.get(r['Label'], 'Sesuaikan dengan profil.')}")

# ============================================================
# HALAMAN 5: DASHBOARD BI
# ============================================================
elif MENU == "5️⃣ Dashboard BI":
    hero('5️⃣ <span class="grad-text-violet">Business Intelligence</span> Dashboard <span class="small" style="color:#dbeafe">(Bobot 15%)</span>',
        "Ringkasan eksekutif untuk manajemen: <b>KPI + tren + produk + Apriori + K-Means + filter</b> dalam satu layar.",
         ["Soal 9 · KPI + penjualan + Apriori + K-Means + filter + bonus integrasi"],
        theme="theme-dash", eyebrow="TAHAP 05 // BI DASHBOARD • KPI • TREN")
    note("📝 <b>Alur halaman ini:</b> ① Baca KPI (filter aktif) → ② Tren bulanan → ③ Top-10 & kategori → ④ Ringkasan mining → ⑤ Bonus pola per segmen. Semua bagian menjelaskan 'apa yang bisa dibaca manajemen'.")

    # KPI memakai filter global
    with st.container(border=True):
        st.markdown(f"### 💰 KPI Utama <span class='small'>(filter: {d0} s/d {d1} • {selected_cat})</span>")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Penjualan", f"{fmt_rp(df_f['Line_Total'].sum())}")
        k2.metric("Jumlah Transaksi", f"{df_f['Transaction_ID'].nunique():,}")
        k3.metric("Pelanggan Unik", f"{df_f['Customer_ID'].nunique():,}")
        k4.metric("Rata-rata Basket", f"{df_f['basket_size'].mean():.2f}")
        st.caption("📝 Definisi (Soal 9a): Total = Σ Line_Total bersih • Transaksi = ID unik • Pelanggan = Customer_ID unik • Basket = mean produk unik/nota. Angka berubah mengikuti filter sidebar.")

    with st.container(border=True):
        st.markdown("### 📈 Tren Penjualan Bulanan (Soal 9b)")
        st.caption("📝 Cara baca: batang biru = Rp penjualan, garis oranye = jumlah nota (sumbu kanan). Keduanya naik = bulan sehat; Rp naik tapi nota turun = belanja per nota membesar.")
        monthly = df_f.groupby('Bulan').agg(Penjualan=('Line_Total', 'sum'), Transaksi=('Transaction_ID', 'nunique')).reset_index()
        monthly['Penjualan_str'] = monthly['Penjualan'].apply(fmt_rp)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=monthly['Bulan'], y=monthly['Penjualan'], name='Penjualan (Rp)', marker_color='#2563eb',
                             customdata=monthly[['Penjualan_str']], hovertemplate='%{x}<br>%{customdata[0]}<extra></extra>',
                             text=monthly['Penjualan_str'], textposition='outside', textfont=dict(size=10), cliponaxis=False))
        fig.add_trace(go.Scatter(x=monthly['Bulan'], y=monthly['Transaksi'], name='Transaksi', yaxis='y2', mode='lines+markers', line=dict(color='#f97316', width=3)))
        _ticks, _labels = ticks_rp(monthly['Penjualan'].tolist(), 5)
        fig.update_layout(yaxis=dict(title='Penjualan (Rp)', tickvals=_ticks, ticktext=_labels), yaxis2=dict(title='Transaksi', overlaying='y', side='right'), hovermode='x unified', height=500,
                          margin=dict(t=50, b=60, l=90, r=40), xaxis=dict(tickangle=-20))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("📝 Semua label batang menampilkan nilai penuh bertitik (bukan singkatan jt/M). Arahkan kursor untuk nilai persis.")
        with st.expander("🔢 Lihat tabel tren bulanan (angka penuh)"):
            _mt = monthly[['Bulan', 'Penjualan', 'Transaksi']].copy()
            _mt['Penjualan'] = _mt['Penjualan'].apply(fmt_rp)
            _mt.columns = ['Bulan', 'Penjualan (Rp, penuh)', 'Transaksi (nota)']
            st.dataframe(_mt, use_container_width=True, hide_index=True)

    with st.container(border=True):
        st.markdown("### 🏆 Top 10 Produk & Distribusi Kategori")
        mode = st.radio("Urutkan Top-10 berdasarkan:", ["Nilai Penjualan (Rp)", "Kuantitas (unit)"], horizontal=True)
        st.caption("📝 Catatan: mode Rp menunjukkan penyumbang revenue; mode unit menunjukkan motor trafik. Keduanya perlu untuk keputusan stok & promo.")
        t1, t2 = st.columns(2)
        with t1:
            if mode.startswith("Nilai"):
                top = df_f.groupby('Product_Name')['Line_Total'].sum().nlargest(10).reset_index()
                top['Rp_str'] = top['Line_Total'].apply(fmt_rp)
                fig = px.bar(top, x='Line_Total', y='Product_Name', orientation='h', color='Line_Total', color_continuous_scale='Blues', title='Top 10 — Revenue',
                             custom_data=['Rp_str'], text='Rp_str')
                fig.update_traces(hovertemplate='%{y}<br>%{custom_data[0]}<extra></extra>', textposition='outside', textfont=dict(size=10), cliponaxis=False)
                _tx, _tl = ticks_rp(top['Line_Total'].tolist(), 5)
                fig.update_layout(xaxis=dict(tickvals=_tx, ticktext=_tl, tickangle=-15),
                                  yaxis={'categoryorder': 'total ascending'}, height=480, margin=dict(t=50, b=60, l=160, r=140))
            else:
                top = df_f.groupby('Product_Name')['Quantity'].sum().nlargest(10).reset_index()
                fig = px.bar(top, x='Quantity', y='Product_Name', orientation='h', color='Quantity', color_continuous_scale='Greens', title='Top 10 — Unit',
                             text='Quantity')
                fig.update_traces(textposition='outside', textfont=dict(size=10), cliponaxis=False)
                fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=480, margin=dict(t=50, b=60, l=160, r=80))
            st.plotly_chart(fig, use_container_width=True)
            st.caption("📝 Label batang menampilkan angka penuh (Rp bertitik, bukan singkatan).")
            with st.expander("🔢 Lihat tabel Top-10 (angka penuh)"):
                _tt = top.copy()
                if 'Line_Total' in _tt.columns:
                    _tt['Line_Total'] = _tt['Line_Total'].apply(fmt_rp)
                    _tt.columns = ['Produk', 'Penjualan (Rp, penuh)']
                else:
                    _tt.columns = ['Produk', 'Kuantitas (unit, penuh)']
                st.dataframe(_tt, use_container_width=True, hide_index=True)
        with t2:
            cat_sales = df_f.groupby('Category')['Line_Total'].sum().reset_index().sort_values('Line_Total', ascending=False)
            cat_sales['Rp_str'] = cat_sales['Line_Total'].apply(fmt_rp)
            _fig_cat = px.bar(cat_sales, x='Category', y='Line_Total', color='Line_Total', color_continuous_scale='Viridis', title='Revenue per Kategori',
                              custom_data=['Rp_str'], text='Rp_str')
            _fig_cat.update_traces(hovertemplate='%{x}<br>%{custom_data[0]}<extra></extra>', textposition='outside', textfont=dict(size=10), cliponaxis=False)
            _cy, _cl = ticks_rp(cat_sales['Line_Total'].tolist(), 5)
            _fig_cat.update_layout(yaxis=dict(tickvals=_cy, ticktext=_cl), xaxis=dict(tickangle=-20), height=480,
                                   margin=dict(t=50, b=90, l=90, r=30))
            st.plotly_chart(_fig_cat, use_container_width=True)
            with st.expander("🔢 Lihat tabel revenue kategori (angka penuh)"):
                _ct = cat_sales[['Category', 'Line_Total']].copy()
                _ct['Line_Total'] = _ct['Line_Total'].apply(fmt_rp)
                _ct.columns = ['Kategori', 'Penjualan (Rp, penuh)']
                st.dataframe(_ct, use_container_width=True, hide_index=True)
            st.caption("📝 Cara baca: kategori teratas = kontributor revenue terbesar — prioritaskan ketersediaan & promo di sana.")

    with st.container(border=True):
        st.markdown("### 🔗 Ringkasan Mining di Dashboard (Soal 9c–9d)")
        # Apriori default untuk dashboard
        rules_d, _ = run_apriori(basket, float(DEFAULT_SUPPORT), float(DEFAULT_CONF), float(DEFAULT_LIFT))
        rfm_seg_d, profile_d = run_kmeans_final(rfm, X_scaled, int(DEFAULT_K))
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**🛒 Aturan asosiasi utama** ({len(rules_d):,} aturan @ support {DEFAULT_SUPPORT}, conf {DEFAULT_CONF})")
            if len(rules_d):
                d = rules_d.copy()
                d['antecedents'] = d['antecedents'].apply(lambda x: ', '.join(sorted(list(x))))
                d['consequents'] = d['consequents'].apply(lambda x: ', '.join(sorted(list(x))))
                st.dataframe(d[['antecedents', 'consequents', 'support', 'confidence', 'lift']].head(5).style.format({'support': '{:.4f}', 'confidence': '{:.3f}', 'lift': '{:.2f}'}), use_container_width=True, hide_index=True)
            else:
                st.info("Tidak ada aturan pada threshold default — lihat halaman 3️⃣ untuk tuning.")
            st.caption("📝 Dibaca manajemen sebagai: 'paket produk apa yang layak di-bundle minggu ini'. Detail penuh di halaman 3️⃣.")
        with c2:
            st.markdown(f"**👥 Komposisi segmen** (k={DEFAULT_K})")
            seg = rfm_seg_d['Segment'].value_counts().reset_index()
            seg.columns = ['Segment', 'Jumlah']
            st.plotly_chart(px.pie(seg, names='Segment', values='Jumlah', hole=.4), use_container_width=True)
            st.caption("📝 Dibaca manajemen sebagai: 'porsi pelanggan setia vs berisiko vs baru'. Detail profil di halaman 4️⃣.")
        with st.expander("🎁 BONUS: Kombinasi produk relevan per segmen (integrasi Apriori × K-Means)"):
            note("📝 <b>Ide bonus soal:</b> tiap segmen punya kantong berbeda — gabungkan peta pelanggan (K-Means) dengan keranjang (Apriori). Di bawah = Top-3 produk per segmen sebagai proksi bundling tersegmentasi.")
            seg_map = rfm_seg_d[['Customer_ID', 'Segment']]
            tmp = df.merge(seg_map, on='Customer_ID', how='left')
            bonus = tmp.groupby(['Segment', 'Product_Name']).size().reset_index(name='Frekuensi')
            bonus = bonus.sort_values(['Segment', 'Frekuensi'], ascending=[True, False]).groupby('Segment').head(3)
            st.dataframe(bonus, use_container_width=True, hide_index=True)
            st.caption("📝 Contoh pakai: bundling untuk 'Loyal' ≠ bundling untuk 'New' — sesuaikan isi paket dengan selera tiap segmen. Analisis lift per segmen dapat dilanjutkan sebagai pengembangan.")
    note("📝 <b>Apa yang bisa dibaca manajemen dari dashboard ini?</b> ① Kesehatan revenue & trafik (KPI + tren) ② Produk/kategori andalan (Top-10) ③ Paket promo siap jalan (Apriori) ④ Siapa disasar promo (K-Means). Untuk tangkapan layar UTS: screenshot tiap kartu + tulis 1 kalimat insight di bawahnya.", "note-green")

# ============================================================
# HALAMAN 6: REKOMENDASI
# ============================================================
elif MENU == "6️⃣ Rekomendasi & Laporan":
    hero('6️⃣ Rekomendasi <span class="grad-text-warm">Keputusan Bisnis</span> <span class="small" style="color:#dbeafe">(Bobot 10%)</span>',
        "Ujung tombak UTS: <b>setiap rekomendasi wajib berbasis bukti data</b>, berprioritas, terukur, dan jujur soal asumsi.",
         ["Soal 10 · Bundling + retensi + prioritas + dampak + rencana + KPI + evaluasi"],
        theme="theme-reco", eyebrow="TAHAP 06 // REKOMENDASI • AKSI BISNIS")
    note("📝 <b>Alur halaman ini:</b> ① Strategi produk (Apriori) ② Strategi pelanggan (K-Means) ③ Prioritas tindakan ④ Dampak + KPI + evaluasi ⑤ Rencana implementasi ⑥ Checklist luaran. Perhatikan pemisahan FAKTA vs ASUMSI.")

    rules_r, _ = run_apriori(basket, float(DEFAULT_SUPPORT), float(DEFAULT_CONF), float(DEFAULT_LIFT))
    rfm_seg_r, profile_r = run_kmeans_final(rfm, X_scaled, int(DEFAULT_K))
    if len(rules_r):
        t1 = rules_r.iloc[0]
        ant1, cons1 = ', '.join(sorted(list(t1['antecedents']))), ', '.join(sorted(list(t1['consequents'])))
        s1, c1_, l1 = f"{t1['support']:.4f}", f"{t1['confidence']:.3f}", f"{t1['lift']:.2f}"
    else:
        ant1, cons1, s1, c1_, l1 = "-", "-", "-", "-", "-"
    seg_counts = rfm_seg_r['Segment'].value_counts()
    n_loyal, n_churn = int(seg_counts.get('Loyal Customers', 0)), int(seg_counts.get('Churn Risk', 0))

    with st.container(border=True):
        st.markdown("### 🛒 A. Strategi Penempatan & Bundling (dari Apriori — Soal 10a)")
        st.markdown(f"""
    1. 🎁 **Bundling `{ant1} + {cons1}`** (aturan terkuat: S={s1}, C={c1_}, L={l1}). Buat paket hemat + label rak 'sering dibeli bersama'.
    2. 🗄️ **Rak berdekatan** untuk pasangan co-occurrence tertinggi (rujuk Heatmap/Network di halaman 3️⃣) — dorong impulse-buying.
    3. 🧾 **Cross-selling kasir & struk:** tawarkan pasangan saat salah satu item terdeteksi di keranjang (aturan Top-10).
        """)
        note("📝 <b>Bukti data:</b> setiap poin merujuk ke lift/confidence terukur, bukan perkiraan. KPI: basket_size naik & attach-rate paket. Cara ukur: bandingkan basket sebelum vs 1 bulan setelah bundling (uji A/B kasir).", "note-green")

    with st.container(border=True):
        st.markdown("### 👥 B. Strategi Retensi per Segmen (dari K-Means — Soal 10b)")
        st.markdown("""
    | Segmen | Bukti profil | Tindakan pemasaran |
    |---|---|---|
    | 👑 **Loyal Customers** | F & M di atas median | Membership premium, poin, early-access; minta review/testimoni |
    | 📲 **Churn Risk** | R tinggi, F rendah | Voucher reaktivasi + pesan personal + survei alasan vakum |
    | 🌱 **New / Potential** | F & M rendah | Onboarding, diskon pembelian ke-2, edukasi produk |
    | 🛒 **Regular** | Di tengah | Upsell premium + cross-sell Apriori untuk naikkan M |
        """)
        note("📝 <b>Bukti data:</b> label segmen diturunkan dari R/F/M rata-rata per klaster (halaman 4️⃣), sehingga strategi menempel pada angka — bukan stereotip.", "note-green")

    with st.container(border=True):
        st.markdown("### 🎯 C. Prioritas Tindakan (Soal 10c — urut dampak × kemudahan)")
        st.caption("📝 Catatan: prioritas = urgensi + potensi manfaat + biaya + kemudahan. Nomor kecil = kerjakan dulu.")
        prio = pd.DataFrame([
            ["1", f"Retensi Loyal ({n_loyal} pelanggan)", "Kontributor revenue utama; kehilangan 1 loyal = kehilangan banyak nota", "Tinggi — jaga 80% revenue", "Rendah (WA + poin)"],
            ["2", f"Reaktivasi Churn ({n_churn} pelanggan)", "Mencegah kebocoran basis pelanggan", "Sedang — pulihkan 5–10% trafik", "Rendah (voucher)"],
            ["3", "Bundling Apriori + rak berdekatan", "Menaikkan basket & attach-rate", "Sedang–Tinggi (+10–15%/nota*)", "Sedang (re-layout rak)"],
            ["4", "Onboarding segmen New/Potential", "Memperbesar pipa calon loyal", "Jangka menengah", "Rendah (otomasi struk)"],
        ], columns=["Prioritas", "Aksi", "Alasan (bukti)", "Potensi manfaat", "Biaya/Kemudahan"])
        st.dataframe(prio, use_container_width=True, hide_index=True)
        st.caption("📝 *Angka % bertanda bintang = ESTIMASI berbasis asumsi pola historis, bukan fakta — lihat bagian D.")

    with st.container(border=True):
        st.markdown("### 📊 D. Estimasi Dampak, Asumsi, KPI & Evaluasi (Soal 10d–10f)")
        f1, f2, f3 = st.columns(3)
        with f1:
            st.markdown("**✅ FAKTA (dari data)**")
            st.markdown(f"- {len(df):,} baris bersih dari {rows_before:,} mentah; {df['Transaction_ID'].nunique():,} nota; basket rata-rata {df['basket_size'].mean():.2f}.  \n- Aturan terkuat lift {l1}; {len(profile_r)} segmen terbentuk.")
        with f2:
            st.markdown("**🔮 ESTIMASI (asumsi)**")
            st.markdown("- Bundling +10–15%/nota; reaktivasi +5–10% trafik.  \n- Asumsi: pola historis berlanjut; promo dieksekusi konsisten; tidak ada guncangan harga/stok.")
        with f3:
            st.markdown("**⏳ DAMPAK AKTUAL (nanti)**")
            st.markdown("- Baru diketahui SETELAH implementasi via eksperimen (A/B, pre-post). Jangan klaim kenaikan sebagai fakta sebelum diuji.")
        note("📝 <b>Kejujuran analisis (wajib UTS):</b> bedakan ketiga kotak di atas. Dosen menilai kemampuan ini — klaim tanpa uji = pengurang nilai.", "note-amber")
        st.markdown("""
    **KPI keberhasilan & cara mengukur:**
    - 💰 Revenue bulanan (target +10%) — dari POS, bandingkan pre vs post.
    - 🧺 Basket_size & attach-rate bundling — dari struk.
    - 🔁 Repeat purchase rate +5% — dari RFM ulang tiap kuartal.
    - 👑 Porsi Loyal +10%, 📉 Churn −5% — dari K-Means ulang.
        """)
        st.markdown("**🗓️ E. Rencana implementasi (Soal 10e)**")
        r1, r2 = st.columns(2)
        with r1:
            st.markdown("**Jangka pendek (1–3 bln)**\n- Terapkan 3 bundling teratas + re-layout rak.\n- Blast reaktivasi Churn + onboarding New.\n- Pasang KPI dashboard mingguan.")
        with r2:
            st.markdown("**Jangka menengah (3–6 bln)**\n- Program loyalitas poin terintegrasi POS.\n- Refresh Apriori & K-Means tiap kuartal.\n- Uji A/B promo tersegmentasi (Loyal vs Regular).")

    with st.container(border=True):
        st.markdown("### 📦 Checklist Luaran UTS (biar tidak ada yang tertinggal)")
        st.markdown("""
    - [ ] **Laporan analisis** — narasi Business Understanding → rekomendasi (setiap angka di dashboard ini = bahan paragraf).
    - [ ] **Dataset hasil preprocessing** — unduh di bawah + dokumentasi perubahan.
    - [ ] **Notebook/kode** — file `app.py` ini + workbook perhitungan.
    - [ ] **Dashboard & tangkapan layar** — screenshot tiap halaman + 1 kalimat insight per gambar.
    - [ ] **Presentasi manajemen** — 5 slide: masalah → temuan → segmen → bundling → rencana.
        """)
        c1, c2, c3, c4 = st.columns(4)
        c1.download_button("⬇️ Data_Clean", df.to_csv(index=False).encode('utf-8'), "Data_Clean.csv", "text/csv", use_container_width=True)
        c2.download_button("⬇️ RFM + Segmen", rfm_seg_r.to_csv(index=False).encode('utf-8'), "RFM_segmen.csv", "text/csv", use_container_width=True)
        if len(rules_r):
            dd = rules_r.copy()
            dd['antecedents'] = dd['antecedents'].apply(lambda x: ', '.join(sorted(list(x))))
            dd['consequents'] = dd['consequents'].apply(lambda x: ', '.join(sorted(list(x))))
            c3.download_button("⬇️ Rules", dd.to_csv(index=False).encode('utf-8'), "association_rules.csv", "text/csv", use_container_width=True)
        c4.download_button("⬇️ Profil Klaster", profile_r.to_csv(index=False).encode('utf-8'), "cluster_profile.csv", "text/csv", use_container_width=True)
        st.caption("📝 Catatan: keempat file di atas + screenshot dashboard + laporan = paket pengumpulan lengkap (Bagian E soal).")

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption("🎓 UTS Business Intelligence Kelas B — Universitas Timor (UNIMOR).  •  ⚠️ Estimasi dampak bersifat asumsi pola historis, bukan hasil eksperimen.  •  Streamlit • mlxtend (Apriori) • scikit-learn (K-Means) • Plotly.")
