# dashboard_kpp_final_v5.py
# PT KALIMANTAN PRIMA PERSADA - TWB Analytics Dashboard
# FINAL VERSION v5.0 - Trend Analysis Diubah Menjadi Financial Analysis
# Run: streamlit run dashboard_kpp_final_v5.py

import warnings
warnings.filterwarnings('ignore')

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from PIL import Image
import os
import base64
import io
from datetime import datetime
from pathlib import Path
import openpyxl
import openpyxl.styles

st.set_page_config(
    page_title="Mining Volume Deviation Monitoring PT. KPP",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Path logo - sesuaikan dengan path logo Anda
BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "assets" / "logo.png"
DEMO_DATA_DIR = BASE_DIR / "demo_data"

def get_available_demo_files():
    if not DEMO_DATA_DIR.exists():
        return {}

    excel_files = sorted(DEMO_DATA_DIR.glob("*.xlsx"))
    return {file.stem: file for file in excel_files}

# Fungsi untuk mengonversi gambar ke base64 dengan resize
def get_base64_image(image_path, size=(120, 120)):  # Diperbesar dari 60 ke 120
    try:
        # Baca gambar
        if Path(image_path).exists():
            img = Image.open(image_path)
            
            # Resize gambar ke ukuran yang diinginkan
            img = img.resize(size, Image.Resampling.LANCZOS)
            
            # Konversi ke base64
            import io
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            return base64.b64encode(buffered.getvalue()).decode()
        else:
            st.warning(f"⚠️ Logo tidak ditemukan di path: {image_path}")
            return None
    except Exception as e:
        st.warning(f"⚠️ Error loading logo: {str(e)}")
        return None

# Fungsi untuk menampilkan logo dengan berbagai metode
def display_logo(logo_path, width=120):
    try:
        logo_path = Path(logo_path)
        if logo_path.exists():
            img = Image.open(logo_path)
            img = img.resize((width, width), Image.Resampling.LANCZOS)
            return img
        else:
            from PIL import ImageDraw, ImageFont
            img = Image.new('RGB', (width, width), color=(0, 40, 80))
            d = ImageDraw.Draw(img)
            try:
                font = ImageFont.truetype("arial.ttf", 40)
            except Exception:
                font = ImageFont.load_default()
            d.text((width // 2 - 40, width // 2 - 20), "KPP", fill=(0, 100, 136), font=font)
            return img
    except Exception as e:
        st.error(f"Error loading logo: {e}")
        from PIL import ImageDraw, ImageFont
        img = Image.new('RGB', (width, width), color=(0, 40, 80))
        d = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("arial.ttf", 40)
        except Exception:
            font = ImageFont.load_default()
        d.text((width // 2 - 40, width // 2 - 20), "KPP", fill=(0, 255, 136), font=font)
        return img
        
    except Exception as e:
        st.error(f"Error loading logo: {e}")
        # Return placeholder
        from PIL import ImageDraw, ImageFont
        img = Image.new('RGB', (width, width), color=(0, 40, 80))
        d = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("arial.ttf", 40)  # Diperbesar font
        except:
            font = ImageFont.load_default()
        d.text((width//2-40, width//2-20), "KPP", fill=(0, 255, 136), font=font)
        return img

# Coba konversi logo ke base64
logo_base64 = get_base64_image(LOGO_PATH, size=(200, 200))

# COMPLETE CSS - DARK MODE PREMIUM (DENGAN WARNA YANG DIPERBAIKI)

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');

* {{
    font-family: 'Inter', sans-serif;
    margin: 0;
    padding: 0;
}}

.block-container {{
    background: linear-gradient(135deg, #0a0a0f 0%, #1a1a2e 50%, #16213e 100%);
    color: #e5e7eb;
    padding: 1rem 2rem 2rem 2rem !important;
    max-width: 100% !important;
}}

[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #0f0f1a 0%, #0a0a0f 100%);
    border-right: 1px solid rgba(255,255,255,0.1);
}}

[data-testid="stSidebar"] * {{
    color: #e5e7eb !important;
}}

/* JANGAN SEMBUNYIKAN HEADER */
footer {{ visibility: hidden; }}

/* Tambahkan style untuk header Streamlit */
header {{
    background: transparent !important;
}}

/* HEADER - DENGAN LOGO BESAR DAN BORDER */
.main-header {{
    background: linear-gradient(135deg, #000000 0%, #1a1a2e 70%, #16213e 100%);
    padding: 1.8rem 2.5rem;  /* Diperbesar padding untuk logo besar */
    border-radius: 0 0 28px 28px;
    margin: -1rem -2rem 2rem -2rem;
    box-shadow: 0 20px 60px rgba(0,0,0,0.6);
    border: 1px solid rgba(255,255,255,0.12);
}}

.header-container {{
    display: flex;
    align-items: center;
    gap: 2rem;  /* Diperbesar gap */
}}

.logo-container {{
    flex-shrink: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 3px;
    background: #ffffff;
    border-radius: 20px;
    border: 3px solid #22c55e;  /* Border hijau solid lebih tebal */
    box-shadow: 
        0 0 20px rgba(0,255,136,0.4),    /* Glow pertama */
        0 0 40px rgba(0,255,136,0.2),    /* Glow kedua */
        0 8px 30px rgba(0,0,0,0.1);      /* Shadow biasa */
}}

.logo-img {{
    width: 120;  /* DIPERBESAR dari 60px ke 120px */
    height: 120px; /* DIPERBESAR dari 60px ke 120px */
    border-radius: 14px;
    object-fit: contain;
    background: rgba(255,255,255,0.03);
    padding: 3px;
}}

/* TITLE SECTION */
.title-wrapper {{
    flex-grow: 1;
    text-align: left;
    padding-left: 0.5rem;
}}

.main-title {{
    font-size: 2.4rem;  /* Sedikit diperbesar */
    font-weight: 900;
    margin: 0 0 0.4rem 0;
    background: linear-gradient(135deg, #22c55e 0%, #00ffcc 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    line-height: 1.2;
    letter-spacing: -0.5px;
}}

.main-subtitle {{
    color: #d1d5db;
    font-size: 1.1rem;  /* Diperbesar */
    font-weight: 400;
    letter-spacing: 0.3px;
    margin: 0;
    padding-left: 0.2rem;
}}

/* STATS SECTION */
.stats-grid {{
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 0.8rem;
    flex-shrink: 0;
    min-width: 220px;
}}

.stat-card {{
    background: linear-gradient(135deg, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0.05) 100%);
    backdrop-filter: blur(20px);
    border-radius: 14px;
    padding: 0.9rem;  /* Sedikit diperbesar */
    text-align: center;
    border: 1px solid rgba(255,255,255,0.15);
    transition: all 0.3s ease;
}}

.stat-card:hover {{
    transform: translateY(-3px);
    box-shadow: 0 10px 30px rgba(0,255,136,0.2);
}}

.stat-value {{
    font-size: 1.5rem;  /* Diperbesar */
    font-weight: 900;
    color: #00ff88;
    display: block;
    line-height: 1;
}}

.stat-label {{
    font-size: 0.75rem;  /* Diperbesar */
    color: #9ca3af;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-top: 0.4rem;
    display: block;
}}

/* PERFORMANCE CARDS */
.perf-card {{
    background: linear-gradient(135deg, rgba(20,20,35,0.95) 0%, rgba(15,15,25,0.95) 100%);
    backdrop-filter: blur(25px);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 22px;
    padding: 2rem 1.8rem;
    box-shadow: 0 18px 55px rgba(0,0,0,0.5);
    transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
    position: relative;
    overflow: hidden;
}}

.perf-card::before {{
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 4px;
    background: var(--accent-color);
}}

.perf-card:hover {{
    transform: translateY(-10px);
    box-shadow: 0 25px 70px rgba(0,0,0,0.6);
}}

.perf-card.normal {{ --accent-color: linear-gradient(90deg, #22c55e, #10b981); }}
.perf-card.caution {{ --accent-color: linear-gradient(90deg, #eab308, #f59e0b); }} /* KUNING */
.perf-card.critical {{ --accent-color: linear-gradient(90deg, #ef4444, #dc2626); }} /* MERAH */
.perf-card.info {{ --accent-color: linear-gradient(90deg, #3b82f6, #2563eb); }}

.card-icon {{
    font-size: 2.6rem;
    margin-bottom: 1rem;
    display: block;
    opacity: 0.95;
}}

.card-label {{
    font-size: 0.74rem;
    color: #9ca3af;
    text-transform: uppercase;
    letter-spacing: 1.1px;
    font-weight: 600;
    margin-bottom: 0.8rem;
}}

.card-value {{
    font-size: 2.6rem;
    font-weight: 900;
    color: #ffffff;
    line-height: 1;
    margin-bottom: 0.6rem;
}}

.card-subtitle {{
    font-size: 0.86rem;
    color: #d1d5db;
    margin-bottom: 1rem;
}}

/* STATUS BADGES - WARNA DIPERBAIKI (Caution=Kuning, Critical=Merah) */
.status-badge {{
    padding: 0.55rem 1.25rem;
    border-radius: 50px;
    font-weight: 700;
    font-size: 0.8rem;
    border: 2px solid;
    display: inline-flex;
    align-items: center;
    gap: 0.6rem;
    transition: all 0.3s ease;
}}

.status-badge::before {{
    content: '';
    width: 8px;
    height: 8px;
    border-radius: 50%;
    animation: pulseDot 2s ease-in-out infinite;
}}

@keyframes pulseDot {{
    0%, 100% {{ opacity: 1; transform: scale(1); }}
    50% {{ opacity: 0.6; transform: scale(1.3); }}
}}

/* WARNA BENAR: Normal=Hijau, Caution=Kuning, Critical=Merah */
.badge-normal {{ 
    background: rgba(34,197,94,0.15); 
    color: #22c55e; 
    border-color: #22c55e; 
}}
.badge-normal::before {{ background: #22c55e; }}

.badge-caution {{ 
    background: rgba(234,179,8,0.15); 
    color: #eab308; 
    border-color: #eab308; 
}}
.badge-caution::before {{ background: #eab308; }}

.badge-critical {{ 
    background: rgba(239,68,68,0.15); 
    color: #ef4444; 
    border-color: #ef4444; 
}}
.badge-critical::before {{ background: #ef4444; }}

/* SECTION HEADERS */
.section-header {{
    display: flex;
    align-items: center;
    margin: 2rem 0 1rem 0;
    padding: 0.65rem 1.2rem;
    background: linear-gradient(135deg, rgba(5,46,22,0.55), rgba(15,23,42,0.7));
    border-left: 3px solid #22c55e;
    border-radius: 0 10px 10px 0;
    gap: 0;
}}
.section-header .section-icon {{ display: none; }}
.section-title {{
    font-size: 1.15rem !important;
    font-weight: 700 !important;
    color: #e2e8f0 !important;
    margin: 0 !important;
    padding: 0 !important;
    letter-spacing: 0.01em;
    line-height: 1.3;
}}

/* MATERIAL FLOW */
.flow-container {{
    background: linear-gradient(135deg, rgba(15,15,30,0.95) 0%, rgba(10,10,20,0.95) 100%);
    backdrop-filter: blur(30px);
    border: 2px solid rgba(255,255,255,0.15);
    border-radius: 26px;
    padding: 2.5rem;
    margin: 1.5rem 0 2rem 0;
    box-shadow: 0 22px 65px rgba(0,0,0,0.5);
}}

.flow-title {{
    text-align: center;
    font-size: 1.3rem;
    font-weight: 800;
    color: #ffffff;
    margin-bottom: 2rem;
    letter-spacing: 0.3px;
}}

.flow-chain {{
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 1.8rem;
    flex-wrap: wrap;
}}

.flow-item {{
    background: linear-gradient(135deg, rgba(255,255,255,0.08) 0%, rgba(255,255,255,0.03) 100%);
    border: 2px solid rgba(255,255,255,0.15);
    border-radius: 20px;
    padding: 1.8rem 1.2rem;
    text-align: center;
    min-width: 145px;
    transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
    box-shadow: 0 10px 30px rgba(0,0,0,0.4);
}}

.flow-item:hover {{
    transform: translateY(-12px) scale(1.05);
    border-color: var(--flow-color);
    box-shadow: 0 20px 50px rgba(0,0,0,0.6);
}}

.flow-1 {{ --flow-color: #3b82f6; }}
.flow-2 {{ --flow-color: #22c55e; }}
.flow-3 {{ --flow-color: #eab308; }}
.flow-4 {{ --flow-color: #ef4444; }}
.flow-5 {{ --flow-color: #8b5cf6; }}

.flow-item-title {{
    font-weight: 900;
    font-size: 1rem;
    color: #ffffff;
    margin-bottom: 0.3rem;
    letter-spacing: 0.2px;
}}

.flow-item-subtitle {{
    font-size: 0.72rem;
    color: #9ca3af;
    margin-bottom: 0.8rem;
}}

.flow-item-value {{
    font-size: 1.4rem;
    font-weight: 800;
    color: var(--flow-color);
    margin-top: 0.5rem;
}}

.flow-arrow {{
    font-size: 2.2rem;
    color: #00ff88;
    animation: arrowSlide 2s ease-in-out infinite;
    filter: drop-shadow(0 0 15px rgba(0,255,136,0.4));
}}

@keyframes arrowSlide {{
    0%, 100% {{ transform: translateX(0); opacity: 0.8; }}
    50% {{ transform: translateX(12px); opacity: 1; }}
}}

.flow-efficiency {{
    text-align: center;
    margin-top: 1.5rem;
    padding-top: 1.5rem;
    border-top: 1px solid rgba(255,255,255,0.1);
}}

.flow-eff-label {{
    color: #9ca3af;
    font-size: 0.95rem;
    margin-bottom: 0.5rem;
}}

.flow-eff-value {{
    color: #00ff88;
    font-size: 1.8rem;
    font-weight: 900;
}}

/* PIE CHARTS */
.pie-grid {{
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 2rem;
    margin: 2rem 0;
}}

.pie-card {{
    background: linear-gradient(135deg, rgba(15,15,30,0.95) 0%, rgba(10,10,20,0.95) 100%);
    border-radius: 20px;
    padding: 1.8rem;
    backdrop-filter: blur(25px);
    border: 1px solid rgba(255,255,255,0.12);
    box-shadow: 0 15px 45px rgba(0,0,0,0.4);
    transition: all 0.3s ease;
}}

.pie-card:hover {{
    transform: translateY(-8px);
    box-shadow: 0 20px 60px rgba(0,0,0,0.5);
    border-color: rgba(0,255,136,0.3);
}}

.pie-title {{
    font-size: 1.1rem;
    font-weight: 800;
    color: #ffffff;
    margin-bottom: 1.2rem;
    text-align: center;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.6rem;
}}

/* MATRIX TABLE */
.matrix-container {{
    background: linear-gradient(135deg, rgba(15,15,30,0.95) 0%, rgba(10,10,20,0.95) 100%);
    backdrop-filter: blur(25px);
    border-radius: 22px;
    overflow: hidden;
    border: 1px solid rgba(255,255,255,0.12);
    margin: 2rem 0;
    box-shadow: 0 20px 60px rgba(0,0,0,0.5);
}}

.matrix-header {{
    background: linear-gradient(135deg, rgba(0,255,136,0.15) 0%, rgba(0,255,136,0.05) 100%);
    padding: 1.3rem 2rem;
    border-bottom: 2px solid rgba(255,255,255,0.1);
}}

.matrix-title {{
    color: #ffffff;
    font-size: 1.2rem;
    font-weight: 800;
    margin: 0;
}}

.matrix-row {{
    display: grid;
    grid-template-columns: 2fr 1fr 1fr 1fr 1fr 1.2fr;
    gap: 1rem;
    padding: 1.4rem 2rem;
    border-bottom: 1px solid rgba(255,255,255,0.05);
    align-items: center;
    transition: all 0.3s ease;
}}

.matrix-row:hover {{
    background: rgba(0,255,136,0.05);
}}

.matrix-row:last-child {{
    border-bottom: none;
}}

.matrix-stage {{
    font-weight: 700;
    color: #ffffff;
    font-size: 1rem;
}}

.matrix-cell {{
    text-align: center;
}}

.matrix-number {{
    font-size: 1.2rem;
    font-weight: 800;
    display: block;
    line-height: 1;
}}

.matrix-label {{
    font-size: 0.7rem;
    color: #9ca3af;
    display: block;
    margin-top: 0.3rem;
    text-transform: uppercase;
}}

/* WARNA MATRIX: Normal=Hijau, Caution=Kuning, Critical=Merah */
.num-normal {{ color: #22c55e; }}
.num-caution {{ color: #eab308; }}
.num-critical {{ color: #ef4444; }}
.num-total {{ color: #ffffff; }}
.num-avg {{ color: #00ff88; }}

/* TABS */
.stTabs [data-baseweb="tab-list"] {{
    gap: 1rem;
    background: transparent;
    border-bottom: 2px solid rgba(255,255,255,0.1);
    padding: 0;
}}

.stTabs [data-baseweb="tab"] {{
    border-radius: 14px 14px 0 0;
    padding: 1rem 2.5rem;
    font-weight: 700;
    font-size: 1.05rem;
    border: none;
    background: rgba(255,255,255,0.05);
    color: #9ca3af;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}}

.stTabs [data-baseweb="tab"]:hover {{
    background: rgba(255,255,255,0.1);
    color: #e5e7eb;
    transform: translateY(-2px);
}}

.stTabs [aria-selected="true"] {{
    background: linear-gradient(135deg, #00ff88 0%, #00cc6a 100%) !important;
    color: #000000 !important;
    font-weight: 900 !important;
    box-shadow: 0 8px 25px rgba(0,255,136,0.4);
    transform: translateY(-2px);
}}

/* GENERAL */
h1, h2, h3, h4, h5, h6 {{
    color: #ffffff !important;
}}

p, span, div, label {{
    color: #e5e7eb;
}}

[data-testid="stMetric"] {{
    background: rgba(15,15,30,0.9);
    padding: 1.3rem;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.1);
}}

[data-testid="stMetricLabel"] {{
    color: #9ca3af !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
}}

[data-testid="stMetricValue"] {{
    color: #ffffff !important;
    font-weight: 800 !important;
}}

/* Data Table Styling */
.data-table-container {{
    background: linear-gradient(135deg, rgba(15,15,30,0.95) 0%, rgba(10,10,20,0.95) 100%);
    backdrop-filter: blur(25px);
    border-radius: 18px;
    padding: 1.5rem;
    border: 1px solid rgba(255,255,255,0.12);
    margin-top: 1.5rem;
    box-shadow: 0 15px 40px rgba(0,0,0,0.4);
}}

/* Financial Cards Styling */
.financial-card {{
    background: linear-gradient(135deg, rgba(20,20,35,0.95) 0%, rgba(15,15,25,0.95) 100%);
    backdrop-filter: blur(25px);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 16px;
    padding: 1.5rem;
    box-shadow: 0 12px 40px rgba(0,0,0,0.4);
    margin: 0.5rem 0;
    text-align: center;
    height: 100%;
}}

.financial-main-value {{
    font-size: 2rem;
    font-weight: 900;
    color: #00ff88;
    margin-bottom: 0.5rem;
}}

.financial-main-label {{
    font-size: 0.9rem;
    color: #9ca3af;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 1rem;
}}

.financial-sub-metric {{
    font-size: 1.2rem;
    font-weight: 700;
    color: #ffffff;
    margin-top: 0.5rem;
}}

.financial-sub-label {{
    font-size: 0.8rem;
    color: #9ca3af;
}}

.aging-container {{
    background: linear-gradient(135deg, rgba(15,15,30,0.95) 0%, rgba(10,10,20,0.95) 100%);
    border-radius: 16px;
    padding: 1.5rem;
    border: 1px solid rgba(255,255,255,0.12);
    margin-top: 1.5rem;
    height: 100%;
}}
</style>
""", unsafe_allow_html=True)

# HELPER FUNCTIONS
def get_kpi_status(dev_pct):
    """Status: Normal ≤2%, Caution 2-3%, Critical >3%"""
    if pd.isna(dev_pct): 
        return 'Unknown', 'info'
    abs_dev = abs(dev_pct)
    if abs_dev <= 2: 
        return 'Normal', 'normal'
    elif abs_dev <= 3: 
        return 'Caution', 'caution'
    else: 
        return 'Critical', 'critical'

def format_number(num, decimals=2):
    if pd.isna(num): 
        return '-'
    return f"{num:,.{decimals}f}"

def format_large(num):
    if pd.isna(num): 
        return '-'
    if abs(num) >= 1e6: 
        return f"{num/1e6:.2f}M"
    elif abs(num) >= 1e3: 
        return f"{num/1e3:.1f}K"
    return f"{num:.0f}"

def get_active_data_source():
    available_demo_files = get_available_demo_files()

    with st.sidebar:
        st.markdown("### DATA SOURCE")

        source_mode = st.radio(
            "Pilih sumber data:",
            ["Data Training", "Upload Excel"],
            index=0,
            key="data_source_mode"
        )

        selected_file = None
        selected_file_name = None
        selected_source_type = None

        if source_mode == "Data Training":
            if not available_demo_files:
                st.warning("Folder demo_data belum berisi file demo (.xlsx).")
            else:
                demo_choice = st.selectbox(
                    "Data Training:",
                    list(available_demo_files.keys()),
                    key="demo_file_choice"
                )
                selected_file = available_demo_files[demo_choice]
                selected_file_name = available_demo_files[demo_choice].name
                selected_source_type = "demo"
                st.success(f"File Aktif: {selected_file_name}")

        else:
            uploaded = st.file_uploader(
                "Upload Excel File",
                type=["xlsx", "xls"],
                key="excel_uploader"
            )
            if uploaded is not None:
                selected_file = uploaded
                selected_file_name = uploaded.name
                selected_source_type = "upload"
                st.success(f"File aktif: {uploaded.name}")

    return selected_file, selected_file_name, selected_source_type

# Tambahkan fungsi ini setelah fungsi format_large
# ========== HELPER FUNCTIONS ==========
def sort_months_chronologically(month_list):
    """Mengurutkan bulan secara kronologis dari Januari-Desember"""
    month_order = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 
                   'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
    
    def month_key(month_str):
        # Format: "Februari-2024" atau "Februari"
        try:
            # Pisahkan bulan dan tahun
            if '-' in month_str:
                parts = month_str.split('-')
                month_part = parts[0]
                # Cari tahun jika ada
                if len(parts) > 1 and parts[1].isdigit():
                    year_part = int(parts[1])
                else:
                    year_part = 2024
            else:
                month_part = month_str
                year_part = 2024
            
            # Cari indeks bulan
            if month_part in month_order:
                month_idx = month_order.index(month_part)
            else:
                # Jika bulan tidak dikenal, coba parse sebagai angka
                try:
                    month_idx = int(month_part) - 1
                except:
                    month_idx = 0
            
            return (year_part, month_idx)
        except:
            return (2024, 0)
    
    return sorted(month_list, key=month_key)

def get_status_color(status):
    """WARNA DIPERBAIKI: Normal=Hijau, Caution=Kuning, Critical=Merah"""
    if status == 'Normal':
        return '#22c55e'  # HIJAU
    elif status == 'Caution':
        return '#eab308'  # KUNING
    elif status == 'Critical':
        return '#ef4444'  # MERAH
    else:
        return '#9ca3af'
def render_status_distribution(status_col, total_count, key_prefix=""):
    """Reusable status distribution component"""
    if total_count > 0:
        status_counts = status_col.value_counts()
        cols1, cols2, cols3 = st.columns(3)
        status_order = ['Normal', 'Caution', 'Critical']
        for idx, status in enumerate(status_order):
            count = status_counts.get(status, 0)
            percentage = (count / total_count * 100) if total_count > 0 else 0
            color = {'Normal': '#22c55e', 'Caution': '#eab308', 'Critical': '#ef4444'}.get(status, '#9ca3af')
            col_status = [cols1, cols2, cols3][idx]
            with col_status:
                st.markdown(f"""
                    <div style='text-align:center;margin:5px 0;padding:12px;
                                background:rgba(255,255,255,0.05);border-radius:10px;
                                border-left:4px solid {color}'>
                        <div style='font-size:0.85rem;font-weight:600;color:{color};
                                    margin-bottom:5px;text-transform:uppercase'>{status}</div>
                        <div style='font-size:1.2rem;font-weight:800;color:#ffffff;line-height:1'>{count}</div>
                        <div style='font-size:0.75rem;color:#9ca3af;margin-top:3px'>{percentage:.1f}% of total</div>
                    </div>
                """, unsafe_allow_html=True)

@st.cache_data
def load_data(file):
    try:
        excel_file = pd.ExcelFile(file)
        required_sheets = ["OB Monthly", "CH CM"]

        missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
        if missing_sheets:
            st.error(f"❌ Sheet tidak ditemukan: {', '.join(missing_sheets)}")
            return None, None

        # Load OB data
        df_ob = pd.read_excel(file, sheet_name="OB Monthly")
        df_ob["Dev_Absolut"] = df_ob["JS"] - df_ob["TC"]
        df_ob["Dev_Relatif_Pct"] = ((df_ob["JS"] - df_ob["TC"]) / df_ob["TC"]) * 100
        df_ob["Status"], df_ob["Status_Color"] = zip(*df_ob["Dev_Relatif_Pct"].apply(get_kpi_status))

        # Load CH/CM data
        df_ch_cm = pd.read_excel(file, sheet_name="CH CM", skiprows=1)
        df_ch_cm.columns = ["Date", "Port_Darat", "Port_Laut", "CPP_Raw", "CPP_Product", "Sales", "CH_WB", "CM_WB"]
        df_ch_cm = df_ch_cm.dropna(subset=["Date"])
        df_ch_cm["Date"] = pd.to_datetime(df_ch_cm["Date"])

        # Hitung total stok Port & CPP
        df_ch_cm["Port_Total"] = df_ch_cm["Port_Darat"] + df_ch_cm["Port_Laut"]
        df_ch_cm["CPP_Total"] = df_ch_cm["CPP_Raw"].fillna(0) + df_ch_cm["CPP_Product"].fillna(0)

        # Baseline dari baris pertama
        if len(df_ch_cm) > 0:
            port_baseline = df_ch_cm["Port_Total"].iloc[0]
            cpp_baseline = df_ch_cm["CPP_Total"].iloc[0]

            df_ch_cm["TWB_CH"] = df_ch_cm["Sales"].fillna(0) + df_ch_cm["Port_Total"] - port_baseline
            df_ch_cm["TWB_CM"] = df_ch_cm["TWB_CH"] + df_ch_cm["CPP_Total"] - cpp_baseline
        else:
            df_ch_cm["TWB_CH"] = 0
            df_ch_cm["TWB_CM"] = 0

        # CH deviation
        df_ch_cm["Dev_CH_Absolut"] = df_ch_cm["TWB_CH"] - df_ch_cm["CH_WB"]
        df_ch_cm["Dev_CH_Relatif_Pct"] = np.where(
            df_ch_cm["CH_WB"] != 0,
            ((df_ch_cm["TWB_CH"] - df_ch_cm["CH_WB"]) / df_ch_cm["CH_WB"]) * 100,
            np.nan
        )
        df_ch_cm["Status_CH"], df_ch_cm["Status_CH_Color"] = zip(*df_ch_cm["Dev_CH_Relatif_Pct"].apply(get_kpi_status))

        # CM deviation
        df_ch_cm["Dev_CM_Absolut"] = df_ch_cm["TWB_CM"] - df_ch_cm["CM_WB"]
        df_ch_cm["Dev_CM_Relatif_Pct"] = np.where(
            df_ch_cm["CM_WB"] != 0,
            ((df_ch_cm["TWB_CM"] - df_ch_cm["CM_WB"]) / df_ch_cm["CM_WB"]) * 100,
            np.nan
        )
        df_ch_cm["Status_CM"], df_ch_cm["Status_CM_Color"] = zip(*df_ch_cm["Dev_CM_Relatif_Pct"].apply(get_kpi_status))

        return df_ob, df_ch_cm

    except Exception as e:
        st.error(f"❌ Error loading data: {str(e)}")
        return None, None

def create_matrix(df_ob, df_ch_cm):
    # OB stats
    ob_normal = len(df_ob[df_ob['Status'] == 'Normal'])
    ob_caution = len(df_ob[df_ob['Status'] == 'Caution'])
    ob_critical = len(df_ob[df_ob['Status'] == 'Critical'])
    ob_avg = df_ob['Dev_Relatif_Pct'].abs().mean()

    # CH stats
    df_ch = df_ch_cm[df_ch_cm['CH_WB'].notna()]
    ch_normal = len(df_ch[df_ch['Status_CH'] == 'Normal'])
    ch_caution = len(df_ch[df_ch['Status_CH'] == 'Caution'])
    ch_critical = len(df_ch[df_ch['Status_CH'] == 'Critical'])
    ch_avg = df_ch['Dev_CH_Relatif_Pct'].abs().mean() if len(df_ch) > 0 else 0

    # CM stats
    df_cm = df_ch_cm[df_ch_cm['CM_WB'].notna()]
    cm_normal = len(df_cm[df_cm['Status_CM'] == 'Normal'])
    cm_caution = len(df_cm[df_cm['Status_CM'] == 'Caution'])
    cm_critical = len(df_cm[df_cm['Status_CM'] == 'Critical'])
    cm_avg = df_cm['Dev_CM_Relatif_Pct'].abs().mean() if len(df_cm) > 0 else 0

    return pd.DataFrame({
        'Tahapan': ['Overburden (OB)', 'Coal Hauling (CH/Port)', 'Coal Mining (CM/CPP33)'],
        'Normal': [ob_normal, ch_normal, cm_normal],
        'Caution': [ob_caution, ch_caution, cm_caution],
        'Critical': [ob_critical, ch_critical, cm_critical],
        'Total': [len(df_ob), len(df_ch), len(df_cm)],
        'Avg': [ob_avg, ch_avg, cm_avg]
    })

def create_flow(df_ch_cm):
    df = df_ch_cm.dropna(subset=["TWB_CM", "TWB_CH"]).copy()
    if len(df) == 0:
        return None

    # Perubahan stok antar periode
    df["Delta CPP Stock"] = df["CPP_Total"].diff().fillna(0)
    df["Delta Port Stock"] = df["Port_Total"].diff().fillna(0)

    records = []
    for _, row in df.iterrows():
        cm_twb = row["TWB_CM"]
        ch_twb = row["TWB_CH"]
        sales = row["Sales"]

        delta_cpp = row["Delta CPP Stock"]
        delta_port = row["Delta Port Stock"]

        # Deviation rekonsiliasi
        # TWB_CM ≈ TWB_CH + Delta CPP Stock
        # TWB_CH ≈ Sales + Delta Port Stock
        dev_cpp = cm_twb - (ch_twb + delta_cpp)
        dev_port = ch_twb - (sales + delta_port)

        # Flow ratio, bukan recovery
        ch_ratio = (ch_twb / cm_twb * 100) if pd.notna(cm_twb) and cm_twb != 0 else 0
        sales_ratio = (sales / ch_twb * 100) if pd.notna(ch_twb) and ch_twb != 0 else 0

        records.append({
            "Date": row["Date"],
            "CPP Stock": row["CPP_Total"],
            "Port Stock": row["Port_Total"],
            "Delta CPP Stock": delta_cpp,
            "Delta Port Stock": delta_port,
            "CM TWB": cm_twb,
            "CH TWB": ch_twb,
            "Sales": sales,
            "Deviation CPP": dev_cpp,
            "Deviation Port": dev_port,
            "CH Flow Ratio (%)": ch_ratio,
            "Sales Flow Ratio (%)": sales_ratio,
        })

    return pd.DataFrame(records)

# CHART THEME
THEME = {
    'layout': {
        'font': {'family': 'Inter, sans-serif', 'color': '#e5e7eb'},
        'plot_bgcolor': 'rgba(0,0,0,0)',
        'paper_bgcolor': 'rgba(0,0,0,0)',
        'colorway': ['#00ff88', '#3b82f6', '#eab308', '#ef4444', '#8b5cf6'],
        'xaxis': {'gridcolor': 'rgba(255,255,255,0.1)', 'zerolinecolor': 'rgba(255,255,255,0.15)'},
        'yaxis': {'gridcolor': 'rgba(255,255,255,0.1)', 'zerolinecolor': 'rgba(255,255,255,0.15)'}
    }
}

def main():
    st.markdown('<div class="main-header">', unsafe_allow_html=True)

    logo_base64_final = get_base64_image(LOGO_PATH, size=(100, 100))

    if logo_base64_final:
        st.markdown(f"""
        <div class="header-container">
            <div class="logo-container">
                <img src="data:image/png;base64,{logo_base64_final}" class="logo-img" alt="KPP Logo">
            </div>
            <div class="title-wrapper">
                <h1 class="main-title">Mining Volume Deviation Monitoring</h1>
                <p class="main-subtitle">PT KALIMANTAN PRIMA PERSADA</p>
            </div>
            <div class="stats-grid">
                <div class="stat-card">
                    <span class="stat-value">3</span>
                    <span class="stat-label">Stages</span>
                </div>
                <div class="stat-card">
                    <span class="stat-value">100%</span>
                    <span class="stat-label">Auto</span>
                </div>
                <div class="stat-card">
                    <span class="stat-value">2026</span>
                    <span class="stat-label">Year</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    selected_file, selected_file_name, selected_source_type = get_active_data_source()

    with st.sidebar:
        st.markdown("---")
        st.markdown("### KPI REFERENCE")
        st.markdown("""
        <div style='background:rgba(255,255,255,0.08);padding:1.3rem;border-radius:16px;border:1px solid rgba(255,255,255,0.1)'>
            <div style='margin:0.9rem 0'>
                <span class="status-badge badge-normal"> Normal</span>
                <p style='font-size:0.82rem;margin:0.4rem 0 0;color:#9ca3af'>|ΔV%| ≤ 2%</p>
            </div>
            <div style='margin:0.9rem 0'>
                <span class="status-badge badge-caution"> Caution</span>
                <p style='font-size:0.82rem;margin:0.4rem 0 0;color:#9ca3af'>2% < |ΔV%| ≤ 3%</p>
            </div>
            <div style='margin:0.9rem 0'>
                <span class="status-badge badge-critical"> Critical</span>
                <p style='font-size:0.82rem;margin:0.4rem 0 0;color:#9ca3af'>|ΔV%| > 3%</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("""
        <div style='text-align:center;color:#9ca3af;margin-top:2rem'>
            <p style='color:#ffffff;font-weight:700;font-size:0.95rem'>Mining Volume Deviation Monitoring</p>
            <p style='font-size:0.85rem'>by : Michael Aragorn Purba</p>
            <p style='font-size:0.85rem'>2026</p>
        </div>
        """, unsafe_allow_html=True)

    if selected_file is None:
        st.info("Silakan Data Training atau upload file Excel dari sidebar.")
        st.markdown("""
        <div style='margin-top:2rem;padding:2rem;background:rgba(255,255,255,0.05);border-radius:20px;border:1px solid rgba(255,255,255,0.1)'>
            <h3 style='color:#00ff88;margin-bottom:1rem'>Required Format:</h3>
            <p style='color:#d1d5db'>• Sheet 1: <strong>"OB (Overburden) Monthly"</strong></p>
            <p style='color:#d1d5db'>• Sheet 2: <strong>"Coal Hauling &amp; Coal Mining"</strong></p>
            <p style='color:#d1d5db'>• Atau pilih file dari folder <strong>demo_data/</strong></p>
        </div>
        """, unsafe_allow_html=True)
        return

    with st.spinner("Processing data..."):
        df_ob, df_ch_cm = load_data(selected_file)

    if df_ob is None or df_ch_cm is None:
        st.error("Failed to load data")
        return

    source_badge_color = "#3b82f6" if selected_source_type == "upload" else "#22c55e"
    source_badge_label = "Uploaded File" if selected_source_type == "upload" else "Data Training"

    st.markdown(f"""
    <div style='margin-bottom:1rem;padding:0.8rem 1rem;background:rgba(255,255,255,0.05);
                border:1px solid rgba(255,255,255,0.1);border-radius:12px;display:flex;
                align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap;'>
        <div>
            <div style='font-size:0.75rem;color:#9ca3af;'>Current Data Source</div>
            <div style='font-size:1rem;font-weight:700;color:#ffffff;'>{selected_file_name}</div>
        </div>
        <div style='padding:0.35rem 0.8rem;border-radius:999px;background:{source_badge_color}20;
                    color:{source_badge_color};border:1px solid {source_badge_color}55;
                    font-size:0.75rem;font-weight:700;letter-spacing:0.05em;'>
            {source_badge_label}
        </div>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs([
        " Overview",
        " OB Analysis",
        " CPP33 & Port",
        " Reports"
    ])

    # ════════════════════════════════════════════════════════════════════════════
    # TAB 1: OVERVIEW
    # ════════════════════════════════════════════════════════════════════════════
    with tab1:
        matrix = create_matrix(df_ob, df_ch_cm)

        # ══════════════════════════════════════════════════════════
        # PERFORMANCE CARDS (4 kolom) — Redesigned
        # ══════════════════════════════════════════════════════════
        st.markdown("""
        <style>
        .perf-card {
            background: linear-gradient(145deg, rgba(22,28,45,0.9), rgba(15,20,35,0.95));
            border: 1px solid rgba(148,163,184,0.08);
            border-radius: 14px;
            padding: 20px 20px 16px;
            position: relative;
            overflow: hidden;
            min-height: 155px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .perf-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 24px rgba(0,0,0,0.3);
        }
        .perf-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 3px;
            border-radius: 14px 14px 0 0;
        }
        .perf-card.normal::before  { background: linear-gradient(90deg, #22c55e, #4ade80); }
        .perf-card.caution::before { background: linear-gradient(90deg, #f59e0b, #fbbf24); }
        .perf-card.critical::before { background: linear-gradient(90deg, #ef4444, #f87171); }
        .perf-card.info::before    { background: linear-gradient(90deg, #3b82f6, #60a5fa); }

        .card-icon {
            font-size: 1.4rem;
            margin-bottom: 10px;
            display: block;
            opacity: 0.9;
        }
        .card-label {
            font-size: 0.68rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: #94a3b8;
            margin-bottom: 4px;
        }
        .card-value {
            font-size: 2.1rem;
            font-weight: 800;
            color: #f1f5f9;
            line-height: 1.1;
            margin-bottom: 2px;
        }
        .card-subtitle {
            font-size: 0.7rem;
            color: #64748b;
            margin-bottom: 10px;
        }
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.7rem;
            font-weight: 600;
            width: fit-content;
        }
        .badge-normal  { background: rgba(34,197,94,0.1); color: #4ade80; border: 1px solid rgba(34,197,94,0.2); }
        .badge-caution { background: rgba(245,158,11,0.1); color: #fbbf24; border: 1px solid rgba(245,158,11,0.2); }
        .badge-critical { background: rgba(239,68,68,0.1); color: #f87171; border: 1px solid rgba(239,68,68,0.2); }
        .badge-info    { background: rgba(59,130,246,0.1); color: #60a5fa; border: 1px solid rgba(59,130,246,0.2); }
        .perf-card.info .card-value { color: #93bbfc; }
        </style>
        """, unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4, gap="large")

        with col1:
            ob = matrix.iloc[0]
            status = 'normal' if ob['Avg'] <= 2 else 'caution' if ob['Avg'] <= 3 else 'critical'
            status_name = 'Normal' if status == 'normal' else 'Caution' if status == 'caution' else 'Critical'
            st.markdown(f"""
            <div class="perf-card {status}">
                <div>
                    <span class="card-icon"></span>
                    <div class="card-label">Overburden (OB)</div>
                    <div class="card-value">{format_number(ob['Avg'], 2)}</div>
                    <div class="card-subtitle">Avg Deviation %</div>
                </div>
                <span class="status-badge badge-{status}">● {status_name}</span>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            ch = matrix.iloc[1]
            status = 'normal' if ch['Avg'] <= 2 else 'caution' if ch['Avg'] <= 3 else 'critical'
            status_name = 'Normal' if status == 'normal' else 'Caution' if status == 'caution' else 'Critical'
            st.markdown(f"""
            <div class="perf-card {status}">
                <div>
                    <span class="card-icon"></span>
                    <div class="card-label">Coal Hauling (CH)</div>
                    <div class="card-value">{format_number(ch['Avg'], 2)}</div>
                    <div class="card-subtitle">Avg Deviation %</div>
                </div>
                <span class="status-badge badge-{status}">● {status_name}</span>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            cm = matrix.iloc[2]
            status = 'normal' if cm['Avg'] <= 2 else 'caution' if cm['Avg'] <= 3 else 'critical'
            status_name = 'Normal' if status == 'normal' else 'Caution' if status == 'caution' else 'Critical'
            st.markdown(f"""
            <div class="perf-card {status}">
                <div>
                    <span class="card-icon"></span>
                    <div class="card-label">Coal Mining (CM)</div>
                    <div class="card-value">{format_number(cm['Avg'], 2)}</div>
                    <div class="card-subtitle">Avg Deviation %</div>
                </div>
                <span class="status-badge badge-{status}">● {status_name}</span>
            </div>
            """, unsafe_allow_html=True)

        with col4:
            total_normal = matrix['Normal'].sum()
            total = matrix['Total'].sum()
            perf = (total_normal / total * 100) if total > 0 else 0
            st.markdown(f"""
            <div class="perf-card info">
                <div>
                    <span class="card-icon"></span>
                    <div class="card-label">Overall Performance</div>
                    <div class="card-value">{format_number(perf, 1)}</div>
                    <div class="card-subtitle">{total_normal}/{total} Periods Normal</div>
                </div>
                <span class="status-badge badge-info">● Performance Index</span>
            </div>
            """, unsafe_allow_html=True)
        # ══════════════════════════════════════════════════════════
        # EXECUTIVE SUMMARY (inserted between cards and flow)
        # ══════════════════════════════════════════════════════════
        total_normal_es = int(matrix['Normal'].sum())
        total_caution_es = int(matrix['Caution'].sum())
        total_critical_es = int(matrix['Critical'].sum())
        total_all_es = int(matrix['Total'].sum())
        perf_es = total_normal_es / total_all_es * 100 if total_all_es > 0 else 0

        ob_disp_es = matrix.iloc[0]['Avg'] if len(matrix) > 0 else 0
        ch_disp_es = matrix.iloc[1]['Avg'] if len(matrix) > 1 else 0
        cm_disp_es = matrix.iloc[2]['Avg'] if len(matrix) > 2 else 0

        worst_stage_es = "Coal Hauling (CH)" if ch_disp_es >= cm_disp_es and ch_disp_es >= ob_disp_es else \
                         "Coal Mining (CM)" if cm_disp_es >= ch_disp_es and cm_disp_es >= ob_disp_es else \
                         "Overburden (OB)"
        worst_val_es = max(ch_disp_es, cm_disp_es, ob_disp_es)
        best_stage_es = "Coal Hauling (CH)" if ch_disp_es <= cm_disp_es and ch_disp_es <= ob_disp_es else \
                        "Coal Mining (CM)" if cm_disp_es <= ch_disp_es and cm_disp_es <= ob_disp_es else \
                        "Overburden (OB)"
        best_val_es = min(ch_disp_es, cm_disp_es, ob_disp_es)

        severity_color_es = "#ef4444" if worst_val_es > 3 else "#f59e0b" if worst_val_es > 2 else "#22c55e"
        severity_text_es = "KRITIS" if worst_val_es > 3 else "PERHATIAN" if worst_val_es > 2 else "NORMAL"
        perf_c_es = '#4ade80' if perf_es >= 70 else '#fbbf24' if perf_es >= 50 else '#f87171'

        st.markdown(f"""
        <div style="background:linear-gradient(135deg,rgba(30,41,59,0.8),rgba(15,23,42,0.95));
            border:1px solid rgba(148,163,184,0.1);border-radius:14px;
            padding:20px 24px;margin:20px 0 8px 0;">
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;">
                <span style="background:{severity_color_es};width:10px;height:10px;border-radius:50%;display:inline-block;"></span>
                <span style="font-size:0.75rem;font-weight:700;text-transform:uppercase;
                    letter-spacing:0.1em;color:{severity_color_es};">Executive Summary — Status: {severity_text_es}</span>
            </div>
            <div style="font-size:0.88rem;color:#e2e8f0;line-height:1.7;">
                Dari <b style="color:#f1f5f9;">{total_all_es} periode</b> yang dianalisis,
                <b style="color:#4ade80;">{total_normal_es}</b> normal,
                <b style="color:#fbbf24;">{total_caution_es}</b> caution, dan
                <b style="color:#f87171;">{total_critical_es}</b> critical.
                <b style="color:{severity_color_es};">{worst_stage_es}</b> menjadi stage dengan deviasi tertinggi
                (<b style="color:{severity_color_es};">{worst_val_es:.2f}%</b>),
                sementara <b style="color:#4ade80;">{best_stage_es}</b> paling stabil
                (<b style="color:#4ade80;">{best_val_es:.2f}%</b>).
                Overall performance berada di <b style="color:{perf_c_es};">{perf_es:.1f}%</b>
                {"— perlu evaluasi menyeluruh." if perf_es < 50 else "— perlu peningkatan." if perf_es < 70 else "— dalam kondisi baik."}
            </div>
        </div>
        """, unsafe_allow_html=True)

        # ══════════════════════════════════════════════════════════
        # PRODUCTION FLOW SUMMARY — Animated Pipeline
        # ══════════════════════════════════════════════════════════
        st.markdown("""
        <div class="section-header">
            <div class="section-icon"></div>
            <h3 class="section-title">Production Flow Summary</h3>
        </div>
        """, unsafe_allow_html=True)

        flow = create_flow(df_ch_cm)

        if flow is not None and len(flow) > 0:
            latest_flow = flow.sort_values("Date").iloc[-1]

            disp_date = pd.to_datetime(latest_flow["Date"]).strftime("%d %b %Y")
            disp_cm = latest_flow["CM TWB"]
            disp_ch = latest_flow["CH TWB"]
            disp_sales = latest_flow["Sales"]

            disp_delta_cpp = latest_flow["Delta CPP Stock"]
            disp_delta_port = latest_flow["Delta Port Stock"]

            disp_dev_cpp = latest_flow["Deviation CPP"]
            disp_dev_port = latest_flow["Deviation Port"]

            disp_ch_ratio = latest_flow["CH Flow Ratio (%)"]
            disp_sales_ratio = latest_flow["Sales Flow Ratio (%)"]

            total_ref = max(abs(disp_cm) + abs(disp_ch) + abs(disp_sales), 1)
            overall_dev_score = max(
                0,
                100 - ((abs(disp_dev_cpp) + abs(disp_dev_port)) / total_ref * 100)
            )

            _r = 30
            _circ = 2 * 3.14159 * _r
            _score_pct = min(max(overall_dev_score, 0), 100)
            _offset = _circ * (1 - _score_pct / 100)
            _ring_color = "#22c55e" if _score_pct >= 98 else "#f59e0b" if _score_pct >= 95 else "#ef4444"

            import streamlit.components.v1 as components

            pipeline_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset="utf-8">
            <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
            * {{ margin: 0; padding: 0; box-sizing: border-box; }}
            body {{ background: transparent; font-family: 'Inter', 'Segoe UI', system-ui, sans-serif; overflow: hidden; }}
            .pipeline-wrap {{
                background: linear-gradient(135deg, rgba(30,41,59,0.85), rgba(15,23,42,0.95));
                border: 1px solid rgba(255,255,255,0.08);
                border-radius: 16px;
                padding: 32px 28px 24px;
                width: 100%;
            }}
            .pipeline-title {{
                text-align: center; font-size: 1.1rem;
                font-weight: 700; color: #e2e8f0;
                margin-bottom: 28px; letter-spacing: 0.02em;
            }}
            .pipe-chain {{
                display: flex; align-items: center;
                justify-content: center; gap: 0;
                width: 100%; padding: 0 8px;
            }}
            .pipe-node {{
                flex: 1; max-width: 220px; min-width: 160px;
                padding: 20px 18px 16px;
                border-radius: 14px; text-align: center;
                border: 1.5px solid rgba(255,255,255,0.1);
                transition: transform 0.25s ease, box-shadow 0.25s ease;
                cursor: default;
            }}
            .pipe-node:hover {{
                transform: translateY(-4px) scale(1.03);
                box-shadow: 0 8px 32px rgba(0,0,0,0.35); z-index: 2;
            }}
            .pipe-node.node-cm   {{ background: linear-gradient(145deg, #064e3b, #065f46); border-color: rgba(34,197,94,0.25); }}
            .pipe-node.node-ch   {{ background: linear-gradient(145deg, #1e3a5f, #1d4ed8); border-color: rgba(59,130,246,0.25); }}
            .pipe-node.node-sales {{ background: linear-gradient(145deg, #4c1d95, #6d28d9); border-color: rgba(167,139,250,0.25); }}

            .pipe-node-label {{ font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.1em; color: #94a3b8; margin-bottom: 2px; }}
            .pipe-node-title {{ font-size: 1.05rem; font-weight: 700; color: #f1f5f9; }}
            .pipe-node-value {{ font-size: 1.5rem; font-weight: 800; margin-top: 8px; }}
            .pipe-node.node-cm .pipe-node-value    {{ color: #4ade80; }}
            .pipe-node.node-ch .pipe-node-value    {{ color: #60a5fa; }}
            .pipe-node.node-sales .pipe-node-value {{ color: #c4b5fd; }}

            .pipe-connector {{
                display: flex; flex-direction: column; align-items: center;
                justify-content: center; flex: 0 0 130px; padding: 0 6px;
                position: relative;
            }}
            .pipe-arrow-wrap {{ position: relative; width: 100%; height: 16px; display: flex; align-items: center; }}
            .pipe-arrow-line {{
                width: 100%; height: 3px;
                background: linear-gradient(90deg, transparent 0%, #475569 12%, #475569 88%, transparent 100%);
                position: relative; overflow: hidden; border-radius: 2px;
            }}
            .pipe-arrow-line::after {{
                content: ''; position: absolute; top: -2px; left: -20px;
                width: 20px; height: 7px; border-radius: 4px;
                background: linear-gradient(90deg, transparent, #22c55e, transparent);
                animation: flowPulse 2.2s ease-in-out infinite;
            }}
            .pipe-connector.loss-connector .pipe-arrow-line::after {{
                background: linear-gradient(90deg, transparent, #3b82f6, transparent);
                animation-delay: 1.1s;
            }}
            @keyframes flowPulse {{
                0%   {{ left: -20px; opacity: 0; }}
                20%  {{ opacity: 1; }}
                80%  {{ opacity: 1; }}
                100% {{ left: calc(100% + 20px); opacity: 0; }}
            }}
            .pipe-arrow-tip {{
                width: 0; height: 0;
                border-top: 7px solid transparent; border-bottom: 7px solid transparent;
                border-left: 11px solid #475569;
                flex-shrink: 0; margin-left: -1px;
            }}
            .pipe-conn-stats {{
                display: flex; flex-direction: column; align-items: center;
                gap: 3px; margin-top: 10px;
            }}
            .pipe-loss-badge {{
                font-size: 0.72rem; padding: 3px 10px;
                border-radius: 10px; font-weight: 600;
                white-space: nowrap;
            }}
            .pipe-loss-badge.loss-val {{ background: rgba(239,68,68,0.15); color: #f87171; }}
            .pipe-loss-badge.eff-val  {{ background: rgba(34,197,94,0.12); color: #4ade80; }}

            .pipe-footer {{
                display: flex; align-items: center; justify-content: center;
                gap: 20px; margin-top: 26px; padding-top: 18px;
                border-top: 1px solid rgba(255,255,255,0.06);
            }}
            .pipe-eff-ring      {{ width: 76px; height: 76px; position: relative; flex-shrink: 0; }}
            .pipe-eff-ring svg  {{ transform: rotate(-90deg); }}
            .ring-bg            {{ fill: none; stroke: rgba(255,255,255,0.08); stroke-width: 6; }}
            .ring-fg            {{ fill: none; stroke-width: 6; stroke-linecap: round; transition: stroke-dashoffset 1s ease; }}
            .ring-label         {{ position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; font-size: 0.9rem; font-weight: 800; color: #f1f5f9; }}

            .pipe-footer-text   {{ display: flex; flex-direction: column; gap: 5px; }}
            .pipe-footer-title  {{ font-size: 0.85rem; color: #94a3b8; font-weight: 600; }}
            .pipe-footer-detail {{ font-size: 0.78rem; color: #64748b; }}
            .pipe-footer-detail .hl-green {{ color: #4ade80; font-weight: 700; }}
            .pipe-footer-detail .hl-red   {{ color: #f87171; font-weight: 700; }}

            @media (max-width: 700px) {{
                .pipe-chain {{ flex-direction: column; gap: 8px; }}
                .pipe-connector {{ transform: rotate(90deg); flex: 0 0 60px; }}
            }}
            </style>
            </head>
            <body>
            <div class="pipeline-wrap">
                <div class="pipeline-title">Production Flow Summary — {disp_date}</div>
                <div class="pipe-chain">
                    <div class="pipe-node node-cm">
                        <div class="pipe-node-label">From PIT To CPP</div>
                        <div class="pipe-node-title">Coal Mining</div>
                        <div class="pipe-node-value">{format_large(disp_cm)}</div>
                    </div>
                    <div class="pipe-connector">
                        <div class="pipe-arrow-wrap">
                            <div class="pipe-arrow-line"></div>
                            <div class="pipe-arrow-tip"></div>
                        </div>
                        <div class="pipe-conn-stats">
                            <span class="pipe-loss-badge loss-val">Δ stok CPP33 {format_large(disp_delta_cpp)}</span>
                            <span class="pipe-loss-badge eff-val">Dev {format_large(disp_dev_cpp)}</span>
                        </div>
                    </div>
                    <div class="pipe-node node-ch">
                        <div class="pipe-node-label">From CPP To Port</div>
                        <div class="pipe-node-title">Coal Hauling</div>
                        <div class="pipe-node-value">{format_large(disp_ch)}</div>
                    </div>
                    <div class="pipe-connector loss-connector">
                        <div class="pipe-arrow-wrap">
                            <div class="pipe-arrow-line"></div>
                            <div class="pipe-arrow-tip"></div>
                        </div>
                        <div class="pipe-conn-stats">
                            <span class="pipe-loss-badge loss-val">Δ stok Port {format_large(disp_delta_port)}</span>
                            <span class="pipe-loss-badge eff-val">Dev {format_large(disp_dev_port)}</span>
                        </div>
                    </div>
                    <div class="pipe-node node-sales">
                        <div class="pipe-node-label">From Port To Customer</div>
                        <div class="pipe-node-title">Sales</div>
                        <div class="pipe-node-value">{format_large(disp_sales)}</div>
                    </div>
                </div>

                <div class="pipe-footer">
                    <div class="pipe-eff-ring">
                        <svg width="76" height="76" viewBox="0 0 76 76">
                            <circle class="ring-bg" cx="38" cy="38" r="{_r}"/>
                            <circle class="ring-fg" cx="38" cy="38" r="{_r}"
                                stroke="{_ring_color}"
                                stroke-dasharray="{_circ:.1f}"
                                stroke-dashoffset="{_offset:.1f}"/>
                        </svg>
                        <div class="ring-label">{format_number(_score_pct, 1)}%</div>
                    </div>
                    <div class="pipe-footer-text">
                        <div class="pipe-footer-title">Reconciliation Deviation Overview</div>
                        <div class="pipe-footer-detail">
                            CPP Dev <span class="hl-red">{format_large(disp_dev_cpp)}</span>
                            · Port Dev <span class="hl-red">{format_large(disp_dev_port)}</span>
                        </div>
                        <div class="pipe-footer-detail">
                            CH Ratio <span class="hl-green">{format_number(disp_ch_ratio, 1)}%</span>
                            · Sales Ratio <span class="hl-green">{format_number(disp_sales_ratio, 1)}%</span>
                        </div>
                    </div>
                </div>
            </div>
            </body>
            </html>
            """

            components.html(pipeline_html, height=340, scrolling=False)

            # ══════════════════════════════════════════════════════════
            # STATUS DISTRIBUTION BY STAGE — Fluid Modern Design
            # ══════════════════════════════════════════════════════════
            st.markdown("""
            <div class="section-header">
                <div class="section-icon"></div>
                <h3 class="section-title">Status Distribution by Stage</h3>
            </div>
            """, unsafe_allow_html=True)


            def make_status_bar(status_series, title, emoji):
                order = ['Normal', 'Caution', 'Critical']
                color_map = {'Normal': '#22C55E', 'Caution': '#F59E0B', 'Critical': '#EF4444'}
                bg_map = {'Normal': 'rgba(34,197,94,0.12)', 'Caution': 'rgba(245,158,11,0.12)', 'Critical': 'rgba(239,68,68,0.12)'}
                counts = status_series.value_counts()
                total = counts.sum()
                if total == 0:
                    return

                segments = []
                for s in order:
                    c = int(counts.get(s, 0))
                    pct = round(c / total * 100, 1) if total > 0 else 0
                    segments.append({'status': s, 'count': c, 'pct': pct, 'color': color_map[s], 'bg': bg_map[s]})

                bar_parts = ""
                for seg in segments:
                    if seg['count'] > 0:
                        bar_parts += (
                            f"<div style='flex:{seg['pct']};background:{seg['color']};"
                            f"height:100%;display:flex;align-items:center;justify-content:center;"
                            f"color:white;font-weight:700;font-size:0.78rem;min-width:32px;"
                            f"transition:flex 0.5s ease;'>"
                            f"{seg['pct']:.0f}%</div>"
                        )

                chips = ""
                for seg in segments:
                    if seg['count'] > 0:
                        chips += (
                            f"<span style='display:inline-flex;align-items:center;gap:5px;"
                            f"background:{seg['bg']};padding:4px 10px;border-radius:12px;"
                            f"font-size:0.75rem;color:#e5e7eb;margin-right:6px;font-weight:500;'>"
                            f"<span style='width:8px;height:8px;border-radius:50%;"
                            f"background:{seg['color']};display:inline-block;'></span>"
                            f"{seg['status']} <b>{seg['count']}</b> ({seg['pct']}%)"
                            f"</span>"
                        )

                html_block = f"""
                <div style="background:linear-gradient(135deg,rgba(20,20,40,0.6),rgba(15,23,42,0.8));
                    border:1px solid rgba(148,163,184,0.1);border-radius:14px;
                    padding:14px 16px;margin-bottom:8px;">
                    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;">
                        <span style="font-size:0.85rem;font-weight:700;color:#e5e7eb;letter-spacing:0.01em;">
                            {emoji} {title}
                        </span>
                        <span style="font-size:0.72rem;color:#64748b;font-weight:500;">{total} data</span>
                    </div>
                    <div style="display:flex;height:28px;border-radius:8px;overflow:hidden;
                        gap:2px;margin-bottom:10px;">
                        {bar_parts}
                    </div>
                    <div style="display:flex;flex-wrap:wrap;gap:4px;">
                        {chips}
                    </div>
                </div>
                """
                st.markdown(html_block, unsafe_allow_html=True)


            col_b1, col_b2, col_b3 = st.columns(3)

            with col_b1:
                make_status_bar(df_ob['Status'], 'Overburden', '🟢')
            with col_b2:
                df_ch = df_ch_cm[df_ch_cm['CH_WB'].notna()]
                if len(df_ch) > 0:
                    make_status_bar(df_ch['Status_CH'], 'Coal Handling', '🔵')
            with col_b3:
                df_cm = df_ch_cm[df_ch_cm['CM_WB'].notna()]
                if len(df_cm) > 0:
                    make_status_bar(df_cm['Status_CM'], 'Coal Mining', '🟡')

            # ══════════════════════════════════════════════════════════
            # MATERIAL FLOW ANALYSIS SECTION
            # ══════════════════════════════════════════════════════════
            st.markdown("""
            <div class="section-header">
                <div class="section-icon"></div>
                <h3 class="section-title">Material Flow Analysis</h3>
            </div>
            """, unsafe_allow_html=True)


            if len(flow) > 0:


                # ══════════════════════════════════════════════════════════
                # PRODUCTION VOLUME + EFFICIENCY (Dual Y-Axis)
                # ══════════════════════════════════════════════════════════
                st.markdown("### Production Volume & Efficiency")


                from plotly.subplots import make_subplots


                flow['CM TWB Plot'] = flow['CM TWB'].replace(0, None)
                flow['CH TWB Plot'] = flow['CH TWB'].replace(0, None)
                flow['Sales Plot'] = flow['Sales'].replace(0, None)
                flow['Eff Plot'] = flow['CH Flow Ratio (%)'].replace(0, None)


                disp_eff_val = flow['CH Flow Ratio (%)'].mean()
                peak_cm = flow['CM TWB'].max()
                below_target = (flow['CH Flow Ratio (%)'] < 100).sum()
                total_p = len(flow)


                vc1, vc2, vc3, vc4 = st.columns(4)

                with vc1:
                    st.markdown(f"""
                    <div style="background:rgba(15,15,30,0.8);border-radius:10px;padding:0.8rem;
                    border:1px solid rgba(148,163,184,0.1);text-align:center">
                        <div style="color:#9ca3af;font-size:0.7rem;margin-bottom:0.2rem">Latest CM TWB</div>
                        <div style="color:#3b82f6;font-size:1.1rem;font-weight:700">{disp_cm:,.0f} t</div>
                    </div>""", unsafe_allow_html=True)

                with vc2:
                    st.markdown(f"""
                    <div style="background:rgba(15,15,30,0.8);border-radius:10px;padding:0.8rem;
                    border:1px solid rgba(148,163,184,0.1);text-align:center">
                        <div style="color:#9ca3af;font-size:0.7rem;margin-bottom:0.2rem">Latest CH TWB</div>
                        <div style="color:#22c55e;font-size:1.1rem;font-weight:700">{disp_ch:,.0f} t</div>
                    </div>""", unsafe_allow_html=True)

                with vc3:
                    ratio_color = '#22c55e' if disp_ch_ratio >= 100 else '#f59e0b' if disp_ch_ratio >= 90 else '#ef4444'
                    st.markdown(f"""
                    <div style="background:rgba(15,15,30,0.8);border-radius:10px;padding:0.8rem;
                    border:1px solid rgba(148,163,184,0.1);text-align:center">
                        <div style="color:#9ca3af;font-size:0.7rem;margin-bottom:0.2rem">CH Flow Ratio</div>
                        <div style="color:{ratio_color};font-size:1.1rem;font-weight:700">{disp_ch_ratio:.1f}%</div>
                    </div>""", unsafe_allow_html=True)

                with vc4:
                    dev_total = abs(disp_dev_cpp) + abs(disp_dev_port)
                    dev_color = '#22c55e' if dev_total < 2000 else '#f59e0b' if dev_total < 10000 else '#ef4444'
                    st.markdown(f"""
                    <div style="background:rgba(15,15,30,0.8);border-radius:10px;padding:0.8rem;
                    border:1px solid rgba(148,163,184,0.1);text-align:center">
                        <div style="color:#9ca3af;font-size:0.7rem;margin-bottom:0.2rem">Total Deviation</div>
                        <div style="color:{dev_color};font-size:1.1rem;font-weight:700">{dev_total:,.0f} t</div>
                    </div>""", unsafe_allow_html=True)


                fig_combined = make_subplots(specs=[[{"secondary_y": True}]])


                cm_max_idx = flow['CM TWB'].idxmax()
                cm_min_idx = flow['CM TWB'].idxmin()


                fig_combined.add_trace(go.Scatter(
                    x=flow['Date'], y=flow['CM TWB Plot'],
                    mode='lines', name='CM TWB',
                    line=dict(color='#3B82F6', width=2),
                    fill='tozeroy',
                    fillcolor='rgba(59,130,246,0.12)',
                    connectgaps=True,
                    hovertemplate='<b>CM TWB</b><br>%{y:,.0f} ton<extra></extra>'
                ), secondary_y=False)


                fig_combined.add_trace(go.Scatter(
                    x=flow['Date'], y=flow['CH TWB Plot'],
                    mode='lines', name='CH TWB',
                    line=dict(color='#22C55E', width=2),
                    fill='tozeroy',
                    fillcolor='rgba(34,197,94,0.10)',
                    connectgaps=True,
                    hovertemplate='<b>CH TWB</b><br>%{y:,.0f} ton<extra></extra>'
                ), secondary_y=False)


                fig_combined.add_trace(go.Scatter(
                    x=flow['Date'], y=flow['Sales Plot'],
                    mode='lines', name='Sales',
                    line=dict(color='#A78BFA', width=2),
                    fill='tozeroy',
                    fillcolor='rgba(167,139,250,0.08)',
                    connectgaps=True,
                    hovertemplate='<b>Sales</b><br>%{y:,.0f} ton<extra></extra>'
                ), secondary_y=False)


                eff_colors = []
                for e in flow['CH Flow Ratio (%)'].values:
                    if e >= 100:
                        eff_colors.append('#22C55E')
                    elif e >= 90:
                        eff_colors.append('#F59E0B')
                    else:
                        eff_colors.append('#EF4444')


                fig_combined.add_trace(go.Scatter(
                    x=flow['Date'], y=flow['Eff Plot'],
                    mode='lines+markers', name='Efficiency (%)',
                    line=dict(color='#F59E0B', width=2.5, dash='dot'),
                    marker=dict(size=9, color=eff_colors, symbol='diamond',
                                line=dict(width=1.5, color='rgba(255,255,255,0.7)')),
                    connectgaps=True,
                    hovertemplate='<b>Efficiency</b><br>%{y:.1f}%<extra></extra>'
                ), secondary_y=True)


                fig_combined.add_hline(
                    y=disp_cm, line_dash='dot',
                    line_color='rgba(59,130,246,0.4)', line_width=1,
                    annotation_text=f'Avg CM {disp_cm:,.0f}',
                    annotation_position='top left',
                    annotation_font=dict(size=9, color='#60A5FA'))

                fig_combined.add_hline(
                    y=100, line_dash='dash',
                    line_color='rgba(245,158,11,0.4)', line_width=1,
                    secondary_y=True,
                    annotation_text='Target 100%',
                    annotation_position='top right',
                    annotation_font=dict(size=9, color='#F59E0B'))

                fig_combined.add_hline(
                    y=disp_eff_val, line_dash='dot',
                    line_color='rgba(245,158,11,0.25)', line_width=1,
                    secondary_y=True,
                    annotation_text=f'Avg Eff {disp_eff_val:.1f}%',
                    annotation_position='bottom right',
                    annotation_font=dict(size=8, color='#FBBF24'))


                fig_combined.add_annotation(
                    x=flow.loc[cm_max_idx, 'Date'],
                    y=flow.loc[cm_max_idx, 'CM TWB'],
                    text=f"▲ Peak {flow.loc[cm_max_idx, 'CM TWB']:,.0f} t",
                    showarrow=True, arrowhead=2,
                    arrowwidth=1.5, arrowcolor='#60A5FA',
                    ax=0, ay=-30,
                    font=dict(size=10, color='#FFF'),
                    bgcolor='rgba(59,130,246,0.75)', borderpad=4)

                fig_combined.add_annotation(
                    x=flow.loc[cm_min_idx, 'Date'],
                    y=flow.loc[cm_min_idx, 'CM TWB'],
                    text=f"▼ Low {flow.loc[cm_min_idx, 'CM TWB']:,.0f} t",
                    showarrow=True, arrowhead=2,
                    arrowwidth=1.5, arrowcolor='#EF4444',
                    ax=0, ay=30,
                    font=dict(size=10, color='#FFF'),
                    bgcolor='rgba(239,68,68,0.75)', borderpad=4)

                eff_max_idx = flow['CH Flow Ratio (%)'].idxmax()
                fig_combined.add_annotation(
                    x=flow.loc[eff_max_idx, 'Date'],
                    y=flow.loc[eff_max_idx, 'CH Flow Ratio (%)'],
                    text=f"⚡ {flow.loc[eff_max_idx, 'CH Flow Ratio (%)']:.1f}%",
                    showarrow=True, arrowhead=2,
                    arrowwidth=1, arrowcolor='#F59E0B',
                    ax=20, ay=-25, yref='y2',
                    font=dict(size=9, color='#FFF'),
                    bgcolor='rgba(245,158,11,0.75)', borderpad=3)


                fig_combined.update_layout(
                    height=500,
                    font=dict(family='Inter, sans-serif', color='#CBD5E1', size=12),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(
                        gridcolor='rgba(148,163,184,0.08)',
                        title='', tickformat='%d %b',
                        tickfont=dict(size=10, color='#94A3B8'),
                        showline=True,
                        linecolor='rgba(148,163,184,0.2)',
                        rangeselector=dict(
                            buttons=list([
                                dict(count=1, label="1M", step="month", stepmode="backward"),
                                dict(count=3, label="3M", step="month", stepmode="backward"),
                                dict(step="all", label="ALL"),
                            ]),
                            bgcolor='rgba(30,30,50,0.8)',
                            activecolor='rgba(59,130,246,0.4)',
                            font=dict(size=10, color='#CBD5E1'),
                        ),
                        rangeslider=dict(visible=True, thickness=0.06),
                        type="date",
                    ),
                    showlegend=True,
                    legend=dict(
                        orientation='h', yanchor='bottom', y=1.08,
                        xanchor='center', x=0.5,
                        font=dict(size=11, color='#E2E8F0'),
                        bgcolor='rgba(0,0,0,0)',
                    ),
                    hovermode='x unified',
                    margin=dict(l=65, r=65, t=70, b=10),
                )

                fig_combined.update_yaxes(
                    title=dict(text='Volume (ton)', font=dict(color='#94A3B8', size=12)),
                    tickformat=',.0f',
                    gridcolor='rgba(148,163,184,0.08)',
                    tickfont=dict(size=10, color='#94A3B8'),
                    showline=False, zeroline=False,
                    secondary_y=False,
                )

                fig_combined.update_yaxes(
                    title=dict(text='Flow Ratio (%)', font=dict(color='#F59E0B', size=12)),
                    tickformat='.0f',
                    ticksuffix='%',
                    gridcolor='rgba(0,0,0,0)',
                    tickfont=dict(size=10, color='#F59E0B'),
                    showline=False,
                    zeroline=False,
                    range=[0, max(130, flow['CH Flow Ratio (%)'].max() + 15)],
                    secondary_y=True,
                )
                st.plotly_chart(fig_combined, use_container_width=True, key='production_vol_eff')


                # ══════════════════════════════════════════════════════════
                # MATERIAL LOSS ANALYSIS — Enhanced Stacked + Cumulative
                # ══════════════════════════════════════════════════════════
                
            # ══════════════════════════════════════════════════════════
            # RECONCILIATION DEVIATION ANALYSIS
            # ══════════════════════════════════════════════════════════
            st.markdown("### Reconciliation Deviation Analysis")

            flow["Net Deviation"] = flow["Deviation CPP"] + flow["Deviation Port"]
            flow["Cum Net Deviation"] = flow["Net Deviation"].cumsum()

            avg_cm_ref = flow["CM TWB"].mean() if len(flow) > 0 else 0
            avg_net = flow["Net Deviation"].mean()
            std_loss = flow["Net Deviation"].std() if len(flow) > 1 else 0
            max_loss_val = flow["Net Deviation"].min()
            cum_total = flow["Cum Net Deviation"].iloc[-1]
            critical_count = (flow["Net Deviation"].abs() > avg_cm_ref * 0.02).sum()
            total_periods = len(flow)

            lc1, lc2, lc3, lc4 = st.columns(4)
            with lc1:
                net_color = '#22c55e' if abs(avg_net) <= avg_cm_ref * 0.01 else '#f59e0b' if abs(avg_net) <= avg_cm_ref * 0.02 else '#ef4444'
                st.markdown(f"""
                <div style="background:rgba(15,15,30,0.8);border-radius:10px;padding:0.8rem;
                border:1px solid rgba(148,163,184,0.1);text-align:center">
                    <div style="color:#9ca3af;font-size:0.7rem;margin-bottom:0.2rem">Avg Net Deviation</div>
                    <div style="color:{net_color};font-size:1.1rem;font-weight:700">{avg_net:,.0f} t</div>
                </div>""", unsafe_allow_html=True)
            with lc2:
                st.markdown(f"""
                <div style="background:rgba(15,15,30,0.8);border-radius:10px;padding:0.8rem;
                border:1px solid rgba(148,163,184,0.1);text-align:center">
                    <div style="color:#9ca3af;font-size:0.7rem;margin-bottom:0.2rem">Max Single Deviation</div>
                    <div style="color:#ef4444;font-size:1.1rem;font-weight:700">{max_loss_val:,.0f} t</div>
                </div>""", unsafe_allow_html=True)
            with lc3:
                cum_color = '#22c55e' if abs(cum_total) <= avg_cm_ref * 0.01 else '#f59e0b' if abs(cum_total) <= avg_cm_ref * 0.02 else '#ef4444'
                st.markdown(f"""
                <div style="background:rgba(15,15,30,0.8);border-radius:10px;padding:0.8rem;
                border:1px solid rgba(148,163,184,0.1);text-align:center">
                    <div style="color:#9ca3af;font-size:0.7rem;margin-bottom:0.2rem">Cumulative Deviation</div>
                    <div style="color:{cum_color};font-size:1.1rem;font-weight:700">{cum_total:,.0f} t</div>
                </div>""", unsafe_allow_html=True)
            with lc4:
                crit_pct = critical_count / total_periods * 100 if total_periods > 0 else 0
                crit_color = '#22c55e' if crit_pct < 20 else '#f59e0b' if crit_pct < 40 else '#ef4444'
                st.markdown(f"""
                <div style="background:rgba(15,15,30,0.8);border-radius:10px;padding:0.8rem;
                border:1px solid rgba(148,163,184,0.1);text-align:center">
                    <div style="color:#9ca3af;font-size:0.7rem;margin-bottom:0.2rem">Critical Periods</div>
                    <div style="color:{crit_color};font-size:1.1rem;font-weight:700">{critical_count}/{total_periods} ({crit_pct:.0f}%)</div>
                </div>""", unsafe_allow_html=True)

            fig_loss = make_subplots(specs=[[{"secondary_y": True}]])

            fig_loss.add_trace(go.Bar(
                x=flow["Date"],
                y=flow["Deviation Port"],
                name="Port Deviation",
                marker=dict(
                    color='rgba(239,68,68,0.7)',
                    line=dict(width=0.5, color='rgba(239,68,68,0.9)')
                ),
                hovertemplate='<b>Port Deviation</b><br>%{x|%d %b %Y}<br>%{y:,.0f} ton<extra></extra>',
            ), secondary_y=False)

            fig_loss.add_trace(go.Bar(
                x=flow["Date"],
                y=flow["Deviation CPP"],
                name="CPP33 Deviation",
                marker=dict(
                    color='rgba(234,179,8,0.7)',
                    line=dict(width=0.5, color='rgba(234,179,8,0.9)')
                ),
                hovertemplate='<b>CPP33 Deviation</b><br>%{x|%d %b %Y}<br>%{y:,.0f} ton<extra></extra>',
            ), secondary_y=False)

            net_colors = []
            for v in flow["Net Deviation"].values:
                abs_v = abs(v)
                if abs_v <= abs(avg_net) + 1 * std_loss:
                    net_colors.append('#22C55E')
                elif abs_v <= abs(avg_net) + 2 * std_loss:
                    net_colors.append('#F59E0B')
                else:
                    net_colors.append('#EF4444')

            fig_loss.add_trace(go.Scatter(
                x=flow["Date"],
                y=flow["Net Deviation"],
                mode='lines+markers',
                name='Net Deviation',
                line=dict(color='rgba(255,255,255,0.4)', width=1.5),
                marker=dict(
                    size=8,
                    color=net_colors,
                    line=dict(width=1.5, color='rgba(255,255,255,0.6)')
                ),
                connectgaps=True,
                hovertemplate='<b>Net Deviation</b><br>%{x|%d %b %Y}<br>%{y:,.0f} ton<extra></extra>',
            ), secondary_y=False)

            fig_loss.add_trace(go.Scatter(
                x=flow["Date"],
                y=flow["Cum Net Deviation"],
                mode='lines',
                name='Cumulative Deviation',
                line=dict(color='#06b6d4', width=2.5, dash='dot'),
                fill='tozeroy',
                fillcolor='rgba(6,182,212,0.06)',
                connectgaps=True,
                hovertemplate='<b>Cumulative Deviation</b><br>%{y:,.0f} ton<extra></extra>',
            ), secondary_y=True)

            fig_loss.add_hline(y=0, line_color='rgba(148,163,184,0.5)', line_width=1.5)

            fig_loss.update_layout(
                height=460,
                barmode='relative',
                font=dict(family='Inter, sans-serif', color='#CBD5E1', size=11),
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(
                    gridcolor='rgba(148,163,184,0.08)',
                    title='',
                    tickformat='%d %b',
                    tickfont=dict(size=10, color='#94A3B8'),
                    showline=True,
                    linecolor='rgba(148,163,184,0.2)',
                    rangeslider=dict(visible=True, thickness=0.06),
                    type="date",
                ),
                showlegend=True,
                legend=dict(
                    orientation='h',
                    yanchor='bottom',
                    y=1.08,
                    xanchor='center',
                    x=0.5,
                    font=dict(size=10, color='#E2E8F0'),
                    bgcolor='rgba(0,0,0,0)',
                ),
                hovermode='x unified',
                margin=dict(l=60, r=60, t=70, b=10),
            )

            fig_loss.update_yaxes(
                title=dict(text='Deviation (ton)', font=dict(color='#94A3B8', size=11)),
                tickformat=',.0f',
                gridcolor='rgba(148,163,184,0.08)',
                tickfont=dict(size=10, color='#94A3B8'),
                showline=False,
                zeroline=False,
                secondary_y=False,
            )

            fig_loss.update_yaxes(
                title=dict(text='Cumulative Deviation (ton)', font=dict(color='#06b6d4', size=11)),
                tickformat=',.0f',
                gridcolor='rgba(0,0,0,0)',
                tickfont=dict(size=10, color='#06b6d4'),
                showline=False,
                zeroline=False,
                secondary_y=True,
            )

            st.plotly_chart(fig_loss, use_container_width=True, key='loss_analysis')

    with tab2:
        # ══════════════════════════════════════════════════════════
        # TAB 2: OVERBURDEN ANALYSIS — v6 (Alert di bawah KPI)
        # ══════════════════════════════════════════════════════════
        st.markdown("""
        <div class="section-header">
            <h3 class="section-title">Overburden Analysis</h3>
        </div>
        """, unsafe_allow_html=True)

        # ── FILTER PERIODE ──
        if 'Bulan' in df_ob.columns and len(df_ob) > 0:
            bulan_list = df_ob['Bulan'].unique().tolist()
            bulan_sorted = sort_months_chronologically(bulan_list)

            st.markdown("### Filter Periode")
            if len(bulan_sorted) > 1:
                total = len(bulan_sorted)
                selected_range = st.slider(
                    "Pilih Range Bulan:",
                    min_value=1, max_value=total,
                    value=(1, total),
                    key='ob_month_slider_v2'
                )
                start_idx = selected_range[0] - 1
                end_idx = selected_range[1]
                selected_months = bulan_sorted[start_idx:end_idx]
                df_ob_filtered = df_ob[df_ob['Bulan'].isin(selected_months)].copy()
                month_order = {month: i for i, month in enumerate(bulan_sorted)}
                df_ob_filtered['Sort_Order'] = df_ob_filtered['Bulan'].map(month_order)
                df_ob_filtered = df_ob_filtered.sort_values('Sort_Order')
                if selected_months:
                    st.info(f"**Bulan yang dipilih:** {len(selected_months)} bulan ({selected_months[0]} - {selected_months[-1]})")
                df_ob_filtered = df_ob_filtered.drop(columns=['Sort_Order'])
            else:
                df_ob_filtered = df_ob.copy()
                st.info(f"Data hanya untuk 1 bulan: {bulan_sorted[0]}")
        else:
            df_ob_filtered = df_ob.copy()
            st.warning("Kolom 'Bulan' tidak ditemukan dalam data")

        # ── KPI CARDS (DULU, sebelum alert) ──
        n_periods = len(df_ob_filtered)
        total_tc = df_ob_filtered['TC'].sum()
        total_js = df_ob_filtered['JS'].sum()
        total_dev_val = df_ob_filtered['Dev_Absolut'].sum()
        disp_dev_pct = df_ob_filtered['Dev_Relatif_Pct'].abs().mean()
        dev_color = '#4ade80' if disp_dev_pct <= 2 else '#fbbf24' if disp_dev_pct <= 3 else '#f87171'

        kpi_data = [
            ("Periods", f"{n_periods}", "#60a5fa"),
            ("Total TC", format_large(total_tc), "#60a5fa"),
            ("Total JS", format_large(total_js), "#4ade80"),
            ("Total Deviation", format_large(total_dev_val), "#fbbf24"),
            ("Avg |Dev%|", f"{format_number(disp_dev_pct, 2)}%", dev_color),
        ]

        cols_kpi = st.columns(5, gap="medium")
        for i, (label, value, color) in enumerate(kpi_data):
            with cols_kpi[i]:
                st.markdown(f"""
                <div style="background:linear-gradient(145deg,rgba(22,28,45,0.9),rgba(15,20,35,0.95));
                    border:1px solid rgba(148,163,184,0.08);border-radius:12px;
                    padding:14px 12px;text-align:center;position:relative;overflow:hidden;">
                    <div style="position:absolute;top:0;left:0;right:0;height:2px;
                        background:{color};border-radius:12px 12px 0 0;"></div>
                    <div style="font-size:0.65rem;font-weight:600;text-transform:uppercase;
                        letter-spacing:0.08em;color:#94a3b8;margin-bottom:4px;">{label}</div>
                    <div style="font-size:1.5rem;font-weight:800;color:{color};line-height:1.1;">{value}</div>
                </div>
                """, unsafe_allow_html=True)

        # ── ALERT BANNER (sekarang DI BAWAH KPI, hanya jika ada Critical) ──
        critical_rows = df_ob_filtered[df_ob_filtered['Status'] == 'Critical']
        if len(critical_rows) > 0:
            alert_chips = " ".join([
                f'<span style="background:rgba(239,68,68,0.15);color:#f87171;'
                f'padding:4px 12px;border-radius:8px;font-size:0.78rem;font-weight:600;">'
                f'{row["Bulan"]}: {row["Dev_Relatif_Pct"]:.1f}%</span>'
                for _, row in critical_rows.iterrows()
            ])
            st.markdown(f"""
            <div style="background:linear-gradient(135deg,rgba(239,68,68,0.08),rgba(15,23,42,0.9));
                border:1px solid rgba(239,68,68,0.2);border-radius:12px;
                padding:14px 20px;margin-top:12px;">
                <div style="color:#f87171;font-size:0.72rem;font-weight:700;
                    text-transform:uppercase;letter-spacing:0.1em;margin-bottom:8px;">
                    Critical Alert — {len(critical_rows)} Period(s) Detected</div>
                <div style="display:flex;flex-wrap:wrap;gap:8px;">{alert_chips}</div>
            </div>
            """, unsafe_allow_html=True)

        # ══ GRAFIK VOLUME & DEVIATION ══
        st.markdown("""
        <div class="section-header">
            <h3 class="section-title">Volume & Deviation Analysis</h3>
        </div>
        """, unsafe_allow_html=True)

        col_left, col_right = st.columns(2)

        with col_left:
            fig_comp = go.Figure()
            fig_comp.add_trace(go.Scatter(
                x=df_ob_filtered['Bulan'], y=df_ob_filtered['TC'],
                mode='lines+markers', name='TC (Target)',
                line=dict(color='#3B82F6', width=2.5),
                marker=dict(size=7, line=dict(width=1.5, color='rgba(255,255,255,0.6)')),
                hovertemplate='<b>TC</b><br>%{y:,.0f}<extra></extra>'
            ))
            fig_comp.add_trace(go.Scatter(
                x=df_ob_filtered['Bulan'], y=df_ob_filtered['JS'],
                mode='lines+markers', name='JS (Survey)',
                line=dict(color='#22C55E', width=2.5),
                marker=dict(size=7, line=dict(width=1.5, color='rgba(255,255,255,0.6)')),
                fill='tonexty', fillcolor='rgba(59,130,246,0.10)',
                hovertemplate='<b>JS</b><br>%{y:,.0f}<extra></extra>'
            ))

            disp_tc_val = df_ob_filtered['TC'].mean()
            fig_comp.add_hline(
                y=disp_tc_val, line_dash='dot',
                line_color='rgba(59,130,246,0.35)', line_width=1,
                annotation_text=f'Avg TC {disp_tc_val:,.0f}',
                annotation_position='top left',
                annotation_font=dict(size=9, color='#60A5FA')
            )

            if len(df_ob_filtered) > 1:
                pk_idx = df_ob_filtered['TC'].idxmax()
                lo_idx = df_ob_filtered['TC'].idxmin()
                fig_comp.add_annotation(
                    x=df_ob_filtered.loc[pk_idx, 'Bulan'],
                    y=df_ob_filtered.loc[pk_idx, 'TC'],
                    text=f"Peak {df_ob_filtered.loc[pk_idx, 'TC']:,.0f}",
                    showarrow=True, arrowhead=2, arrowwidth=1.5,
                    arrowcolor='#60A5FA', ax=0, ay=-30,
                    font=dict(size=10, color='#FFF'),
                    bgcolor='rgba(59,130,246,0.75)', borderpad=4
                )
                fig_comp.add_annotation(
                    x=df_ob_filtered.loc[lo_idx, 'Bulan'],
                    y=df_ob_filtered.loc[lo_idx, 'TC'],
                    text=f"Low {df_ob_filtered.loc[lo_idx, 'TC']:,.0f}",
                    showarrow=True, arrowhead=2, arrowwidth=1.5,
                    arrowcolor='#EF4444', ax=0, ay=30,
                    font=dict(size=10, color='#FFF'),
                    bgcolor='rgba(239,68,68,0.75)', borderpad=4
                )

            fig_comp.update_layout(
                height=420,
                font=dict(family='Inter, sans-serif', color='#CBD5E1', size=11),
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(
                    gridcolor='rgba(148,163,184,0.08)', title='',
                    tickfont=dict(size=10, color='#94A3B8'),
                    showline=True, linecolor='rgba(148,163,184,0.2)'
                ),
                yaxis=dict(
                    title=dict(text='Volume', font=dict(color='#94A3B8', size=12)),
                    tickformat=',.0f', gridcolor='rgba(148,163,184,0.08)',
                    tickfont=dict(size=10, color='#94A3B8'),
                    showline=False, zeroline=False
                ),
                showlegend=True,
                legend=dict(
                    orientation='h', yanchor='bottom', y=1.05,
                    xanchor='center', x=0.5,
                    font=dict(size=10, color='#E2E8F0'),
                    bgcolor='rgba(0,0,0,0)'
                ),
                hovermode='x unified',
                margin=dict(l=65, r=20, t=50, b=40)
            )
            st.plotly_chart(fig_comp, use_container_width=True, key='ob_comp')

        with col_right:
            colors_bar = [get_status_color(s) for s in df_ob_filtered['Status']]
            fig_dev = go.Figure()

            fig_dev.add_shape(type='rect', xref='paper', yref='y',
                x0=0, x1=1, y0=-2, y1=2,
                fillcolor='rgba(34,197,94,0.04)', line_width=0, layer='below')
            fig_dev.add_shape(type='rect', xref='paper', yref='y',
                x0=0, x1=1, y0=2, y1=3,
                fillcolor='rgba(245,158,11,0.04)', line_width=0, layer='below')
            fig_dev.add_shape(type='rect', xref='paper', yref='y',
                x0=0, x1=1, y0=-3, y1=-2,
                fillcolor='rgba(245,158,11,0.04)', line_width=0, layer='below')

            fig_dev.add_trace(go.Bar(
                x=df_ob_filtered['Bulan'], y=df_ob_filtered['Dev_Relatif_Pct'],
                marker_color=colors_bar, showlegend=False,
                text=df_ob_filtered['Dev_Relatif_Pct'].apply(lambda x: f"{x:.1f}%"),
                textposition='outside',
                textfont=dict(size=10, color='#e2e8f0'),
                hovertemplate='<b>%{x}</b><br>Deviation: %{y:.2f}%<extra></extra>'
            ))

            fig_dev.add_hline(y=0, line_color='rgba(148,163,184,0.5)', line_width=1.5)
            fig_dev.add_hline(y=2, line_color='#22c55e', line_width=1, line_dash='dash')
            fig_dev.add_hline(y=-2, line_color='#22c55e', line_width=1, line_dash='dash')
            fig_dev.add_hline(y=3, line_color='#ef4444', line_width=1, line_dash='dash')
            fig_dev.add_hline(y=-3, line_color='#ef4444', line_width=1, line_dash='dash')

            fig_dev.add_annotation(xref='paper', yref='y', x=1.02, y=2,
                text='2%', showarrow=False, font=dict(size=8, color='#22c55e'))
            fig_dev.add_annotation(xref='paper', yref='y', x=1.02, y=3,
                text='3%', showarrow=False, font=dict(size=8, color='#ef4444'))

            if len(df_ob_filtered) > 0:
                w_idx = df_ob_filtered['Dev_Relatif_Pct'].abs().idxmax()
                w_val = df_ob_filtered.loc[w_idx, 'Dev_Relatif_Pct']
                fig_dev.add_annotation(
                    x=df_ob_filtered.loc[w_idx, 'Bulan'], y=w_val,
                    text=f"Worst {w_val:.1f}%",
                    showarrow=True, arrowhead=2, arrowwidth=1.5,
                    arrowcolor='#EF4444',
                    ax=35, ay=-25 if w_val > 0 else 25,
                    font=dict(size=10, color='#FFF'),
                    bgcolor='rgba(239,68,68,0.85)', borderpad=4
                )

            fig_dev.update_layout(
                height=420,
                font=dict(family='Inter, sans-serif', color='#CBD5E1', size=11),
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(
                    gridcolor='rgba(148,163,184,0.08)', title='',
                    tickfont=dict(size=10, color='#94A3B8'),
                    showline=True, linecolor='rgba(148,163,184,0.2)'
                ),
                yaxis=dict(
                    title=dict(text='Deviation %', font=dict(color='#94A3B8', size=12)),
                    gridcolor='rgba(148,163,184,0.08)',
                    tickfont=dict(size=10, color='#94A3B8'),
                    showline=False, zeroline=False
                ),
                hovermode='x unified',
                margin=dict(l=60, r=40, t=40, b=40)
            )
            st.plotly_chart(fig_dev, use_container_width=True, key='ob_dev')

        # ══ DEVIATION HEATMAP (langsung tanpa header) ══
        if len(df_ob_filtered) > 0:
            hm_cells = ""
            for _, r in df_ob_filtered.iterrows():
                ad = abs(r['Dev_Relatif_Pct'])
                if ad <= 2:
                    bg, tc = "rgba(34,197,94,0.2)", "#4ade80"
                elif ad <= 3:
                    bg, tc = "rgba(245,158,11,0.2)", "#fbbf24"
                else:
                    bg, tc = "rgba(239,68,68,0.2)", "#f87171"
                hm_cells += f"""<div style="flex:1;background:{bg};border-radius:8px;
                    padding:10px 4px;text-align:center;min-width:50px;">
                    <div style="font-size:0.6rem;color:#94a3b8;margin-bottom:3px;">{r['Bulan'][:3]}</div>
                    <div style="font-size:0.82rem;font-weight:700;color:{tc};">{r['Dev_Relatif_Pct']:.1f}%</div>
                </div>"""
            st.markdown(f"""<div style="display:flex;gap:5px;margin:0 0 20px 0;">{hm_cells}</div>""",
                unsafe_allow_html=True)

        # ══ OB DATA TABLE — styled HTML (replace st.dataframe) ══
        st.markdown("""
        <div class="section-header" style="margin-top:2.5rem;">
            <h3 class="section-title">OB Data Table</h3>
        </div>
        """, unsafe_allow_html=True)

        if not df_ob_filtered.empty:
            ob_cols = ['Bulan', 'TC', 'JS', 'Dev_Absolut', 'Dev_Relatif_Pct', 'Status']
            ob_header = ''.join(
                f'<th style="padding:10px 10px;text-align:center;color:#e2e8f0;font-weight:700;'
                f'font-size:0.68rem;text-transform:uppercase;letter-spacing:0.06em;'
                f'position:sticky;top:0;background:rgba(22,101,52,0.95);z-index:1;">{c}</th>'
                for c in ob_cols
            )
            ob_rows = ""
            for _, row in df_ob_filtered.iterrows():
                status = row['Status']
                if status == 'Critical':
                    row_bg, bl, sc, sbg = 'rgba(239,68,68,0.08)', '3px solid #ef4444', '#f87171', 'rgba(239,68,68,0.15)'
                elif status == 'Caution':
                    row_bg, bl, sc, sbg = 'rgba(245,158,11,0.08)', '3px solid #f59e0b', '#fbbf24', 'rgba(245,158,11,0.15)'
                else:
                    row_bg, bl, sc, sbg = 'rgba(34,197,94,0.04)', '3px solid #22c55e', '#4ade80', 'rgba(34,197,94,0.15)'

                dev_val = row['Dev_Relatif_Pct']
                dev_color = '#f87171' if abs(dev_val) > 3 else '#fbbf24' if abs(dev_val) > 2 else '#4ade80'

                ob_rows += f"""
                <tr style="background:{row_bg};border-left:{bl};border-bottom:1px solid rgba(148,163,184,0.06);">
                    <td style="padding:9px 10px;text-align:center;color:#e2e8f0;font-weight:600;font-size:0.8rem;">{row['Bulan']}</td>
                    <td style="padding:9px 10px;text-align:center;color:#cbd5e1;font-size:0.8rem;">{row['TC']:,.2f}</td>
                    <td style="padding:9px 10px;text-align:center;color:#cbd5e1;font-size:0.8rem;">{row['JS']:,.2f}</td>
                    <td style="padding:9px 10px;text-align:center;color:#cbd5e1;font-size:0.8rem;">{row['Dev_Absolut']:,.2f}</td>
                    <td style="padding:9px 10px;text-align:center;color:{dev_color};font-weight:700;font-size:0.8rem;">{dev_val:+.2f}%</td>
                    <td style="padding:9px 10px;text-align:center;">
                        <span style="background:{sbg};color:{sc};padding:3px 10px;border-radius:20px;font-size:0.7rem;font-weight:700;">{status}</span>
                    </td>
                </tr>"""

            st.markdown(f"""
            <div style="background:linear-gradient(135deg,rgba(20,20,40,0.6),rgba(15,23,42,0.8));
                border:1px solid rgba(148,163,184,0.1);border-radius:10px;overflow:hidden;
                max-height:400px;overflow-y:auto;">
            <table style="width:100%;border-collapse:collapse;">
            <thead><tr style="background:rgba(22,101,52,0.95);">{ob_header}</tr></thead>
            <tbody>{ob_rows}</tbody>
            </table></div>
            """, unsafe_allow_html=True)
        else:
            st.info("Tidak ada data untuk ditampilkan")

    with tab3:
        # ══════════════════════════════════════════════════════════
        # TAB 3: CPP33 & PORT ANALYSIS — v6 (Samakan dengan Tab 2)
        # Perubahan: 1) Status Distribution dihilangkan
        #            2) Insight dihilangkan
        #            3) Alert di bawah KPI
        #            4) Data table diberi warna row sesuai status
        # ══════════════════════════════════════════════════════════
        st.markdown("""
        <div class="section-header">
            <h3 class="section-title">CPP33 & Port Analysis</h3>
        </div>
        """, unsafe_allow_html=True)

        df_valid = df_ch_cm.dropna(subset=['CH_WB', 'CM_WB'])

        if len(df_valid) > 0:
            df_valid_copy = df_valid.copy()
            df_valid_copy['YearMonth'] = df_valid_copy['Date'].dt.strftime('%b-%Y')
            available_months = sorted(df_valid_copy['YearMonth'].unique(),
                                      key=lambda x: pd.to_datetime(x, format='%b-%Y'))

            st.markdown("### Filter Periode")
            if len(available_months) > 1:
                total_months = len(available_months)
                selected_range_cpp = st.slider(
                    "Pilih Range Periode",
                    min_value=1, max_value=total_months,
                    value=(1, total_months),
                    key='cpp_port_slider')
                start_idx = selected_range_cpp[0] - 1
                end_idx = selected_range_cpp[1]
                selected_cpp_months = available_months[start_idx:end_idx]
                df_valid_filtered = df_valid_copy[df_valid_copy['YearMonth'].isin(selected_cpp_months)].copy()
                if selected_cpp_months:
                    st.info(f"Periode: **{len(selected_cpp_months)} bulan** ({selected_cpp_months[0]} - {selected_cpp_months[-1]})")
            else:
                df_valid_filtered = df_valid_copy.copy()
                st.info(f"Data hanya untuk 1 bulan: {available_months[0]}")
            df_valid_filtered = df_valid_filtered.drop(columns=['YearMonth'], errors='ignore')

            # ── KPI CARDS (dulu, sebelum alert) ──
            ch_avg = df_valid_filtered['Dev_CH_Relatif_Pct'].abs().mean()
            cm_avg = df_valid_filtered['Dev_CM_Relatif_Pct'].abs().mean()
            twb_ch = df_valid_filtered['TWB_CH'].mean()
            twb_cm = df_valid_filtered['TWB_CM'].mean()
            ch_color = '#4ade80' if ch_avg <= 2 else '#fbbf24' if ch_avg <= 3 else '#f87171'
            cm_color = '#4ade80' if cm_avg <= 2 else '#fbbf24' if cm_avg <= 3 else '#f87171'

            kpi_cpp = [
                ("CH Avg |Dev%|", f"{format_number(ch_avg, 2)}%", ch_color),
                ("CM Avg |Dev%|", f"{format_number(cm_avg, 2)}%", cm_color),
                ("Avg TWB CH", format_large(twb_ch), "#60a5fa"),
                ("Avg TWB CM", format_large(twb_cm), "#eab308"),
            ]
            cols_cpp = st.columns(4, gap="medium")
            for i, (label, value, color) in enumerate(kpi_cpp):
                with cols_cpp[i]:
                    st.markdown(f"""
                    <div style="background:linear-gradient(145deg,rgba(22,28,45,0.9),rgba(15,20,35,0.95));
                        border:1px solid rgba(148,163,184,0.08);border-radius:12px;
                        padding:14px 12px;text-align:center;position:relative;overflow:hidden;">
                        <div style="position:absolute;top:0;left:0;right:0;height:2px;
                            background:{color};border-radius:12px 12px 0 0;"></div>
                        <div style="font-size:0.65rem;font-weight:600;text-transform:uppercase;
                            letter-spacing:0.08em;color:#94a3b8;margin-bottom:4px;">{label}</div>
                        <div style="font-size:1.5rem;font-weight:800;color:{color};line-height:1.1;">{value}</div>
                    </div>
                    """, unsafe_allow_html=True)

            # ── ALERT BANNER (di bawah KPI, hanya jika ada Critical) ──
            crit_ch = df_valid_filtered[df_valid_filtered['Status_CH'] == 'Critical']
            crit_cm = df_valid_filtered[df_valid_filtered['Status_CM'] == 'Critical']
            n_crit = len(crit_ch) + len(crit_cm)

            if n_crit > 0:
                parts = []
                if len(crit_ch) > 0:
                    parts.append(f'<span style="background:rgba(59,130,246,0.15);color:#60a5fa;'
                        f'padding:4px 12px;border-radius:8px;font-size:0.78rem;font-weight:600;">'
                        f'CH: {len(crit_ch)} period(s)</span>')
                if len(crit_cm) > 0:
                    parts.append(f'<span style="background:rgba(245,158,11,0.15);color:#fbbf24;'
                        f'padding:4px 12px;border-radius:8px;font-size:0.78rem;font-weight:600;">'
                        f'CM: {len(crit_cm)} period(s)</span>')
                st.markdown(f"""
                <div style="background:linear-gradient(135deg,rgba(239,68,68,0.08),rgba(15,23,42,0.9));
                    border:1px solid rgba(239,68,68,0.2);border-radius:12px;
                    padding:14px 20px;margin-top:12px;">
                    <div style="color:#f87171;font-size:0.72rem;font-weight:700;
                        text-transform:uppercase;letter-spacing:0.1em;margin-bottom:8px;">
                        Critical Alert — {n_crit} Period(s) Detected</div>
                    <div style="display:flex;flex-wrap:wrap;gap:8px;">{" ".join(parts)}</div>
                </div>
                """, unsafe_allow_html=True)

            # ── CHARTS (langsung setelah alert, tanpa status distribution & insight) ──
            col_left, col_right = st.columns([1, 1])

            with col_left:
                st.markdown("""
                <div class="section-header">
                    <h3 class="section-title">Total Weight Bridge vs Weight Bridge Trend</h3>
                </div>
                """, unsafe_allow_html=True)

                # Coal Hauling
                fig_ch = go.Figure()
                fig_ch.add_trace(go.Scatter(
                    x=df_valid_filtered['Date'], y=df_valid_filtered['TWB_CH'],
                    name='TWB (Actual)',
                    line=dict(color='#3b82f6', width=2.5),
                    mode='lines+markers',
                    marker=dict(size=6, line=dict(width=1, color='rgba(255,255,255,0.5)')),
                    hovertemplate='<b>TWB CH</b><br>%{y:,.0f}<extra></extra>'
                ))
                fig_ch.add_trace(go.Scatter(
                    x=df_valid_filtered['Date'], y=df_valid_filtered['CH_WB'],
                    name='WB (Target)',
                    line=dict(color='#22c55e', width=2.5, dash='dash'),
                    mode='lines+markers', marker=dict(size=6),
                    fill='tonexty', fillcolor='rgba(59,130,246,0.08)',
                    hovertemplate='<b>WB Target</b><br>%{y:,.0f}<extra></extra>'
                ))
                if len(df_valid_filtered) > 1:
                    ch_pk = df_valid_filtered['TWB_CH'].idxmax()
                    fig_ch.add_annotation(
                        x=df_valid_filtered.loc[ch_pk, 'Date'],
                        y=df_valid_filtered.loc[ch_pk, 'TWB_CH'],
                        text=f"Peak {df_valid_filtered.loc[ch_pk, 'TWB_CH']:,.0f}",
                        showarrow=True, arrowhead=2, arrowwidth=1.5,
                        arrowcolor='#60a5fa', ax=0, ay=-28,
                        font=dict(size=9, color='#FFF'),
                        bgcolor='rgba(59,130,246,0.75)', borderpad=3)
                    ch_mean = df_valid_filtered['TWB_CH'].mean()
                    fig_ch.add_hline(y=ch_mean, line_dash='dot',
                        line_color='rgba(59,130,246,0.3)', line_width=1,
                        annotation_text=f'Avg {ch_mean:,.0f}',
                        annotation_position='top left',
                        annotation_font=dict(size=9, color='#60a5fa'))
                fig_ch.update_layout(
                    title=dict(text='Coal Hauling: TWB vs WB (ton)', font=dict(size=13, color='#ffffff')),
                    height=320,
                    font=dict(family='Inter, sans-serif', color='#e5e7eb'),
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(gridcolor='rgba(255,255,255,0.08)', title=''),
                    yaxis=dict(gridcolor='rgba(255,255,255,0.08)', title='Ton', tickformat=',.0f'),
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                    hovermode='x unified',
                    margin=dict(l=60, r=20, t=50, b=40)
                )
                st.plotly_chart(fig_ch, use_container_width=True, key='cpp_ch_comp')

                # Coal Mining
                fig_cm = go.Figure()
                fig_cm.add_trace(go.Scatter(
                    x=df_valid_filtered['Date'], y=df_valid_filtered['TWB_CM'],
                    name='TWB (Actual)',
                    line=dict(color='#eab308', width=2.5),
                    mode='lines+markers',
                    marker=dict(size=6, line=dict(width=1, color='rgba(255,255,255,0.5)')),
                    hovertemplate='<b>TWB CM</b><br>%{y:,.0f}<extra></extra>'
                ))
                fig_cm.add_trace(go.Scatter(
                    x=df_valid_filtered['Date'], y=df_valid_filtered['CM_WB'],
                    name='WB (Target)',
                    line=dict(color='#ef4444', width=2.5, dash='dash'),
                    mode='lines+markers', marker=dict(size=6),
                    fill='tonexty', fillcolor='rgba(234,179,8,0.08)',
                    hovertemplate='<b>WB Target</b><br>%{y:,.0f}<extra></extra>'
                ))
                if len(df_valid_filtered) > 1:
                    cm_pk = df_valid_filtered['TWB_CM'].idxmax()
                    fig_cm.add_annotation(
                        x=df_valid_filtered.loc[cm_pk, 'Date'],
                        y=df_valid_filtered.loc[cm_pk, 'TWB_CM'],
                        text=f"Peak {df_valid_filtered.loc[cm_pk, 'TWB_CM']:,.0f}",
                        showarrow=True, arrowhead=2, arrowwidth=1.5,
                        arrowcolor='#eab308', ax=0, ay=-28,
                        font=dict(size=9, color='#FFF'),
                        bgcolor='rgba(234,179,8,0.75)', borderpad=3)
                    cm_mean = df_valid_filtered['TWB_CM'].mean()
                    fig_cm.add_hline(y=cm_mean, line_dash='dot',
                        line_color='rgba(234,179,8,0.3)', line_width=1,
                        annotation_text=f'Avg {cm_mean:,.0f}',
                        annotation_position='top left',
                        annotation_font=dict(size=9, color='#eab308'))
                fig_cm.update_layout(
                    title=dict(text='Coal Mining: TWB vs WB (ton)', font=dict(size=13, color='#ffffff')),
                    height=320,
                    font=dict(family='Inter, sans-serif', color='#e5e7eb'),
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(gridcolor='rgba(255,255,255,0.08)', title=''),
                    yaxis=dict(gridcolor='rgba(255,255,255,0.08)', title='Ton', tickformat=',.0f'),
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                    hovermode='x unified',
                    margin=dict(l=60, r=20, t=50, b=40)
                )
                st.plotly_chart(fig_cm, use_container_width=True, key='cpp_cm_comp')

            with col_right:
                st.markdown("""
                <div class="section-header">
                    <h3 class="section-title">Deviation Pattern</h3>
                </div>
                """, unsafe_allow_html=True)

                # CH Deviation
                c_ch = [get_status_color(s) for s in df_valid_filtered['Status_CH']]
                fig_chd = go.Figure()
                fig_chd.add_shape(type='rect', xref='paper', yref='y',
                    x0=0, x1=1, y0=-2, y1=2,
                    fillcolor='rgba(34,197,94,0.04)', line_width=0, layer='below')
                fig_chd.add_trace(go.Bar(
                    x=df_valid_filtered['Date'],
                    y=df_valid_filtered['Dev_CH_Relatif_Pct'],
                    marker_color=c_ch, showlegend=False,
                    hovertemplate='<b>CH Dev</b><br>%{y:.2f}%<extra></extra>'
                ))
                fig_chd.add_hline(y=0, line_color="#9ca3af", line_width=2)
                fig_chd.add_hline(y=2, line_color="#22c55e", line_width=1, line_dash="dash",
                    annotation_text="2%", annotation_position="top right",
                    annotation_font=dict(size=8, color='#22c55e'))
                fig_chd.add_hline(y=-2, line_color="#22c55e", line_width=1, line_dash="dash")
                fig_chd.add_hline(y=3, line_color="#eab308", line_width=1, line_dash="dash",
                    annotation_text="3%", annotation_position="top right",
                    annotation_font=dict(size=8, color='#eab308'))
                fig_chd.add_hline(y=-3, line_color="#eab308", line_width=1, line_dash="dash")
                if len(df_valid_filtered) > 0:
                    cw = df_valid_filtered['Dev_CH_Relatif_Pct'].abs().idxmax()
                    cwv = df_valid_filtered.loc[cw, 'Dev_CH_Relatif_Pct']
                    fig_chd.add_annotation(
                        x=df_valid_filtered.loc[cw, 'Date'], y=cwv,
                        text=f"Worst {cwv:.1f}%",
                        showarrow=True, arrowhead=2, arrowwidth=1.5,
                        arrowcolor='#ef4444',
                        ax=30, ay=-25 if cwv > 0 else 25,
                        font=dict(size=9, color='#FFF'),
                        bgcolor='rgba(239,68,68,0.8)', borderpad=3)
                fig_chd.update_layout(
                    title=dict(text='CH Daily Deviation (%)', font=dict(size=13, color='#ffffff')),
                    height=320,
                    font=dict(family='Inter, sans-serif', color='#e5e7eb'),
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(gridcolor='rgba(255,255,255,0.08)', title=''),
                    yaxis=dict(gridcolor='rgba(255,255,255,0.08)', title='Deviation (%)'),
                    hovermode='x unified',
                    margin=dict(l=60, r=20, t=50, b=40)
                )
                st.plotly_chart(fig_chd, use_container_width=True, key='cpp_ch_dev')

                # CM Deviation
                c_cm = [get_status_color(s) for s in df_valid_filtered['Status_CM']]
                fig_cmd = go.Figure()
                fig_cmd.add_shape(type='rect', xref='paper', yref='y',
                    x0=0, x1=1, y0=-2, y1=2,
                    fillcolor='rgba(34,197,94,0.04)', line_width=0, layer='below')
                fig_cmd.add_trace(go.Bar(
                    x=df_valid_filtered['Date'],
                    y=df_valid_filtered['Dev_CM_Relatif_Pct'],
                    marker_color=c_cm, showlegend=False,
                    hovertemplate='<b>CM Dev</b><br>%{y:.2f}%<extra></extra>'
                ))
                fig_cmd.add_hline(y=0, line_color="#9ca3af", line_width=2)
                fig_cmd.add_hline(y=2, line_color="#22c55e", line_width=1, line_dash="dash",
                    annotation_text="2%", annotation_position="top right",
                    annotation_font=dict(size=8, color='#22c55e'))
                fig_cmd.add_hline(y=-2, line_color="#22c55e", line_width=1, line_dash="dash")
                fig_cmd.add_hline(y=3, line_color="#eab308", line_width=1, line_dash="dash",
                    annotation_text="3%", annotation_position="top right",
                    annotation_font=dict(size=8, color='#eab308'))
                fig_cmd.add_hline(y=-3, line_color="#eab308", line_width=1, line_dash="dash")
                if len(df_valid_filtered) > 0:
                    mw = df_valid_filtered['Dev_CM_Relatif_Pct'].abs().idxmax()
                    mwv = df_valid_filtered.loc[mw, 'Dev_CM_Relatif_Pct']
                    fig_cmd.add_annotation(
                        x=df_valid_filtered.loc[mw, 'Date'], y=mwv,
                        text=f"Worst {mwv:.1f}%",
                        showarrow=True, arrowhead=2, arrowwidth=1.5,
                        arrowcolor='#ef4444',
                        ax=30, ay=-25 if mwv > 0 else 25,
                        font=dict(size=9, color='#FFF'),
                        bgcolor='rgba(239,68,68,0.8)', borderpad=3)
                fig_cmd.update_layout(
                    title=dict(text='CM Daily Deviation (%)', font=dict(size=13, color='#ffffff')),
                    height=320,
                    font=dict(family='Inter, sans-serif', color='#e5e7eb'),
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(gridcolor='rgba(255,255,255,0.08)', title=''),
                    yaxis=dict(gridcolor='rgba(255,255,255,0.08)', title='Deviation (%)'),
                    hovermode='x unified',
                    margin=dict(l=60, r=20, t=50, b=40)
                )
                st.plotly_chart(fig_cmd, use_container_width=True, key='cpp_cm_dev')

            # ── CH-CM DATA TABLE — styled HTML (replace st.dataframe) ──
            st.markdown("""
            <div class="section-header" style="margin-top:2.5rem;">
                <h3 class="section-title">CH-CM Data Table</h3>
            </div>
            """, unsafe_allow_html=True)

            if not df_valid_filtered.empty:
                chcm_cols_def = [
                    ('Date', 'Date'), ('Port_Darat', 'Port Darat'), ('Port_Laut', 'Port Laut'),
                    ('Port_Total', 'Port Total'), ('CPP_Raw', 'CPP Raw'), ('CPP_Product', 'CPP Product'),
                    ('CPP_Total', 'CPP Total'), ('Sales', 'Sales'),
                    ('CH_WB', 'CH WB'), ('TWB_CH', 'TWB CH'), ('Dev_CH_Relatif_Pct', 'CH Dev%'), ('Status_CH', 'Status CH'),
                    ('CM_WB', 'CM WB'), ('TWB_CM', 'TWB CM'), ('Dev_CM_Relatif_Pct', 'CM Dev%'), ('Status_CM', 'Status CM'),
                ]
                chcm_cols = [(k, v) for k, v in chcm_cols_def if k in df_valid_filtered.columns]

                chcm_header = ''.join(
                    f'<th style="padding:9px 8px;text-align:center;color:#e2e8f0;font-weight:700;'
                    f'font-size:0.62rem;text-transform:uppercase;letter-spacing:0.04em;'
                    f'position:sticky;top:0;background:rgba(22,101,52,0.95);z-index:1;white-space:nowrap;">{label}</th>'
                    for _, label in chcm_cols
                )

                chcm_rows = ""
                for _, row in df_valid_filtered.iterrows():
                    ch_st = row.get('Status_CH', 'Normal')
                    cm_st = row.get('Status_CM', 'Normal')
                    worst = 'Critical' if 'Critical' in (ch_st, cm_st) else 'Caution' if 'Caution' in (ch_st, cm_st) else 'Normal'

                    if worst == 'Critical':
                        row_bg, bl = 'rgba(239,68,68,0.06)', '3px solid #ef4444'
                    elif worst == 'Caution':
                        row_bg, bl = 'rgba(245,158,11,0.06)', '3px solid #f59e0b'
                    else:
                        row_bg, bl = 'rgba(34,197,94,0.03)', '3px solid #22c55e'

                    cells = ""
                    for col_key, _ in chcm_cols:
                        val = row.get(col_key, '')

                        if col_key in ('Status_CH', 'Status_CM'):
                            s = str(val)
                            if s == 'Critical':
                                sc, sbg = '#f87171', 'rgba(239,68,68,0.15)'
                            elif s == 'Caution':
                                sc, sbg = '#fbbf24', 'rgba(245,158,11,0.15)'
                            else:
                                sc, sbg = '#4ade80', 'rgba(34,197,94,0.15)'
                            cells += (f'<td style="padding:8px 6px;text-align:center;">'
                                      f'<span style="background:{sbg};color:{sc};padding:2px 8px;'
                                      f'border-radius:20px;font-size:0.65rem;font-weight:700;">{s}</span></td>')

                        elif col_key in ('Dev_CH_Relatif_Pct', 'Dev_CM_Relatif_Pct'):
                            dv = float(val) if pd.notnull(val) else 0
                            dc = '#f87171' if abs(dv) > 3 else '#fbbf24' if abs(dv) > 2 else '#4ade80'
                            cells += (f'<td style="padding:8px 6px;text-align:center;color:{dc};'
                                      f'font-weight:700;font-size:0.75rem;">{dv:+.2f}%</td>')

                        elif col_key == 'Date':
                            cells += (f'<td style="padding:8px 6px;text-align:center;color:#cbd5e1;'
                                      f'font-size:0.72rem;white-space:nowrap;">{str(val)[:10]}</td>')

                        elif isinstance(val, (int, float)) and pd.notnull(val):
                            cells += (f'<td style="padding:8px 6px;text-align:center;color:#cbd5e1;'
                                      f'font-size:0.72rem;">{val:,.2f}</td>')
                        else:
                            cells += (f'<td style="padding:8px 6px;text-align:center;color:#cbd5e1;'
                                      f'font-size:0.72rem;">{val}</td>')

                    chcm_rows += (f'<tr style="background:{row_bg};border-left:{bl};'
                                  f'border-bottom:1px solid rgba(148,163,184,0.06);">{cells}</tr>')

                st.markdown(f"""
                <div style="background:linear-gradient(135deg,rgba(20,20,40,0.6),rgba(15,23,42,0.8));
                    border:1px solid rgba(148,163,184,0.1);border-radius:10px;overflow:hidden;
                    max-height:420px;overflow-y:auto;overflow-x:auto;">
                <table style="width:100%;border-collapse:collapse;min-width:1200px;">
                <thead><tr style="background:rgba(22,101,52,0.95);">{chcm_header}</tr></thead>
                <tbody>{chcm_rows}</tbody>
                </table></div>
                """, unsafe_allow_html=True)
            else:
                st.info("Tidak ada data untuk ditampilkan")

    with tab4:
        # ══════════════════════════════════════════════════════════
        # TAB 4: PERFORMANCE REPORTS
        # ══════════════════════════════════════════════════════════
        matrix = create_matrix(df_ob, df_ch_cm)
        flow = create_flow(df_ch_cm)
        df_ch = df_ch_cm[df_ch_cm['CH_WB'].notna()].copy()
        df_cm_data = df_ch_cm[df_ch_cm['CM_WB'].notna()].copy()

        if flow is not None and len(flow) > 0:
            flow = flow.copy()
            if 'Net Deviation' not in flow.columns:
                flow['Net Deviation'] = flow['Deviation CPP'] + flow['Deviation Port']
            if 'Cum Net Deviation' not in flow.columns:
                flow['Cum Net Deviation'] = flow['Net Deviation'].cumsum()

        ob_avg = matrix.iloc[0]['Avg'] if len(matrix) > 0 else 0
        ch_avg = matrix.iloc[1]['Avg'] if len(matrix) > 1 else 0
        cm_avg = matrix.iloc[2]['Avg'] if len(matrix) > 2 else 0

        total_normal = int(matrix['Normal'].sum())
        total_caution = int(matrix['Caution'].sum())
        total_critical = int(matrix['Critical'].sum())
        total_all = int(matrix['Total'].sum())
        perf = total_normal / total_all * 100 if total_all > 0 else 0

        ob_disp_dev = df_ob['Dev_Relatif_Pct'].abs().mean()
        ch_disp_dev = df_ch['Dev_CH_Relatif_Pct'].abs().mean() if len(df_ch) > 0 else 0
        cm_disp_dev = df_cm_data['Dev_CM_Relatif_Pct'].abs().mean() if len(df_cm_data) > 0 else 0

        disp_ch_ratio = flow['CH Flow Ratio (%)'].mean() if (flow is not None and len(flow) > 0) else 0
        disp_sales_ratio = flow['Sales Flow Ratio (%)'].mean() if (flow is not None and len(flow) > 0) else 0

        disp_cm = flow['CM TWB'].mean() if (flow is not None and len(flow) > 0) else 0
        disp_ch = flow['CH TWB'].mean() if (flow is not None and len(flow) > 0) else 0
        disp_sales = flow['Sales'].mean() if (flow is not None and len(flow) > 0) else 0

        disp_dev_cpp = flow['Deviation CPP'].mean() if (flow is not None and len(flow) > 0) else 0
        disp_dev_port = flow['Deviation Port'].mean() if (flow is not None and len(flow) > 0) else 0
        disp_net_dev = flow['Net Deviation'].mean() if (flow is not None and len(flow) > 0) else 0

        report_date = datetime.now().strftime("%d %B %Y")

        def get_status_label(val):
            if val <= 2:
                return "Normal", "#4ade80"
            elif val <= 3:
                return "Caution", "#fbbf24"
            else:
                return "Critical", "#f87171"

        def scolor(status):
            if status == 'Normal':
                return '#22C55E'
            elif status == 'Caution':
                return '#F59E0B'
            return '#EF4444'

        def sbadge(val):
            if val <= 2:
                return '#22C55E', 'Normal'
            elif val <= 3:
                return '#F59E0B', 'Caution'
            return '#EF4444', 'Critical'

        ob_lbl, ob_c = get_status_label(ob_avg)
        ch_lbl, ch_c = get_status_label(ch_avg)
        cm_lbl, cm_c = get_status_label(cm_avg)
        perf_c = '#4ade80' if perf >= 70 else '#fbbf24' if perf >= 50 else '#f87171'
        perf_lbl = "Good" if perf >= 70 else "Fair" if perf >= 50 else "Poor"

        ob_clr, ob_lbl2 = sbadge(ob_disp_dev)
        ch_clr, ch_lbl2 = sbadge(ch_disp_dev)
        cm_clr, cm_lbl2 = sbadge(cm_disp_dev)

        st.markdown("""
        <div class="section-header" style="margin-top:2.5rem;">
            <h3 class="section-title">Performance Matrix</h3>
        </div>
        """, unsafe_allow_html=True)

        if not matrix.empty:
            perf_html = """
            <div style="background:linear-gradient(135deg,rgba(20,20,40,0.6),rgba(15,23,42,0.8));
                border:1px solid rgba(148,163,184,0.1);border-radius:12px;overflow:hidden;">
            <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
            <thead>
            <tr style="background:rgba(22,101,52,0.4);">
                <th style="padding:12px 16px;text-align:left;color:#e2e8f0;font-weight:700;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.08em;">Stage</th>
                <th style="padding:12px 14px;text-align:center;color:#e2e8f0;font-weight:700;font-size:0.72rem;text-transform:uppercase;">Avg |Dev%|</th>
                <th style="padding:12px 14px;text-align:center;color:#4ade80;font-weight:700;font-size:0.72rem;text-transform:uppercase;">Normal</th>
                <th style="padding:12px 14px;text-align:center;color:#fbbf24;font-weight:700;font-size:0.72rem;text-transform:uppercase;">Caution</th>
                <th style="padding:12px 14px;text-align:center;color:#f87171;font-weight:700;font-size:0.72rem;text-transform:uppercase;">Critical</th>
                <th style="padding:12px 14px;text-align:center;color:#e2e8f0;font-weight:700;font-size:0.72rem;text-transform:uppercase;">Total</th>
                <th style="padding:12px 14px;text-align:center;color:#e2e8f0;font-weight:700;font-size:0.72rem;text-transform:uppercase;">Performance</th>
            </tr>
            </thead><tbody>"""

            stage_names = ['Overburden (OB)', 'Coal Hauling (CH/Port)', 'Coal Mining (CM/CPP33)']
            for idx, row in matrix.iterrows():
                disp_d = row['Avg']
                dc = '#4ade80' if disp_d <= 2 else '#fbbf24' if disp_d <= 3 else '#f87171'
                norm = int(row['Normal'])
                caut = int(row['Caution'])
                crit = int(row['Critical'])
                tot = int(row['Total'])
                p = norm / tot * 100 if tot > 0 else 0
                pc = '#4ade80' if p >= 70 else '#fbbf24' if p >= 50 else '#f87171'
                name = stage_names[idx] if idx < len(stage_names) else f"Stage {idx}"
                perf_html += f"""
                <tr style="border-bottom:1px solid rgba(148,163,184,0.06);">
                    <td style="padding:11px 16px;color:#e2e8f0;font-weight:600;">{name}</td>
                    <td style="padding:11px 14px;text-align:center;color:{dc};font-weight:700;">{disp_d:.2f}%</td>
                    <td style="padding:11px 14px;text-align:center;color:#4ade80;font-weight:700;">{norm}</td>
                    <td style="padding:11px 14px;text-align:center;color:#fbbf24;font-weight:700;">{caut}</td>
                    <td style="padding:11px 14px;text-align:center;color:#f87171;font-weight:700;">{crit}</td>
                    <td style="padding:11px 14px;text-align:center;color:#94a3b8;">{tot}</td>
                    <td style="padding:11px 14px;text-align:center;color:{pc};font-weight:700;">{p:.1f}%</td>
                </tr>"""

            perf_html += f"""
            <tr style="border-top:2px solid #22c55e;background:rgba(34,197,94,0.06);">
                <td style="padding:12px 16px;color:#22c55e;font-weight:800;">OVERALL</td>
                <td style="padding:12px 14px;text-align:center;color:#64748b;">—</td>
                <td style="padding:12px 14px;text-align:center;color:#22c55e;font-weight:800;">{total_normal}</td>
                <td style="padding:12px 14px;text-align:center;color:#fbbf24;font-weight:800;">{total_caution}</td>
                <td style="padding:12px 14px;text-align:center;color:#f87171;font-weight:800;">{total_critical}</td>
                <td style="padding:12px 14px;text-align:center;color:#e2e8f0;font-weight:800;">{total_all}</td>
                <td style="padding:12px 14px;text-align:center;color:{perf_c};font-weight:800;">{perf:.1f}%</td>
            </tr>"""

            perf_html += "</tbody></table></div>"
            st.markdown(perf_html, unsafe_allow_html=True)

        st.markdown("""
        <div class="section-header" style="margin-top:2.5rem;">
            <h3 class="section-title">Critical / Caution Periods</h3>
        </div>
        """, unsafe_allow_html=True)

        def render_alert_table(df_alert, cols, dev_col, status_col):
            if df_alert.empty:
                return '<div style="color:#4ade80;font-size:0.8rem;padding:12px;">Semua periode Normal</div>'

            header_html = ''.join(
                f'<th style="padding:10px 8px;text-align:center;color:#e2e8f0;font-weight:700;'
                f'font-size:0.65rem;text-transform:uppercase;letter-spacing:0.06em;'
                f'position:sticky;top:0;background:rgba(22,101,52,0.95);z-index:1;">{c}</th>'
                for c in cols
            )

            rows_html = ""
            for _, row in df_alert.iterrows():
                status = row[status_col]
                if status == 'Critical':
                    row_bg = 'rgba(239,68,68,0.08)'
                    border_left = '3px solid #ef4444'
                    status_color = '#f87171'
                    status_bg = 'rgba(239,68,68,0.15)'
                elif status == 'Caution':
                    row_bg = 'rgba(245,158,11,0.08)'
                    border_left = '3px solid #f59e0b'
                    status_color = '#fbbf24'
                    status_bg = 'rgba(245,158,11,0.15)'
                else:
                    row_bg = 'transparent'
                    border_left = '3px solid #22c55e'
                    status_color = '#4ade80'
                    status_bg = 'rgba(34,197,94,0.15)'

                cells = ""
                for c in cols:
                    val = row[c]
                    if c == status_col:
                        cells += (
                            f'<td style="padding:8px 6px;text-align:center;">'
                            f'<span style="background:{status_bg};color:{status_color};'
                            f'padding:3px 8px;border-radius:20px;font-size:0.65rem;font-weight:700;">'
                            f'{val}</span></td>'
                        )
                    elif c == dev_col:
                        dev_val = float(val) if not pd.isna(val) else 0
                        dev_color = '#f87171' if abs(dev_val) > 3 else '#fbbf24' if abs(dev_val) > 2 else '#4ade80'
                        cells += (
                            f'<td style="padding:8px 6px;text-align:center;color:{dev_color};'
                            f'font-weight:700;font-size:0.75rem;">{dev_val:+.2f}%</td>'
                        )
                    elif isinstance(val, (int, float)) and not pd.isna(val):
                        cells += (
                            f'<td style="padding:8px 6px;text-align:center;color:#cbd5e1;'
                            f'font-size:0.72rem;">{val:,.2f}</td>'
                        )
                    else:
                        display_val = str(val)[:10] if 'Date' in c or 'date' in str(c).lower() else str(val)
                        cells += (
                            f'<td style="padding:8px 6px;text-align:center;color:#cbd5e1;'
                            f'font-size:0.72rem;">{display_val}</td>'
                        )

                rows_html += (
                    f'<tr style="background:{row_bg};border-left:{border_left};'
                    f'border-bottom:1px solid rgba(148,163,184,0.06);">{cells}</tr>'
                )

            return f"""
            <div style="background:linear-gradient(135deg,rgba(20,20,40,0.6),rgba(15,23,42,0.8));
                border:1px solid rgba(148,163,184,0.1);border-radius:10px;overflow:hidden;
                max-height:320px;overflow-y:auto;">
            <table style="width:100%;border-collapse:collapse;">
            <thead><tr style="background:rgba(22,101,52,0.95);">{header_html}</tr></thead>
            <tbody>{rows_html}</tbody>
            </table></div>"""

        col_cp1, col_cp2, col_cp3 = st.columns(3)

        with col_cp1:
            st.markdown('<div style="font-size:0.82rem;font-weight:700;color:#60a5fa;margin-bottom:10px;">Overburden (OB)</div>', unsafe_allow_html=True)
            ob_alert = df_ob[df_ob['Status'].isin(['Critical', 'Caution'])][['Bulan', 'TC', 'JS', 'Dev_Relatif_Pct', 'Status']].copy()
            st.markdown(render_alert_table(ob_alert, ['Bulan', 'TC', 'JS', 'Dev_Relatif_Pct', 'Status'], 'Dev_Relatif_Pct', 'Status'), unsafe_allow_html=True)

        with col_cp2:
            st.markdown('<div style="font-size:0.82rem;font-weight:700;color:#60a5fa;margin-bottom:10px;">Coal Hauling (CH)</div>', unsafe_allow_html=True)
            ch_alert = df_ch[df_ch['Status_CH'].isin(['Critical', 'Caution'])][['Date', 'TWB_CH', 'CH_WB', 'Dev_CH_Relatif_Pct', 'Status_CH']].copy() if len(df_ch) > 0 else pd.DataFrame()
            st.markdown(render_alert_table(ch_alert, ['Date', 'TWB_CH', 'CH_WB', 'Dev_CH_Relatif_Pct', 'Status_CH'], 'Dev_CH_Relatif_Pct', 'Status_CH'), unsafe_allow_html=True)

        with col_cp3:
            st.markdown('<div style="font-size:0.82rem;font-weight:700;color:#60a5fa;margin-bottom:10px;">Coal Mining (CM)</div>', unsafe_allow_html=True)
            cm_alert = df_cm_data[df_cm_data['Status_CM'].isin(['Critical', 'Caution'])][['Date', 'TWB_CM', 'CM_WB', 'Dev_CM_Relatif_Pct', 'Status_CM']].copy() if len(df_cm_data) > 0 else pd.DataFrame()
            st.markdown(render_alert_table(cm_alert, ['Date', 'TWB_CM', 'CM_WB', 'Dev_CM_Relatif_Pct', 'Status_CM'], 'Dev_CM_Relatif_Pct', 'Status_CM'), unsafe_allow_html=True)

        # ══════════════════════════════════════════════════════════
        # EXPORT REPORTS
        # ══════════════════════════════════════════════════════════
        st.markdown("""
        <div class="section-header">
            <h3 class="section-title"> Export Reports</h3>
        </div>
        """, unsafe_allow_html=True)

        chart_theme_export = dict(
            font=dict(family='Inter, Arial, sans-serif', color='#E2E8F0', size=12),
            plot_bgcolor='#1E293B',
            paper_bgcolor='#1E293B',
            margin=dict(l=60, r=40, t=55, b=45),
            hovermode='x unified'
        )
        chart_theme_st = dict(
            font=dict(family='Inter, Arial, sans-serif', color='#E2E8F0', size=11),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=50, r=30, t=50, b=40),
            hovermode='x unified'
        )
        axis_style = dict(
            gridcolor='rgba(148,163,184,0.15)',
            tickfont=dict(size=11, color='#CBD5E1'),
            linecolor='rgba(148,163,184,0.25)',
            showline=True,
            title_font=dict(size=12, color='#CBD5E1')
        )
        legend_cfg = dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1,
            font=dict(size=10, color='#CBD5E1'),
            bgcolor='rgba(0,0,0,0)'
        )

        import textwrap

        def fmt_period_label(value):
            try:
                return pd.to_datetime(value).strftime("%d %b %Y")
            except Exception:
                return str(value)

        latest_flow = None
        latest_period_label = "-"
        if flow is not None and len(flow) > 0:
            latest_flow = flow.sort_values("Date").iloc[-1]
            latest_period_label = fmt_period_label(latest_flow["Date"])

        stage_meta = {
            "OB": {
                "avg_dev": float(ob_disp_dev),
                "critical": int(matrix.iloc[0]["Critical"]) if len(matrix) > 0 else 0,
                "caution": int(matrix.iloc[0]["Caution"]) if len(matrix) > 0 else 0,
                "normal": int(matrix.iloc[0]["Normal"]) if len(matrix) > 0 else 0,
                "total": int(matrix.iloc[0]["Total"]) if len(matrix) > 0 else 0,
            },
            "CH": {
                "avg_dev": float(ch_disp_dev),
                "critical": int(matrix.iloc[1]["Critical"]) if len(matrix) > 1 else 0,
                "caution": int(matrix.iloc[1]["Caution"]) if len(matrix) > 1 else 0,
                "normal": int(matrix.iloc[1]["Normal"]) if len(matrix) > 1 else 0,
                "total": int(matrix.iloc[1]["Total"]) if len(matrix) > 1 else 0,
            },
            "CM": {
                "avg_dev": float(cm_disp_dev),
                "critical": int(matrix.iloc[2]["Critical"]) if len(matrix) > 2 else 0,
                "caution": int(matrix.iloc[2]["Caution"]) if len(matrix) > 2 else 0,
                "normal": int(matrix.iloc[2]["Normal"]) if len(matrix) > 2 else 0,
                "total": int(matrix.iloc[2]["Total"]) if len(matrix) > 2 else 0,
            },
        }

        def kpi_color(value):
            if value <= 2:
                return "#22C55E"
            if value <= 3:
                return "#F59E0B"
            return "#EF4444"

        def stage_note(stage):
            if stage == "OB":
                return "Review JS vs TC and survey consistency"
            if stage == "CH":
                return "Check WB CH, port stock movement, and shipment timing"
            return "Check WB CM, CPP stock movement, and survey / density inputs"

        def collect_top_alerts(limit=8):
            rows = []

            for _, r in df_ob[df_ob["Status"].isin(["Critical", "Caution"])].iterrows():
                rows.append({
                    "Priority": 0 if r["Status"] == "Critical" else 1,
                    "Stage": "OB",
                    "Period": str(r["Bulan"]),
                    "Deviation": abs(float(r["Dev_Relatif_Pct"])),
                    "DeviationLabel": f'{float(r["Dev_Relatif_Pct"]):+.2f}%',
                    "Status": r["Status"],
                    "Metric": "JS vs TC",
                    "Action": stage_note("OB"),
                })

            if len(df_ch) > 0:
                for _, r in df_ch[df_ch["Status_CH"].isin(["Critical", "Caution"])].iterrows():
                    rows.append({
                        "Priority": 0 if r["Status_CH"] == "Critical" else 1,
                        "Stage": "CH",
                        "Period": fmt_period_label(r["Date"]),
                        "Deviation": abs(float(r["Dev_CH_Relatif_Pct"])),
                        "DeviationLabel": f'{float(r["Dev_CH_Relatif_Pct"]):+.2f}%',
                        "Status": r["Status_CH"],
                        "Metric": "TWB CH vs WB",
                        "Action": stage_note("CH"),
                    })

            if len(df_cm_data) > 0:
                for _, r in df_cm_data[df_cm_data["Status_CM"].isin(["Critical", "Caution"])].iterrows():
                    rows.append({
                        "Priority": 0 if r["Status_CM"] == "Critical" else 1,
                        "Stage": "CM",
                        "Period": fmt_period_label(r["Date"]),
                        "Deviation": abs(float(r["Dev_CM_Relatif_Pct"])),
                        "DeviationLabel": f'{float(r["Dev_CM_Relatif_Pct"]):+.2f}%',
                        "Status": r["Status_CM"],
                        "Metric": "TWB CM vs WB",
                        "Action": stage_note("CM"),
                    })

            if not rows:
                return pd.DataFrame(columns=[
                    "Stage", "Period", "Deviation", "DeviationLabel",
                    "Status", "Metric", "Action"
                ])

            alerts_df = pd.DataFrame(rows)
            alerts_df = alerts_df.sort_values(
                ["Priority", "Deviation"],
                ascending=[True, False]
            ).drop(columns=["Priority"]).reset_index(drop=True)

            return alerts_df.head(limit)

        top_alerts = collect_top_alerts(limit=8)

        def build_key_findings():
            findings = []

            worst_stage = max(stage_meta.keys(), key=lambda k: stage_meta[k]["avg_dev"])
            findings.append(
                f"{worst_stage} menjadi fokus utama dengan average deviation "
                f"{stage_meta[worst_stage]['avg_dev']:.2f}% dan "
                f"{stage_meta[worst_stage]['critical']} periode critical."
            )

            findings.append(
                f"Overall performance saat ini {perf:.1f}%, sehingga prioritas utama adalah "
                f"menurunkan exception pada CH dan menjaga CM tetap di bawah threshold 3%."
            )

            if latest_flow is not None:
                dominant_dev_name = "Port" if abs(float(latest_flow["Deviation Port"])) >= abs(float(latest_flow["Deviation CPP"])) else "CPP33"
                dominant_dev_value = float(latest_flow["Deviation Port"]) if dominant_dev_name == "Port" else float(latest_flow["Deviation CPP"])
                findings.append(
                    f"Snapshot {latest_period_label}: deviation terbesar berada di {dominant_dev_name} "
                    f"sebesar {dominant_dev_value:,.0f} ton, dengan CH ratio {float(latest_flow['CH Flow Ratio (%)']):.1f}% "
                    f"dan Sales ratio {float(latest_flow['Sales Flow Ratio (%)']):.1f}%."
                )
            else:
                findings.append(
                    "Belum ada snapshot flow yang tersedia untuk dihitung dari data CH/CM."
                )

            return findings

        key_findings = build_key_findings()

        def build_recommended_actions():
            actions = []

            if ch_disp_dev > 3:
                actions.append("Prioritaskan audit Coal Hauling: validasi WB CH, port stock movement, dan timing shipment pada periode critical.")
            elif ch_disp_dev > 2:
                actions.append("Coal Hauling masih caution: lakukan review berkala pada periode dengan deviasi mendekati 3%.")

            if cm_disp_dev > 2:
                actions.append("Review Coal Mining / CPP33: cek konsistensi WB CM, perubahan stok CPP, dan input survey / densitas.")

            if latest_flow is not None and abs(float(latest_flow["Deviation Port"])) > abs(float(latest_flow["Deviation CPP"])):
                actions.append("Fokus investigasi tambahan di Port karena deviation snapshot terbaru lebih dominan dibanding CPP33.")
            else:
                actions.append("Fokus investigasi tambahan di CPP33 bila deviation snapshot terbaru lebih dominan di hulu.")

            if ob_disp_dev > 2:
                actions.append("Untuk Overburden, validasi kembali JS vs TC pada bulan-bulan exception agar deviasi kembali ke rentang normal.")

            return actions[:4]

        recommended_actions = build_recommended_actions()

        def make_fig_stage_summary(theme):
            order = sorted(
                [("Coal Hauling (CH)", ch_disp_dev), ("Coal Mining (CM)", cm_disp_dev), ("Overburden (OB)", ob_disp_dev)],
                key=lambda x: x[1],
                reverse=True
            )
            labels = [x[0] for x in order]
            values = [x[1] for x in order]
            colors = [kpi_color(v) for v in values]

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=values,
                y=labels,
                orientation="h",
                marker=dict(color=colors),
                text=[f"{v:.2f}%" for v in values],
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Avg Dev: %{x:.2f}%<extra></extra>"
            ))
            fig.add_vline(x=2, line_dash="dash", line_color="#F59E0B", line_width=1)
            fig.add_vline(x=3, line_dash="dash", line_color="#EF4444", line_width=1)

            fig.update_layout(
                **theme,
                height=360,
                showlegend=False,
                title=dict(text="Average Deviation by Stage", font=dict(size=16, color="#F1F5F9")),
                xaxis=dict(**axis_style, title="Average |Deviation| (%)", zeroline=False, range=[0, max(5, max(values) + 1)]),
                yaxis=dict(**axis_style, title="", categoryorder="array", categoryarray=labels[::-1]),
            )
            return fig

        def make_fig_dev_timeline(theme):
            fig = go.Figure()

            ch_plot = df_ch.sort_values("Date").tail(12) if len(df_ch) > 0 else pd.DataFrame()
            cm_plot = df_cm_data.sort_values("Date").tail(12) if len(df_cm_data) > 0 else pd.DataFrame()

            if len(ch_plot) > 0:
                fig.add_trace(go.Scatter(
                    x=ch_plot["Date"],
                    y=ch_plot["Dev_CH_Relatif_Pct"],
                    mode="lines+markers",
                    name="CH Dev",
                    line=dict(color="#EF4444", width=2.5),
                    marker=dict(size=7, color="#EF4444"),
                    hovertemplate="<b>CH</b><br>%{x|%d %b %Y}<br>%{y:.2f}%<extra></extra>"
                ))

            if len(cm_plot) > 0:
                fig.add_trace(go.Scatter(
                    x=cm_plot["Date"],
                    y=cm_plot["Dev_CM_Relatif_Pct"],
                    mode="lines+markers",
                    name="CM Dev",
                    line=dict(color="#60A5FA", width=2.5),
                    marker=dict(size=7, color="#60A5FA"),
                    hovertemplate="<b>CM</b><br>%{x|%d %b %Y}<br>%{y:.2f}%<extra></extra>"
                ))

            for yv, clr in [(2, "#F59E0B"), (3, "#EF4444"), (-2, "#F59E0B"), (-3, "#EF4444")]:
                fig.add_hline(y=yv, line_dash="dash", line_color=clr, line_width=1)

            layout_cfg = dict(theme)
            layout_cfg.update({
                height=360,
                title=dict(text="Deviation Trend (Last 12 Periods)", font=dict(size=16, color="#F1F5F9")),
                xaxis=dict(**axis_style, title="Period", tickformat="%d %b"),
                yaxis=dict(**axis_style, title="Deviation (%)", zeroline=False),
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="center",
                    x=0.5,
                    font=dict(size=10, color="#E2E8F0"),
                    bgcolor="rgba(0,0,0,0)"
                ),
            )
            
            fig.update_layout(**layout_cfg)
            return fig

def make_fig_latest_flow(theme):
    fig = go.Figure()

    if latest_flow is not None:
        cats = ["CM TWB", "CH TWB", "Sales"]
        vals = [
            float(latest_flow["CM TWB"]),
            float(latest_flow["CH TWB"]),
            float(latest_flow["Sales"]),
        ]
        cols = ["#60A5FA", "#34D399", "#A78BFA"]

        fig.add_trace(go.Bar(
            x=cats,
            y=vals,
            marker=dict(color=cols),
            text=[f"{v:,.0f}" for v in vals],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>%{y:,.0f} ton<extra></extra>"
        ))

        fig.add_annotation(
            x=0.5, y=1.16, xref="paper", yref="paper",
            text=(
                f"Latest period: {latest_period_label}"
                f"<br>Δ CPP: {float(latest_flow['Delta CPP Stock']):,.0f} ton"
                f" · Dev CPP: {float(latest_flow['Deviation CPP']):,.0f} ton"
                f"<br>Δ Port: {float(latest_flow['Delta Port Stock']):,.0f} ton"
                f" · Dev Port: {float(latest_flow['Deviation Port']):,.0f} ton"
            ),
            showarrow=False,
            font=dict(size=11, color="#CBD5E1"),
            align="center"
        )

    layout_cfg = dict(theme)
    layout_cfg.update({
        "height": 360,
        "showlegend": False,
        "title": dict(
            text="Latest Reconciliation Snapshot",
            font=dict(size=16, color="#F1F5F9")
        ),
        "xaxis": dict(**axis_style, title="Flow"),
        "yaxis": dict(**axis_style, title="Ton", zeroline=False),
        "margin": dict(l=60, r=40, t=95, b=55),
    })

    fig.update_layout(**layout_cfg)
    return fig

        def make_fig_ratio_trend(theme):
            fig = go.Figure()

            if flow is not None and len(flow) > 0:
                ratio_plot = flow.sort_values("Date").tail(12)

                fig.add_trace(go.Scatter(
                    x=ratio_plot["Date"],
                    y=ratio_plot["CH Flow Ratio (%)"],
                    mode="lines+markers",
                    name="CH Ratio",
                    line=dict(color="#34D399", width=2.5),
                    marker=dict(size=6, color="#34D399"),
                    hovertemplate="<b>CH Ratio</b><br>%{x|%d %b %Y}<br>%{y:.1f}%<extra></extra>"
                ))

                fig.add_trace(go.Scatter(
                    x=ratio_plot["Date"],
                    y=ratio_plot["Sales Flow Ratio (%)"],
                    mode="lines+markers",
                    name="Sales Ratio",
                    line=dict(color="#60A5FA", width=2.5, dash="dot"),
                    marker=dict(size=6, color="#60A5FA"),
                    hovertemplate="<b>Sales Ratio</b><br>%{x|%d %b %Y}<br>%{y:.1f}%<extra></extra>"
                ))

            fig.add_hline(y=100, line_dash="dash", line_color="#F59E0B", line_width=1.2)

            layout_cfg = dict(theme)
            layout_cfg.update({
                **theme,
                height=360,
                title=dict(text="Flow Ratio Trend (Last 12 Periods)", font=dict(size=16, color="#F1F5F9")),
                xaxis=dict(**axis_style, title="Period", tickformat="%d %b"),
                yaxis=dict(**axis_style, title="Flow Ratio (%)", zeroline=False),
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="center",
                    x=0.5,
                    font=dict(size=10, color="#E2E8F0"),
                    bgcolor="rgba(0,0,0,0)"
                ),
            )
            return fig

        # WAJIB ADA SEBELUM with exp1/exp2/exp3
        exp1, exp2, exp3 = st.columns(3)

        with exp1:
            st.markdown("""
            <div style="background:rgba(22,101,52,0.08);padding:10px 12px;border-radius:8px;
                border-left:3px solid #16A34A;margin-bottom:8px;">
                <div style="font-size:0.82rem;font-weight:700;color:#e5e7eb;">📊 Professional Excel Report</div>
                <div style="font-size:0.72rem;color:#94A3B8;margin-top:2px;">Executive summary + focused flow/deviation export</div>
            </div>
            """, unsafe_allow_html=True)

            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
            from openpyxl.utils import get_column_letter

            excel_buffer = io.BytesIO()

            C = {
                'kpp_dk': '166534', 'kpp_md': '16A34A', 'kpp_lt': 'DCFCE7', 'kpp_bg': 'F0FDF4',
                'gold_md': 'D97706',
                'gray1': '1F2937', 'gray2': '6B7280', 'gray3': 'F3F4F6',
                'green_txt': '059669', 'green_bg': 'D1FAE5',
                'amber_txt': 'D97706', 'amber_bg': 'FEF3C7',
                'red_txt': 'DC2626', 'red_bg': 'FEE2E2',
                'white': 'FFFFFF', 'bdr': 'D1D5DB',
            }

            hf = Font(bold=True, color=C['white'], size=10)
            bd = Border(
                left=Side(style='thin', color=C['bdr']), right=Side(style='thin', color=C['bdr']),
                top=Side(style='thin', color=C['bdr']), bottom=Side(style='thin', color=C['bdr'])
            )
            al_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
            al_l = Alignment(horizontal='left', vertical='center', wrap_text=True)
            af = PatternFill(start_color=C['kpp_bg'], end_color=C['kpp_bg'], fill_type='solid')

            sfill = {
                'Normal': PatternFill(start_color=C['green_bg'], end_color=C['green_bg'], fill_type='solid'),
                'Good': PatternFill(start_color=C['green_bg'], end_color=C['green_bg'], fill_type='solid'),
                'Caution': PatternFill(start_color=C['amber_bg'], end_color=C['amber_bg'], fill_type='solid'),
                'Warning': PatternFill(start_color=C['amber_bg'], end_color=C['amber_bg'], fill_type='solid'),
                'Critical': PatternFill(start_color=C['red_bg'], end_color=C['red_bg'], fill_type='solid'),
            }
            sfont = {
                'Normal': C['green_txt'], 'Good': C['green_txt'],
                'Caution': C['amber_txt'], 'Warning': C['amber_txt'],
                'Critical': C['red_txt'],
            }

            def style_data_sheet(ws, title_text, hdr_row, data_start_row, merge_cols,
                                 status_col_idx=None, freeze_cell='A2',
                                 num_fmt_cols=None, pct_fmt_cols=None):
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=merge_cols)
                ws['A1'] = title_text
                ws['A1'].font = Font(size=14, bold=True, color=C['white'])
                ws['A1'].fill = PatternFill(start_color=C['kpp_dk'], end_color=C['kpp_dk'], fill_type='solid')
                ws['A1'].alignment = al_c
                ws.row_dimensions[1].height = 30

                for col in range(1, merge_cols + 1):
                    cell = ws.cell(row=hdr_row, column=col)
                    cell.font = hf
                    cell.fill = PatternFill(start_color=C['kpp_md'], end_color=C['kpp_md'], fill_type='solid')
                    cell.alignment = al_c
                    cell.border = bd

                for ri, row in enumerate(ws.iter_rows(min_row=data_start_row, max_row=ws.max_row, max_col=merge_cols)):
                    for cell in row:
                        cell.border = bd
                        cell.alignment = al_c
                        if ri % 2 == 1:
                            cell.fill = af
                        if status_col_idx and cell.column == status_col_idx:
                            sv = str(cell.value) if cell.value else ''
                            if sv in sfill:
                                cell.fill = sfill[sv]
                                cell.font = Font(bold=True, size=9, color=sfont.get(sv, C['gray1']))
                        if num_fmt_cols and cell.column in num_fmt_cols:
                            cell.number_format = '#,##0'
                        if pct_fmt_cols and cell.column in pct_fmt_cols:
                            cell.number_format = '0.00'

                for ci in range(1, merge_cols + 1):
                    ml = 0
                    for row in ws.iter_rows(min_col=ci, max_col=ci):
                        for cell in row:
                            if cell.value:
                                ml = max(ml, len(str(cell.value)[:40]))
                    ws.column_dimensions[get_column_letter(ci)].width = max(min(ml + 3, 22), 10)

                ws.freeze_panes = freeze_cell

            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                ed = []
                ed.append(['KPP MINING — EXECUTIVE DEVIATION REPORT', '', '', '', '', '', ''])
                ed.append([f'PT Kalimantan Prima Persada | Generated: {datetime.now().strftime("%d %b %Y, %H:%M")}', '', '', '', '', '', ''])
                ed.append(['', '', '', '', '', '', ''])
                ed.append(['KEY PERFORMANCE INDICATORS', '', '', '', '', '', ''])
                ed.append(['Overall Perf', 'Total Periods', 'Critical Alerts', 'CH Ratio', 'OB Avg Dev', 'CH Avg Dev', 'CM Avg Dev'])
                ed.append([
                    round(perf, 1), total_all,
                    total_critical, round(disp_ch_ratio, 1),
                    round(ob_disp_dev, 2), round(ch_disp_dev, 2), round(cm_disp_dev, 2)
                ])
                ed.append(['', '', '', '', '', '', ''])
                ed.append(['FLOW SNAPSHOT', 'Value', '', '', 'LATEST PERIOD', latest_period_label, ''])
                if latest_flow is not None:
                    ed.append(['CM TWB', round(float(latest_flow['CM TWB'])), '', '', 'CH TWB', round(float(latest_flow['CH TWB'])), ''])
                    ed.append(['Sales', round(float(latest_flow['Sales'])), '', '', 'Δ CPP Stock', round(float(latest_flow['Delta CPP Stock'])), ''])
                    ed.append(['Δ Port Stock', round(float(latest_flow['Delta Port Stock'])), '', '', 'Dev CPP', round(float(latest_flow['Deviation CPP'])), ''])
                    ed.append(['Dev Port', round(float(latest_flow['Deviation Port'])), '', '', 'Sales Ratio', round(float(latest_flow['Sales Flow Ratio (%)']), 1), ''])
                else:
                    ed.append(['No flow data', '', '', '', '', '', ''])

                pd.DataFrame(ed).to_excel(writer, sheet_name='Executive Dashboard', index=False, header=False)

                df_ob_exp = df_ob[['Bulan','TC','JS','Dev_Absolut','Dev_Relatif_Pct','Status']].copy()
                df_ob_exp['Dev_Relatif_Pct'] = df_ob_exp['Dev_Relatif_Pct'].round(2)
                df_ob_exp['Dev_Absolut'] = df_ob_exp['Dev_Absolut'].round(0)
                df_ob_exp.columns = ['Month', 'TC (BCM)', 'JS (BCM)', 'Dev (Absolute)', 'Dev (%)', 'Status']
                df_ob_exp.to_excel(writer, sheet_name='OB Analysis', index=False, startrow=2)

                df_ch_exp = df_ch_cm.dropna(subset=['CH_WB','TWB_CH']).copy()
                ch_cols = [c for c in ['Date','Port_Darat','Port_Laut','Port_Total','CH_WB','TWB_CH','Dev_CH_Relatif_Pct','Status_CH'] if c in df_ch_exp.columns]
                df_ch_exp = df_ch_exp[ch_cols].copy()
                if 'Dev_CH_Relatif_Pct' in df_ch_exp.columns:
                    df_ch_exp['Dev_CH_Relatif_Pct'] = df_ch_exp['Dev_CH_Relatif_Pct'].round(2)
                if 'Date' in df_ch_exp.columns:
                    df_ch_exp['Date'] = pd.to_datetime(df_ch_exp['Date']).dt.strftime('%Y-%m-%d')
                col_map_ch = {'Date':'Date','Port_Darat':'Port Darat','Port_Laut':'Port Laut','Port_Total':'Port Total','CH_WB':'WB Target','TWB_CH':'TWB Actual','Dev_CH_Relatif_Pct':'Dev (%)','Status_CH':'Status'}
                df_ch_exp.columns = [col_map_ch.get(c, c) for c in ch_cols]
                df_ch_exp.to_excel(writer, sheet_name='CH Analysis', index=False, startrow=2)

                df_cm_exp = df_ch_cm.dropna(subset=['CM_WB','TWB_CM']).copy()
                cm_cols = [c for c in ['Date','CPP_Raw','CPP_Product','CPP_Total','Sales','CM_WB','TWB_CM','Dev_CM_Relatif_Pct','Status_CM'] if c in df_cm_exp.columns]
                df_cm_exp = df_cm_exp[cm_cols].copy()
                if 'Dev_CM_Relatif_Pct' in df_cm_exp.columns:
                    df_cm_exp['Dev_CM_Relatif_Pct'] = df_cm_exp['Dev_CM_Relatif_Pct'].round(2)
                if 'Date' in df_cm_exp.columns:
                    df_cm_exp['Date'] = pd.to_datetime(df_cm_exp['Date']).dt.strftime('%Y-%m-%d')
                col_map_cm = {'Date':'Date','CPP_Raw':'CPP Raw','CPP_Product':'CPP Product','CPP_Total':'CPP Total','Sales':'Sales','CM_WB':'WB Target','TWB_CM':'TWB Actual','Dev_CM_Relatif_Pct':'Dev (%)','Status_CM':'Status'}
                df_cm_exp.columns = [col_map_cm.get(c, c) for c in cm_cols]
                df_cm_exp.to_excel(writer, sheet_name='CM Analysis', index=False, startrow=2)

                if flow is not None and len(flow) > 0:
                    df_fl = flow.copy()
                    fl_cols = [c for c in [
                        'Date','CM TWB','Delta CPP Stock','CH TWB','Deviation CPP',
                        'Delta Port Stock','Sales','Deviation Port',
                        'CH Flow Ratio (%)','Sales Flow Ratio (%)','Net Deviation'
                    ] if c in df_fl.columns]
                    df_fl = df_fl[fl_cols].copy()
                    if 'Date' in df_fl.columns:
                        df_fl['Date'] = pd.to_datetime(df_fl['Date']).dt.strftime('%Y-%m-%d')
                    for ec in ['CH Flow Ratio (%)','Sales Flow Ratio (%)']:
                        if ec in df_fl.columns:
                            df_fl[ec] = df_fl[ec].round(1)
                    for lc in ['Deviation CPP','Deviation Port','Net Deviation','Delta CPP Stock','Delta Port Stock']:
                        if lc in df_fl.columns:
                            df_fl[lc] = df_fl[lc].round(0)
                    rename_fl = {
                        'CM TWB':'CM TWB',
                        'Delta CPP Stock':'Δ CPP Stock',
                        'CH TWB':'CH TWB',
                        'Deviation CPP':'Dev CPP',
                        'Delta Port Stock':'Δ Port Stock',
                        'Sales':'Sales',
                        'Deviation Port':'Dev Port',
                        'CH Flow Ratio (%)':'CH Ratio (%)',
                        'Sales Flow Ratio (%)':'Sales Ratio (%)',
                        'Net Deviation':'Net Deviation'
                    }
                    df_fl = df_fl.rename(columns=rename_fl)
                    df_fl.to_excel(writer, sheet_name='Material Flow', index=False, startrow=2)

                df_mx = matrix.copy()
                df_mx.columns = ['Stage', 'Normal', 'Caution', 'Critical', 'Total', 'Avg Dev (%)']
                df_mx['Avg Dev (%)'] = df_mx['Avg Dev (%)'].round(2)
                total_row = pd.DataFrame([{
                    'Stage': 'TOTAL',
                    'Normal': df_mx['Normal'].sum(),
                    'Caution': df_mx['Caution'].sum(),
                    'Critical': df_mx['Critical'].sum(),
                    'Total': df_mx['Total'].sum(),
                    'Avg Dev (%)': round(df_mx['Avg Dev (%)'].mean(), 2)
                }])
                df_mx = pd.concat([df_mx, total_row], ignore_index=True)
                df_mx.to_excel(writer, sheet_name='Performance Matrix', index=False, startrow=2)

                if not top_alerts.empty:
                    top_alerts.to_excel(writer, sheet_name='Top Exceptions', index=False, startrow=2)

            excel_buffer.seek(0)
            st.download_button(
                label="Download Excel Report",
                data=excel_buffer.getvalue(),
                file_name=f"KPP_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='dl_excel',
                use_container_width=True
            )

        with exp2:
            st.markdown("""
            <div style="background:rgba(34,197,94,0.06);padding:10px 12px;border-radius:8px;
                border-left:3px solid #22C55E;margin-bottom:8px;">
                <div style="font-size:0.82rem;font-weight:700;color:#e5e7eb;"> HTML Report</div>
                <div style="font-size:0.72rem;color:#94A3B8;margin-top:2px;">Executive summary + exceptions + focused charts</div>
            </div>
            """, unsafe_allow_html=True)

            fig_stage = make_fig_stage_summary(chart_theme_st)
            fig_dev = make_fig_dev_timeline(chart_theme_st)
            fig_flow = make_fig_latest_flow(chart_theme_st)
            fig_ratio = make_fig_ratio_trend(chart_theme_st)

            f_stage = fig_stage.to_html(full_html=False, include_plotlyjs=False)
            f_dev = fig_dev.to_html(full_html=False, include_plotlyjs=False)
            f_flow = fig_flow.to_html(full_html=False, include_plotlyjs=False)
            f_ratio = fig_ratio.to_html(full_html=False, include_plotlyjs=False)

            findings_html = "".join(
                f'<li><span class="bullet-dot"></span><span>{item}</span></li>'
                for item in key_findings
            )

            actions_html = "".join(
                f'<li><span class="bullet-dot"></span><span>{item}</span></li>'
                for item in recommended_actions
            )

            if top_alerts.empty:
                alert_rows_html = """
                <tr>
                    <td colspan="5" style="padding:16px;text-align:center;color:#94A3B8;">No critical / caution periods found.</td>
                </tr>
                """
            else:
                alert_rows_html = ""
                for _, r in top_alerts.head(8).iterrows():
                    status_color = "#EF4444" if r["Status"] == "Critical" else "#F59E0B"
                    badge_bg = "rgba(239,68,68,0.14)" if r["Status"] == "Critical" else "rgba(245,158,11,0.14)"
                    alert_rows_html += f"""
                    <tr>
                        <td>{r["Stage"]}</td>
                        <td>{r["Period"]}</td>
                        <td style="font-weight:700;color:{status_color};">{r["DeviationLabel"]}</td>
                        <td><span class="status-badge" style="background:{badge_bg};color:{status_color};">{r["Status"]}</span></td>
                        <td>{r["Metric"]}</td>
                    </tr>
                    """

            latest_cards_html = ""
            if latest_flow is not None:
                latest_cards = [
                    ("CM TWB", f'{float(latest_flow["CM TWB"]):,.0f} ton', "#60A5FA"),
                    ("CH TWB", f'{float(latest_flow["CH TWB"]):,.0f} ton', "#34D399"),
                    ("Sales", f'{float(latest_flow["Sales"]):,.0f} ton', "#A78BFA"),
                    ("Δ CPP", f'{float(latest_flow["Delta CPP Stock"]):,.0f} ton', "#F59E0B"),
                    ("Δ Port", f'{float(latest_flow["Delta Port Stock"]):,.0f} ton', "#F59E0B"),
                    ("Dev CPP", f'{float(latest_flow["Deviation CPP"]):,.0f} ton', "#EF4444"),
                    ("Dev Port", f'{float(latest_flow["Deviation Port"]):,.0f} ton', "#EF4444"),
                    ("CH Ratio", f'{float(latest_flow["CH Flow Ratio (%)"]):.1f}%', "#22C55E"),
                    ("Sales Ratio", f'{float(latest_flow["Sales Flow Ratio (%)"]):.1f}%', "#60A5FA"),
                ]
                latest_cards_html = "".join(
                    f'<div class="mini-card"><div class="mini-label">{lb}</div><div class="mini-value" style="color:{cl};">{vl}</div></div>'
                    for lb, vl, cl in latest_cards
                )

            html_report = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>KPP Executive Deviation Report {report_date}</title>
<script src="https://cdn.plot.ly/plotly-2.35.0.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:Inter,sans-serif;background:#0F172A;color:#E2E8F0}}
.wrap{{max-width:1440px;margin:0 auto;padding:28px}}
.hero{{background:linear-gradient(135deg,#0F172A,#172554);border:1px solid rgba(148,163,184,0.12);border-radius:18px;padding:28px 30px;margin-bottom:22px}}
.hero h1{{font-size:2rem;font-weight:900;color:#F8FAFC;margin-bottom:6px}}
.hero .sub{{color:#22C55E;font-size:0.95rem;font-weight:700;letter-spacing:.04em;text-transform:uppercase}}
.hero .meta{{margin-top:10px;color:#94A3B8;font-size:0.9rem}}

.kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:22px}}
.kpi-card{{background:#162235;border:1px solid rgba(148,163,184,0.12);border-radius:16px;padding:20px 18px;position:relative;overflow:hidden}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:var(--ac)}}
.kpi-label{{font-size:0.78rem;color:#94A3B8;text-transform:uppercase;letter-spacing:.08em;font-weight:700}}
.kpi-value{{font-size:2rem;font-weight:900;color:#F8FAFC;margin-top:10px}}
.kpi-status{{font-size:0.9rem;font-weight:700;margin-top:6px}}
.kpi-note{{font-size:0.78rem;color:#64748B;margin-top:8px}}

.grid-2{{display:grid;grid-template-columns:1.05fr .95fr;gap:16px;margin-bottom:18px}}
.grid-2-eq{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:18px}}
.panel{{background:#162235;border:1px solid rgba(148,163,184,0.12);border-radius:16px;padding:18px}}
.panel h2{{font-size:1.05rem;color:#F8FAFC;font-weight:800;margin-bottom:14px}}
.panel h3{{font-size:0.86rem;color:#94A3B8;font-weight:700;text-transform:uppercase;letter-spacing:.08em;margin-bottom:10px}}

.finding-list,.action-list{{list-style:none;display:flex;flex-direction:column;gap:12px}}
.finding-list li,.action-list li{{display:flex;gap:12px;align-items:flex-start;padding:12px 12px;background:rgba(15,23,42,0.55);border-radius:12px;border:1px solid rgba(148,163,184,0.08)}}
.bullet-dot{{width:10px;height:10px;border-radius:999px;background:#22C55E;flex:0 0 auto;margin-top:6px}}

.alert-table{{width:100%;border-collapse:collapse}}
.alert-table th{{text-align:left;font-size:0.72rem;text-transform:uppercase;color:#94A3B8;padding:10px 8px;border-bottom:1px solid rgba(148,163,184,0.12)}}
.alert-table td{{padding:11px 8px;border-bottom:1px solid rgba(148,163,184,0.08);font-size:0.9rem;color:#E2E8F0}}
.status-badge{{display:inline-flex;padding:4px 10px;border-radius:999px;font-size:0.75rem;font-weight:800}}

.mini-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}}
.mini-card{{background:rgba(15,23,42,0.55);border:1px solid rgba(148,163,184,0.08);border-radius:12px;padding:12px}}
.mini-label{{font-size:0.72rem;color:#94A3B8;text-transform:uppercase;letter-spacing:.06em;font-weight:700}}
.mini-value{{font-size:1.05rem;font-weight:800;margin-top:6px}}

.chart-panel{{padding:12px 12px 4px 12px}}
.chart-full{{margin-bottom:18px}}
.footer{{margin-top:16px;padding-top:18px;border-top:1px solid rgba(148,163,184,0.12);text-align:center;color:#64748B;font-size:0.82rem}}

@media (max-width: 1100px) {{
  .kpi-grid{{grid-template-columns:repeat(2,1fr)}}
  .grid-2,.grid-2-eq{{grid-template-columns:1fr}}
  .mini-grid{{grid-template-columns:repeat(2,1fr)}}
}}
@media (max-width: 700px) {{
  .kpi-grid{{grid-template-columns:1fr}}
  .mini-grid{{grid-template-columns:1fr}}
}}
</style>
</head>
<body>
<div class="wrap">

<div class="hero">
    <div class="sub">PT Kalimantan Prima Persada</div>
    <h1>Executive Deviation & Reconciliation Report</h1>
    <div class="meta">Generated: {report_date} · Latest flow snapshot: {latest_period_label}</div>
</div>

<div class="kpi-grid">
    <div class="kpi-card" style="--ac:{ob_clr}">
        <div class="kpi-label">Overburden (OB)</div>
        <div class="kpi-value">{ob_disp_dev:.2f}%</div>
        <div class="kpi-status" style="color:{ob_clr};">{ob_lbl2}</div>
        <div class="kpi-note">{stage_meta["OB"]["critical"]} critical · {stage_meta["OB"]["total"]} periods</div>
    </div>
    <div class="kpi-card" style="--ac:{ch_clr}">
        <div class="kpi-label">Coal Hauling (CH)</div>
        <div class="kpi-value">{ch_disp_dev:.2f}%</div>
        <div class="kpi-status" style="color:{ch_clr};">{ch_lbl2}</div>
        <div class="kpi-note">{stage_meta["CH"]["critical"]} critical · {stage_meta["CH"]["total"]} periods</div>
    </div>
    <div class="kpi-card" style="--ac:{cm_clr}">
        <div class="kpi-label">Coal Mining (CM)</div>
        <div class="kpi-value">{cm_disp_dev:.2f}%</div>
        <div class="kpi-status" style="color:{cm_clr};">{cm_lbl2}</div>
        <div class="kpi-note">{stage_meta["CM"]["critical"]} critical · {stage_meta["CM"]["total"]} periods</div>
    </div>
    <div class="kpi-card" style="--ac:{'#22C55E' if perf >= 70 else '#F59E0B' if perf >= 50 else '#EF4444'}">
        <div class="kpi-label">Overall Performance</div>
        <div class="kpi-value">{perf:.1f}%</div>
        <div class="kpi-status" style="color:{'#22C55E' if perf >= 70 else '#F59E0B' if perf >= 50 else '#EF4444'};">{'Good' if perf >= 70 else 'Fair' if perf >= 50 else 'Poor'}</div>
        <div class="kpi-note">CH Ratio {disp_ch_ratio:.1f}% · Sales Ratio {disp_sales_ratio:.1f}%</div>
    </div>
</div>

<div class="grid-2">
    <div class="panel">
        <h2>Key Findings</h2>
        <ul class="finding-list">{findings_html}</ul>
    </div>
    <div class="panel">
        <h2>Top Critical / Caution Periods</h2>
        <table class="alert-table">
            <thead>
                <tr>
                    <th>Stage</th>
                    <th>Period</th>
                    <th>Deviation</th>
                    <th>Status</th>
                    <th>Metric</th>
                </tr>
            </thead>
            <tbody>
                {alert_rows_html}
            </tbody>
        </table>
    </div>
</div>

<div class="grid-2-eq">
    <div class="panel chart-panel">
        <h2>Stage Summary</h2>
        {f_stage}
    </div>
    <div class="panel chart-panel">
        <h2>Latest Reconciliation Snapshot</h2>
        {f_flow}
    </div>
</div>

<div class="panel chart-panel chart-full">
    <h2>Deviation Trend</h2>
    {f_dev}
</div>

<div class="grid-2">
    <div class="panel chart-panel">
        <h2>Flow Ratio Trend</h2>
        {f_ratio}
    </div>
    <div class="panel">
        <h2>Latest Snapshot Metrics</h2>
        <div class="mini-grid">
            {latest_cards_html}
        </div>
        <div style="height:14px"></div>
        <h2>Recommended Actions</h2>
        <ul class="action-list">{actions_html}</ul>
    </div>
</div>

<div class="footer">
    PT Kalimantan Prima Persada — Mining Volume Deviation Monitoring System<br>
    Report generated: {report_date}
</div>

</div>
</body>
</html>
"""
            html_bytes = html_report.encode("utf-8")
            st.download_button(
                label="Download HTML",
                data=html_bytes,
                file_name=f"KPP_Executive_Report_{datetime.now().strftime('%Y%m%d')}.html",
                mime="text/html",
                key="dl_html",
                use_container_width=True
            )

        with exp3:
            st.markdown("""
            <div style="background:rgba(22,101,52,0.08);padding:10px 12px;border-radius:8px;
                border-left:3px solid #16A34A;margin-bottom:8px;">
                <div style="font-size:0.82rem;font-weight:700;color:#e5e7eb;"> PNG Report</div>
                <div style="font-size:0.72rem;color:#94A3B8;margin-top:2px;">
                    One-page executive summary • larger typography • focused charts</div>
            </div>
            """, unsafe_allow_html=True)

            try:
                import kaleido
                kaleido_ok = True
            except ImportError:
                kaleido_ok = False
                st.warning("Library 'kaleido' tidak terinstall. Tambahkan kaleido==0.2.1 ke requirements.txt")

            if kaleido_ok:
                if st.button(" Generate PNG", key="gen_png", use_container_width=True):
                    with st.spinner("Generating executive PNG report..."):
                        try:
                            from PIL import Image as PILImage
                            from PIL import ImageDraw, ImageFont

                            png_theme = dict(
                                font=dict(family="Inter, Arial, sans-serif", color="#E2E8F0", size=12),
                                plot_bgcolor="#182332",
                                paper_bgcolor="#182332",
                                margin=dict(l=55, r=35, t=70, b=50),
                                hovermode="x unified"
                            )

                            fig_stage_png = make_fig_stage_summary(png_theme)
                            fig_dev_png = make_fig_dev_timeline(png_theme)
                            fig_flow_png = make_fig_latest_flow(png_theme)

                            BGCOLOR = (13, 20, 33)
                            PANEL = (22, 34, 53)
                            PANEL_ALT = (18, 29, 45)
                            BORDER = (44, 60, 85)
                            WHITE = (241, 245, 249)
                            TEXT = (226, 232, 240)
                            MUTED = (148, 163, 184)
                            DARK = (100, 116, 139)
                            GREEN = (34, 197, 94)
                            BLUE = (96, 165, 250)
                            PURPLE = (167, 139, 250)
                            YELLOW = (245, 158, 11)
                            RED = (239, 68, 68)

                            W = 2200
                            PAD = 42
                            GAP = 18
                            HEADER_H = 130
                            KPI_H = 150
                            INFO_H = 330
                            CHART_H = 430
                            FOOT_H = 70
                            CONTENT_W = W - PAD * 2

                            H = PAD + HEADER_H + GAP + KPI_H + GAP + INFO_H + GAP + CHART_H + GAP + CHART_H + GAP + FOOT_H + PAD
                            canvas = PILImage.new("RGB", (W, H), BGCOLOR)
                            draw = ImageDraw.Draw(canvas)

                            try:
                                font_title = ImageFont.truetype("arial.ttf", 44)
                                font_sub = ImageFont.truetype("arial.ttf", 20)
                                font_h2 = ImageFont.truetype("arial.ttf", 24)
                                font_body = ImageFont.truetype("arial.ttf", 18)
                                font_small = ImageFont.truetype("arial.ttf", 14)
                                font_kpi_value = ImageFont.truetype("arial.ttf", 34)
                                font_kpi_label = ImageFont.truetype("arial.ttf", 15)
                                font_kpi_status = ImageFont.truetype("arial.ttf", 18)
                                font_table_head = ImageFont.truetype("arial.ttf", 15)
                                font_table_cell = ImageFont.truetype("arial.ttf", 16)
                            except Exception:
                                font_title = ImageFont.load_default()
                                font_sub = font_h2 = font_body = font_small = font_title
                                font_kpi_value = font_kpi_label = font_kpi_status = font_title
                                font_table_head = font_table_cell = font_title

                            def rounded_box(x, y, w, h, fill, outline=BORDER, radius=18, width=1):
                                draw.rounded_rectangle([x, y, x+w, y+h], radius=radius, fill=fill, outline=outline, width=width)

                            def draw_multiline(text, x, y, max_width, font, fill, line_gap=6):
                                lines = []
                                words = str(text).split()
                                current = ""
                                for word in words:
                                    test = word if not current else current + " " + word
                                    bbox = draw.textbbox((0, 0), test, font=font)
                                    if bbox[2] - bbox[0] <= max_width:
                                        current = test
                                    else:
                                        if current:
                                            lines.append(current)
                                        current = word
                                if current:
                                    lines.append(current)

                                cy = y
                                for line in lines:
                                    draw.text((x, cy), line, font=font, fill=fill)
                                    bbox = draw.textbbox((0, 0), line, font=font)
                                    cy += (bbox[3] - bbox[1]) + line_gap
                                return cy

                            def fig_to_img(fig, width_px, height_px):
                                fig.update_layout(width=width_px, height=height_px)
                                raw = fig.to_image(format="png", engine="kaleido", scale=2)
                                img = PILImage.open(io.BytesIO(raw)).convert("RGB")
                                return img.resize((width_px, height_px), PILImage.LANCZOS)

                            def kpi_card(x, y, w, h, label, value, status, note, accent_hex):
                                rounded_box(x, y, w, h, PANEL)
                                accent = tuple(int(accent_hex.strip("#")[i:i+2], 16) for i in (0, 2, 4))
                                draw.rounded_rectangle([x+12, y, x+w-12, y+5], radius=3, fill=accent)
                                draw.text((x+18, y+18), label.upper(), font=font_kpi_label, fill=MUTED)
                                draw.text((x+18, y+52), value, font=font_kpi_value, fill=WHITE)
                                draw.text((x+18, y+98), status, font=font_kpi_status, fill=accent)
                                draw.text((x+18, y+123), note, font=font_small, fill=DARK)

                            y = PAD
                            rounded_box(PAD, y, CONTENT_W, HEADER_H, PANEL)
                            draw.text((PAD + 28, y + 24), "Executive Deviation & Reconciliation Report", font=font_title, fill=WHITE)
                            draw.text((PAD + 28, y + 78), "PT Kalimantan Prima Persada", font=font_sub, fill=GREEN)
                            draw.text((W - PAD - 320, y + 32), f"Generated: {report_date}", font=font_sub, fill=MUTED)
                            draw.text((W - PAD - 320, y + 68), f"Latest snapshot: {latest_period_label}", font=font_sub, fill=MUTED)

                            y += HEADER_H + GAP
                            kpi_w = (CONTENT_W - 3 * GAP) // 4

                            kpi_card(PAD + 0 * (kpi_w + GAP), y, kpi_w, KPI_H, "Overburden (OB)", f"{ob_disp_dev:.2f}%", ob_lbl2, f'{stage_meta["OB"]["critical"]} critical · {stage_meta["OB"]["total"]} periods', ob_clr)
                            kpi_card(PAD + 1 * (kpi_w + GAP), y, kpi_w, KPI_H, "Coal Hauling (CH)", f"{ch_disp_dev:.2f}%", ch_lbl2, f'{stage_meta["CH"]["critical"]} critical · {stage_meta["CH"]["total"]} periods', ch_clr)
                            kpi_card(PAD + 2 * (kpi_w + GAP), y, kpi_w, KPI_H, "Coal Mining (CM)", f"{cm_disp_dev:.2f}%", cm_lbl2, f'{stage_meta["CM"]["critical"]} critical · {stage_meta["CM"]["total"]} periods', cm_clr)
                            perf_accent = "#22C55E" if perf >= 70 else "#F59E0B" if perf >= 50 else "#EF4444"
                            perf_status = "Good" if perf >= 70 else "Fair" if perf >= 50 else "Poor"
                            kpi_card(PAD + 3 * (kpi_w + GAP), y, kpi_w, KPI_H, "Overall Performance", f"{perf:.1f}%", perf_status, f"CH Ratio {disp_ch_ratio:.1f}% · Sales Ratio {disp_sales_ratio:.1f}%", perf_accent)

                            y += KPI_H + GAP
                            left_w = int(CONTENT_W * 0.52)
                            right_w = CONTENT_W - left_w - GAP

                            rounded_box(PAD, y, left_w, INFO_H, PANEL)
                            rounded_box(PAD + left_w + GAP, y, right_w, INFO_H, PANEL)

                            draw.text((PAD + 20, y + 18), "Key Findings", font=font_h2, fill=WHITE)
                            fy = y + 62
                            for item in key_findings:
                                draw.ellipse((PAD + 22, fy + 5, PAD + 32, fy + 15), fill=GREEN)
                                fy = draw_multiline(item, PAD + 42, fy, left_w - 70, font_body, TEXT, line_gap=5) + 14

                            draw.text((PAD + left_w + GAP + 20, y + 18), "Top Critical / Caution Periods", font=font_h2, fill=WHITE)

                            tx = PAD + left_w + GAP + 20
                            ty = y + 58
                            headers = [("Stage", 90), ("Period", 150), ("Dev", 100), ("Status", 110), ("Metric", right_w - 20 - 90 - 150 - 100 - 110 - 30)]
                            cx = tx
                            for head, width in headers:
                                draw.text((cx, ty), head.upper(), font=font_table_head, fill=MUTED)
                                cx += width
                            ty += 28

                            if top_alerts.empty:
                                draw.text((tx, ty + 16), "No critical / caution periods found.", font=font_body, fill=MUTED)
                            else:
                                for _, r in top_alerts.head(6).iterrows():
                                    draw.line((tx, ty - 6, PAD + left_w + GAP + right_w - 20, ty - 6), fill=BORDER, width=1)
                                    cx = tx
                                    row_vals = [
                                        str(r["Stage"]),
                                        str(r["Period"]),
                                        str(r["DeviationLabel"]),
                                        str(r["Status"]),
                                        str(r["Metric"]),
                                    ]
                                    for idx, ((_, width), val) in enumerate(zip(headers, row_vals)):
                                        fill_color = TEXT
                                        if idx == 2:
                                            fill_color = RED if r["Status"] == "Critical" else YELLOW
                                        draw.text((cx, ty), val, font=font_table_cell, fill=fill_color)
                                        cx += width
                                    ty += 34

                            y += INFO_H + GAP
                            chart_w = (CONTENT_W - GAP) // 2
                            chart_stage = fig_to_img(fig_stage_png, chart_w, CHART_H)
                            chart_flow = fig_to_img(fig_flow_png, chart_w, CHART_H)

                            rounded_box(PAD, y, chart_w, CHART_H, PANEL_ALT)
                            rounded_box(PAD + chart_w + GAP, y, chart_w, CHART_H, PANEL_ALT)
                            canvas.paste(chart_stage, (PAD, y))
                            canvas.paste(chart_flow, (PAD + chart_w + GAP, y))

                            y += CHART_H + GAP
                            chart_dev = fig_to_img(fig_dev_png, CONTENT_W, CHART_H)
                            rounded_box(PAD, y, CONTENT_W, CHART_H, PANEL_ALT)
                            canvas.paste(chart_dev, (PAD, y))

                            y += CHART_H + GAP
                            rounded_box(PAD, y, CONTENT_W, FOOT_H, PANEL)
                            snapshot_items = []
                            if latest_flow is not None:
                                snapshot_items = [
                                    ("CM TWB", f'{float(latest_flow["CM TWB"]):,.0f}'),
                                    ("CH TWB", f'{float(latest_flow["CH TWB"]):,.0f}'),
                                    ("Sales", f'{float(latest_flow["Sales"]):,.0f}'),
                                    ("Δ CPP", f'{float(latest_flow["Delta CPP Stock"]):,.0f}'),
                                    ("Δ Port", f'{float(latest_flow["Delta Port Stock"]):,.0f}'),
                                    ("Dev CPP", f'{float(latest_flow["Deviation CPP"]):,.0f}'),
                                    ("Dev Port", f'{float(latest_flow["Deviation Port"]):,.0f}'),
                                    ("CH Ratio", f'{float(latest_flow["CH Flow Ratio (%)"]):.1f}%'),
                                    ("Sales Ratio", f'{float(latest_flow["Sales Flow Ratio (%)"]):.1f}%'),
                                ]
                            else:
                                snapshot_items = [("Snapshot", "No data")]

                            slot_w = CONTENT_W / len(snapshot_items)
                            for i, (lb, vl) in enumerate(snapshot_items):
                                sx = int(PAD + i * slot_w + 10)
                                draw.text((sx, y + 12), lb.upper(), font=font_small, fill=MUTED)
                                draw.text((sx, y + 34), vl, font=font_body, fill=WHITE)

                            buf = io.BytesIO()
                            canvas.save(buf, format="PNG", quality=95, optimize=True)
                            buf.seek(0)

                            st.success(f"✅ PNG generated — {canvas.size[0]}×{canvas.size[1]}px")
                            st.image(canvas, caption="Preview — Executive PNG Report", use_container_width=True)
                            st.download_button(
                                label=" Download PNG Report",
                                data=buf.getvalue(),
                                file_name=f"KPP_Executive_Report_{datetime.now().strftime('%Y%m%d')}.png",
                                mime="image/png",
                                key="dl_png",
                                use_container_width=True
                            )

                        except Exception as e:
                            st.error(f"Error: {str(e)}")
                            import traceback
                            st.code(traceback.format_exc())
                
if __name__ == "__main__":
    main()
