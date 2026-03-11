from __future__ import annotations

import base64
import io
import re
import time
from typing import Any, Dict, List

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from scipy.stats import spearmanr

import mcdm_engine as me

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt
    from docx.enum.section import WD_SECTION
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False

st.set_page_config(
    page_title="MCDM Toolbox — Prof. Dr. Ömer Faruk Rençber",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Space+Mono:wght@400;700&display=swap');

        :root {
            --bg-main: #F0F5FA;
            --bg-sidebar: #E5EDF5;
            --card-bg: #FFFFFF;
            --text-main: #1C1C1E;
            --text-muted: #556070;
            --text-light: #7A8A9A;
            --border: #C8D8E8;
            --border-strong: #B0C8DC;
            --accent: #1F4A73;
            --accent-hover: #173754;
            --accent-soft: #EAF0F5;
            --button-soft: #87CEEB;
            --button-soft-hover: #5BB8E0;
            --button-soft-text: #000000;
            --header-bg-from: #1C3252;
            --header-bg-to: #263D5A;
            --header-text: #EEF2FA;
            --success-bg: #DFF0E6;
            --success-text: #1A5C3A;
            --warn-bg: #FEF0CC;
            --warn-text: #7A5018;
            --danger-bg: #FAE0E0;
            --danger-text: #7B1E1E;
        }

        .stApp {
            background:
                radial-gradient(circle at 0% 0%, rgba(135, 206, 235, 0.12) 0%, rgba(135, 206, 235, 0.00) 28%),
                radial-gradient(circle at 100% 0%, rgba(31, 74, 115, 0.08) 0%, rgba(31, 74, 115, 0.00) 26%),
                linear-gradient(180deg, #F6FAFE 0%, var(--bg-main) 50%, #EEF4FA 100%) !important;
            color: var(--text-main) !important;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        .main .block-container {
            padding-top: 0.18rem !important;
            padding-bottom: 2.5rem;
            max-width: 1500px;
        }

        /* ────── Streamlit'in kendi header'ını gizle ────── */
        header[data-testid="stHeader"] { display: none !important; }
        .stDeployButton { display: none !important; }
        #MainMenu { display: none !important; }
        footer { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
        button[aria-label="Collapse sidebar"] { display: none !important; }
        button[aria-label="Expand sidebar"] { display: none !important; }

        /* ────── GLOBAL HEADER (yalnızca sağ panel) ────── */
        .global-header {
            position: relative;
            width: 100%;
            min-height: 108px;
            background:
                linear-gradient(135deg, rgba(255,255,255,0.06) 0%, rgba(255,255,255,0.0) 42%),
                linear-gradient(120deg, #132742 0%, #1E3A5F 52%, #12253D 100%);
            color: #EEF2FA !important;
            display: flex;
            flex-direction: column;
            justify-content: center;
            gap: 0.35rem;
            padding: 0.9rem 2.2rem 0.8rem 2.2rem;
            margin: 0 0 0.95rem 0;
            border-radius: 0 0 22px 22px;
            border-bottom: 1px solid rgba(184, 154, 92, 0.48);
            box-shadow: 0 10px 28px rgba(5, 13, 24, 0.28);
        }
        .header-title {
            font-size: 1.14rem;
            font-weight: 700;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            color: #E7EDF7 !important;
            line-height: 1.2;
            text-align: center;
            text-shadow: 0 1px 10px rgba(0, 0, 0, 0.18);
        }
        .header-professor {
            width: 100%;
            font-size: 1.78rem;
            font-weight: 700;
            color: #FFFFFF !important;
            font-family: 'Monotype Corsiva', 'Apple Chancery', 'URW Chancery L', 'Brush Script MT', cursive;
            font-style: italic;
            line-height: 1.12;
            text-align: center;
            letter-spacing: 0.01em;
            text-shadow: 0 3px 18px rgba(0, 0, 0, 0.24);
        }
        .header-meta {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 1rem;
            margin-top: 0.15rem;
            padding-top: 0.45rem;
            border-top: 1px solid rgba(255,255,255,0.14);
        }
        .header-dedication {
            font-size: 0.82rem;
            font-style: italic;
            color: #D6E0EE !important;
            letter-spacing: 0.03em;
        }
        .header-url {
            font-size: 0.82rem;
            font-weight: 600;
            color: #DCE5F2 !important;
            letter-spacing: 0.05em;
        }
        .header-url a {
            color: #DCE5F2 !important;
            text-decoration: none;
            transition: color 0.15s ease;
        }
        .header-url a:hover {
            color: #FFFFFF !important;
        }

        /* ────── SIDEBAR BRANDING ────── */
        .sidebar-brand {
            text-align: center;
            padding: 0 0.35rem 0.5rem 0.35rem;
            border-bottom: 1px solid var(--border-strong);
            margin-bottom: 0.5rem;
            margin-top: -0.75rem;
        }
        .sidebar-brand-logo {
            display: block;
            width: 100%;
            max-width: 282px;
            margin: 0 auto;
            height: auto;
            filter: drop-shadow(0 10px 18px rgba(0, 0, 0, 0.16));
        }
        .sidebar-brand-name {
            margin-top: 0.45rem;
            font-size: 0.88rem;
            font-weight: 700;
            color: #1C3252 !important;
            font-family: "Georgia", "Times New Roman", serif;
            font-style: italic;
            letter-spacing: 0.01em;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .sidebar-brand-title {
            font-size: 1.55rem;
            font-weight: 900;
            letter-spacing: 2px;
            color: #1C3252 !important;
            line-height: 1.15;
            font-family: 'Inter', sans-serif;
        }
        .sidebar-brand-sub {
            font-size: 0.68rem;
            font-weight: 600;
            letter-spacing: 0.4px;
            color: #5C5650 !important;
            text-transform: uppercase;
            margin-top: 0.1rem;
        }

        /* ────── SIDEBAR ────── */
        section[data-testid="stSidebar"],
        section[data-testid="stSidebar"] > div,
        aside[data-testid="stSidebar"],
        aside[data-testid="stSidebar"] > div {
            background-color: var(--bg-sidebar) !important;
            border-right: 1px solid var(--border-strong) !important;
        }
        section[data-testid="stSidebar"],
        aside[data-testid="stSidebar"] {
            min-width: 315px !important;
            width: 315px !important;
        }
        section[data-testid="stSidebar"][aria-expanded="false"],
        aside[data-testid="stSidebar"][aria-expanded="false"] {
            min-width: 315px !important;
            max-width: 315px !important;
            width: 315px !important;
            transform: translateX(0) !important;
            margin-left: 0 !important;
        }
        section[data-testid="stSidebar"][aria-expanded="false"] > div,
        aside[data-testid="stSidebar"][aria-expanded="false"] > div {
            min-width: 315px !important;
            max-width: 315px !important;
            width: 315px !important;
        }
        section[data-testid="stSidebar"] *,
        aside[data-testid="stSidebar"] *,
        [data-testid="stSidebar"] * {
            color: var(--text-main) !important;
        }
        /* Sidebar global font sizes — compact & readable */
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] .stRadio label,
        section[data-testid="stSidebar"] .stCheckbox label,
        section[data-testid="stSidebar"] .stSelectbox label,
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] span {
            font-size: 0.75rem !important;
            line-height: 1.25 !important;
        }
        section[data-testid="stSidebar"] .stButton > button {
            font-size: 0.73rem !important;
            padding: 0.25rem 0.6rem !important;
            min-height: 1.8rem !important;
        }
        section[data-testid="stSidebar"] .stExpander {
            margin-bottom: 0.3rem !important;
        }
        section[data-testid="stSidebar"] [data-testid="stExpander"] summary {
            padding: 0.25rem 0.5rem !important;
            min-height: 1.6rem !important;
        }
        section[data-testid="stSidebar"] .element-container {
            margin-bottom: 0.15rem !important;
        }

        /* Sidebar section labels */
        .sb-section-label {
            font-size: 0.7rem;
            font-weight: 700;
            letter-spacing: 0.5px;
            text-transform: uppercase;
            color: #5C5650 !important;
            margin: 0.6rem 0 0.25rem 0;
            border-bottom: 1px solid var(--border);
            padding-bottom: 0.15rem;
        }

        /* Compact guide/method expanders in sidebar */
        section[data-testid="stSidebar"] [data-testid="stExpander"] summary {
            font-size: 0.73rem !important;
            padding: 0.3rem 0.5rem !important;
            min-height: 1.8rem !important;
            background: rgba(90,141,200,0.08) !important;
            border-radius: 6px !important;
            border: 1px solid var(--border) !important;
        }
        section[data-testid="stSidebar"] [data-testid="stExpander"] summary p {
            font-size: 0.73rem !important;
        }

        /* Sidebar footer */
        .sidebar-footer {
            text-align: center;
            padding: 0.6rem 0.4rem 0.4rem 0.4rem;
            margin-top: 0.5rem;
            border-top: 1px solid var(--border-strong);
            font-size: 0.68rem;
            font-weight: 600;
            color: #5C5650 !important;
            letter-spacing: 0.2px;
        }

        /* ────── CARDS ────── */
        .section-card {
            background: var(--card-bg);
            border-radius: 10px;
            padding: 1rem 1.2rem;
            border: 1px solid var(--border);
            margin-bottom: 0.8rem;
        }

        /* ────── BUTTONS — açık gökyüzü mavisi, siyah yazı ────── */
        .stButton > button {
            border-radius: 8px !important;
            font-weight: 600 !important;
            border: 1px solid rgba(70, 150, 210, 0.40) !important;
            background: #87CEEB !important;
            color: #000000 !important;
            padding: 0.44rem 1rem !important;
            font-size: 0.875rem !important;
            letter-spacing: 0.15px !important;
            box-shadow: 0 2px 8px rgba(70, 140, 200, 0.15) !important;
            transition: background 0.12s ease, box-shadow 0.12s ease, border-color 0.12s ease !important;
        }
        [data-testid="stDownloadButton"] > button {
            border-radius: 8px !important;
            font-weight: 700 !important;
            border: 1px solid rgba(18, 39, 62, 0.18) !important;
            background: linear-gradient(180deg, #274F78 0%, var(--accent) 100%) !important;
            color: #FFFFFF !important;
            padding: 0.44rem 1rem !important;
            font-size: 0.875rem !important;
            letter-spacing: 0.15px !important;
            box-shadow: 0 8px 18px rgba(21, 43, 66, 0.16) !important;
            transition: background 0.12s ease, box-shadow 0.12s ease, border-color 0.12s ease !important;
        }
        .stButton > button:hover,
        .stButton > button[kind="secondary"]:hover {
            background: #5BB8E0 !important;
            color: #000000 !important;
            border-color: rgba(70, 150, 210, 0.60) !important;
            box-shadow: 0 4px 12px rgba(70, 140, 200, 0.25) !important;
        }
        [data-testid="stDownloadButton"] > button:hover {
            background: linear-gradient(180deg, #1E4367 0%, var(--accent-hover) 100%) !important;
            color: #FFFFFF !important;
            box-shadow: 0 10px 20px rgba(21, 43, 66, 0.22) !important;
            border-color: rgba(18, 39, 62, 0.26) !important;
        }
        .stButton > button:focus,
        .stButton > button:active,
        [data-testid="stDownloadButton"] > button:focus,
        [data-testid="stDownloadButton"] > button:active {
            box-shadow: 0 0 0 0.16rem rgba(70, 150, 210, 0.30) !important;
            outline: none !important;
        }
        .stButton > button[kind="primary"] {
            background: linear-gradient(180deg, #274F78 0%, var(--accent) 100%) !important;
            color: #FFFFFF !important;
            border: 1px solid rgba(18, 39, 62, 0.18) !important;
            box-shadow: 0 8px 18px rgba(21, 43, 66, 0.16) !important;
        }
        .stButton > button[kind="primary"]:hover {
            background: linear-gradient(180deg, #1E4367 0%, var(--accent-hover) 100%) !important;
            color: #FFFFFF !important;
            border-color: rgba(18, 39, 62, 0.26) !important;
            box-shadow: 0 10px 20px rgba(21, 43, 66, 0.22) !important;
        }

        /* ────── FILE UPLOADER ────── */
        [data-testid="stFileUploaderDropzone"] {
            border: 1.5px dashed var(--border-strong) !important;
            background: var(--accent-soft) !important;
            border-radius: 10px !important;
        }
        [data-testid="stFileUploaderDropzone"] button {
            background: var(--accent) !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 7px !important;
            font-weight: 600 !important;
        }

        /* ────── TABLES — light background ────── */
        [data-testid="stTable"] table {
            background: #FFFFFF !important;
            color: var(--text-main) !important;
            border: 1px solid var(--border) !important;
            border-radius: 8px !important;
        }
        [data-testid="stTable"] thead th {
            background: #EEF2F8 !important;
            color: #1A2A3A !important;
            border: 1px solid var(--border) !important;
            font-weight: 600 !important;
            font-size: 0.84rem !important;
        }
        [data-testid="stTable"] tbody td {
            background: #FFFFFF !important;
            color: var(--text-main) !important;
            border: 1px solid #EDE8DF !important;
        }
        .stTable { color: var(--text-main) !important; }

        /* ────── EXPANDERS ────── */
        [data-testid="stExpander"] {
            border: 1px solid var(--border) !important;
            border-radius: 10px !important;
            background: var(--card-bg) !important;
            margin-bottom: 0.65rem !important;
            box-shadow: none !important;
            transition: border-color 0.12s ease !important;
        }
        [data-testid="stExpander"]:hover {
            border-color: var(--accent) !important;
        }
        [data-testid="stExpander"] details {
            border: none !important;
            border-radius: 10px !important;
            background: var(--card-bg) !important;
        }
        [data-testid="stExpander"] summary {
            border-radius: 10px !important;
            font-weight: 600 !important;
            color: var(--text-main) !important;
            font-size: 0.9rem !important;
            padding: 0.65rem 0.8rem !important;
        }

        /* ────── TABS ────── */
        .stTabs [data-baseweb="tab-list"] {
            gap: 3px;
            background: #D6E8F5;
            border-radius: 10px;
            border: 1px solid var(--border-strong);
            padding: 4px;
        }
        .stTabs [data-baseweb="tab"] {
            background: transparent;
            border-radius: 8px;
            border: none;
            padding: 0.38rem 0.82rem;
            font-size: 0.83rem;
            font-weight: 500;
            color: var(--text-main);
            transition: all 0.12s ease;
        }
        .stTabs [data-baseweb="tab"]:hover {
            background: rgba(135, 206, 235, 0.45) !important;
        }
        .stTabs [aria-selected="true"] {
            background: #87CEEB !important;
            color: #000000 !important;
            font-weight: 700 !important;
            border: none !important;
        }

        /* ────── ASSISTANT / INFO BOXES ────── */
        .assistant-box {
            background: var(--accent-soft);
            border: 1px solid var(--border);
            border-left: 3px solid var(--accent);
            border-radius: 10px;
            padding: 1rem 1.1rem;
            margin-bottom: 0.85rem;
        }
        .tab-assistant {
            background: var(--accent-soft);
            border: 1px solid var(--border);
            border-left: 4px solid var(--accent);
            border-radius: 10px;
            padding: 0.9rem 1rem;
            font-size: 0.86rem;
            line-height: 1.68;
            color: var(--text-main);
        }
        .tab-assistant strong { color: #1A2A3A; }

        /* ────── COMMENTARY BOX ────── */
        .commentary-box {
            background: #F0F6E8;
            border: 1px solid #C8D8B4;
            border-left: 3px solid #5A7A3A;
            border-radius: 8px;
            padding: 0.65rem 0.9rem;
            font-size: 0.84rem;
            line-height: 1.62;
            color: var(--text-main);
            margin-top: 0.4rem;
        }

        /* ────── LAYER CARDS ────── */
        .layer-card {
            border-radius: 8px;
            padding: 0.75rem 1rem;
            margin-bottom: 0.45rem;
            border: 1px solid var(--border);
            line-height: 1.52;
        }
        .layer-desc { background: #EDF4FF; }
        .layer-analytic { background: #EDF7F0; }
        .layer-norm { background: #FFF8EC; }
        .layer-label {
            font-size: 0.69rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.4px;
            margin-bottom: 0.25rem;
        }
        .layer-text { font-size: 0.87rem; color: var(--text-main); margin: 0; }

        /* ────── BADGES ────── */
        .badge-benefit { background:#D4EFE2; color:#1A5C40; border-radius:5px; padding:0.1rem 0.35rem; font-size:0.65rem; font-weight:700; }
        .badge-cost    { background:#FAE0E0; color:#7B1E1E; border-radius:5px; padding:0.1rem 0.35rem; font-size:0.65rem; font-weight:700; }

        /* ────── NETWORK ────── */
        .network-wrap {
            background: #FFFFFF;
            border-radius: 10px;
            border: 1px solid var(--border);
            padding: 0.5rem;
        }

        .mc-high  { color: #1A5C3A; font-weight: 700; }
        .mc-mid   { color: #7A5018; font-weight: 700; }
        .mc-low   { color: #7B1E1E; font-weight: 700; }

        /* ────── STEPPER ────── */
        .stepper-wrap {
            background: #FFFFFF;
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 0.55rem 0.75rem;
            margin-bottom: 0.6rem;
            display: flex;
            flex-direction: column;
            gap: 0;
        }
        .step-item { display:flex; align-items:center; gap:0.4rem; margin:0.2rem 0; font-size:0.74rem; color: var(--text-main); }
        .step-dot { width:9px; height:9px; border-radius:50%; display:inline-block; }
        .step-done { background:#3D7A5E; }
        .step-pending { background:#BFBAB4; }
        .step-active { background:var(--accent); box-shadow:0 0 0 3px rgba(61,100,145,0.2); }

        /* ────── KPI STRIP ────── */
        .kpi-strip { background:#FFFFFF; border:1px solid var(--border); border-radius:10px; padding:0.8rem 1rem; margin:0.35rem 0 1rem 0; }
        .kpi-grid { display:grid; grid-template-columns:repeat(5, minmax(0,1fr)); gap:0.55rem; }
        .kpi-item { background:var(--card-bg); border:1px solid var(--border); border-radius:8px; padding:0.5rem 0.7rem; }
        .kpi-label { font-size:0.70rem; color: var(--text-muted); text-transform:uppercase; letter-spacing:.35px; }
        .kpi-value { font-size:0.92rem; font-weight:700; color: var(--text-main); margin-top:0.1rem; }

        /* ────── ASSISTANT GRID ────── */
        .assistant-grid { display:grid; grid-template-columns:repeat(3, minmax(0,1fr)); gap:0.6rem; margin-bottom:0.6rem; }
        .assistant-card2 { background:#FFFFFF; border:1px solid var(--border); border-radius:10px; padding:0.65rem 0.75rem; }
        .assistant-title2 { font-size:0.72rem; font-weight:700; color: var(--text-main); text-transform:uppercase; letter-spacing:.3px; margin-bottom:0.25rem; }
        .assistant-body2 { font-size:0.82rem; color: var(--text-main); line-height:1.5; }

        /* ────── DIAG BADGES ────── */
        .diag-badge { display:inline-block; font-size:0.68rem; font-weight:700; border-radius:999px; padding:0.15rem 0.45rem; margin-left:0.3rem; }
        .diag-badge-good { background:#D4EFE2; color:#1A5C40; }
        .diag-badge-mid  { background:#FEF0CC; color:#7A5018; }
        .diag-badge-bad  { background:#FAE0E0; color:#7B1E1E; }

        .sidebar-small-note { font-size: 0.80rem; line-height: 1.5; color: var(--text-main); }

        h1, h2, h3, h4, h5, h6 { color: var(--text-main) !important; }
        p, span, label, div { color: var(--text-main); }

        /* ────── RESPONSIVE ────── */
        @media (max-width: 980px) {
            .main .block-container { padding-top: 1rem !important; }
            .global-header {
                padding: 0.8rem 1rem;
                min-height: 132px;
            }
            .header-title { font-size: 0.9rem; }
            .header-professor { font-size: 1.26rem; }
            .header-meta {
                flex-direction: column;
                align-items: flex-start;
                gap: 0.28rem;
            }
            .kpi-grid { grid-template-columns:repeat(2, minmax(0,1fr)); }
            .assistant-grid { grid-template-columns:1fr; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

def init_state() -> None:
    defaults = {
        "analysis_result": None,
        "clean_data": None,
        "raw_data": None,
        "report_docx": None,
        "editor_df": None,
        "data_source_id": None,
        "prep_done": False,
        "alt_names": {},        
        "crit_dir": {},         
        "crit_include": {},     
        "weight_method_pref": None,
        "ranking_prefs": [],
        "ui_lang": "TR",
        "analysis_scope": "single",
        "panel_year_column": None,
        "panel_results": None,
        "panel_years": [],
        "panel_view_choice": None,
        "step1_done": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_state()

def tt(tr_text: str, en_text: str) -> str:
    return en_text if st.session_state.get("ui_lang", "TR") == "EN" else tr_text

def _method_help_text(method_name: str) -> str:
    base_method = str(method_name or "").replace("Fuzzy ", "").strip()
    fallback = {
        "Eşit Ağırlık": {
            "simple": "Eşit ağırlık yaklaşımı, tüm kriterlere aynı başlangıç önemini verir.",
            "academic": "Bu seçim, veri temelli farklılaştırma yapmadan dengeli ve normatif bir başlangıç çerçevesi kurar.",
        },
        "Manuel Ağırlık": {
            "simple": "Manuel ağırlık yaklaşımı, kriter önemini doğrudan kullanıcı tercihine göre tanımlar.",
            "academic": "Bu seçim, karar verici yargısını modele açık biçimde taşıdığı için sonuçların öznel önceliklerle birlikte okunmasını gerektirir.",
        },
    }
    ph = (
        me.METHOD_PHILOSOPHY.get(str(method_name), {})
        or me.METHOD_PHILOSOPHY.get(base_method, {})
        or fallback.get(str(method_name), {})
        or fallback.get(base_method, {})
    )
    simple = ph.get("simple", tt("Seçilen yöntem, problemi kendi değerlendirme mantığıyla okur.", "The selected method reads the problem through its own evaluation logic."))
    academic = ph.get("academic", tt("Bu nedenle sonuçlar, seçilen yöntemin varsayımları içinde yorumlanmalıdır.", "Therefore, results should be interpreted within the assumptions of the selected method."))
    return f"{simple}\n{academic}"

_COL_TR_EN = {
    "Kriter": "Criterion",
    "Ağırlık": "Weight",
    "HamAğırlık": "RawWeight",
    "Alternatif": "Alternative",
    "Skor": "Score",
    "Sıra": "Rank",
    "Yöntem": "Method",
    "BirincilikOranı": "FirstPlaceRate",
    "OrtalamaSıra": "MeanRank",
    "AğırlıkDeğişimi": "WeightChange",
    "YeniAğırlık": "NewWeight",
    "TopKriterUyumu": "TopCriterionMatch",
    "MaksMutlakFark": "MaxAbsoluteDiff",
    "OrtalamaAğırlık": "MeanWeight",
    "AltYuzde5": "Lower5Pct",
    "UstYuzde95": "Upper95Pct",
    "StdSapma": "StdDev",
    "Bilgiİçeriği": "InformationContent",
    "Etkisi": "Effect",
    "Enİyi": "Best",
    "EnKötü": "Worst",
    "Ortalama": "Average",
    "BirinciAlternatif": "TopAlternative",
    "Değer": "Value",
    "Parametre": "Parameter",
    "Yıl": "Year",
    "AğırlıkYöntemi": "WeightMethod",
    "SıralamaYöntemi": "RankingMethod",
    "LiderAlternatif": "TopAlternative",
    "LiderSkor": "TopScore",
    "LiderKararlılığı": "LeaderStability",
    "EnİyiAlternatif": "BestAlternative",
    "EntropyAğırlığı": "EntropyWeight",
    "CILOSAğırlığı": "CILOSWeight",
    "IDOCRIWAğırlığı": "IDOCRIWWeight",
    "OrtalamaEntropy": "MeanEntropy",
    "OrtalamaCILOS": "MeanCILOS",
    "Senaryo": "Scenario",
}

_WEIGHT_METHOD_TR_EN = {
    "Standart Sapma": "Standard Deviation",
    "Eşit Ağırlık": "Equal Weights",
    "Manuel Ağırlık": "Manual Weights",
}

def col_key(df: pd.DataFrame, tr_name: str, en_name: str) -> str:
    if tr_name in df.columns:
        return tr_name
    if en_name in df.columns:
        return en_name
    return tr_name

def localize_df(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame) or st.session_state.get("ui_lang") != "EN":
        return df
    return df.rename(columns={k: v for k, v in _COL_TR_EN.items() if k in df.columns})

def localize_df_lang(df: pd.DataFrame, lang: str) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame) or lang != "EN":
        return df
    return df.rename(columns={k: v for k, v in _COL_TR_EN.items() if k in df.columns})

def method_display_name(name: str) -> str:
    if st.session_state.get("ui_lang") == "EN":
        return _WEIGHT_METHOD_TR_EN.get(name, name)
    return name

def method_internal_name(name: str) -> str:
    if name in _WEIGHT_METHOD_TR_EN:
        return name
    for tr_name, en_name in _WEIGHT_METHOD_TR_EN.items():
        if name == en_name:
            return tr_name
    return name


def weight_method_groups() -> List[tuple[str, List[str]]]:
    return [
        (
            tt("Bilgi ve dağılım temelli", "Information and dispersion based"),
            ["Entropy", "Standart Sapma", "LOPCOW"],
        ),
        (
            tt("İlişki ve yapı temelli", "Relation and structure based"),
            ["CRITIC", "PCA"],
        ),
        (
            tt("Etki ve hibrit temelli", "Impact and hybrid based"),
            ["MEREC", "IDOCRIW", "Fuzzy IDOCRIW"],
        ),
    ]


def _wm_single_select_cb(selected_name: str, all_methods: List[str]) -> None:
    """Weight method checkbox on_change: enforce single selection."""
    if st.session_state.get(f"weight_cb_{selected_name}", False):
        for _m in all_methods:
            if _m != selected_name:
                st.session_state[f"weight_cb_{_m}"] = False
        st.session_state["weight_method_pref"] = selected_name
    else:
        # Prevent deselecting the only active checkbox
        if not any(st.session_state.get(f"weight_cb_{_m}", False) for _m in all_methods if _m != selected_name):
            st.session_state[f"weight_cb_{selected_name}"] = True


def ranking_method_groups(layer_key: str) -> List[tuple[str, List[str]]]:
    base_groups: List[tuple[str, List[str]]] = [
        (
            tt("İdeal ve uzlaşı odaklı", "Ideal and compromise oriented"),
            ["TOPSIS", "VIKOR", "MARCOS", "ARAS"],
        ),
        (
            tt("Sapma ve mesafe odaklı", "Deviation and distance oriented"),
            ["EDAS", "CODAS", "MABAC", "GRA"],
        ),
        (
            tt("Fayda toplulaştırma odaklı", "Utility aggregation oriented"),
            ["SAW", "WPM", "MAUT", "WASPAS", "CoCoSo"],
        ),
        (
            tt("Göreli üstünlük ve rekabet odaklı", "Relative dominance and competitiveness oriented"),
            ["COPRAS", "OCRA", "MOORA", "PROMETHEE"],
        ),
    ]
    if layer_key != "fuzzy":
        return base_groups
    fuzzy_methods = set(me.FUZZY_MCDM_METHODS)
    return [
        (label, [f"Fuzzy {method}" for method in methods if f"Fuzzy {method}" in fuzzy_methods])
        for label, methods in base_groups
    ]

def sample_dataset() -> pd.DataFrame:
    rng = np.random.default_rng(42)
    n = 16
    df = pd.DataFrame({
        "Alternatif": [f"A{idx:02d}" for idx in range(1, n + 1)],
        "Karlılık": rng.uniform(10, 28, n),
        "Likidite": rng.uniform(0.9, 2.6, n),
        "PazarPayı": rng.uniform(8, 22, n),
        "MüşteriMemnuniyeti": rng.uniform(55, 95, n),
        "KaliteSkoru": rng.uniform(60, 98, n),
        "TeslimSüresi": rng.uniform(2, 15, n),
        "Risk": rng.uniform(1, 12, n),
        "İşletmeMaliyeti": rng.uniform(80, 420, n),
    })
    return df

def guess_direction(col_name: str) -> str:
    lowered = col_name.lower()
    cost_keywords = ["maliyet", "risk", "süre", "borç", "gider", "hata", "şikayet", "kayip", "loss", "cost", "time", "defect"]
    return "Min (Maliyet)" if any(k in lowered for k in cost_keywords) else "Max (Fayda)"

def clean_dataframe(df: pd.DataFrame, missing_strategy: str, clip_outliers: bool) -> pd.DataFrame:
    out = df.copy()
    numeric_cols = out.select_dtypes(include=[np.number]).columns.tolist()
    if missing_strategy in {"Sil", "Drop"}:
        out = out.dropna(subset=numeric_cols)
    else:
        for col in numeric_cols:
            if out[col].isna().any():
                if missing_strategy in {"Medyan", "Median"}:
                    out[col] = out[col].fillna(out[col].median())
                elif missing_strategy in {"Ortalama", "Mean"}:
                    out[col] = out[col].fillna(out[col].mean())
                elif missing_strategy in {"Interpolasyon", "Interpolation"}:
                    out[col] = out[col].interpolate(method="linear", limit_direction="both")
                    out[col] = out[col].fillna(out[col].median())
                else:
                    out[col] = out[col].fillna(0)
    for col in numeric_cols:
        if clip_outliers:
            q1, q3 = out[col].quantile(0.25), out[col].quantile(0.75)
            iqr = q3 - q1
            out[col] = np.clip(out[col], q1 - 1.5 * iqr, q3 + 1.5 * iqr)
    return out

def recommend_parameter_defaults(selected_methods: List[str], n_alt: int, n_crit: int, uses_fuzzy: bool) -> Dict[str, float]:
    defaults = {
        "vikor_v": 0.50,
        "waspas_lambda": 0.50,
        "codas_tau": 0.02,
        "cocoso_lambda": 0.50,
        "gra_rho": 0.50,
        "promethee_q": 0.05,
        "promethee_p": 0.30,
        "promethee_s": 0.20,
        "promethee_pref_func": "linear",
        "fuzzy_spread": 0.10,
        "sensitivity_iterations": 800,
        "sensitivity_sigma": 0.12,
    }
    if n_alt <= 8:
        defaults["sensitivity_iterations"] = 500
        defaults["sensitivity_sigma"] = 0.14
    elif n_alt >= 30:
        defaults["sensitivity_iterations"] = 1500
        defaults["sensitivity_sigma"] = 0.10
    if n_crit >= 10:
        defaults["sensitivity_sigma"] = max(0.08, defaults["sensitivity_sigma"] - 0.01)
    if uses_fuzzy:
        defaults["fuzzy_spread"] = 0.08 if n_alt >= 20 else 0.10
    if any("CODAS" in m for m in selected_methods) and n_alt > 20:
        defaults["codas_tau"] = 0.015
    return defaults

def _safe_spearman(x: np.ndarray, y: np.ndarray) -> float:
    try:
        rho, _ = spearmanr(x, y)
        if rho is None or not np.isfinite(rho):
            return 0.0
        return float(rho)
    except Exception:
        return 0.0

def compute_weight_robustness(
    data: pd.DataFrame,
    criteria: List[str],
    criteria_types: Dict[str, str],
    weight_method: str,
    base_weights: Dict[str, float],
    bootstrap_n: int = 160,
    fuzzy_spread: float = 0.10,
) -> Dict[str, Any]:
    base_vec = np.asarray([float(base_weights[c]) for c in criteria], dtype=float)
    concentration = float(np.sum(base_vec ** 2))
    eff_n = float(1.0 / concentration) if concentration > 0 else 0.0
    sorted_w = sorted(base_weights.items(), key=lambda kv: kv[1], reverse=True)
    top1, top1_w = sorted_w[0]
    top2_w = sorted_w[1][1] if len(sorted_w) > 1 else np.nan
    dominance_ratio = float(top1_w / top2_w) if np.isfinite(top2_w) and top2_w > 0 else np.nan

    loo_rows: List[Dict[str, Any]] = []
    if len(data) >= 5:
        for alt in data.index:
            sub = data.drop(index=alt)
            if len(sub) < 4:
                continue
            try:
                w_sub, _ = me.compute_objective_weights(sub, criteria, criteria_types, weight_method, fuzzy_spread=fuzzy_spread)
                sub_vec = np.asarray([float(w_sub[c]) for c in criteria], dtype=float)
                rho = _safe_spearman(base_vec, sub_vec)
                top_sub = max(w_sub, key=w_sub.get)
                loo_rows.append(
                    {
                        "Alternatif": str(alt),
                        "SpearmanRho": rho,
                        "TopKriterUyumu": int(top_sub == top1),
                        "MaksMutlakFark": float(np.max(np.abs(sub_vec - base_vec))),
                    }
                )
            except Exception:
                continue
    loo_df = pd.DataFrame(loo_rows)
    loo_mean = float(loo_df["SpearmanRho"].mean()) if not loo_df.empty else np.nan
    loo_min = float(loo_df["SpearmanRho"].min()) if not loo_df.empty else np.nan
    loo_top_match = float(loo_df["TopKriterUyumu"].mean()) if not loo_df.empty else np.nan

    rng = np.random.default_rng(42)
    boot_rows: List[Dict[str, float]] = []
    n_boot = int(max(80, min(bootstrap_n, 320)))
    for _ in range(n_boot):
        idx = rng.integers(0, len(data), len(data))
        sample = data.iloc[idx]
        try:
            w_b, _ = me.compute_objective_weights(sample, criteria, criteria_types, weight_method, fuzzy_spread=fuzzy_spread)
            boot_rows.append({c: float(w_b[c]) for c in criteria})
        except Exception:
            continue
    boot_df = pd.DataFrame(boot_rows)
    boot_summary_rows: List[Dict[str, Any]] = []
    if not boot_df.empty:
        for c in criteria:
            col = boot_df[c]
            boot_summary_rows.append(
                {
                    "Kriter": c,
                    "OrtalamaAğırlık": float(col.mean()),
                    "StdSapma": float(col.std(ddof=0)),
                    "AltYuzde5": float(col.quantile(0.05)),
                    "UstYuzde95": float(col.quantile(0.95)),
                }
            )
    boot_summary = pd.DataFrame(boot_summary_rows).sort_values("OrtalamaAğırlık", ascending=False) if boot_summary_rows else pd.DataFrame()

    return {
        "base_top_criterion": top1,
        "concentration": concentration,
        "effective_criteria_n": eff_n,
        "dominance_ratio": dominance_ratio,
        "leave_one_out": loo_df,
        "loo_mean_rho": loo_mean,
        "loo_min_rho": loo_min,
        "loo_top_match_rate": loo_top_match,
        "bootstrap_summary": boot_summary,
        "bootstrap_n": len(boot_df),
    }

def build_ranking_robustness_table(result: Dict[str, Any]) -> pd.DataFrame:
    ranking = result.get("ranking", {})
    rt = ranking.get("table")
    primary = ranking.get("method")
    if rt is None or rt.empty or not primary:
        return pd.DataFrame()

    alt_col = col_key(rt, "Alternatif", "Alternative")
    primary_top = str(rt.iloc[0][alt_col])
    sens = result.get("sensitivity") or {}
    primary_stab = sens.get("top_stability")

    rows: List[Dict[str, Any]] = []
    comp = result.get("comparison", {}) or {}
    spdf = comp.get("spearman_matrix")
    tops = comp.get("top_alternatives")
    top_map: Dict[str, str] = {}
    if isinstance(tops, pd.DataFrame) and not tops.empty:
        t_m_col = col_key(tops, "Yöntem", "Method")
        t_a_col = col_key(tops, "BirinciAlternatif", "TopAlternative")
        top_map = dict(zip(tops[t_m_col].astype(str), tops[t_a_col].astype(str)))
    top_map.setdefault(primary, primary_top)

    rho_map: Dict[str, float] = {}
    if isinstance(spdf, pd.DataFrame) and not spdf.empty:
        m_col = col_key(spdf, "Yöntem", "Method")
        mat = spdf.set_index(m_col)
        for m in mat.columns:
            if primary in mat.index and m in mat.columns:
                val = mat.loc[primary, m]
                rho_map[str(m)] = float(val) if np.isfinite(val) else np.nan

    methods = list(top_map.keys()) if top_map else [primary]
    if primary not in methods:
        methods = [primary] + methods

    for m in methods:
        rho = rho_map.get(m, 1.0 if m == primary else np.nan)
        top_alt = top_map.get(m, "")
        same_top = str(top_alt) == str(primary_top)
        if np.isfinite(rho):
            if rho >= 0.85:
                mean_txt = tt("Bu yöntem ana sonuca çok yakın davranıyor.", "This method behaves very close to the main result.")
            elif rho >= 0.70:
                mean_txt = tt("Genel tablo benzer, alt sıralarda fark olabilir.", "The overall picture is similar, with possible differences in lower ranks.")
            else:
                mean_txt = tt("Bu yöntem farklı bir bakış sunuyor; sonucu dikkatle karşılaştırın.", "This method gives a different perspective; compare results carefully.")
        else:
            mean_txt = tt("Bu yöntem için yeterli karşılaştırma verisi oluşmadı.", "There is not enough comparison data for this method.")

        top_txt = (
            tt("Lider alternatif ana yöntemle aynı.", "The top alternative is the same as in the main method.")
            if same_top
            else tt("Lider alternatif farklı. Karar öncesi iki yöntemi birlikte okuyun.", "The top alternative is different. Review both methods together before deciding.")
        )
        if m == primary and primary_stab is not None:
            stab_txt = tt(
                f"Monte Carlo'da birincilik koruma oranı %{float(primary_stab)*100:.1f}.",
                f"Monte Carlo first-place retention is {float(primary_stab)*100:.1f}%.",
            )
        else:
            stab_txt = tt("Bu yöntemde Monte Carlo ayrı çalıştırılmadı.", "Monte Carlo was not run separately for this method.")

        rows.append(
            {
                tt("Yöntem", "Method"): m,
                tt("Ana Yöntemle Uyum (ρ)", "Agreement with Primary (ρ)"): (f"{rho:.3f}" if np.isfinite(rho) else "—"),
                tt("Lider Aynı mı?", "Same Top Alternative?"): (tt("Evet", "Yes") if same_top else tt("Hayır", "No")),
                tt("Ne Anlama Geliyor?", "What Does It Mean?"): f"{mean_txt} {top_txt} {stab_txt}",
            }
        )
    return pd.DataFrame(rows)

def get_method_steps(method: str, stage: str) -> str:
    is_en = st.session_state.get("ui_lang") == "EN"

    if is_en:
        weight_steps = {
            "Entropy": "1) Normalize\n2) Compute entropy\n3) Compute divergence\n4) Normalize final weights",
            "CRITIC": "1) Min-max normalization\n2) Compute standard deviations\n3) Build correlation structure\n4) Generate information-content weights",
            "Standart Sapma": "1) Normalize\n2) Compute criterion standard deviations\n3) Convert relative dispersion to weights",
            "Eşit Ağırlık": "1) Assign equal weight to each criterion\n2) Normalize (sum=1)\n3) Proceed to ranking",
            "Manuel Ağırlık": "1) Collect raw user-defined importance values\n2) Normalize valid ratios automatically\n3) Proceed to ranking",
            "MEREC": "1) Positive transformation\n2) Log-based performance matrix\n3) Criterion-removal effect\n4) Normalize effects to weights",
            "LOPCOW": "1) Normalize\n2) Compute RMS/std terms\n3) Logarithmic percentage variability\n4) Normalize weights",
            "PCA": "1) Standardize\n2) Extract principal components\n3) Combine loadings and explained variance\n4) Produce criterion weights",
            "IDOCRIW": "1) Compute entropy weights\n2) Build CILOS loss matrix\n3) Integrate entropy and CILOS components\n4) Normalize final hybrid weights",
            "Fuzzy IDOCRIW": "1) Build lower/middle/upper fuzzy scenarios\n2) Compute entropy and CILOS in each scenario\n3) Average hybrid scenario weights\n4) Normalize final fuzzy IDOCRIW weights",
        }
        ranking_steps = {
            "TOPSIS": "1) Vector normalization\n2) Weighted matrix\n3) Define PIS/NIS\n4) Distances and closeness score",
            "VIKOR": "1) Best/worst criterion values\n2) Compute S and R measures\n3) Compromise index Q\n4) Lower Q ranks higher",
            "EDAS": "1) Average solution\n2) PDA/NDA distances\n3) Weighted totals\n4) Combined normalized score",
            "CODAS": "1) Normalize\n2) Distances to negative ideal\n3) Euclidean + Manhattan separation\n4) Normalize H score",
            "COPRAS": "1) Column-based normalization\n2) Weighted normalized matrix\n3) Compute S+ and S-\n4) Rank by Q_i",
            "OCRA": "1) Min-max relative performance\n2) Benefit/cost competitiveness components\n3) Weighted total score\n4) Final ranking",
            "ARAS": "1) Add ideal row\n2) Normalize\n3) Weighted totals\n4) Rank by utility ratio K",
            "SAW": "1) Benefit/cost normalization\n2) Weighted additive utility\n3) Aggregate total utility\n4) Rank by utility score",
            "WPM": "1) Benefit/cost normalization\n2) Weighted multiplicative utility\n3) Compute product utility\n4) Rank by utility score",
            "MAUT": "1) Min-max utility transformation\n2) Weighted utility matrix\n3) Aggregate expected utility\n4) Rank by utility score",
            "WASPAS": "1) Normalize\n2) Compute WSM and WPM\n3) Combine with lambda\n4) Higher combined score is better",
            "MOORA": "1) Vector normalization\n2) Weighting\n3) Benefit-cost difference\n4) Rank by score",
            "MABAC": "1) Min-max normalization\n2) Border approximation area\n3) Distance matrices\n4) Rank by total score",
            "MARCOS": "1) Add ideal/anti-ideal rows\n2) Normalize + weight\n3) Utility ratios\n4) Final utility ranking",
            "CoCoSo": "1) Normalize\n2) Compute S and P components\n3) Ka, Kb, Kc compromises\n4) Rank by composite score",
            "PROMETHEE": "1) Pairwise comparisons\n2) Preference indices\n3) Phi+ / Phi- flows\n4) Complete ranking by net flow",
            "GRA": "1) Min-max normalization\n2) Distance to reference\n3) Grey relational coefficients\n4) Rank by relational grade",
        }
        if stage == "weight":
            return weight_steps.get(method, "Detailed steps are available for this weighting method.")
        if method.startswith("Fuzzy "):
            base = method.replace("Fuzzy ", "")
            return "1) Fuzzify data (TFN)\n2) Build defuzzified representation\n3) " + ranking_steps.get(base, "Core ranking steps") + "\n4) Interpret final ranking"
        return ranking_steps.get(method, "Detailed steps are available for this ranking method.")

    weight_steps = {
        "Entropy": "1) Normalize et\n2) Entropi hesapla\n3) Ayrışma derecesi bul\n4) Ağırlıkları normalize et",
        "CRITIC": "1) Min-max normalize et\n2) Standart sapmaları bul\n3) Korelasyon yapısını çıkar\n4) Bilgi içeriğinden ağırlık üret",
        "Standart Sapma": "1) Normalize et\n2) Kriter std sapmalarını hesapla\n3) Std oranlarını ağırlığa çevir",
        "Eşit Ağırlık": "1) Her kritere eşit önem ver\n2) Toplamı 1'e normalize et\n3) Sıralama aşamasına geç",
        "Manuel Ağırlık": "1) Kullanıcıdan ham önem puanlarını al\n2) Geçerli oranları otomatik normalize et\n3) Sıralama aşamasına geç",
        "MEREC": "1) Pozitif dönüşüm\n2) Log tabanlı performans\n3) Kriter çıkarma etkisi\n4) Etkiyi normalize et",
        "LOPCOW": "1) Normalize et\n2) RMS ve std oranlarını hesapla\n3) Logaritmik yüzde değişim gücü\n4) Ağırlıkları normalize et",
        "PCA": "1) Standardize et\n2) Temel bileşenleri çıkar\n3) Yükleri ve açıklanan varyansı birleştir\n4) Kriter ağırlıklarını üret",
        "IDOCRIW": "1) Entropi ağırlıklarını hesapla\n2) CILOS kayıp matrisini kur\n3) Entropi ve CILOS'u birleştir\n4) Hibrit ağırlıkları normalize et",
        "Fuzzy IDOCRIW": "1) Alt/orta/üst bulanık senaryoları kur\n2) Her senaryoda Entropi ve CILOS hesapla\n3) Senaryo hibrit ağırlıklarını ortala\n4) Nihai bulanık IDOCRIW ağırlıklarını normalize et",
    }
    ranking_steps = {
        "TOPSIS": "1) Vektör normalize et\n2) Ağırlıklı matris\n3) PIS/NIS belirle\n4) Uzaklıklar ve yakınlık skoru",
        "VIKOR": "1) En iyi/en kötü değerler\n2) S ve R ölçütleri\n3) Q uzlaşı indeksi\n4) Q küçük olan üst sırada",
        "EDAS": "1) Ortalama çözüm\n2) PDA/NDA\n3) Ağırlıklı toplamlar\n4) Normalize birleşik skor",
        "CODAS": "1) Normalize et\n2) Negatif ideal mesafeler\n3) Öklid+Manhattan ayrışımı\n4) H skorunu normalize et",
        "COPRAS": "1) Sütun bazlı normalize et\n2) Ağırlıklı normalize matris\n3) S+ ve S- hesapla\n4) Q_i ile sırala",
        "OCRA": "1) Min-max göreli puanlar\n2) Fayda ve maliyet rekabet bileşenleri\n3) Ağırlıklı toplam skor\n4) Sıralama",
        "ARAS": "1) İdeal satır ekle\n2) Normalize et\n3) Ağırlıklı toplam\n4) Fayda oranı K ile sırala",
        "SAW": "1) Fayda/maliyet normalize et\n2) Ağırlıklı toplamsal fayda\n3) Toplam faydayı hesapla\n4) Fayda skoruna göre sırala",
        "WPM": "1) Fayda/maliyet normalize et\n2) Ağırlıklı çarpımsal fayda\n3) Çarpımsal faydayı hesapla\n4) Fayda skoruna göre sırala",
        "MAUT": "1) Min-max fayda dönüşümü\n2) Ağırlıklı fayda matrisi\n3) Toplam beklenen fayda\n4) Fayda skoruna göre sırala",
        "WASPAS": "1) Normalize et\n2) WSM ve WPM hesapla\n3) λ ile birleştir\n4) Büyük skor daha iyi",
        "MOORA": "1) Vektör normalize et\n2) Ağırlıklandır\n3) Fayda-maliyet farkı\n4) Skoru sırala",
        "MABAC": "1) Min-max normalize et\n2) Sınır yaklaşım alanı\n3) Uzaklık matrisleri\n4) Toplam skorla sırala",
        "MARCOS": "1) İdeal/anti-ideal ekle\n2) Normalize + ağırlıklandır\n3) Yararlılık oranları\n4) Nihai utility ile sırala",
        "CoCoSo": "1) Normalize et\n2) S ve P bileşenleri\n3) Ka, Kb, Kc uzlaşıları\n4) Bileşik skorla sırala",
        "PROMETHEE": "1) İkili karşılaştırma\n2) Tercih indeksleri\n3) Phi+ / Phi- akışları\n4) Net akış ile tam sıralama",
        "GRA": "1) Min-max normalize et\n2) Referansa uzaklık\n3) Gri katsayılar\n4) İlişkisel dereceyi sırala",
    }
    if stage == "weight":
        return weight_steps.get(method, "Ağırlık yöntemi için detay adımlar mevcut.")
    if method.startswith("Fuzzy "):
        base = method.replace("Fuzzy ", "")
        return "1) Veriyi bulanıklaştır (TFN)\n2) Durulaştırılmış temsil üret\n3) " + ranking_steps.get(base, "Temel sıralama adımları") + "\n4) Sonuçları yorumla"
    return ranking_steps.get(method, "Sıralama yöntemi için detay adımlar mevcut.")

def _extract_detail_tables(details: Dict[str, Any]) -> Dict[str, pd.DataFrame]:
    out: Dict[str, pd.DataFrame] = {}
    for key, val in (details or {}).items():
        if isinstance(val, pd.DataFrame) and not val.empty:
            out[key] = val.copy()
        elif isinstance(val, np.ndarray):
            arr = np.asarray(val)
            if arr.ndim == 1 and arr.size > 0:
                out[key] = pd.DataFrame({key: arr})
            elif arr.ndim == 2 and arr.size > 0:
                out[key] = pd.DataFrame(arr)
    return out

def diag_rec_text(rec: Dict[str, str]) -> tuple[str, str]:
    if st.session_state.get("ui_lang") == "EN" and rec.get("text_en"):
        return rec.get("text_en", ""), rec.get("action_en", "")
    text = rec.get("text", "")
    action = rec.get("action", "")
    if st.session_state.get("ui_lang") != "EN":
        return text, action
    low = text.lower()
    if "sabit kriter" in low:
        return (
            "Constant criterion/criteria detected. These criteria do not provide discrimination power.",
            "Remove constant criteria from the active analysis set.",
        )
    if "yüksek korelasyon" in low:
        return (
            "High correlation detected among criteria, indicating information redundancy.",
            "Prefer correlation-aware methods such as CRITIC/PCA and reconsider overlapping criteria.",
        )
    if "orta düzey korelasyon" in low:
        return (
            "Moderate correlation detected among criteria.",
            "Use conflict-aware weighting methods and review criterion independence.",
        )
    if "yüksek değişkenlik" in low:
        return (
            "High dispersion detected in criteria values.",
            "Entropy or Standard Deviation weighting can effectively capture this discriminative structure.",
        )
    if "dengeli veri yapısı" in low:
        return (
            "Data profile appears balanced in dispersion and correlation.",
            "Most objective weighting and ranking methods can be applied safely.",
        )
    if "alternatif sayısı az" in low:
        return (
            "The number of alternatives is limited.",
            "Interpret Monte Carlo robustness more cautiously and cross-check with additional methods.",
        )
    if "geniş alternatif seti" in low:
        return (
            "A large alternative set is detected.",
            "Distance-based methods and robustness simulations are particularly informative in this setting.",
        )
    if "tüm kriterler fayda" in low:
        return (
            "All criteria are currently marked as benefit criteria.",
            "Review directions and verify whether at least one criterion should be modeled as cost.",
        )
    return (
        "No additional critical anomaly was detected for this item.",
        "Proceed with the current workflow and validate results with sensitivity analysis.",
    )

def reset_all() -> None:
    st.session_state.clear()
    init_state()

def _diag_score_and_label(diag: Dict[str, Any]) -> tuple[int, str]:
    score = 100
    if diag.get("constant_criteria"):
        score -= 35
    max_corr = float(diag.get("max_corr", 0.0))
    if max_corr >= 0.85:
        score -= 30
    elif max_corr >= 0.70:
        score -= 15
    mean_cv = float(diag.get("mean_cv", 0.0))
    if mean_cv < 0.08:
        score -= 10
    if int(diag.get("n_alt", 0)) < 5:
        score -= 10
    score = int(max(0, min(score, 100)))
    if score >= 80:
        return score, tt("Yüksek Uygunluk", "High Suitability")
    if score >= 60:
        return score, tt("Orta Uygunluk", "Moderate Suitability")
    return score, tt("Düşük Uygunluk", "Low Suitability")

# ── MATEMATİKSEL METİN OLUŞTURUCU ──────────────────────────────────────────────
def get_math_formulation(w_method: str, r_methods: List[str]) -> str:
    text = "Kullanılan Yöntemlerin Matematiksel Altyapısı:\n\n"
    
    if ("Entropi" in w_method) or ("Entropy" in w_method):
        text += "Entropi (Entropy) Ağırlıklandırma Yöntemi:\nShannon bilgi kuramına dayanan bu yöntemde, veri setindeki zıtlık ve çeşitlilik ölçülerek nesnel ağırlıklar elde edilir.\n1. Karar matrisi normalize edilir: P_ij = x_ij / Σ x_ij\n2. Her kriter için entropi değeri (e_j) hesaplanır: e_j = -k * Σ [P_ij * ln(P_ij)] (Burada k=1/ln(m))\n3. Çeşitlilik derecesi bulunur: d_j = 1 - e_j\n4. Nihai objektif ağırlıklar elde edilir: w_j = d_j / Σ d_j\n\n"
    elif "CRITIC" in w_method:
        text += "CRITIC Ağırlıklandırma Yöntemi:\nKriterlerin standart sapmasını ve birbiriyle olan korelasyonlarını (çakışmalarını) dikkate alan objektif bir yöntemdir.\n1. Karar matrisi min-max yöntemiyle doğrusal olarak normalize edilir.\n2. Kriterlerin standart sapması (σ_j) hesaplanır.\n3. Kriterler arası korelasyon matrisi (r_jk) oluşturulur.\n4. Kapsanan bilgi miktarı hesaplanır: C_j = σ_j * Σ (1 - r_jk)\n5. Nihai ağırlıklar elde edilir: w_j = C_j / Σ C_j\n\n"
    elif "Standart Sapma" in w_method:
        text += "Standart Sapma Yöntemi:\n1. Karar matrisi doğrusal olarak normalize edilir.\n2. Her bir kriterin standart sapması (σ_j) hesaplanır.\n3. Ağırlıklar, standart sapmaların oransal payına göre belirlenir: w_j = σ_j / Σ σ_j\n\n"
    
    if r_methods:
        text += f"Alternatiflerin sıralanmasında ve performanslarının değerlendirilmesinde {', '.join(r_methods)} algoritması/algoritmaları uygulanmıştır.\n\n"
        for rm in r_methods:
            if "TOPSIS" in rm and "Fuzzy" not in rm:
                text += "TOPSIS Yöntemi:\n1. Karar matrisi vektör normalizasyon yöntemi ile dönüştürülür: r_ij = x_ij / √(Σ x_ij²)\n2. Ağırlıklı matris oluşturulur: v_ij = w_j * r_ij\n3. Pozitif İdeal (A*) ve Negatif İdeal (A⁻) çözümler belirlenir.\n4. İdeallere olan Öklid uzaklıkları (S_i* ve S_i⁻) hesaplanır.\n5. Göreceli yakınlık katsayısı bulunur: C_i* = S_i⁻ / (S_i* + S_i⁻). Katsayısı 1'e en yakın olan alternatif optimum çözüm olarak kabul edilir.\n\n"
            elif "VIKOR" in rm and "Fuzzy" not in rm:
                text += "VIKOR Yöntemi:\n1. Her kriter için en iyi (f_j*) ve en kötü (f_j⁻) değerler tespit edilir.\n2. Grup faydası (S_i) ve Bireysel pişmanlık (R_i) hesaplanır.\n3. VIKOR uzlaşı indeksi hesaplanır: Q_i = v*(S_i - S*) / (S⁻ - S*) + (1-v)*(R_i - R*) / (R⁻ - R*).\n4. Q_i değeri en küçük (minimum) olan alternatif en iyi uzlaşı çözümü olarak sıralanır.\n\n"
            elif "WASPAS" in rm and "Fuzzy" not in rm:
                text += "WASPAS Yöntemi:\n1. Karar matrisi fayda/maliyet yönlerine göre doğrusal normalize edilir.\n2. Ağırlıklı Toplam Modeli (WSM) ve Ağırlıklı Çarpım Modeli (WPM) çalıştırılır.\n3. Birleşik optimizasyon skoru bulunur: Q_i = λ * WSM_i + (1 - λ) * WPM_i. Q değeri en büyük olan alternatif birinci seçilir.\n\n"
            elif "EDAS" in rm:
                text += "EDAS Yöntemi:\n1. Her bir kriterin aritmetik ortalaması (AV) hesaplanır.\n2. Ortalama değere göre Pozitif Uzaklık (PDA) ve Negatif Uzaklık (NDA) hesaplanır.\n3. Bu uzaklıklar kriter ağırlıklarıyla çarpılarak toplam skorlar (SP ve SN) elde edilir.\n4. Normalize edilmiş değerler birleştirilerek Değerlendirme Skoru (AS_i) bulunur.\n\n"
            elif "COPRAS" in rm and "Fuzzy" not in rm:
                text += "COPRAS Yöntemi:\n1. Karar matrisi sütun toplamlarına bölünerek normalize edilir.\n2. Kriter ağırlıkları ile çarpılarak ağırlıklı normalize matris elde edilir.\n3. Fayda kriterleri için S+ ve maliyet kriterleri için S- toplamları hesaplanır.\n4. Göreli önem düzeyi (Q_i) ile sıralama yapılır; büyük Q_i daha iyidir.\n\n"
            elif "OCRA" in rm and "Fuzzy" not in rm:
                text += "OCRA Yöntemi:\n1. Her kriter için min-max tabanlı göreli performans hesaplanır.\n2. Fayda ve maliyet kriterlerinin rekabet bileşen puanları ayrı üretilir.\n3. Bileşenler ağırlıklarla toplanarak operasyonel rekabet skoru elde edilir.\n4. En yüksek toplam skor en iyi alternatifi verir.\n\n"
            elif "ARAS" in rm and "Fuzzy" not in rm:
                text += "ARAS Yöntemi:\n1. Karar matrisine, en iyi performans göstergelerinden oluşan 'Optimum Alternatif (A0)' eklenir.\n2. Matris normalize edilir ve kriter ağırlıkları ile çarpılır.\n3. Her bir alternatifin optimalite fonksiyonu (S_i) hesaplanır.\n4. Fayda derecesi (K_i = S_i / S_0) formülü ile belirlenip sıralama yapılır.\n\n"
            elif "SAW" in rm and "Fuzzy" not in rm:
                text += "SAW (Simple Additive Weighting) Yöntemi:\n1. Kriterler fayda/maliyet yönüne göre normalize edilir.\n2. Normalize değerler kriter ağırlıkları ile çarpılır.\n3. Her alternatif için ağırlıklı toplam fayda skoru hesaplanır.\n4. En yüksek toplam fayda değerine sahip alternatif en iyi olarak seçilir.\n\n"
            elif "WPM" in rm and "Fuzzy" not in rm:
                text += "WPM (Weighted Product Model) Yöntemi:\n1. Kriterler fayda/maliyet yönüne göre normalize edilir.\n2. Normalize değerler kriter ağırlıkları kadar üs alınır.\n3. Her alternatif için çarpımsal fayda skoru hesaplanır.\n4. En yüksek çarpımsal fayda skoruna sahip alternatif üstte sıralanır.\n\n"
            elif "MAUT" in rm and "Fuzzy" not in rm:
                text += "MAUT (Multi-Attribute Utility Theory) Yöntemi:\n1. Her kriter doğrusal fayda fonksiyonu ile ortak ölçeğe dönüştürülür.\n2. Fayda değerleri kriter ağırlıkları ile çarpılır.\n3. Alternatifler için toplam beklenen fayda hesaplanır.\n4. Toplam faydası en yüksek alternatif en uygun seçenek kabul edilir.\n\n"
            elif "PROMETHEE" in rm:
                text += "PROMETHEE II Yöntemi:\n1. Alternatifler kriter bazında ikili olarak karşılaştırılır.\n2. Her kriter için tercih fonksiyonu ile 0-1 aralığında tercih derecesi hesaplanır.\n3. Kriter ağırlıklarıyla küresel tercih indeksi elde edilir.\n4. Pozitif akış (phi+) ve negatif akış (phi-) hesaplanır.\n5. Net akış (phi = phi+ - phi-) en büyük olan alternatif en üstte sıralanır.\n\n"
            elif "Fuzzy" in rm:
                text += f"{rm} Yöntemi:\nKlasik (kesin) sayıların ölçüm belirsizliği veya insan hata payı barındırabileceği varsayımıyla, veriler tolerans sınırları içeren bulanık (fuzzy) sayılara dönüştürülür. Değerlendirme, üçgensel bulanık kümeler kullanılarak aralıklar üzerinden yapılır ve sonuçlar durulaştırılarak (defuzzification) kesin sıralamaya çevrilir.\n\n"
            
    return text.strip()

# ── EKRAN VE ÇIKTI FONKSİYONLARI ──────────────────────────────────────────────
def render_3layer(layers: Dict[str, str], title: str = "") -> None:
    if not any(layers.values()): return
    if title: st.markdown(f'<p style="font-size:0.8rem;font-weight:700;color:#718096;text-transform:uppercase;">{title}</p>', unsafe_allow_html=True)
    if layers.get("descriptive"): st.markdown(f'<div class="layer-card layer-desc"><div class="layer-label" style="color:#3182CE">🔵 {tt("Tanımlayıcı — Ne bulundu?", "Descriptive — What was found?")}</div><p class="layer-text">{layers["descriptive"]}</p></div>', unsafe_allow_html=True)
    if layers.get("analytic"): st.markdown(f'<div class="layer-card layer-analytic"><div class="layer-label" style="color:#38A169">🟢 {tt("Analitik — Ne anlama gelir?", "Analytic — What does it mean?")}</div><p class="layer-text">{layers["analytic"]}</p></div>', unsafe_allow_html=True)
    if layers.get("normative"): st.markdown(f'<div class="layer-card layer-norm"><div class="layer-label" style="color:#D69E2E">🟡 {tt("Normatif — Ne yapılmalı?", "Normative — What should be done?")}</div><p class="layer-text">{layers["normative"]}</p></div>', unsafe_allow_html=True)

def load_uploaded_file(uploaded_file) -> pd.DataFrame:
    return pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith(".csv") else pd.read_excel(uploaded_file)

def _panel_label(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        return str(int(value))
    return str(value).strip()

def _guess_year_columns(df: pd.DataFrame) -> List[str]:
    if not isinstance(df, pd.DataFrame) or df.empty:
        return []
    scored: List[tuple[int, str]] = []
    for col in df.columns:
        series = df[col]
        col_low = str(col).strip().lower()
        score = 0
        if any(token in col_low for token in ["year", "yil", "yıl", "donem", "dönem", "period"]):
            score += 4
        numeric = pd.to_numeric(series, errors="coerce")
        non_na = numeric.dropna()
        if len(non_na) >= max(2, int(len(series) * 0.6)):
            unique_vals = sorted(set(non_na.astype(float).tolist()))
            if 2 <= len(unique_vals) <= min(60, max(2, len(series))):
                if all(abs(v - round(v)) < 1e-9 for v in unique_vals):
                    min_v, max_v = min(unique_vals), max(unique_vals)
                    if 1900 <= min_v <= 2100 and 1900 <= max_v <= 2100:
                        score += 3
        if score > 0:
            scored.append((score, str(col)))
    scored.sort(key=lambda item: (-item[0], item[1]))
    return [col for _, col in scored]

def _sorted_panel_years(series: pd.Series) -> List[str]:
    values = {_panel_label(v) for v in series.dropna().tolist() if _panel_label(v)}
    def _sort_key(label: str) -> tuple[int, float | str]:
        try:
            return (0, float(label))
        except Exception:
            return (1, label)
    return sorted(values, key=_sort_key)

def _panel_mask(df: pd.DataFrame, year_col: str, year_label: str) -> pd.Series:
    return df[year_col].map(_panel_label) == year_label

def _set_docx_run_style(
    run,
    lang_code: str = "tr-TR",
    font_name: str = "Times New Roman",
    font_size: float = 12,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
) -> None:
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), font_name)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    set_docx_language(run, lang_code)

def _style_docx_paragraph(paragraph, lang_code: str = "tr-TR", font_size: float = 12) -> None:
    for run in paragraph.runs:
        _set_docx_run_style(run, lang_code=lang_code, font_size=font_size)
    paragraph.paragraph_format.space_after = Pt(6)

def _configure_apa_doc(doc: Document) -> None:
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)

def add_table_to_doc(doc: Document, df: pd.DataFrame, max_rows: int = 25, lang_code: str = "tr-TR") -> None:
    if df is None or df.empty: return
    use_df = df.copy().head(max_rows)
    table = doc.add_table(rows=1, cols=len(use_df.columns))
    table.style = "Table Grid"
    for i, col in enumerate(use_df.columns):
        table.rows[0].cells[i].text = str(col)
        for paragraph in table.rows[0].cells[i].paragraphs:
            for run in paragraph.runs:
                _set_docx_run_style(run, lang_code=lang_code, bold=True)
    for _, row in use_df.iterrows():
        cells = table.add_row().cells
        for j, value in enumerate(row):
            cells[j].text = f"{value:.4f}" if isinstance(value, (float, np.floating)) else str(value)
            for paragraph in cells[j].paragraphs:
                for run in paragraph.runs:
                    _set_docx_run_style(run, lang_code=lang_code)
    doc.add_paragraph("")

def render_table(df: pd.DataFrame, max_rows: int = 300) -> None:
    if not isinstance(df, pd.DataFrame):
        st.write(df)
        return
    if df.empty:
        st.info(tt("Tablo boş.", "Table is empty."))
        return
    view = df.copy()
    if len(view) > max_rows:
        st.caption(tt(f"Tablo ilk {max_rows} satırla sınırlandı.", f"Table limited to first {max_rows} rows."))
        view = view.head(max_rows)
    _height = min(540, max(180, 36 + len(view) * 28))
    view = view.replace([np.inf, -np.inf], np.nan).fillna("")
    _html = view.to_html(index=False, escape=True)
    st.markdown(
        f"""
        <div style="max-height:{_height}px; overflow:auto; border:1px solid #D1D5DB; border-radius:8px; background:#FFFFFF;">
            <style>
                .mcdm-inline-table {{
                    width:100%;
                    border-collapse:collapse;
                    font-size:0.86rem;
                    color:#111827;
                }}
                .mcdm-inline-table thead th {{
                    position:sticky;
                    top:0;
                    background:#F3F4F6;
                    color:#111827;
                    border:1px solid #D1D5DB;
                    padding:6px 8px;
                    text-align:left;
                    white-space:nowrap;
                }}
                .mcdm-inline-table tbody td {{
                    background:#FFFFFF;
                    color:#111827;
                    border:1px solid #E5E7EB;
                    padding:6px 8px;
                    white-space:nowrap;
                }}
            </style>
            {_html.replace('<table border="1" class="dataframe">', '<table class="mcdm-inline-table">')}
        </div>
        """,
        unsafe_allow_html=True,
    )

def set_docx_language(run, lang_code: str = "tr-TR") -> None:
    rpr = run._element.get_or_add_rPr()
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), lang_code)
    rpr.append(lang)

def _html_to_plain(text: str) -> str:
    if not text:
        return ""
    txt = text.replace("<br><br>", "\n\n").replace("<br>", "\n")
    txt = re.sub(r"<[^>]+>", "", txt)
    return txt.strip()

def get_math_formulation_en(w_method: str, r_methods: List[str]) -> str:
    lines: List[str] = ["Mathematical and Algorithmic Framework", ""]
    lines.append(f"Objective weighting method: {method_display_name(w_method)}")
    lines.append(get_method_steps(w_method, "weight"))
    lines.append("")
    if r_methods:
        lines.append(f"Ranking methods applied: {', '.join(r_methods)}")
        lines.append("")
        for rm in r_methods:
            lines.append(f"{rm} method:")
            lines.append(get_method_steps(rm, "ranking"))
            lines.append("")
    return "\n".join(lines).strip()

def _build_doc_sections(result: Dict[str, Any], lang: str) -> Dict[str, Any]:
    base = result.get("report_sections", {})
    w_method = result["weights"]["method"]
    w_method_disp = _WEIGHT_METHOD_TR_EN.get(w_method, w_method) if lang == "EN" else w_method
    r_method = result.get("ranking", {}).get("method")
    r_method_disp = r_method or ("Not applied" if lang == "EN" else "Uygulanmadı")
    data = result.get("selected_data", pd.DataFrame())
    n_alt, n_crit = data.shape if isinstance(data, pd.DataFrame) else (0, 0)
    comp = result.get("comparison", {}) or {}
    sens = result.get("sensitivity") or {}
    spdf = comp.get("spearman_matrix")
    n_comp = int(spdf.shape[0]) if isinstance(spdf, pd.DataFrame) and not spdf.empty else 0
    n_iter = int(sens.get("n_iterations", 0)) if sens else 0
    w_ph = result.get("method_philosophy", {}).get("weight", {}) or {}
    r_ph = result.get("method_philosophy", {}).get("ranking", {}) or {}
    findings_parts = [
        ("Data Profile", "Veri Profili"),
        ("Weight Findings", "Ağırlık Bulguları"),
    ]
    findings_texts = [
        _html_to_plain(gen_stat_commentary(result)),
        _html_to_plain(gen_weight_commentary(result)),
    ]
    if r_method:
        findings_parts.append(("Ranking Findings", "Sıralama Bulguları"))
        findings_texts.append(_html_to_plain(gen_ranking_commentary(result, {})))
    if n_comp >= 2:
        findings_parts.append(("Method Comparison", "Yöntem Karşılaştırması"))
        findings_texts.append(_html_to_plain(gen_comparison_commentary(result)))
    if sens:
        findings_parts.append(("Robustness", "Sağlamlık"))
        findings_texts.append(_html_to_plain(gen_mc_commentary(result)))
    refs = base.get("Kaynakça", []) or base.get("References", [])

    if lang == "EN":
        findings_body = []
        for (en_title, _), text in zip(findings_parts, findings_texts):
            if text:
                findings_body.append(f"{en_title}\n{text}")
        philosophy_lines = [
            f"Weighting approach: {w_method_disp}. {w_ph.get('simple', '').strip()}",
        ]
        if w_ph.get("academic"):
            philosophy_lines.append(w_ph["academic"].strip())
        if r_method:
            philosophy_lines.append(f"Ranking approach: {r_method_disp}. {r_ph.get('simple', '').strip()}")
            if r_ph.get("academic"):
                philosophy_lines.append(r_ph["academic"].strip())
        return {
            "Objective of the Study": (
                f"This report evaluates {n_alt} alternatives across {n_crit} criteria. "
                f"The scope includes criterion weighting with {w_method_disp}"
                + (f", final ranking with {r_method_disp}" if r_method else "")
                + (f", comparison across {n_comp} methods" if n_comp >= 2 else "")
                + (f", and robustness checks based on {n_iter:,} Monte Carlo scenarios." if n_iter > 0 else ".")
            ),
            "Philosophy of the Study": "\n\n".join([line for line in philosophy_lines if line.strip()]),
            "Methodology": (
                f"The workflow was kept application-focused: data profile review, weighting, ranking"
                + (", cross-method comparison" if n_comp >= 2 else "")
                + (", and robustness testing" if sens else "")
                + ". This report intentionally excludes theoretical formulas and focuses on the chosen methods, their findings, and the supporting tables."
            ),
            "Findings": "\n\n".join(findings_body),
            "References": refs,
        }

    findings_body = []
    for (_, tr_title), text in zip(findings_parts, findings_texts):
        if text:
            findings_body.append(f"{tr_title}\n{text}")
    philosophy_lines = [
        f"Ağırlıklandırma yaklaşımı: {w_method_disp}. {w_ph.get('simple', '').strip()}",
    ]
    if w_ph.get("academic"):
        philosophy_lines.append(w_ph["academic"].strip())
    if r_method:
        philosophy_lines.append(f"Sıralama yaklaşımı: {r_method_disp}. {r_ph.get('simple', '').strip()}")
        if r_ph.get("academic"):
            philosophy_lines.append(r_ph["academic"].strip())
    return {
        "Çalışmanın Amacı": (
            f"Bu rapor {n_alt} alternatif ve {n_crit} kriterden oluşan veri setini değerlendirmektedir. "
            f"Kapsamda {w_method_disp} ile kriter ağırlıkları"
            + (f", {r_method_disp} ile nihai sıralama" if r_method else "")
            + (f", {n_comp} yöntem arasında karşılaştırma" if n_comp >= 2 else "")
            + (f" ve {n_iter:,} Monte Carlo senaryosuna dayalı sağlamlık incelemesi yer almaktadır." if n_iter > 0 else ".")
        ),
        "Çalışmanın Felsefesi": "\n\n".join([line for line in philosophy_lines if line.strip()]),
        "Metodoloji": (
            "İş akışı uygulama odaklı tutulmuştur: veri profilinin okunması, ağırlıkların üretilmesi, sıralamanın kurulması"
            + (", yöntemler arası karşılaştırma" if n_comp >= 2 else "")
            + (", sağlamlık testleri" if sens else "")
            + ". Bu raporda teorik formüller özellikle çıkarılmış; yalnızca seçilen yöntemler, bulgular ve bunları destekleyen tablolar korunmuştur."
        ),
        "Bulgular": "\n\n".join(findings_body),
        "Kaynakça": refs,
    }

def _reference_notice_lines(lang: str) -> List[str]:
    if lang == "EN":
        return [
            "The references listed below are recommended sources. Responsibility for checking their accuracy, suitability, and correct use in a study belongs to the author of that study; users are advised to review all references before use.",
            "The methods presented in this system are free of charge. If you benefit from this system or my related publications in your academic or professional work, appropriate citation is kindly requested as a matter of academic courtesy.",
        ]
    return [
        "Aşağıda listelenen kaynaklar önerilen kaynaklar niteliğindedir. Bu kaynakların doğruluğunu, uygunluğunu ve bir çalışmada doğru kullanımını kontrol etme sorumluluğu o çalışmanın yazarına aittir; kullanıcıların tüm kaynakları kullanmadan önce gözden geçirmesi önerilir.",
        "Bu sistemde sunulan yöntemler ücretsizdir. Akademik veya mesleki çalışmalarınızda bu sistemden ya da ilgili yayınlarımdan yararlanmanız halinde, akademik nezaket gereği uygun atıf yapmanız kibarca rica olunur.",
    ]

def _reference_notice_html(lang: str) -> str:
    title = "Important Note" if lang == "EN" else "Önemli Not"
    body = "<br><br>".join(_reference_notice_lines(lang))
    return f'<div class="commentary-box"><strong>{title}:</strong><br>{body}</div>'

def _docx_method_name(method: str | None, lang: str) -> str:
    if not method:
        return "Not Applied" if lang == "EN" else "Uygulanmadı"
    if lang == "EN":
        return _WEIGHT_METHOD_TR_EN.get(method, method)
    return method

def _docx_heading_map(lang: str) -> Dict[str, str]:
    if lang == "EN":
        return {
            "objective": "Objective of the Study",
            "scope": "Scope of the Study",
            "philosophy": "Philosophy of the Method",
            "tables": "Tables and Interpretations",
            "conclusion": "Conclusion",
            "references": "References",
        }
    return {
        "objective": "Çalışmanın Amacı",
        "scope": "Çalışmanın Kapsamı",
        "philosophy": "Yöntemin Felsefesi",
        "tables": "Tablolar ve Yorumlar",
        "conclusion": "Sonuç",
        "references": "Kaynakça",
    }

def _docx_preview_df(selected_data: pd.DataFrame, lang: str) -> pd.DataFrame:
    if not isinstance(selected_data, pd.DataFrame) or selected_data.empty:
        return pd.DataFrame()
    preview = selected_data.copy()
    if "Alternatif" not in preview.columns and "Alternative" not in preview.columns and not isinstance(preview.index, pd.RangeIndex):
        preview = preview.reset_index()
        first_col = str(preview.columns[0])
        if first_col == "index":
            preview = preview.rename(columns={first_col: "Alternative" if lang == "EN" else "Alternatif"})
    return localize_df_lang(preview.head(15), lang)

def _docx_mean_spearman(comp: Dict[str, Any]) -> float | None:
    spdf = comp.get("spearman_matrix")
    if not isinstance(spdf, pd.DataFrame) or spdf.empty or spdf.shape[0] < 2:
        return None
    method_col = col_key(spdf, "Yöntem", "Method")
    mat = spdf.set_index(method_col)
    vals = mat.where(~np.eye(mat.shape[0], dtype=bool)).stack()
    if len(vals) == 0:
        return None
    vals = pd.to_numeric(vals, errors="coerce").dropna()
    return float(vals.mean()) if len(vals) else None

def _docx_comparison_methods(comp: Dict[str, Any]) -> List[str]:
    rank_table = comp.get("rank_table")
    if not isinstance(rank_table, pd.DataFrame) or rank_table.empty:
        return []
    alt_col = col_key(rank_table, "Alternatif", "Alternative")
    return [str(c) for c in rank_table.columns if str(c) != alt_col]

def _docx_detail_interpretation(method: str | None, lang: str) -> str:
    if lang == "EN":
        mapping = {
            "PROMETHEE": "The PROMETHEE flow table makes the internal structure of the ranking visible through positive, negative, and net flows, thereby clarifying whether the leading alternative dominates pairwise comparisons broadly or only selectively.",
            "TOPSIS": "The TOPSIS distance table supports the ranking by jointly reporting closeness to the ideal solution and distance from the negative ideal, which helps explain whether leadership is driven by proximity to the optimum rather than by a single dominant criterion.",
            "VIKOR": "The VIKOR compromise table reveals how the selected leader balances group utility, individual regret, and compromise performance, making it possible to evaluate whether the result reflects a stable compromise rather than a one-sided optimum.",
            "EDAS": "The EDAS detail table documents positive and negative deviations from the average solution, showing whether the leading alternative preserves its position by producing advantage while limiting adverse deviation.",
            "CODAS": "The CODAS distance table interprets the result through separation from the negative ideal, enabling the reader to see whether the leading alternative remains strong under both Euclidean and taxicab distance logic.",
            "GRA": "The GRA detail table demonstrates how strongly each alternative is associated with the reference profile, allowing the ranking to be interpreted through relational closeness rather than raw magnitude alone.",
            "MARCOS": "The MARCOS detail table frames the result with simultaneous reference to ideal and anti-ideal benchmarks, which strengthens the interpretability of the utility ratio underlying the final order.",
            "CoCoSo": "The CoCoSo detail table shows how additive and multiplicative compromise logics converge in the final score, making the composite nature of the leadership pattern explicit.",
            "MABAC": "The MABAC detail table shows how far each alternative stands from the border approximation area, which helps explain whether the final ranking is supported by stable criterion-level positioning.",
        }
    else:
        mapping = {
            "PROMETHEE": "PROMETHEE akış tablosu, pozitif, negatif ve net akış bileşenleri üzerinden sıralamanın iç yapısını görünür kılmakta; böylece lider alternatifin ikili karşılaştırmalarda geniş tabanlı mı yoksa sınırlı mı üstünlük kurduğu daha açık biçimde izlenebilmektedir.",
            "TOPSIS": "TOPSIS uzaklık tablosu, ideal çözüme yakınlık ile negatif idealden uzaklık bilgisini birlikte sunarak liderliğin tek bir baskın kritere değil, optimum profile görece yakınlığa dayanıp dayanmadığını açıklamaktadır.",
            "VIKOR": "VIKOR uzlaşı tablosu, seçilen liderin grup faydası, bireysel pişmanlık ve uzlaşı performansı arasında nasıl bir denge kurduğunu ortaya koymakta; böylece sonucun tek yönlü bir optimumdan mı yoksa savunulabilir bir uzlaşı çözümünden mi beslendiği anlaşılmaktadır.",
            "EDAS": "EDAS detay tablosu, ortalama çözüme göre pozitif ve negatif sapmaları birlikte raporlayarak lider alternatifin avantaj üretirken olumsuz sapmayı ne ölçüde sınırladığını göstermektedir.",
            "CODAS": "CODAS uzaklık tablosu, sonucu negatif idealden ayrışma mantığıyla yorumlamakta ve lider alternatifin hem Öklid hem Manhattan uzaklığı altında ne ölçüde güçlü kaldığını görünür kılmaktadır.",
            "GRA": "GRA detay tablosu, her alternatifin referans profile olan ilişkisel yakınlığını ortaya koymakta; böylece sıralama ham büyüklükten çok referans örüntüye benzerlik üzerinden yorumlanabilmektedir.",
            "MARCOS": "MARCOS detay tablosu, ideal ve anti-ideal referanslarla eş zamanlı karşılaştırma sağlayarak nihai yararlılık oranının hangi eksende oluştuğunu daha açık hale getirmektedir.",
            "CoCoSo": "CoCoSo detay tablosu, toplamsal ve çarpımsal uzlaşı mantıklarının nihai skorda nasıl birleştiğini göstererek liderlik deseninin bileşik yapısını açıklamaktadır.",
            "MABAC": "MABAC detay tablosu, alternatiflerin sınır yaklaşım alanına göre konumunu gösterdiği için nihai sıralamanın kriter bazlı yerleşim bakımından ne kadar tutarlı olduğunu desteklemektedir.",
        }
    if method and method.startswith("Fuzzy "):
        return (
            "The fuzzy extension preserves the logic of the selected method while explicitly incorporating uncertainty through fuzzy representations, which broadens interpretability under imprecise evaluations."
            if lang == "EN"
            else "Bulanık uzantı, seçilen yöntemin temel mantığını korurken belirsizliği bulanık temsiller üzerinden açık biçimde modele kattığı için yorum gücünü kesin olmayan değerlendirmeler altında genişletmektedir."
        )
    return mapping.get(method or "", "")

def _build_academic_doc_sections(result: Dict[str, Any], selected_data: pd.DataFrame, lang: str) -> Dict[str, Any]:
    headings = _docx_heading_map(lang)
    refs = (result.get("report_sections", {}) or {}).get("Kaynakça", []) or (result.get("report_sections", {}) or {}).get("References", [])
    data = selected_data if isinstance(selected_data, pd.DataFrame) else result.get("selected_data", pd.DataFrame())
    validation = result.get("validation", {}) or {}
    shape = validation.get("shape", (0, 0))
    n_alt = int(data.shape[0]) if isinstance(data, pd.DataFrame) and not data.empty else int(shape[0] or 0)
    n_crit = int(data.shape[1]) if isinstance(data, pd.DataFrame) and not data.empty else int(shape[1] or 0)
    direction = validation.get("direction_summary", {}) or {}
    benefit_n = len(direction.get("benefit", []) or [])
    cost_n = len(direction.get("cost", []) or [])

    weight_method = result.get("weights", {}).get("method")
    ranking_method = result.get("ranking", {}).get("method")
    weight_name = _docx_method_name(weight_method, lang)
    ranking_name = _docx_method_name(ranking_method, lang) if ranking_method else None
    w_ph = (result.get("method_philosophy", {}) or {}).get("weight", {}) or {}
    r_ph = (result.get("method_philosophy", {}) or {}).get("ranking", {}) or {}

    weight_df = result.get("weights", {}).get("table")
    ranking_df = result.get("ranking", {}).get("table")
    comp = result.get("comparison", {}) or {}
    sens = result.get("sensitivity") or {}
    comp_methods = _docx_comparison_methods(comp)
    mean_rho = _docx_mean_spearman(comp)
    stability = sens.get("top_stability")
    n_iter = int(sens.get("n_iterations", 0)) if sens else 0

    top_criterion = ""
    top_weight = None
    top3_share = None
    if isinstance(weight_df, pd.DataFrame) and not weight_df.empty:
        c_col = col_key(weight_df, "Kriter", "Criterion")
        w_col = col_key(weight_df, "Ağırlık", "Weight")
        w_sorted = weight_df.sort_values(w_col, ascending=False).reset_index(drop=True)
        top_criterion = str(w_sorted.iloc[0][c_col])
        top_weight = float(pd.to_numeric(w_sorted.iloc[0][w_col], errors="coerce"))
        top3_share = float(pd.to_numeric(w_sorted[w_col], errors="coerce").head(3).sum())

    top_alt = ""
    second_alt = ""
    top_score = None
    score_gap = None
    if isinstance(ranking_df, pd.DataFrame) and not ranking_df.empty:
        alt_col = col_key(ranking_df, "Alternatif", "Alternative")
        score_col = col_key(ranking_df, "Skor", "Score")
        r_sorted = ranking_df.sort_values(col_key(ranking_df, "Sıra", "Rank")).reset_index(drop=True)
        top_alt = str(r_sorted.iloc[0][alt_col])
        top_score = float(pd.to_numeric(r_sorted.iloc[0][score_col], errors="coerce"))
        if len(r_sorted) > 1:
            second_alt = str(r_sorted.iloc[1][alt_col])
            second_score = float(pd.to_numeric(r_sorted.iloc[1][score_col], errors="coerce"))
            score_gap = top_score - second_score

    if lang == "EN":
        objective = (
            f"The purpose of this report is to examine a decision problem composed of {n_alt} alternatives and {n_crit} criteria "
            f"by deriving criterion weights through {weight_name}"
            + (f" and constructing the final ranking through {ranking_name}" if ranking_name else "")
            + ". The analytical focus is to identify the criteria that structure the problem and"
            + (f" to present the leading alternative produced by {ranking_name} within an academically defensible frame." if ranking_name else " to establish the importance structure of the criteria within an academically defensible frame.")
        )

        scope_parts = [f"The scope of the study is limited to a decision matrix containing {n_alt} alternatives and {n_crit} criteria."]
        if benefit_n or cost_n:
            scope_parts.append(f"The criterion set consists of {benefit_n} benefit-oriented and {cost_n} cost-oriented variables.")
        scope_parts.append(
            f"The analytical design is based on the user-selected {weight_name} weighting approach"
            + (f" together with the {ranking_name} ranking procedure." if ranking_name else ".")
        )
        if len(comp_methods) >= 2:
            scope_parts.append(
                f"In addition, the agreement among the user-selected methods ({', '.join(comp_methods)}) is reported through a Spearman rank-correlation table."
            )
        if n_iter > 0 and stability is not None:
            scope_parts.append(f"Result robustness against weight perturbation is further examined through {n_iter:,} Monte Carlo scenarios.")
        scope = " ".join(scope_parts)

        philosophy_parts = []
        if w_ph.get("academic"):
            philosophy_parts.append(w_ph["academic"].strip())
        else:
            philosophy_parts.append(f"{weight_name} serves as the weighting backbone of the analysis.")
        if ranking_name:
            if r_ph.get("academic"):
                philosophy_parts.append(r_ph["academic"].strip())
            else:
                philosophy_parts.append(f"{ranking_name} provides the final ranking logic of the analysis.")
            philosophy_parts.append("Taken together, the selected weighting and ranking stages separate criterion importance from alternative dominance and thereby improve interpretability at the reporting stage.")
        philosophy_parts.append("Accordingly, the findings should be interpreted within the epistemic assumptions of the selected method set rather than as method-independent truths.")
        philosophy = "\n\n".join(philosophy_parts)

        table_parts = []
        if top_criterion and top_weight is not None and top3_share is not None:
            table_parts.append(
                f"The weight table shows that {top_criterion} is the dominant criterion with a weight of {top_weight:.4f}. "
                f"The cumulative weight of the first three criteria reaches {top3_share:.2%}, suggesting that the decision structure concentrates around a limited number of decisive dimensions."
            )
        if ranking_name and top_alt and top_score is not None:
            text = f"According to {ranking_name}, {top_alt} ranks first with a score of {top_score:.4f}."
            if second_alt and score_gap is not None:
                text += f" The score gap of {score_gap:.4f} over {second_alt} indicates the degree to which the leading alternative differentiates itself from the immediate follower."
            table_parts.append(text)
        detail_text = _docx_detail_interpretation(ranking_method, lang)
        if detail_text:
            table_parts.append(detail_text)
        if mean_rho is not None and len(comp_methods) >= 2:
            consistency = "high" if mean_rho >= 0.85 else "moderate" if mean_rho >= 0.70 else "limited"
            table_parts.append(
                f"The mean Spearman agreement across the user-selected methods is {mean_rho:.3f}, which points to {consistency} structural consistency among the reported rankings."
            )
        if stability is not None:
            robustness = "strong" if float(stability) >= 0.75 else "moderate" if float(stability) >= 0.50 else "limited"
            table_parts.append(
                f"The Monte Carlo analysis yields a first-place retention rate of {float(stability):.1%} for the leading alternative, indicating {robustness} robustness against weight perturbation."
            )
        tables_text = "\n\n".join(table_parts)

        conclusion_parts = []
        if ranking_name and top_alt:
            conclusion_parts.append(
                f"In conclusion, the weighting structure derived through {weight_name} and the ranking logic established by {ranking_name} jointly identify {top_alt} as the leading alternative."
            )
        else:
            conclusion_parts.append(
                f"In conclusion, the {weight_name} weighting structure identifies the main criteria that shape the decision problem."
            )
        if top_criterion and top_weight is not None:
            conclusion_parts.append(
                f"The prominence of {top_criterion} confirms that the decision is primarily driven by criterion-specific differentiation rather than by uniform criterion effects."
            )
        if stability is not None:
            conclusion_parts.append(
                "The robustness evidence supports the transfer of the result into practice, provided that the final decision is still read together with contextual expert judgment and data recency."
            )
        else:
            conclusion_parts.append(
                "The result should be transferred into practice together with contextual expert judgment and problem-specific constraints."
            )
        conclusion = " ".join(conclusion_parts)
    else:
        objective = (
            f"Bu raporun amacı, {n_alt} alternatif ve {n_crit} kriterden oluşan karar problemini "
            f"{weight_name} ile kriter ağırlıklarını türeterek"
            + (f" ve {ranking_name} ile nihai sıralamayı kurarak" if ranking_name else "")
            + " sistematik biçimde incelemektir. Analitik odak, karar yapısını sürükleyen kriterleri görünür kılmak ve"
            + (f" {ranking_name} altında öne çıkan alternatifi akademik olarak savunulabilir bir çerçevede ortaya koymaktır." if ranking_name else " kriter önem yapısını akademik olarak savunulabilir bir çerçevede ortaya koymaktır.")
        )

        scope_parts = [f"Çalışmanın kapsamı, {n_alt} alternatif ve {n_crit} kriter içeren karar matrisi ile sınırlıdır."]
        if benefit_n or cost_n:
            scope_parts.append(f"Kriter seti {benefit_n} fayda yönlü ve {cost_n} maliyet yönlü değişkenden oluşmaktadır.")
        scope_parts.append(
            f"Analitik kurgu kullanıcı tarafından seçilen {weight_name} ağırlıklandırmasına"
            + (f" ve {ranking_name} sıralama yaklaşımına dayanmaktadır." if ranking_name else " dayanmaktadır.")
        )
        if len(comp_methods) >= 2:
            scope_parts.append(
                f"Ek olarak kullanıcı tarafından seçilen {', '.join(comp_methods)} yöntemleri arasındaki uyum, Spearman sıra korelasyonu tablosu üzerinden raporlanmıştır."
            )
        if n_iter > 0 and stability is not None:
            scope_parts.append(f"Sonuçların ağırlık bozulmalarına karşı dayanıklılığı ayrıca {n_iter:,} Monte Carlo senaryosu ile incelenmiştir.")
        scope = " ".join(scope_parts)

        philosophy_parts = []
        if w_ph.get("academic"):
            philosophy_parts.append(w_ph["academic"].strip())
        else:
            philosophy_parts.append(f"{weight_name} bu raporda ağırlıklandırma omurgasını oluşturmaktadır.")
        if ranking_name:
            if r_ph.get("academic"):
                philosophy_parts.append(r_ph["academic"].strip())
            else:
                philosophy_parts.append(f"{ranking_name} bu raporda nihai sıralama mantığını üretmektedir.")
            philosophy_parts.append("Seçilen ağırlıklandırma ve sıralama aşamalarının birlikte kullanılması, kriter önemini alternatif üstünlüğünden ayırarak yorum izlenebilirliğini artırmaktadır.")
        philosophy_parts.append("Bu nedenle bulgular, yöntemden bağımsız mutlak doğrular olarak değil, seçilen yöntem setinin epistemik varsayımları içinde okunmalıdır.")
        philosophy = "\n\n".join(philosophy_parts)

        table_parts = []
        if top_criterion and top_weight is not None and top3_share is not None:
            table_parts.append(
                f"Ağırlık tablosu, {top_criterion} kriterinin {top_weight:.4f} ağırlık ile en baskın boyut olduğunu göstermektedir. "
                f"İlk üç kriterin toplam ağırlığı %{top3_share * 100:.2f} düzeyine ulaşmakta; bu görünüm karar yapısının sınırlı sayıda belirleyici eksen etrafında yoğunlaştığını düşündürmektedir."
            )
        if ranking_name and top_alt and top_score is not None:
            text = f"{ranking_name} sonuçlarına göre {top_alt}, {top_score:.4f} skoruyla ilk sırada yer almıştır."
            if second_alt and score_gap is not None:
                text += f" İkinci sıradaki {second_alt} ile oluşan {score_gap:.4f} puanlık fark, lider alternatifin en yakın rakibine göre ne ölçüde ayrıştığını göstermektedir."
            table_parts.append(text)
        detail_text = _docx_detail_interpretation(ranking_method, lang)
        if detail_text:
            table_parts.append(detail_text)
        if mean_rho is not None and len(comp_methods) >= 2:
            consistency = "yüksek" if mean_rho >= 0.85 else "orta" if mean_rho >= 0.70 else "sınırlı"
            table_parts.append(
                f"Kullanıcı tarafından seçilen yöntemler arasındaki ortalama Spearman uyumu {mean_rho:.3f} düzeyindedir; bu değer yöntemler arası {consistency} yapısal tutarlılığa işaret etmektedir."
            )
        if stability is not None:
            robustness = "güçlü" if float(stability) >= 0.75 else "orta" if float(stability) >= 0.50 else "sınırlı"
            table_parts.append(
                f"Monte Carlo incelemesinde lider alternatifin birinci sırayı koruma oranı %{float(stability) * 100:.1f} olarak hesaplanmıştır. Bu bulgu, sonucun ağırlık bozulmalarına karşı {robustness} bir dayanıklılık sergilediğini göstermektedir."
            )
        tables_text = "\n\n".join(table_parts)

        conclusion_parts = []
        if ranking_name and top_alt:
            conclusion_parts.append(
                f"Sonuç olarak, {weight_name} ile üretilen ağırlık yapısı ve {ranking_name} ile kurulan sıralama mantığı birlikte değerlendirildiğinde {top_alt} alternatifi en güçlü seçenek olarak öne çıkmaktadır."
            )
        else:
            conclusion_parts.append(
                f"Sonuç olarak, {weight_name} ile elde edilen ağırlık yapısı karar problemini taşıyan başlıca kriterleri açık biçimde ortaya koymaktadır."
            )
        if top_criterion and top_weight is not None:
            conclusion_parts.append(
                f"{top_criterion} kriterinin baskınlığı, kararın homojen bir yapıdan çok kriterler arası ayırt edicilik üzerinden şekillendiğini doğrulamaktadır."
            )
        if stability is not None:
            conclusion_parts.append(
                "Sağlamlık bulguları, sonucun uygulamaya aktarılabilirliğini desteklemektedir; bununla birlikte nihai karar yine bağlamsal uzman görüşü ve veri güncelliği ile birlikte değerlendirilmelidir."
            )
        else:
            conclusion_parts.append(
                "Bulguların uygulamaya aktarımı, bağlamsal uzman görüşü ve problem alanına özgü sınırlılıklar ile birlikte ele alınmalıdır."
            )
        conclusion = " ".join(conclusion_parts)

    return {
        headings["objective"]: objective,
        headings["scope"]: scope,
        headings["philosophy"]: philosophy,
        headings["tables"]: tables_text,
        headings["conclusion"]: conclusion,
        headings["references"]: refs,
    }

def _preferred_doc_detail_table(result: Dict[str, Any], lang: str) -> tuple[str, pd.DataFrame] | None:
    details = result.get("ranking", {}).get("details", {}) or {}
    preferred = [
        ("promethee_flows", ("PROMETHEE Flow Table", "PROMETHEE Akış Tablosu")),
        ("distance_table", ("TOPSIS Distance Table", "TOPSIS Uzaklık Tablosu")),
        ("vikor_table", ("VIKOR Compromise Table", "VIKOR Uzlaşı Tablosu")),
        ("edas_table", ("EDAS Distance Table", "EDAS Sapma Tablosu")),
        ("codas_table", ("CODAS Distance Table", "CODAS Uzaklık Tablosu")),
        ("gra_table", ("GRA Detail Table", "GRA Detay Tablosu")),
        ("marcos_table", ("MARCOS Detail Table", "MARCOS Detay Tablosu")),
        ("cocoso_table", ("CoCoSo Detail Table", "CoCoSo Detay Tablosu")),
        ("mabac_table", ("MABAC Detail Table", "MABAC Detay Tablosu")),
    ]
    for key, labels in preferred:
        df = details.get(key)
        if isinstance(df, pd.DataFrame) and not df.empty:
            return (labels[0] if lang == "EN" else labels[1], df)
    return None

def generate_apa_docx(result: Dict[str, Any], selected_data: pd.DataFrame, lang: str = "TR") -> bytes | None:
    if not DOCX_AVAILABLE: return None
    doc = Document()
    _configure_apa_doc(doc)
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1), Inches(1)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(
        "Çok Kriterli Karar Verme Akademik Analiz Raporu"
        if lang != "EN"
        else "Multi-Criteria Decision-Making Academic Analysis Report"
    )
    _set_docx_run_style(title_run, "en-US" if lang == "EN" else "tr-TR", bold=True)
    doc.add_paragraph("")

    sections = _build_academic_doc_sections(result, selected_data, lang)
    heading_map = _docx_heading_map(lang)
    heading_order = [
        heading_map["objective"],
        heading_map["scope"],
        heading_map["philosophy"],
        heading_map["tables"],
        heading_map["conclusion"],
        heading_map["references"],
    ]
    ref_heading = heading_map["references"]
    tables_heading = heading_map["tables"]
    body_lang = "en-US" if lang == "EN" else "tr-TR"

    for heading in heading_order:
        if heading == ref_heading:
            for note_line in _reference_notice_lines(lang):
                note_p = doc.add_paragraph()
                note_run = note_p.add_run(note_line)
                _set_docx_run_style(note_run, body_lang, italic=True)
        p = doc.add_paragraph()
        run = p.add_run(heading)
        _set_docx_run_style(run, body_lang, bold=True)
        if heading != ref_heading:
            section_text = str(sections.get(heading, "") or "").strip()
            if section_text:
                for block in [b.strip() for b in section_text.split("\n\n") if b.strip()]:
                    body = doc.add_paragraph()
                    body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    body_run = body.add_run(block)
                    _set_docx_run_style(body_run, body_lang)
            if heading == tables_heading:
                _dtitle = "Decision Matrix Snapshot" if lang == "EN" else "Karar Matrisi Özeti"
                dtitle = doc.add_paragraph()
                _set_docx_run_style(dtitle.add_run(_dtitle), body_lang, bold=True)
                add_table_to_doc(doc, _docx_preview_df(selected_data, lang), lang_code=body_lang)
                _wtitle = "Weight Table" if lang == "EN" else "Ağırlık Tablosu"
                wtitle = doc.add_paragraph()
                _set_docx_run_style(wtitle.add_run(_wtitle), body_lang, bold=True)
                add_table_to_doc(doc, localize_df_lang(result["weights"]["table"], lang), lang_code=body_lang)
                if result["ranking"]["table"] is not None:
                    _rtitle = "Ranking Table" if lang == "EN" else "Sıralama Tablosu"
                    rtitle = doc.add_paragraph()
                    _set_docx_run_style(rtitle.add_run(_rtitle), body_lang, bold=True)
                    add_table_to_doc(doc, localize_df_lang(result["ranking"]["table"], lang), lang_code=body_lang)
                detail_info = _preferred_doc_detail_table(result, lang)
                if detail_info is not None:
                    detail_title, detail_df = detail_info
                    dpar = doc.add_paragraph()
                    _set_docx_run_style(dpar.add_run(detail_title), body_lang, bold=True)
                    add_table_to_doc(doc, localize_df_lang(detail_df, lang), lang_code=body_lang)
                comp_df = (result.get("comparison") or {}).get("spearman_matrix")
                if isinstance(comp_df, pd.DataFrame) and not comp_df.empty and comp_df.shape[0] >= 2:
                    _ctitle = "Method Comparison Table" if lang == "EN" else "Yöntem Karşılaştırma Tablosu"
                    ctitle = doc.add_paragraph()
                    _set_docx_run_style(ctitle.add_run(_ctitle), body_lang, bold=True)
                    add_table_to_doc(doc, localize_df_lang(comp_df, lang), lang_code=body_lang)
                mc_df = (result.get("sensitivity") or {}).get("monte_carlo_summary")
                if isinstance(mc_df, pd.DataFrame) and not mc_df.empty:
                    _mtitle = "Monte Carlo Summary" if lang == "EN" else "Monte Carlo Özeti"
                    mtitle = doc.add_paragraph()
                    _set_docx_run_style(mtitle.add_run(_mtitle), body_lang, bold=True)
                    add_table_to_doc(doc, localize_df_lang(mc_df, lang), lang_code=body_lang)
        else:
            for ref in sections.get(ref_heading, []):
                ref_p = doc.add_paragraph()
                ref_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                ref_p.paragraph_format.left_indent = Inches(0.5)
                ref_p.paragraph_format.first_line_indent = Inches(-0.5)
                ref_run = ref_p.add_run(ref)
                _set_docx_run_style(ref_run, body_lang)

    doc.add_paragraph("\n")
    sig = doc.add_paragraph()
    sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_docx_run_style(sig.add_run('Prof. Dr. Ömer Faruk Rençber'), body_lang, italic=True)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def generate_excel(result: Dict[str, Any], selected_data: pd.DataFrame, lang: str = "TR") -> bytes:
    output = io.BytesIO()
    sheet_map = (
        {"decision": "Decision_Matrix", "stats": "Statistics", "weights": "Weights", "ranking": "Ranking"}
        if lang == "EN"
        else {"decision": "Karar_Matrisi", "stats": "Istatistikler", "weights": "Agirliklar", "ranking": "Siralama"}
    )
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        localize_df_lang(selected_data, lang).to_excel(writer, sheet_name=sheet_map["decision"])
        localize_df_lang(result["stats"], lang).to_excel(writer, sheet_name=sheet_map["stats"], index=False)
        localize_df_lang(result["weights"]["table"], lang).to_excel(writer, sheet_name=sheet_map["weights"], index=False)
        if result["ranking"]["table"] is not None:
            localize_df_lang(result["ranking"]["table"], lang).to_excel(writer, sheet_name=sheet_map["ranking"], index=False)
        workbook = writer.book
        fmt = workbook.add_format({"bold": True, "bg_color": "#1d3557", "font_color": "white"})
        for ws in writer.sheets.values():
            ws.freeze_panes(1, 0)
            ws.set_row(0, 22, fmt)
            ws.set_column(0, 18, 18)
    return output.getvalue()

def _panel_summary_rows(panel_results: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for year_label, year_result in panel_results.items():
        ranking_table = year_result.get("ranking", {}).get("table")
        top_alt = "—"
        top_score = np.nan
        if isinstance(ranking_table, pd.DataFrame) and not ranking_table.empty:
            alt_col = col_key(ranking_table, "Alternatif", "Alternative")
            score_col = col_key(ranking_table, "Skor", "Score")
            top_alt = str(ranking_table.iloc[0][alt_col])
            top_score = float(pd.to_numeric(ranking_table.iloc[0][score_col], errors="coerce"))
        rows.append(
            {
                "Yıl": year_label,
                "AğırlıkYöntemi": year_result.get("weights", {}).get("method"),
                "SıralamaYöntemi": year_result.get("ranking", {}).get("method") or "—",
                "LiderAlternatif": top_alt,
                "LiderSkor": top_score,
                "LiderKararlılığı": (year_result.get("sensitivity") or {}).get("top_stability"),
            }
        )
    return pd.DataFrame(rows)

def generate_panel_excel(panel_results: Dict[str, Dict[str, Any]], lang: str = "TR") -> bytes:
    output = io.BytesIO()
    summary = _panel_summary_rows(panel_results)
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        localize_df_lang(summary, lang).to_excel(
            writer,
            sheet_name=("Panel_Summary" if lang == "EN" else "Panel_Ozet"),
            index=False,
        )
        for year_label, year_result in panel_results.items():
            prefix = str(year_label)[:12]
            selected = year_result.get("selected_data", pd.DataFrame())
            localize_df_lang(selected, lang).to_excel(writer, sheet_name=f"{prefix}_Data"[:31], index=False)
            localize_df_lang(year_result["weights"]["table"], lang).to_excel(writer, sheet_name=f"{prefix}_W"[:31], index=False)
            if year_result.get("ranking", {}).get("table") is not None:
                localize_df_lang(year_result["ranking"]["table"], lang).to_excel(writer, sheet_name=f"{prefix}_R"[:31], index=False)
            mc_df = (year_result.get("sensitivity") or {}).get("monte_carlo_summary")
            if isinstance(mc_df, pd.DataFrame) and not mc_df.empty:
                localize_df_lang(mc_df, lang).to_excel(writer, sheet_name=f"{prefix}_MC"[:31], index=False)
        workbook = writer.book
        fmt = workbook.add_format({"bold": True, "bg_color": "#1d3557", "font_color": "white"})
        for ws in writer.sheets.values():
            ws.freeze_panes(1, 0)
            ws.set_row(0, 22, fmt)
            ws.set_column(0, 18, 18)
    return output.getvalue()

def generate_panel_apa_docx(panel_results: Dict[str, Dict[str, Any]], lang: str = "TR") -> bytes | None:
    if not DOCX_AVAILABLE:
        return None
    doc = Document()
    _configure_apa_doc(doc)
    body_lang = "en-US" if lang == "EN" else "tr-TR"
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1), Inches(1)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_text = (
        "Panel Data Academic Analysis Report"
        if lang == "EN"
        else "Panel Veri Akademik Analiz Raporu"
    )
    _set_docx_run_style(title.add_run(title_text), body_lang, bold=True)
    doc.add_paragraph("")

    refs_union: List[str] = []
    seen_refs: set[str] = set()
    heading_map = _docx_heading_map(lang)
    section_order = [
        heading_map["objective"],
        heading_map["scope"],
        heading_map["philosophy"],
        heading_map["tables"],
        heading_map["conclusion"],
    ]
    for year_label, year_result in panel_results.items():
        y_head = doc.add_paragraph()
        _set_docx_run_style(
            y_head.add_run(
                f"Year {year_label}" if lang == "EN" else f"Yıl {year_label}"
            ),
            body_lang,
            bold=True,
        )
        sections = _build_academic_doc_sections(year_result, year_result.get("selected_data", pd.DataFrame()), lang)
        for heading in section_order:
            p = doc.add_paragraph()
            _set_docx_run_style(p.add_run(heading), body_lang, bold=True)
            section_text = str(sections.get(heading, "") or "").strip()
            for block in [b.strip() for b in section_text.split("\n\n") if b.strip()]:
                body = doc.add_paragraph()
                body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                _set_docx_run_style(body.add_run(block), body_lang)
            if heading == heading_map["tables"]:
                preview_title = "Decision Matrix Snapshot" if lang == "EN" else "Karar Matrisi Özeti"
                preview_head = doc.add_paragraph()
                _set_docx_run_style(preview_head.add_run(preview_title), body_lang, bold=True)
                add_table_to_doc(doc, _docx_preview_df(year_result.get("selected_data", pd.DataFrame()), lang), lang_code=body_lang)
                weight_title = "Weight Table" if lang == "EN" else "Ağırlık Tablosu"
                weight_head = doc.add_paragraph()
                _set_docx_run_style(weight_head.add_run(weight_title), body_lang, bold=True)
                add_table_to_doc(doc, localize_df_lang(year_result["weights"]["table"], lang), lang_code=body_lang)
                ranking_table = year_result.get("ranking", {}).get("table")
                if isinstance(ranking_table, pd.DataFrame) and not ranking_table.empty:
                    rank_title = "Ranking Table" if lang == "EN" else "Sıralama Tablosu"
                    rank_head = doc.add_paragraph()
                    _set_docx_run_style(rank_head.add_run(rank_title), body_lang, bold=True)
                    add_table_to_doc(doc, localize_df_lang(ranking_table, lang), lang_code=body_lang)
                detail_info = _preferred_doc_detail_table(year_result, lang)
                if detail_info is not None:
                    detail_title, detail_df = detail_info
                    detail_head = doc.add_paragraph()
                    _set_docx_run_style(detail_head.add_run(detail_title), body_lang, bold=True)
                    add_table_to_doc(doc, localize_df_lang(detail_df, lang), lang_code=body_lang)
                comp_df = (year_result.get("comparison") or {}).get("spearman_matrix")
                if isinstance(comp_df, pd.DataFrame) and not comp_df.empty and comp_df.shape[0] >= 2:
                    comp_title = "Method Comparison Table" if lang == "EN" else "Yöntem Karşılaştırma Tablosu"
                    comp_head = doc.add_paragraph()
                    _set_docx_run_style(comp_head.add_run(comp_title), body_lang, bold=True)
                    add_table_to_doc(doc, localize_df_lang(comp_df, lang), lang_code=body_lang)
                mc_df = (year_result.get("sensitivity") or {}).get("monte_carlo_summary")
                if isinstance(mc_df, pd.DataFrame) and not mc_df.empty:
                    mc_title = "Monte Carlo Summary" if lang == "EN" else "Monte Carlo Özeti"
                    mc_head = doc.add_paragraph()
                    _set_docx_run_style(mc_head.add_run(mc_title), body_lang, bold=True)
                    add_table_to_doc(doc, localize_df_lang(mc_df, lang), lang_code=body_lang)
        for ref in sections.get(heading_map["references"], []):
            if ref not in seen_refs:
                seen_refs.add(ref)
                refs_union.append(ref)
        doc.add_paragraph("")

    for note_line in _reference_notice_lines(lang):
        note_p = doc.add_paragraph()
        _set_docx_run_style(note_p.add_run(note_line), body_lang, italic=True)
    ref_heading = doc.add_paragraph()
    _set_docx_run_style(ref_heading.add_run(heading_map["references"]), body_lang, bold=True)
    for ref in refs_union:
        ref_p = doc.add_paragraph()
        ref_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        ref_p.paragraph_format.left_indent = Inches(0.5)
        ref_p.paragraph_format.first_line_indent = Inches(-0.5)
        _set_docx_run_style(ref_p.add_run(ref), body_lang)

    sig = doc.add_paragraph()
    sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_docx_run_style(sig.add_run("Prof. Dr. Ömer Faruk Rençber"), body_lang, italic=True)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def _run_single_analysis_bundle(
    data_slice: pd.DataFrame,
    criteria: List[str],
    criteria_types: Dict[str, str],
    config,
    weight_mode_key: str,
    weight_method: str,
    main_rank: str | None,
) -> Dict[str, Any]:
    result = me.run_full_analysis(data_slice, config)
    if main_rank:
        actual_method = result.get("ranking", {}).get("method")
        if actual_method != main_rank:
            raise ValueError(tt(f"Yöntem uyumsuzluğu: seçilen '{main_rank}', çalışan '{actual_method}'.", f"Method mismatch: selected '{main_rank}', executed '{actual_method}'."))
        rt_chk = result.get("ranking", {}).get("table")
        if rt_chk is None or rt_chk.empty:
            raise ValueError(tt("Sıralama çıktısı boş döndü.", "Ranking output is empty."))
        if rt_chk["Skor"].isna().any() or (~np.isfinite(rt_chk["Skor"])).any():
            raise ValueError(tt(f"{main_rank} yöntemi geçersiz skor üretti (NaN/Inf).", f"{main_rank} produced invalid scores (NaN/Inf)."))
        if "Sıra" not in rt_chk.columns or rt_chk["Sıra"].isna().any():
            raise ValueError(tt(f"{main_rank} yöntemi sıra kolonunu üretemedi.", f"{main_rank} did not produce a valid rank column."))
    result["selected_data"] = data_slice[criteria].copy()
    if weight_mode_key == "objective":
        try:
            result["weight_robustness"] = compute_weight_robustness(
                data=data_slice[criteria].copy(),
                criteria=criteria,
                criteria_types=criteria_types,
                weight_method=weight_method,
                base_weights=result["weights"]["values"],
                bootstrap_n=180,
                fuzzy_spread=float(getattr(config, "fuzzy_spread", 0.10)),
            )
        except Exception as _w_exc:
            result["weight_robustness"] = {"error": str(_w_exc)}
    else:
        result["weight_robustness"] = {
            "info": tt(
                "Sağlamlık testi objektif ağırlık yöntemleri için tasarlanmıştır.",
                "Robustness test is designed for objective weighting methods.",
            )
        }
    return result

_THEME = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#F8FAFC", margin=dict(l=20, r=20, t=44, b=20))

def fig_weight_bar(weight_df: pd.DataFrame) -> go.Figure:
    fig = px.bar(weight_df.sort_values("Ağırlık", ascending=False), x="Kriter", y="Ağırlık", text_auto=".4f", color="Ağırlık", color_continuous_scale="Blues")
    fig.update_layout(
        height=400,
        title=tt("Kriter Ağırlıkları", "Criterion Weights"),
        coloraxis_colorbar=dict(title=tt("Ağırlık", "Weight")),
        **_THEME,
    )
    return fig

def fig_rank_bar(rank_df: pd.DataFrame) -> go.Figure:
    fig = px.bar(rank_df.sort_values("Skor"), x="Skor", y="Alternatif", orientation="h", text_auto=".4f", color="Skor", color_continuous_scale="Blues")
    fig.update_layout(
        height=450,
        title=tt("Nihai Sıralama Skorları", "Final Ranking Scores"),
        coloraxis_colorbar=dict(title=tt("Skor", "Score")),
        **_THEME,
    )
    return fig

def fig_box_plots(data: pd.DataFrame, criteria: List[str]) -> go.Figure:
    colors = ["#2E7D9E", "#1B365D", "#48BB78", "#ED8936", "#9F7AEA", "#F56565", "#38B2AC", "#ECC94B"]
    fig = go.Figure()
    for i, c in enumerate(criteria):
        hex_c = colors[i % len(colors)]
        r, g, b = int(hex_c[1:3], 16), int(hex_c[3:5], 16), int(hex_c[5:7], 16)
        
        fig.add_trace(go.Box(
            y=data[c], name=c, boxpoints="all", jitter=0.35, pointpos=-1.6,
            marker=dict(size=5, color=hex_c, opacity=0.65),
            line=dict(color=hex_c),
            fillcolor=f"rgba({r},{g},{b},0.2)",
        ))
    fig.update_layout(height=420, title=tt("Kriter Dağılım Profilleri (Kutu + Nokta)", "Criterion Distribution Profiles (Box + Points)"), showlegend=False, **_THEME)
    return fig

def fig_weight_radar(weight_df: pd.DataFrame) -> go.Figure:
    cats = weight_df["Kriter"].tolist()
    vals = weight_df["Ağırlık"].tolist()
    fig = go.Figure(go.Scatterpolar(
        r=vals + [vals[0]], theta=cats + [cats[0]],
        fill="toself", fillcolor="rgba(27,54,93,0.12)",
        line=dict(color="#1B365D", width=2),
        marker=dict(size=7, color="#2E7D9E"),
        name=tt("Ağırlık", "Weight"),
    ))
    fig.update_layout(
        height=420, title=tt("Ağırlık Dağılımı — Radar Görünümü", "Weight Distribution — Radar View"),
        polar=dict(bgcolor="#F8FAFC", radialaxis=dict(visible=True, gridcolor="#E2E8F0"), angularaxis=dict(gridcolor="#E2E8F0")),
        **_THEME,
    )
    return fig

def fig_parallel_coords(ranking_table: pd.DataFrame, contrib_table: pd.DataFrame, criteria: List[str]) -> go.Figure:
    merged = ranking_table.merge(contrib_table, on="Alternatif", how="inner")
    avail = [c for c in criteria if c in merged.columns]
    if not avail:
        return go.Figure()
    dims = [dict(label=c, values=merged[c].tolist()) for c in avail]
    dims.append(dict(label=tt("Skor", "Score"), values=merged["Skor"].tolist()))
    fig = go.Figure(go.Parcoords(
        line=dict(color=merged["Skor"].tolist(), colorscale="Blues", showscale=True, colorbar=dict(title=tt("Skor", "Score"))),
        dimensions=dims,
    ))
    fig.update_layout(height=460, title=tt("Paralel Koordinat — Alternatif Profilleri", "Parallel Coordinates — Alternative Profiles"), **_THEME)
    return fig

def fig_network_alternatives(ranking_table: pd.DataFrame) -> go.Figure:
    n = len(ranking_table)
    if n == 0:
        return go.Figure()
    alts   = ranking_table["Alternatif"].tolist()
    scores = ranking_table["Skor"].tolist()
    ranks  = ranking_table["Sıra"].tolist() if "Sıra" in ranking_table.columns else list(range(1, n + 1))

    angles = [2 * np.pi * i / n for i in range(n)]
    xs = [np.cos(a) for a in angles]
    ys = [np.sin(a) for a in angles]

    cx, cy = 0.0, 0.0
    edge_x, edge_y = [], []
    for i in range(n):
        edge_x += [xs[i], cx, None]
        edge_y += [ys[i], cy, None]

    for i in range(n):
        j = (i + 1) % n
        edge_x += [xs[i], xs[j], None]
        edge_y += [ys[i], ys[j], None]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=edge_x, y=edge_y, mode="lines",
        line=dict(color="#CBD5E0", width=1.0), hoverinfo="none", showlegend=False,
    ))
    fig.add_trace(go.Scatter(
        x=[cx], y=[cy], mode="markers",
        marker=dict(size=18, color="#1B365D", symbol="star"),
        text=[tt("En Yüksek", "Highest")], hoverinfo="text", showlegend=False,
    ))
    node_sizes = [max(20, 55 - (r - 1) * 3) for r in ranks]
    fig.add_trace(go.Scatter(
        x=xs, y=ys, mode="markers+text",
        marker=dict(
            size=node_sizes,
            color=scores, colorscale="Blues", showscale=True,
            colorbar=dict(title=tt("Skor", "Score"), thickness=12),
            line=dict(color="#1B365D", width=1.5),
        ),
        text=alts, textposition="top center",
        hovertemplate=f"<b>%{{text}}</b><br>{tt('Skor', 'Score')}: %{{marker.color:.4f}}<extra></extra>",
        showlegend=False,
    ))
    fig.update_layout(
        height=520, title=tt("Alternatif Ağ Diyagramı (Dairesel Yerleşim)", "Alternative Network Diagram (Circular Layout)"),
        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False, range=[-1.5, 1.5]),
        yaxis=dict(showgrid=False, zeroline=False, showticklabels=False, range=[-1.5, 1.5]),
        **_THEME,
    )
    return fig

def fig_mc_stability_bubble(mc_df: pd.DataFrame) -> go.Figure:
    sdf = mc_df.sort_values("BirincilikOranı", ascending=False).copy()
    sdf["Yüzde"] = sdf["BirincilikOranı"] * 100
    fig = go.Figure(go.Scatter(
        x=sdf["OrtalamaSıra"], y=sdf["Yüzde"],
        mode="markers+text",
        text=sdf["Alternatif"],
        textposition="top center",
        marker=dict(
            size=sdf["Yüzde"].clip(lower=5) * 1.2,
            color=sdf["Yüzde"],
            colorscale="Blues", showscale=True,
            colorbar=dict(title=tt("Birincilik %", "First-place %")),
            line=dict(color="#1B365D", width=1),
        ),
        hovertemplate=f"<b>%{{text}}</b><br>{tt('Ort. Sıra', 'Mean Rank')}: %{{x:.2f}}<br>{tt('Birincilik', 'First-place')}: %{{y:.1f}}%<extra></extra>",
    ))
    fig.update_layout(
        height=460, title=tt("Monte Carlo — Birincilik Oranı vs Ortalama Sıra (Balon)", "Monte Carlo — First-Place Rate vs Mean Rank (Bubble)"),
        xaxis_title=tt("Ortalama Sıra (düşük = daha iyi)", "Mean Rank (lower = better)"),
        yaxis_title=tt("Birincilik Oranı (%)", "First-Place Rate (%)"),
        **_THEME,
    )
    fig.update_xaxes(autorange="reversed")
    return fig

def fig_mc_rank_bar(mc_df: pd.DataFrame) -> go.Figure:
    sdf = mc_df.sort_values("BirincilikOranı", ascending=False)
    colors_mc = ["#1B365D" if v >= 0.5 else "#2E7D9E" if v >= 0.2 else "#90CDF4" for v in sdf["BirincilikOranı"]]
    fig = go.Figure(go.Bar(
        x=sdf["Alternatif"], y=sdf["BirincilikOranı"] * 100,
        text=[f"%{v*100:.1f}" for v in sdf["BirincilikOranı"]],
        textposition="outside",
        marker_color=colors_mc,
    ))
    fig.update_layout(
        height=420, title=tt("Monte Carlo — Birincilik Oranı (%)", "Monte Carlo — First-Place Rate (%)"),
        yaxis_title=tt("Birincilik Oranı (%)", "First-Place Rate (%)"), xaxis_title="",
        **_THEME,
    )
    return fig

def fig_local_sensitivity(local_df: pd.DataFrame) -> go.Figure:
    cat_order = ["-20%", "-10%", "+10%", "+20%"]
    local_df = local_df.copy()
    crit_col = "Kriter"
    wchg_col = "AğırlıkDeğişimi"
    if st.session_state.get("ui_lang") == "EN":
        local_df = local_df.rename(columns={"Kriter": "Criterion", "AğırlıkDeğişimi": "WeightChange"})
        crit_col = "Criterion"
        wchg_col = "WeightChange"
    local_df[wchg_col] = pd.Categorical(local_df[wchg_col], categories=cat_order, ordered=True)
    fig = px.line(
        local_df.sort_values([crit_col, wchg_col]),
        x=wchg_col, y="SpearmanRho", color=crit_col,
        markers=True, title=tt("Lokal Duyarlılık — Ağırlık Değişimine Tepki", "Local Sensitivity — Response to Weight Changes"),
    )
    fig.update_traces(line_width=2, marker_size=8)
    fig.update_layout(height=420, **_THEME)
    return fig

def _map_alt_names(values, alt_names: Dict[str, str] | None) -> List[str]:
    if not alt_names:
        return [str(v) for v in values]
    return [alt_names.get(str(v), str(v)) for v in values]

def _map_alt_names_in_df(df: pd.DataFrame, alt_names: Dict[str, str] | None) -> pd.DataFrame:
    out = df.copy()
    if not alt_names:
        return out
    for col in ["Alternatif", "Alternative", "BirinciAlternatif", "TopAlternative"]:
        if col in out.columns:
            out[col] = out[col].astype(str).map(lambda x: alt_names.get(x, x))
    return out

def _map_alt_names_in_matrix(df: pd.DataFrame, alt_names: Dict[str, str] | None) -> pd.DataFrame:
    out = df.copy()
    if not alt_names:
        return out
    out.index = [alt_names.get(str(v), str(v)) for v in out.index]
    out.columns = [alt_names.get(str(v), str(v)) for v in out.columns]
    return out

def fig_topsis_distance_scatter(distance_df: pd.DataFrame, alt_names: Dict[str, str] | None = None) -> go.Figure:
    sdf = distance_df.copy().sort_values("Skor", ascending=False)
    sdf["AltLabel"] = _map_alt_names(sdf["Alternatif"].tolist(), alt_names)
    fig = px.scatter(
        sdf,
        x="D+",
        y="D-",
        text="AltLabel",
        color="Skor",
        color_continuous_scale="Blues",
        title=tt("TOPSIS Uzaklık Haritası", "TOPSIS Distance Map"),
    )
    fig.update_traces(marker=dict(size=13, line=dict(color="#1B365D", width=1.2)), textposition="top center")
    fig.update_layout(
        height=440,
        xaxis_title=tt("İdeal Çözüme Uzaklık D+ (düşük daha iyi)", "Distance to Ideal D+ (lower is better)"),
        yaxis_title=tt("Negatif İdeale Uzaklık D- (yüksek daha iyi)", "Distance to Negative Ideal D- (higher is better)"),
        **_THEME,
    )
    fig.update_xaxes(autorange="reversed")
    return fig

def fig_vikor_components(vikor_df: pd.DataFrame, alt_names: Dict[str, str] | None = None) -> go.Figure:
    sdf = vikor_df.copy().sort_values("Q", ascending=True)
    labels = _map_alt_names(sdf["Alternatif"].tolist(), alt_names)
    fig = go.Figure()
    fig.add_trace(go.Bar(name="S", x=labels, y=sdf["S"], marker_color="#7AA6D1"))
    fig.add_trace(go.Bar(name="R", x=labels, y=sdf["R"], marker_color="#315D8A"))
    fig.add_trace(go.Bar(name="Q", x=labels, y=sdf["Q"], marker_color="#17324D"))
    fig.update_layout(
        barmode="group",
        height=440,
        title=tt("VIKOR Bileşen Dengesi (S, R, Q)", "VIKOR Component Balance (S, R, Q)"),
        xaxis_title="",
        yaxis_title=tt("Bileşen Değeri", "Component Value"),
        **_THEME,
    )
    return fig

def fig_edas_balance(edas_df: pd.DataFrame, alt_names: Dict[str, str] | None = None) -> go.Figure:
    sdf = edas_df.copy().sort_values("Skor", ascending=False)
    labels = _map_alt_names(sdf["Alternatif"].tolist(), alt_names)
    fig = go.Figure()
    fig.add_trace(go.Bar(name="NSP", x=labels, y=sdf["NSP"], marker_color="#2E7D9E"))
    fig.add_trace(go.Bar(name="NSN", x=labels, y=sdf["NSN"], marker_color="#1B365D"))
    fig.update_layout(
        barmode="group",
        height=440,
        title=tt("EDAS Pozitif/Negatif Sapma Dengesi", "EDAS Positive/Negative Distance Balance"),
        xaxis_title="",
        yaxis_title=tt("Normalize Bileşen", "Normalized Component"),
        **_THEME,
    )
    return fig

def fig_codas_distance_map(codas_df: pd.DataFrame, alt_names: Dict[str, str] | None = None) -> go.Figure:
    sdf = codas_df.copy().sort_values("H", ascending=False)
    sdf["AltLabel"] = _map_alt_names(sdf["Alternatif"].tolist(), alt_names)
    fig = px.scatter(
        sdf,
        x="E",
        y="T",
        text="AltLabel",
        color="H",
        size=np.abs(sdf["H"]).clip(lower=0.05),
        color_continuous_scale="Blues",
        title=tt("CODAS Negatif İdeale Uzaklık Haritası", "CODAS Distance-from-Negative-Ideal Map"),
    )
    fig.update_traces(marker=dict(line=dict(color="#1B365D", width=1.1)), textposition="top center")
    fig.update_layout(
        height=440,
        xaxis_title=tt("Öklid Uzaklığı E", "Euclidean Distance E"),
        yaxis_title=tt("Manhattan Uzaklığı T", "Taxicab Distance T"),
        **_THEME,
    )
    return fig

def fig_promethee_flows(prom_df: pd.DataFrame, alt_names: Dict[str, str] | None = None) -> go.Figure:
    sdf = prom_df.copy().sort_values("PhiNet", ascending=False)
    labels = _map_alt_names(sdf["Alternatif"].tolist(), alt_names)
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Phi+", x=labels, y=sdf["PhiPlus"], marker_color="#2E7D9E"))
    fig.add_trace(go.Bar(name="Phi-", x=labels, y=sdf["PhiMinus"], marker_color="#9BB9D6"))
    fig.add_trace(go.Scatter(name="PhiNet", x=labels, y=sdf["PhiNet"], mode="lines+markers", line=dict(color="#10283F", width=3), marker=dict(size=8)))
    fig.update_layout(
        barmode="group",
        height=460,
        title=tt("PROMETHEE Akış Ayrışımı", "PROMETHEE Flow Decomposition"),
        xaxis_title="",
        yaxis_title=tt("Akış Değeri", "Flow Value"),
        **_THEME,
    )
    return fig

def fig_preference_heatmap(pref_df: pd.DataFrame, alt_names: Dict[str, str] | None = None, title: str | None = None) -> go.Figure:
    disp = _map_alt_names_in_matrix(pref_df, alt_names)
    fig = px.imshow(
        disp,
        text_auto=".2f",
        aspect="auto",
        color_continuous_scale="Blues",
        zmin=0,
        zmax=1,
        title=title or tt("Tercih Matrisi", "Preference Matrix"),
    )
    fig.update_layout(height=max(420, len(disp) * 28 + 120), **_THEME)
    return fig

def render_method_specific_insights(result: Dict[str, Any], alt_names: Dict[str, str]) -> None:
    ranking = result.get("ranking", {}) or {}
    method = ranking.get("method")
    details = ranking.get("details", {}) or {}
    if not method or not details:
        st.info(tt("Yöntem-özel içgörü üretilemedi.", "Method-specific insights are not available."))
        return

    base_method = method.replace("Fuzzy ", "")
    shown = False

    if base_method == "PROMETHEE":
        flows = details.get("promethee_flows")
        pref = details.get("promethee_pref_matrix")
        if isinstance(flows, pd.DataFrame) and not flows.empty:
            shown = True
            st.markdown(f"##### 🔀 {tt('PROMETHEE Akış Tablosu', 'PROMETHEE Flow Table')}")
            _flows_disp = _map_alt_names_in_df(flows, alt_names)
            for col in ["PhiPlus", "PhiMinus", "PhiNet", "Skor"]:
                if col in _flows_disp.columns:
                    _flows_disp[col] = pd.to_numeric(_flows_disp[col], errors="coerce").round(4)
            render_table(localize_df(_flows_disp))
            fg1, fg2 = st.columns(2)
            with fg1:
                st.plotly_chart(fig_promethee_flows(flows, alt_names), use_container_width=True)
                with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                    st.markdown(f'<div class="commentary-box">{tt("PROMETHEE akış grafiği, her alternatifin diğerlerini ne ölçüde geçtiğini (Phi+) ve ne ölçüde geçildiğini (Phi-) birlikte gösterir. PhiNet çizgisi sıralamayı doğrudan belirler; yüksek net akış, ikili karşılaştırmalarda yapısal üstünlüğe işaret eder.", "The PROMETHEE flow chart shows how much each alternative outranks others (Phi+) and is outranked by others (Phi-) simultaneously. The PhiNet line directly determines the ranking; a high net flow indicates structural superiority in pairwise comparisons.")}</div>', unsafe_allow_html=True)
            with fg2:
                if isinstance(pref, pd.DataFrame) and not pref.empty:
                    st.plotly_chart(
                        fig_preference_heatmap(pref, alt_names, tt("PROMETHEE Tercih Matrisi", "PROMETHEE Preference Matrix")),
                        use_container_width=True,
                    )
                    with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                        st.markdown(f'<div class="commentary-box">{tt("Tercih matrisi ısı haritası, her alternatifin diğerine karşı ne kadar güçlü tercih ürettiğini gösterir. Satırdaki koyu hücreler baskınlığı, sütundaki koyu hücreler ise zayıf kalınan eşleşmeleri ortaya çıkarır.", "The preference matrix heatmap shows how strongly each alternative is preferred over another. Dark cells across a row indicate dominance, while dark cells down a column reveal pairings where the alternative is weak.")}</div>', unsafe_allow_html=True)
        elif isinstance(pref, pd.DataFrame) and not pref.empty:
            shown = True
            st.plotly_chart(
                fig_preference_heatmap(pref, alt_names, tt("PROMETHEE Tercih Matrisi", "PROMETHEE Preference Matrix")),
                use_container_width=True,
            )

    elif base_method == "TOPSIS":
        dist_df = details.get("distance_table")
        if isinstance(dist_df, pd.DataFrame) and not dist_df.empty:
            shown = True
            st.markdown(f"##### 📐 {tt('TOPSIS Uzaklık Tablosu', 'TOPSIS Distance Table')}")
            _dist_disp = _map_alt_names_in_df(dist_df, alt_names)
            for col in ["D+", "D-", "Skor"]:
                if col in _dist_disp.columns:
                    _dist_disp[col] = pd.to_numeric(_dist_disp[col], errors="coerce").round(4)
            render_table(localize_df(_dist_disp))
            st.plotly_chart(fig_topsis_distance_scatter(dist_df, alt_names), use_container_width=True)
            with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                st.markdown(f'<div class="commentary-box">{tt("TOPSIS uzaklık haritasında sağ üst bölgeye yaklaşan alternatifler ideala daha yakın ve negatif ideale daha uzaktır. Bu görünüm liderin yalnızca yüksek skor değil, aynı zamanda güçlü ayrışma ürettiğini göstermeye yarar.", "In the TOPSIS distance map, alternatives closer to the upper-right region are nearer to the ideal and farther from the negative ideal. This view shows whether the leader not only scores high but also separates strongly from the rest.")}</div>', unsafe_allow_html=True)

    elif base_method == "VIKOR":
        vikor_df = details.get("vikor_table")
        if isinstance(vikor_df, pd.DataFrame) and not vikor_df.empty:
            shown = True
            st.markdown(f"##### ⚖️ {tt('VIKOR Uzlaşı Tablosu', 'VIKOR Compromise Table')}")
            _vikor_disp = _map_alt_names_in_df(vikor_df, alt_names)
            for col in ["S", "R", "Q", "Skor"]:
                if col in _vikor_disp.columns:
                    _vikor_disp[col] = pd.to_numeric(_vikor_disp[col], errors="coerce").round(4)
            render_table(localize_df(_vikor_disp))
            st.plotly_chart(fig_vikor_components(vikor_df, alt_names), use_container_width=True)
            with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                st.markdown(f'<div class="commentary-box">{tt("VIKOR bileşen grafiği grup faydası (S), bireysel pişmanlık (R) ve uzlaşı skoru (Q) arasındaki dengeyi görünür kılar. Düşük Q ile birlikte makul S ve R değerleri, lider alternatifin uzlaşı çözümü olarak daha savunulabilir olduğunu gösterir.", "The VIKOR component chart makes the balance among group utility (S), individual regret (R), and compromise score (Q) visible. A low Q together with reasonable S and R values indicates the leading alternative is more defensible as a compromise solution.")}</div>', unsafe_allow_html=True)

    elif base_method == "EDAS":
        edas_df = details.get("edas_table")
        if isinstance(edas_df, pd.DataFrame) and not edas_df.empty:
            shown = True
            st.markdown(f"##### 📏 {tt('EDAS Sapma Tablosu', 'EDAS Distance Table')}")
            _edas_disp = _map_alt_names_in_df(edas_df, alt_names)
            for col in ["SP", "SN", "NSP", "NSN", "Skor"]:
                if col in _edas_disp.columns:
                    _edas_disp[col] = pd.to_numeric(_edas_disp[col], errors="coerce").round(4)
            render_table(localize_df(_edas_disp))
            st.plotly_chart(fig_edas_balance(edas_df, alt_names), use_container_width=True)
            with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                st.markdown(f'<div class="commentary-box">{tt("EDAS grafiği, her alternatifin ortalama çözüme göre pozitif ve negatif sapma dengesini gösterir. Yüksek NSP ve yüksek NSN birlikte görülüyorsa alternatif hem avantaj yaratıyor hem de olumsuz sapmayı sınırlıyor demektir.", "The EDAS chart shows the balance of positive and negative distances from the average solution for each alternative. When high NSP and high NSN appear together, the alternative both creates advantage and limits adverse deviation.")}</div>', unsafe_allow_html=True)

    elif base_method == "CODAS":
        codas_df = details.get("codas_table")
        if isinstance(codas_df, pd.DataFrame) and not codas_df.empty:
            shown = True
            st.markdown(f"##### 🧭 {tt('CODAS Uzaklık Tablosu', 'CODAS Distance Table')}")
            _codas_disp = _map_alt_names_in_df(codas_df, alt_names)
            for col in ["E", "T", "H", "Skor"]:
                if col in _codas_disp.columns:
                    _codas_disp[col] = pd.to_numeric(_codas_disp[col], errors="coerce").round(4)
            render_table(localize_df(_codas_disp))
            st.plotly_chart(fig_codas_distance_map(codas_df, alt_names), use_container_width=True)
            with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                st.markdown(f'<div class="commentary-box">{tt("CODAS haritası, alternatiflerin negatif idealden hem Öklid hem Manhattan uzaklığıyla nasıl ayrıştığını gösterir. Sağ üst bölgeye taşınan ve yüksek H skoru alan alternatifler riskli kötü senaryolardan daha güçlü biçimde ayrışır.", "The CODAS map shows how alternatives separate from the negative ideal in both Euclidean and Taxicab distance terms. Alternatives moving toward the upper-right region with high H scores separate more strongly from adverse worst-case conditions.")}</div>', unsafe_allow_html=True)

    if not shown:
        _detail_frames = _extract_detail_tables(details)
        _detail_frames.pop("result_table", None)
        if _detail_frames:
            st.markdown(f"##### 🧾 {tt('Yöntem-Özel Detay Tablosu', 'Method-Specific Detail Table')}")
            _first_key = next(iter(_detail_frames))
            render_table(localize_df(_map_alt_names_in_df(_detail_frames[_first_key], alt_names).head(250)))
        else:
            st.info(tt("Bu yöntem için ek yöntem-özel görsel bulunmuyor.", "No additional method-specific visualization is available for this method."))

def render_tab_assistant(commentary: str, key: str = "") -> None:
    with st.expander(tt("💬 Analiz Asistanı — Yorum", "💬 Analysis Assistant — Commentary"), expanded=True):
        st.markdown(f'<div class="tab-assistant">{commentary}</div>', unsafe_allow_html=True)
    st.caption(tt("Detaylar için lütfen aşağıdaki tablo ve şekilleri inceleyin.", "For details, please review the tables and figures below."))

def _top_cv_crit(result: Dict[str, Any]) -> tuple[str, float]:
    data = result["selected_data"]
    cv = (data.std() / data.mean().abs())
    return cv.idxmax(), float(cv.max())

def _max_corr_pair(result: Dict[str, Any]) -> tuple[str, str, float]:
    corr = result.get("correlation_matrix")
    if corr is None or corr.shape[0] < 2:
        return ("—", "—", 0.0)
    mask = ~np.eye(len(corr), dtype=bool)
    vals = corr.abs().where(mask).stack()
    if vals.empty: return ("—", "—", 0.0)
    pair = vals.idxmax()
    return pair[0], pair[1], float(vals.max())

def _compose_commentary(simple: str, academic: str, action: str, example: str = "") -> str:
    is_en = st.session_state.get("ui_lang") == "EN"
    main_commentary = academic or simple
    labels = {
        "commentary": "Commentary" if is_en else "Yorum",
        "example": "Short Example" if is_en else "Kısa Örnek",
        "action": "What To Do Next" if is_en else "Ne Yapmalı",
    }
    parts = [f"<strong>{labels['commentary']}:</strong> {main_commentary}"]
    if example:
        parts.append(f"<strong>{labels['example']}:</strong> {example}")
    if action:
        parts.append(f"<strong>{labels['action']}:</strong> {action}")
    return "<br><br>".join(parts)

def _compose_structured_commentary(did: str, why: str, found: str, example: str = "", action: str = "") -> str:
    is_en = st.session_state.get("ui_lang") == "EN"
    labels = {
        "did": "What Was Done" if is_en else "Ne Yapıldı",
        "why": "Why It Was Done" if is_en else "Neden Yapıldı",
        "found": "What Was Found" if is_en else "Ne Bulundu",
        "example": "Short Example" if is_en else "Kısa Örnek",
        "action": "How To Read It" if is_en else "Nasıl Okunmalı",
    }
    parts = [
        f"<strong>{labels['did']}:</strong> {did}",
        f"<strong>{labels['why']}:</strong> {why}",
        f"<strong>{labels['found']}:</strong> {found}",
    ]
    if example:
        parts.append(f"<strong>{labels['example']}:</strong> {example}")
    if action:
        parts.append(f"<strong>{labels['action']}:</strong> {action}")
    return "<br><br>".join(parts)

def gen_weight_robustness_commentary(result: Dict[str, Any]) -> str:
    wr = result.get("weight_robustness", {}) or {}
    if not wr:
        return tt("Ağırlık sağlamlık verisi mevcut değil.", "Weight robustness data is not available.")
    if wr.get("info"):
        return str(wr.get("info"))
    if wr.get("error"):
        return tt("Ağırlık sağlamlık testleri üretilemedi.", "Weight robustness tests could not be generated.")

    loo_mean = wr.get("loo_mean_rho")
    loo_min = wr.get("loo_min_rho")
    top_match = wr.get("loo_top_match_rate")
    eff_n = wr.get("effective_criteria_n")
    dom = wr.get("dominance_ratio")
    loo_df = wr.get("leave_one_out")
    boot_df = wr.get("bootstrap_summary")

    worst_alt = "—"
    worst_rho = np.nan
    if isinstance(loo_df, pd.DataFrame) and not loo_df.empty:
        rho_col = col_key(loo_df, "SpearmanRho", "SpearmanRho")
        alt_col = col_key(loo_df, "Alternatif", "Alternative")
        worst_idx = pd.to_numeric(loo_df[rho_col], errors="coerce").idxmin()
        if pd.notna(worst_idx):
            worst_alt = str(loo_df.loc[worst_idx, alt_col])
            worst_rho = float(pd.to_numeric(loo_df.loc[worst_idx, rho_col], errors="coerce"))

    uncertain_crit = "—"
    uncertain_span = np.nan
    if isinstance(boot_df, pd.DataFrame) and not boot_df.empty:
        crit_col = col_key(boot_df, "Kriter", "Criterion")
        low_col = col_key(boot_df, "AltYuzde5", "Lower5Pct")
        high_col = col_key(boot_df, "UstYuzde95", "Upper95Pct")
        boot_tmp = boot_df.copy()
        boot_tmp["_span"] = pd.to_numeric(boot_tmp[high_col], errors="coerce") - pd.to_numeric(boot_tmp[low_col], errors="coerce")
        idx = boot_tmp["_span"].idxmax()
        if pd.notna(idx):
            uncertain_crit = str(boot_tmp.loc[idx, crit_col])
            uncertain_span = float(pd.to_numeric(boot_tmp.loc[idx, "_span"], errors="coerce"))

    is_en = st.session_state.get("ui_lang") == "EN"
    did = (
        "We tested the weight vector in two ways: first, each alternative was removed one by one (leave-one-out); second, the data was repeatedly resampled to build bootstrap confidence intervals for each criterion weight."
        if is_en else
        "Ağırlık vektörü iki yoldan sınandı: önce her alternatif tek tek çıkarıldı (leave-one-out), ardından veri tekrar örneklenerek her kriter ağırlığı için bootstrap güven aralıkları üretildi."
    )
    why = (
        "The goal is to see whether the reported criterion priorities depend too much on a single row or on small sample fluctuations."
        if is_en else
        "Amaç, raporlanan kriter önceliklerinin tek bir satıra ya da küçük örnek oynamalarına aşırı bağlı olup olmadığını görmektir."
    )

    if loo_mean is not None and np.isfinite(loo_mean) and top_match is not None and np.isfinite(top_match):
        if loo_mean >= 0.90 and top_match >= 0.75:
            stability_text = (
                "Overall robustness is high."
                if is_en else
                "Genel sağlamlık düzeyi yüksektir."
            )
            action = (
                "You can report the weight structure with stronger confidence; still keep the most uncertain criterion under observation in the discussion section."
                if is_en else
                "Ağırlık yapısını daha güçlü biçimde raporlayabilirsiniz; yine de tartışma bölümünde en belirsiz kriteri ayrıca izleyin."
            )
        elif loo_mean >= 0.75 and top_match >= 0.50:
            stability_text = (
                "Overall robustness is moderate."
                if is_en else
                "Genel sağlamlık düzeyi ortadır."
            )
            action = (
                "Present the weights together with a caution note, and emphasize which criterion becomes unstable under data perturbation."
                if is_en else
                "Ağırlıkları temkin notuyla birlikte sunun ve veri oynadığında hangi kriterin daha çabuk bozulduğunu özellikle belirtin."
            )
        else:
            stability_text = (
                "Overall robustness is low."
                if is_en else
                "Genel sağlamlık düzeyi düşüktür."
            )
            action = (
                "Avoid presenting the weights as fixed truth; compare alternative weighting methods and treat the current profile as sensitive."
                if is_en else
                "Ağırlıkları kesin gerçek gibi sunmayın; alternatif ağırlık yöntemleriyle karşılaştırın ve mevcut profili duyarlı kabul edin."
            )
    else:
        stability_text = tt("Kararlılık özeti üretilemedi.", "Stability summary could not be generated.")
        action = tt("Veri yapısını yeniden kontrol edin.", "Recheck the data structure.")

    found = (
        f"{stability_text} Mean leave-one-out agreement is {float(loo_mean):.3f}, worst-case agreement is {float(loo_min):.3f}, and top-criterion consistency is {float(top_match)*100:.1f}%. "
        f"The effective number of criteria is {float(eff_n):.2f}. "
        + (
            f"The top-to-second criterion weight ratio is {float(dom):.2f}, which suggests a {'concentrated' if float(dom) >= 1.80 else 'balanced'} weight profile. "
            if dom is not None and np.isfinite(dom) else ""
        )
        + (
            f"The weakest leave-one-out scenario appears when {worst_alt} is removed (ρ = {float(worst_rho):.3f}). "
            if np.isfinite(worst_rho) else ""
        )
        + (
            f"The widest bootstrap uncertainty band belongs to {uncertain_crit} (span ≈ {float(uncertain_span):.4f}), so the first movement is most likely to appear there."
            if np.isfinite(uncertain_span) else ""
        )
        if is_en else
        f"{stability_text} Leave-one-out ortalama uyumu {float(loo_mean):.3f}, en zayıf senaryo uyumu {float(loo_min):.3f} ve lider kriter tutarlılığı %{float(top_match)*100:.1f} düzeyindedir. "
        f"Etkin kriter sayısı {float(eff_n):.2f} olarak görünmektedir. "
        + (
            f"Lider/ikinci kriter ağırlık oranı {float(dom):.2f} olduğu için ağırlık yapısı {'yoğunlaşmış' if float(dom) >= 1.80 else 'daha dengeli'} okunabilir. "
            if dom is not None and np.isfinite(dom) else ""
        )
        + (
            f"En zayıf leave-one-out senaryosu {worst_alt} çıkarıldığında oluşmuştur (ρ = {float(worst_rho):.3f}). "
            if np.isfinite(worst_rho) else ""
        )
        + (
            f"Bootstrap tarafında en geniş belirsizlik bandı {uncertain_crit} kriterinde görülmektedir (aralık ≈ {float(uncertain_span):.4f}); bu nedenle ilk oynama en çok burada beklenir."
            if np.isfinite(uncertain_span) else ""
        )
    )
    example = (
        f"For example, if removing one alternative still leaves the same criterion on top, the model keeps telling the same story. If the story changes immediately, the weighting result depends too much on that observation. Here the most fragile removal is {worst_alt}."
        if is_en else
        f"Örneğin bir alternatif çıkarıldığında en önemli kriter yine aynı kalıyorsa model aynı hikâyeyi anlatmaya devam eder. Hikâye hemen değişiyorsa ağırlık sonucu o gözleme fazla bağımlıdır. Bu çalışmada en hassas çıkarma senaryosu {worst_alt} ile görülmüştür."
    )
    return _compose_structured_commentary(did, why, found, example, action)

def _ranking_method_commentary(result: Dict[str, Any], top_alt: str, top_disp: str) -> tuple[str, str, str]:
    ranking = result.get("ranking", {}) or {}
    method = str(ranking.get("method") or "")
    base_method = method.replace("Fuzzy ", "")
    details = ranking.get("details", {}) or {}
    is_en = st.session_state.get("ui_lang") == "EN"

    if base_method == "TOPSIS" and isinstance(details.get("distance_table"), pd.DataFrame):
        dist_df = details["distance_table"]
        row = dist_df[dist_df["Alternatif"].astype(str) == str(top_alt)]
        if not row.empty:
            d_plus = float(row.iloc[0]["D+"])
            d_minus = float(row.iloc[0]["D-"])
            if is_en:
                return (
                    f"{method} looks for the option closest to the ideal and farthest from the weakest profile. On that basis, {top_disp} stands out.",
                    f"For the leading option, D+ = {d_plus:.4f} and D- = {d_minus:.4f}. This means it is relatively close to the desired profile while staying far from the undesirable profile.",
                    "If the first two options are close, read this map together with the Monte Carlo and comparison tabs before making a final decision.",
                )
            return (
                f"{method}, en iyi profile en yakın ve en zayıf profile en uzak seçeneği arar. Bu ölçüte göre {top_disp} öne çıkıyor.",
                f"Lider seçenek için D+ = {d_plus:.4f} ve D- = {d_minus:.4f}. Yani {top_disp}, istenen profile görece yakın, zayıf profile ise görece uzaktır.",
                "İlk iki seçenek birbirine yakınsa son kararı vermeden önce bu haritayı Monte Carlo ve yöntem karşılaştırması sekmeleriyle birlikte okuyun.",
            )

    if base_method == "VIKOR" and isinstance(details.get("vikor_table"), pd.DataFrame):
        vikor_df = details["vikor_table"]
        row = vikor_df[vikor_df["Alternatif"].astype(str) == str(top_alt)]
        if not row.empty:
            s_val = float(row.iloc[0]["S"])
            r_val = float(row.iloc[0]["R"])
            q_val = float(row.iloc[0]["Q"])
            if is_en:
                return (
                    f"{method} tries to find the most balanced compromise. Here, {top_disp} appears as the most acceptable middle ground.",
                    f"The leading option has S = {s_val:.4f}, R = {r_val:.4f}, and Q = {q_val:.4f}. A lower Q means the option balances collective benefit and individual regret more successfully in this dataset.",
                    "This result is especially useful when you want a defensible compromise rather than an aggressive winner-takes-all choice.",
                )
            return (
                f"{method}, en dengeli uzlaşı çözümünü bulmaya çalışır. Bu veri setinde {top_disp} en kabul edilebilir orta yol olarak görünüyor.",
                f"Lider seçenek için S = {s_val:.4f}, R = {r_val:.4f} ve Q = {q_val:.4f}. Düşük Q değeri, bu alternatifin grup faydası ile bireysel pişmanlık arasında daha iyi denge kurduğunu gösterir.",
                "Keskin bir kazanan yerine savunulabilir bir uzlaşı arıyorsanız bu sonuç daha anlamlıdır.",
            )

    if base_method == "PROMETHEE" and isinstance(details.get("promethee_flows"), pd.DataFrame):
        flows = details["promethee_flows"]
        row = flows[flows["Alternatif"].astype(str) == str(top_alt)]
        if not row.empty:
            phi_plus = float(row.iloc[0]["PhiPlus"])
            phi_minus = float(row.iloc[0]["PhiMinus"])
            phi_net = float(row.iloc[0]["PhiNet"])
            if is_en:
                return (
                    f"{method} compares options pair by pair. In these direct matchups, {top_disp} defeats the others more often than it is defeated.",
                    f"For the leading option, Phi+ = {phi_plus:.4f}, Phi- = {phi_minus:.4f}, and PhiNet = {phi_net:.4f}. A high net flow means structural superiority across pairwise comparisons in this dataset.",
                    "Use the flow chart and preference matrix together if you want to see not only who leads, but against whom that lead is strong or weak.",
                )
            return (
                f"{method}, seçenekleri ikili karşılaştırmalarla okur. Bu eşleşmelerde {top_disp}, diğerlerini geçtiği durumlarda daha baskın görünüyor.",
                f"Lider seçenek için Phi+ = {phi_plus:.4f}, Phi- = {phi_minus:.4f} ve PhiNet = {phi_net:.4f}. Yüksek net akış, bu veri setinde ikili karşılaştırmaların genelinde yapısal üstünlük olduğunu gösterir.",
                "Sadece kimin önde olduğuna değil, hangi rakiplere karşı güçlü veya zayıf olduğuna da bakmak istiyorsanız akış grafiği ile tercih matrisini birlikte okuyun.",
            )

    if base_method == "EDAS" and isinstance(details.get("edas_table"), pd.DataFrame):
        edas_df = details["edas_table"]
        row = edas_df[edas_df["Alternatif"].astype(str) == str(top_alt)]
        if not row.empty:
            nsp = float(row.iloc[0]["NSP"])
            nsn = float(row.iloc[0]["NSN"])
            if is_en:
                return (
                    f"{method} checks who stays favorably above the average profile while avoiding negative deviation. On that basis, {top_disp} is the strongest candidate.",
                    f"For the leading option, NSP = {nsp:.4f} and NSN = {nsn:.4f}. The joint strength of these two components indicates a balanced performance relative to the dataset average.",
                    "This method is useful when you want to compare each option against the general market or sample average, not only against the best case.",
                )
            return (
                f"{method}, ortalama profile göre kimin avantajlı kaldığını ve kimin olumsuz sapmayı daha iyi sınırladığını ölçer. Bu açıdan {top_disp} en güçlü adaydır.",
                f"Lider seçenek için NSP = {nsp:.4f} ve NSN = {nsn:.4f}. Bu iki bileşenin birlikte güçlü olması, veri seti ortalamasına göre dengeli bir üstünlüğe işaret eder.",
                "En iyi duruma değil, genel ortalamaya göre kimin iyi kaldığını görmek istiyorsanız bu yöntem özellikle faydalıdır.",
            )

    if base_method == "CODAS" and isinstance(details.get("codas_table"), pd.DataFrame):
        codas_df = details["codas_table"]
        row = codas_df[codas_df["Alternatif"].astype(str) == str(top_alt)]
        if not row.empty:
            h_val = float(row.iloc[0]["H"])
            e_val = float(row.iloc[0]["E"])
            t_val = float(row.iloc[0]["T"])
            if is_en:
                return (
                    f"{method} rewards options that move far away from the weakest profile. On this criterion, {top_disp} creates the strongest separation.",
                    f"For the leading option, E = {e_val:.4f}, T = {t_val:.4f}, and H = {h_val:.4f}. A high H score means the option stays farther from adverse scenarios in multiple distance senses.",
                    "This reading is especially useful when you care about avoiding weak outcomes, not only chasing the best average score.",
                )
            return (
                f"{method}, en zayıf profile uzaklaşan seçenekleri ödüllendirir. Bu açıdan {top_disp} en güçlü ayrışmayı üretmiştir.",
                f"Lider seçenek için E = {e_val:.4f}, T = {t_val:.4f} ve H = {h_val:.4f}. Yüksek H skoru, olumsuz senaryolardan birden fazla uzaklık ölçüsünde uzak kalındığını gösterir.",
                "Yalnızca ortalama performansı değil, zayıf sonuçlardan kaçınmayı önemsiyorsanız bu okuma daha değerlidir.",
            )

    ph = result.get("method_philosophy", {}).get("ranking", {}) or {}
    if is_en:
        return (
            f"{method} places {top_disp} in the first position for this dataset.",
            ph.get("academic", "The selected method interprets the same data through its own ranking logic and identifies the strongest candidate accordingly."),
            "Read this result together with the score gap, comparison, and robustness outputs before turning it into a final decision.",
        )
    return (
        f"{method} bu veri setinde {top_disp} alternatifini ilk sıraya yerleştirmiştir.",
        ph.get("academic", "Seçilen yöntem aynı veriyi kendi karar mantığıyla okuyarak en güçlü adayı belirlemiştir."),
        "Bu sonucu son karara dönüştürmeden önce skor farkı, yöntem karşılaştırması ve sağlamlık çıktılarıyla birlikte okuyun.",
    )

def gen_stat_commentary(result: Dict[str, Any]) -> str:
    n_alt = result["selected_data"].shape[0]
    n_crit = result["selected_data"].shape[1]
    cv_crit, cv_val = _top_cv_crit(result)
    c1, c2, max_rho = _max_corr_pair(result)
    if st.session_state.get("ui_lang") == "EN":
        simple = f"This dataset compares <strong>{n_alt} options</strong> across <strong>{n_crit} criteria</strong>. The criterion that separates the options most clearly is <strong>{cv_crit}</strong>."
        example = (
            f"For a non-technical reading: if two criteria move almost together, the model may be hearing the same story twice. Here the strongest overlap is between <strong>{c1}</strong> and <strong>{c2}</strong>."
            if max_rho > 0.75 else
            f"For a non-technical reading: the criteria are not strongly repeating each other, so the model can look at the problem from multiple angles."
        )
        if max_rho > 0.75:
            academic = f"<strong>{cv_crit}</strong> has the highest coefficient of variation (Coeff. of Variation ≈{cv_val:.2f}), so it carries the strongest separating signal. At the same time, the high correlation between <strong>{c1}</strong> and <strong>{c2}</strong> (|ρ| = {max_rho:.2f}) suggests partial information overlap."
            action = "If these two criteria are conceptually similar, consider a correlation-aware weighting method such as CRITIC or simplify the active criterion set."
        else:
            academic = f"<strong>{cv_crit}</strong> has the highest coefficient of variation (Coeff. of Variation ≈{cv_val:.2f}), meaning it contributes the strongest discrimination in the ranking problem. Maximum inter-criterion correlation remains limited at |ρ| = {max_rho:.2f}, so redundancy risk is controlled."
            action = "The current data profile is appropriate for objective weighting and multi-method reading without an obvious structural warning."
        return _compose_commentary(simple, academic, action, example)
    simple = f"Bu veri setinde <strong>{n_alt} seçenek</strong>, <strong>{n_crit} kritere</strong> göre karşılaştırılıyor. Seçenekleri birbirinden en çok ayıran kriter <strong>{cv_crit}</strong> görünüyor."
    example = (
        f"Basit bir örnekle: <strong>{c1}</strong> ve <strong>{c2}</strong> neredeyse birlikte hareket ediyorsa sistem aynı bilgiyi iki kez okuyor olabilir."
        if max_rho > 0.75 else
        "Basit bir örnekle: kriterler birbirinin kopyası gibi davranmıyorsa sistem probleme farklı açılardan bakabilir."
    )
    if max_rho > 0.75:
        academic = f"<strong>{cv_crit}</strong> kriteri varyasyon katsayısı bakımından en yüksek ayrıştırıcı gücü sergiliyor (Varyasyon Katsayısı ≈{cv_val:.2f}). Bununla birlikte <strong>{c1}</strong> ile <strong>{c2}</strong> arasındaki yüksek ilişki (|ρ| = {max_rho:.2f}), kısmi bilgi tekrarına işaret ediyor."
        action = "Bu iki kriter kavramsal olarak da benzerse CRITIC gibi korelasyon duyarlı bir ağırlıklandırma tercih edin veya aktif kriter setini sadeleştirin."
    else:
        academic = f"<strong>{cv_crit}</strong> kriteri varyasyon katsayısı bakımından en güçlü ayrıştırıcı sinyali taşıyor (Varyasyon Katsayısı ≈{cv_val:.2f}). Kriterler arası en yüksek ilişki |ρ| = {max_rho:.2f} düzeyinde kaldığı için tekrar riski sınırlı görünüyor."
        action = "Mevcut veri profili, belirgin bir yapısal uyarı vermeden objektif ağırlıklandırma ve çoklu yöntem okumasına uygundur."
    return _compose_commentary(simple, academic, action, example)

def gen_weight_commentary(result: Dict[str, Any]) -> str:
    w_method = result["weights"]["method"]
    w_method_disp = method_display_name(w_method)
    w_vals = result["weights"]["values"]
    top_k = max(w_vals, key=w_vals.get)
    top_v = float(w_vals[top_k])
    concentration = sum(v**2 for v in w_vals.values())
    mode = result.get("weights", {}).get("details", {}).get("mode", "objective")
    if st.session_state.get("ui_lang") == "EN":
        if mode == "equal":
            simple = "All criteria were given equal importance. In other words, the model was instructed not to privilege any one criterion over the others."
            academic = f"This is a normative baseline rather than a data-derived weighting scheme. It is useful when you want the ranking to reflect balanced treatment across all criteria."
            example = "Example: if price, quality, and speed are all treated equally, none of them gets special leverage in the final score."
            action = "Use this mode when you want a neutral baseline, then compare it with an objective weighting result to see whether the story changes."
        elif mode == "manual":
            simple = f"The final ranking follows the priorities you entered manually. The strongest emphasis was placed on <strong>{top_k}</strong>."
            academic = f"Manual weighting preserves user judgment and then normalizes it for computation. The highest declared emphasis is <strong>{top_k}</strong> with normalized weight {top_v:.4f}."
            example = "Example: if you care much more about reliability than cost, the system will deliberately push reliability to the front of the decision."
            action = "Use this mode when domain preference is intentional, but validate the outcome with sensitivity analysis because the result depends directly on your preference structure."
        else:
            simple = f"The system says the most influential criterion is <strong>{top_k}</strong>. That means changes in this criterion will shape the final ranking more than the others."
            example = f"Example: if <strong>{top_k}</strong> worsens for a leading option, that option may lose ground faster than expected."
            if top_v > 0.35:
                academic = f"<strong>{w_method_disp}</strong> assigns the highest weight to <strong>{top_k}</strong> (w = {top_v:.4f}). This is a concentrated structure, so the final ranking may depend strongly on one dominant signal."
                action = f"Interpret the ranking with explicit attention to <strong>{top_k}</strong> and examine how sensitive the result is when this weight changes."
            else:
                academic = f"<strong>{w_method_disp}</strong> assigns the highest weight to <strong>{top_k}</strong> (w = {top_v:.4f}), while the overall concentration index remains moderate at ≈ {concentration:.2f}. This indicates a relatively balanced weighting profile."
                action = f"The weighting structure is reasonably distributed; still, prioritize <strong>{top_k}</strong> in sensitivity checks because it remains the main driver."
        return _compose_commentary(simple, academic, action, example)
    if mode == "equal":
        simple = "Tüm kriterlere eşit önem verildi. Yani sistem, hiçbir kriteri diğerlerinden özellikle daha baskın kabul etmedi."
        academic = "Bu yaklaşım veriden türetilmiş değil, bilinçli bir denge varsayımıdır. Tüm kriterlerin eşit söz hakkına sahip olduğu nötr bir başlangıç senaryosu sunar."
        example = "Örnek: fiyat, kalite ve hız aynı önemde kabul edilirse son puanda hiçbiri özel ayrıcalık kazanmaz."
        action = "Bu modu nötr referans senaryosu olarak kullanın; ardından objektif ağırlıklandırma ile karşılaştırıp hikayenin değişip değişmediğine bakın."
    elif mode == "manual":
        simple = f"Son sıralama, sizin girdiğiniz önceliklere göre şekillendi. En güçlü vurgu <strong>{top_k}</strong> üzerinde kaldı."
        academic = f"Manuel ağırlıklandırma kullanıcı tercihlerini korur ve hesaplama için normalize eder. En yüksek kullanıcı önceliği <strong>{top_k}</strong> kriterinde {top_v:.4f} normalize ağırlıkla görülmektedir."
        example = "Örnek: güvenilirliği maliyetten çok daha önemli görüyorsanız sistem bu tercihi özellikle öne taşır."
        action = "Alan bilgisini bilinçli biçimde yansıtmak istiyorsanız bu mod uygundur; ancak sonuç doğrudan tercih yapınıza bağlı olduğu için duyarlılık analizi ile birlikte okunmalıdır."
    else:
        simple = f"Sistem en çok <strong>{top_k}</strong> kriterini önemsemiş görünüyor. Yani nihai sıralamayı en fazla etkileyen başlık şu anda bu kriter."
        example = f"Örnek: <strong>{top_k}</strong> değeri kötüleşirse önde görünen bir seçenek beklenenden daha hızlı geriye düşebilir."
        if top_v > 0.35:
            academic = f"<strong>{w_method}</strong> yöntemi <strong>{top_k}</strong> kriterine en yüksek ağırlığı atamıştır (w = {top_v:.4f}). Bu yoğun yapı, nihai sıralamanın tek bir baskın sinyale duyarlı olabileceğini düşündürür."
            action = f"Sonucu yorumlarken <strong>{top_k}</strong> kriterini özellikle izleyin ve bu ağırlık değiştiğinde sıralamanın ne kadar oynadığını mutlaka kontrol edin."
        else:
            academic = f"<strong>{w_method}</strong> yöntemi <strong>{top_k}</strong> kriterini en güçlü sinyal olarak işaret etse de genel yoğunlaşma indeksi ≈ {concentration:.2f} düzeyinde kalmıştır. Bu, ağırlık yapısının görece dengeli olduğunu gösterir."
            action = f"Ağırlık profili dengeli görünüyor; yine de ana sürükleyici kriter olarak <strong>{top_k}</strong> üzerinde duyarlılık kontrolü yapmak yerinde olacaktır."
    return _compose_commentary(simple, academic, action, example)

def gen_ranking_commentary(result: Dict[str, Any], alt_names: Dict[str, str] = None) -> str:
    rt = result["ranking"]["table"]
    if rt is None or rt.empty:
        return tt("Sıralama tablosu mevcut değil.", "Ranking table is not available.")
    method   = result["ranking"]["method"]
    alt_col  = col_key(rt, "Alternatif", "Alternative")
    score_col = col_key(rt, "Skor", "Score")
    top_alt  = rt.iloc[0][alt_col]
    top_disp = (alt_names or {}).get(str(top_alt), str(top_alt))
    n_alt    = len(rt)
    score_range = float(rt[score_col].max()) - float(rt[score_col].min())
    second_disp = top_disp
    gap = 0.0
    if len(rt) > 1:
        second_alt = rt.iloc[1][alt_col]
        second_disp = (alt_names or {}).get(str(second_alt), str(second_alt))
        gap = float(rt.iloc[0][score_col]) - float(rt.iloc[1][score_col])
    method_simple, method_academic, method_action = _ranking_method_commentary(result, str(top_alt), top_disp)
    if st.session_state.get("ui_lang") == "EN":
        simple = f"Among <strong>{n_alt} options</strong>, the current leader is <strong>{top_disp}</strong>. {method_simple}"
        example = (
            f"A practical example: if the gap between the first and second options is small, the final choice may still depend on field knowledge. Here the top-two gap is only {gap:.4f} between <strong>{top_disp}</strong> and <strong>{second_disp}</strong>."
            if gap < 0.05 else
            f"A practical example: when the leader pulls away clearly, the model is not just saying 'first place' but also 'noticeably ahead'. Here <strong>{top_disp}</strong> separates from <strong>{second_disp}</strong> by {gap:.4f}."
        )
        if score_range < 0.05:
            academic = f"{method_academic} The overall score range remains narrow at {score_range:.4f}, so ranking uncertainty is still material."
        else:
            academic = f"{method_academic} The overall score range is {score_range:.4f}, which indicates meaningful separation across alternatives."
        action = (
            f"{method_action} Because the top-two distance is small, confirm the final choice with Monte Carlo robustness and method comparison."
            if gap < 0.05 else
            f"{method_action} Use the ranking table and method-specific visual to communicate why the leading option is substantively ahead."
        )
        return _compose_commentary(simple, academic, action, example)
    simple = f"<strong>{n_alt} seçenek</strong> içinde şu an öne çıkan alternatif <strong>{top_disp}</strong>. {method_simple}"
    example = (
        f"Pratik bir örnek: ilk iki seçenek birbirine çok yakınsa son karar hâlâ saha bilgisine bağlı kalabilir. Burada <strong>{top_disp}</strong> ile <strong>{second_disp}</strong> arasındaki fark yalnızca {gap:.4f}."
        if gap < 0.05 else
        f"Pratik bir örnek: lider seçenek belirgin biçimde ayrıştığında model sadece 'birinci' demiyor, aynı zamanda 'gözle görülür şekilde önde' diyor. Burada <strong>{top_disp}</strong>, <strong>{second_disp}</strong> karşısında {gap:.4f} fark yaratıyor."
    )
    if score_range < 0.05:
        academic = f"{method_academic} Genel skor aralığı {score_range:.4f} gibi dar bir bantta kaldığı için sıralama belirsizliği hâlâ anlamlı düzeydedir."
    else:
        academic = f"{method_academic} Genel skor aralığı {score_range:.4f} düzeyindedir; bu da alternatifler arasında anlamlı bir ayrışma olduğunu destekler."
    action = (
        f"{method_action} İlk iki seçenek yakın olduğu için nihai kararı vermeden önce Monte Carlo ve yöntem karşılaştırması sonuçlarını birlikte okuyun."
        if gap < 0.05 else
        f"{method_action} Liderliğin neden oluştuğunu göstermek için sıralama tablosu ile yöntem-özel grafiği birlikte kullanın."
    )
    return _compose_commentary(simple, academic, action, example)

def gen_comparison_commentary(result: Dict[str, Any]) -> str:
    comp = result.get("comparison", {})
    if not comp or "spearman_matrix" not in comp:
        return tt("Yöntem karşılaştırması verisi mevcut değil.", "Method comparison data is not available.")
    spdf = comp["spearman_matrix"]
    if not isinstance(spdf, pd.DataFrame) or spdf.shape[0] < 2:
        return tt("Karşılaştırma için en az iki yöntem gereklidir.", "At least two methods are required for comparison.")
    method_col = col_key(spdf, "Yöntem", "Method")
    mat = spdf.set_index(method_col)
    mask = ~np.eye(mat.shape[0], dtype=bool)
    off_diag = mat.where(mask).stack()
    mean_rho = float(off_diag.mean()) if len(off_diag) > 0 else 1.0
    n_meth = mat.shape[0]
    if st.session_state.get("ui_lang") == "EN":
        simple = f"We compared <strong>{n_meth} methods</strong> to see whether they tell a similar story. The average agreement is <strong>ρ = {mean_rho:.3f}</strong>."
        example = (
            "In plain terms: if different methods still point to a similar winner, confidence in the decision grows."
            if mean_rho >= 0.70 else
            "In plain terms: if methods disagree strongly, the result depends on the method family you chose."
        )
        if mean_rho >= 0.85:
            academic = "This is high inter-method agreement, indicating that the ranking structure is stable across methodological families."
            action = "You can report the leading result with stronger confidence, while still attaching the Spearman matrix for transparency."
        elif mean_rho >= 0.70:
            academic = "Agreement is moderate. This usually means the upper ranks are relatively stable while lower positions remain method-sensitive."
            action = "Report the common message across methods and avoid over-claiming fine differences in the lower part of the ranking."
        else:
            academic = "Agreement is low, which indicates methodological sensitivity rather than a single clear ranking narrative."
            action = "Revisit weighting and normalization choices, and give more attention to compromise-oriented methods before drawing a firm conclusion."
        return _compose_commentary(simple, academic, action, example)
    simple = f"<strong>{n_meth} yöntemi</strong> yan yana koyduğumuzda ortalama uyum <strong>ρ = {mean_rho:.3f}</strong> çıktı. Yani farklı yöntemler karar hikayesini ne kadar benzer okuyor, bunu görüyoruz."
    example = (
        "Basitçe: farklı yöntemler yine benzer kazananı gösteriyorsa sonuca güven artar."
        if mean_rho >= 0.70 else
        "Basitçe: yöntemler birbirinden çok farklı konuşuyorsa sonuç, seçtiğiniz yöntem ailesine daha bağımlı hale gelir."
    )
    if mean_rho >= 0.85:
        academic = "Bu yüksek yöntemler-arası uyum, sıralama yapısının yöntem ailesinden bağımsız biçimde kararlı kaldığını gösterir."
        action = "Ana bulguyu daha güçlü savunabilirsiniz; yine de Spearman matrisini şeffaflık için rapora ekleyin."
    elif mean_rho >= 0.70:
        academic = "Uyum orta düzeydedir. Bu durumda özellikle üst sıralar görece kararlı kalırken alt sıralarda yöntem duyarlılığı devam edebilir."
        action = "Yöntemlerin ortak mesajını öne çıkarın; alt sıralardaki ince farkları kesin gerçek gibi sunmayın."
    else:
        academic = "Uyum düşüktür; bu durum tek bir kesin sıralama anlatısından çok metodolojik duyarlılığa işaret eder."
        action = "Ağırlıklandırma ve normalleştirme tercihlerini yeniden gözden geçirin; kesin sonuç vermeden önce uzlaşı odaklı yöntemlere daha fazla dikkat edin."
    return _compose_commentary(simple, academic, action, example)

def gen_mc_commentary(result: Dict[str, Any]) -> str:
    sens = result.get("sensitivity")
    if not sens:
        return tt("Sağlamlık analizi verisi mevcut değil.", "Robustness analysis data is not available.")
    stab = float(sens.get("top_stability", 0.0))
    mc_df = sens.get("monte_carlo_summary")
    n_iter = int(sens.get("n_iterations", 0))
    base_top = str(sens.get("base_top", "—"))
    stable_leader = base_top
    challenger = "—"
    if isinstance(mc_df, pd.DataFrame) and not mc_df.empty:
        mc_alt_col = col_key(mc_df, "Alternatif", "Alternative")
        stable_leader = str(mc_df.iloc[0][mc_alt_col])
        if len(mc_df) > 1:
            challenger = str(mc_df.iloc[1][mc_alt_col])
    local_df = sens.get("local_scenarios")
    sensitive_crit = "—"
    sensitive_delta = "—"
    sensitive_rho = np.nan
    if isinstance(local_df, pd.DataFrame) and not local_df.empty:
        rho_col = col_key(local_df, "SpearmanRho", "SpearmanRho")
        crit_col = col_key(local_df, "Kriter", "Criterion")
        delta_col = col_key(local_df, "AğırlıkDeğişimi", "WeightChange")
        idx = pd.to_numeric(local_df[rho_col], errors="coerce").idxmin()
        if pd.notna(idx):
            sensitive_crit = str(local_df.loc[idx, crit_col])
            sensitive_delta = str(local_df.loc[idx, delta_col])
            sensitive_rho = float(pd.to_numeric(local_df.loc[idx, rho_col], errors="coerce"))

    is_en = st.session_state.get("ui_lang") == "EN"
    did = (
        f"We reran the ranking under {n_iter:,} Monte Carlo weight perturbations and also checked local ±10% / ±20% changes for each criterion."
        if is_en else
        f"Sıralama, {n_iter:,} adet Monte Carlo ağırlık bozma senaryosu altında yeniden çalıştırıldı; ayrıca her kriter için yerel ±%10 / ±%20 değişim etkisi de kontrol edildi."
    )
    why = (
        "The goal is to see whether the reported winner survives reasonable weight uncertainty and to identify which criterion is most likely to destabilize the ranking."
        if is_en else
        "Amaç, raporlanan kazananın makul ağırlık belirsizliği altında korunup korunmadığını ve sıralamayı en çok hangi kriterin bozduğunu görmektir."
    )
    if stab >= 0.80:
        stability_text = tt("Dayanıklılık düzeyi yüksektir.", "Robustness is high.")
        action = (
            "You can present the leader with stronger confidence; still keep the local sensitivity result as supporting evidence in the discussion."
            if is_en else
            "Lider alternatifi daha güçlü güvenle sunabilirsiniz; yine de yerel duyarlılık sonucunu tartışma bölümünde destekleyici kanıt olarak koruyun."
        )
    elif stab >= 0.60:
        stability_text = tt("Dayanıklılık düzeyi ortadır.", "Robustness is moderate.")
        action = (
            "Present the leader with caution and report the challenger together with the most sensitive criterion."
            if is_en else
            "Lider alternatifi temkinli sunun ve en hassas kriterle birlikte yakın rakibi de raporlayın."
        )
    else:
        stability_text = tt("Dayanıklılık düzeyi düşüktür.", "Robustness is low.")
        action = (
            f"Do not rely on a single winner yet; compare {stable_leader} and {challenger}, and revisit the weighting logic."
            if is_en else
            f"Henüz tek bir kazanana güvenmeyin; {stable_leader} ile {challenger} seçeneklerini birlikte karşılaştırın ve ağırlık mantığını yeniden gözden geçirin."
        )

    found = (
        f"{stability_text} The original leader {base_top} remains first in {stab*100:.1f}% of the simulations, while the most frequent simulated leader is {stable_leader}. "
        + (
            f"The main challenger is {challenger}. "
            if challenger != "—" else ""
        )
        + (
            f"The most sensitive local scenario occurs in criterion {sensitive_crit} at {sensitive_delta}, where rank agreement falls to ρ = {sensitive_rho:.3f}. "
            if np.isfinite(sensitive_rho) else ""
        )
        + "This means the reported ranking is being tested both globally and criterion by criterion."
        if is_en else
        f"{stability_text} Başlangıç lideri {base_top}, simülasyonların %{stab*100:.1f} bölümünde birinci kalmıştır; simülasyonlarda en sık öne çıkan lider ise {stable_leader} olmuştur. "
        + (
            f"Başlıca yakın rakip {challenger} olarak görünmektedir. "
            if challenger != "—" else ""
        )
        + (
            f"Yerel olarak en hassas senaryo {sensitive_crit} kriterinde {sensitive_delta} değişimde oluşmuş ve sıra uyumu ρ = {sensitive_rho:.3f} düzeyine kadar düşmüştür. "
            if np.isfinite(sensitive_rho) else ""
        )
        + "Bu tablo, sıralamanın hem genel belirsizlik altında hem de kriter bazında tek tek sınandığını gösterir."
    )
    example = (
        f"For example, if a hospital stays first in 85 out of 100 simulations, we say the ranking is durable. If another option frequently takes over when one criterion moves, the conclusion should be read more cautiously. Here the most pressure comes from {sensitive_crit}."
        if is_en else
        f"Örneğin bir hastane 100 simülasyonun 85'inde yine birinci kalıyorsa sıralama dayanıklıdır deriz. Bir kriter oynadığında başka bir seçenek sık sık öne geçiyorsa sonuç daha temkinli okunmalıdır. Bu çalışmada en büyük baskı {sensitive_crit} kriterinden gelmektedir."
    )
    return _compose_structured_commentary(did, why, found, example, action)

# ---------------------------------------------------------
# GLOBAL BANNER
# ---------------------------------------------------------
st.markdown(
    f"""
    <div class="global-header">
        <div class="header-title">MCDM- Profesyonel Karar Destek Sistemi</div>
        <div class="header-professor">Prof. Dr. Ömer Faruk Rençber</div>
        <div class="header-meta">
            <div class="header-dedication">{tt("Çocuklarım M. Eymen ve H. Serra'ya İthafen..", "Dedicated to My Children M. Eymen and H. Serra..")}</div>
            <div class="header-url"><a href="https://www.ofrencber.com" target="_blank">www.ofrencber.com</a></div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# SOL PANEL: KONTROL, AMAC VE TEMİZLİK
# ---------------------------------------------------------
with st.sidebar:
    with open("logo.png", "rb") as _logo_file:
        _sidebar_logo_b64 = base64.b64encode(_logo_file.read()).decode("utf-8")
    st.markdown(
        f"""
        <div class="sidebar-brand">
            <img src="data:image/png;base64,{_sidebar_logo_b64}" class="sidebar-brand-logo" alt="MCDM logo">
            <div class="sidebar-brand-name">Prof. Dr. Ömer Faruk Rençber</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _lang_idx = 1 if st.session_state.get("ui_lang") == "EN" else 0
    _lang_pick = st.radio(
        "Language",
        ["Türkçe", "English"],
        index=_lang_idx,
        horizontal=True,
        label_visibility="collapsed",
    )
    _new_lang = "EN" if _lang_pick == "English" else "TR"
    if _new_lang != st.session_state.get("ui_lang"):
        st.session_state["ui_lang"] = _new_lang
        st.rerun()

    with st.expander(tt("📘 Uygulama Rehberi", "📘 User Guide"), expanded=False):
        st.markdown(
            f"<p style='font-size:0.78rem; line-height:1.55; color:#2C2C2C; margin:0;'>"
            f"{tt('1️⃣ Sol panelden veriyi yükleyin veya örnek veriyi kullanın.<br>'
                  '2️⃣ Sağ panelde analiz amacını seçin.<br>'
                  '3️⃣ Kriterlerin yönünü (fayda/maliyet) doğrulayın.<br>'
                  '4️⃣ Ağırlıklandırma ve sıralama yöntemlerini belirleyin.<br>'
                  '5️⃣ &#34;Analiz Zamanı&#34; butonuna basın.<br>'
                  '6️⃣ Sonuçları sekmeler arasında inceleyin ve raporu indirin.',
                  '1️⃣ Upload data from the left panel or use sample data.<br>'
                  '2️⃣ Select your analysis objective in the right panel.<br>'
                  '3️⃣ Confirm criterion directions (benefit/cost).<br>'
                  '4️⃣ Select weighting and ranking methods.<br>'
                  '5️⃣ Click &#34;Run Analysis&#34;.<br>'
                  '6️⃣ Explore results across tabs and download the report.')}"
            f"</p>",
            unsafe_allow_html=True,
        )

    with st.expander(tt("🧭 Metodolojik Yardım", "🧭 Methodological Help"), expanded=False):
        st.markdown(
            f"<p style='font-size:0.78rem; line-height:1.55; color:#2C2C2C; margin:0;'>"
            f"{tt('<b>Objektif Ağırlık Yöntemleri:</b> Kriter önemini verinin kendi yapısından türetir; araştırmacı müdahalesi gerekmez.<br><br>'
                  '<b>Klasik ÇKKV:</b> Kesin sayısal veriler için tasarlanmış; TOPSIS, VIKOR, EDAS vb. 17 yöntem.<br><br>'
                  '<b>Fuzzy ÇKKV:</b> Ölçüm belirsizliği varsa tercih edilir; üçgensel bulanık sayılar kullanır.<br><br>'
                  '<b>Monte Carlo:</b> Ağırlıklar rastgele bozularak sıralamanın ne kadar kararlı olduğu test edilir.',
                  '<b>Objective Weighting:</b> Derives criterion importance from data structure; no researcher intervention needed.<br><br>'
                  '<b>Classical MCDM:</b> Designed for crisp numerical data; 17 methods incl. TOPSIS, VIKOR, EDAS.<br><br>'
                  '<b>Fuzzy MCDM:</b> Preferred when measurement uncertainty exists; uses triangular fuzzy numbers.<br><br>'
                  '<b>Monte Carlo:</b> Weights are randomly perturbed to test how stable the ranking is.')}"
            f"</p>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div style='margin-top:0.6rem; padding:0.5rem 0.65rem; background:#EAF3FB; border-left:3px solid #5A9CC5; border-radius:4px;'>"
            f"<p style='font-size:0.75rem; font-weight:700; color:#1F5F9A; margin:0 0 0.3rem 0;'>"
            f"🎯 {tt('Yöntem Önerisi Nasıl Yapılır?', 'How Are Methods Recommended?')}"
            f"</p>"
            f"<p style='font-size:0.74rem; line-height:1.5; color:#2C3E50; margin:0;'>"
            f"{tt('<b>Yüksek korelasyon</b> → CRITIC veya PCA ağırlıklandırma; uzlaşı için VIKOR önerilir.<br>'
                  '<b>Yüksek değişkenlik (varyans)</b> → Entropi veya Standart Sapma ağırlıklandırma öne çıkar.<br>'
                  '<b>Dengeli veri yapısı</b> → TOPSIS, VIKOR, EDAS güvenle uygulanabilir.<br>'
                  '<b>Az alternatif (&lt;6)</b> → Yöntem seçimi daha kritik; Monte Carlo ihtiyatla yorumlanmalı.<br>'
                  '<b>Geniş alternatif seti</b> → Mesafe tabanlı yöntemler (TOPSIS, EDAS) özellikle etkin.',
                  '<b>High correlation</b> → CRITIC or PCA weighting; VIKOR recommended for consensus.<br>'
                  '<b>High dispersion (variance)</b> → Entropy or Standard Deviation weighting stands out.<br>'
                  '<b>Balanced data structure</b> → TOPSIS, VIKOR, EDAS can be applied safely.<br>'
                  '<b>Few alternatives (&lt;6)</b> → Method choice is more critical; interpret Monte Carlo cautiously.<br>'
                  '<b>Large alternative set</b> → Distance-based methods (TOPSIS, EDAS) are especially effective.')}"
            f"</p></div>",
            unsafe_allow_html=True,
        )

    is_data_loaded = st.session_state.get("raw_data") is not None

    if st.button(tt("🔄 Yeni Analize Başla (Sıfırla)", "🔄 Start New Analysis (Reset)"), use_container_width=True):
        reset_all()
        st.rerun()

    # ── Veri Girişi ──
    with st.expander(
        f"📂 {tt('Veri Girişi', 'Data Input')} {'✅' if is_data_loaded else '⏳'}",
        expanded=not is_data_loaded,
    ):
        uploaded = st.file_uploader(tt("CSV veya XLSX yükleyin", "Upload CSV or XLSX"), type=["csv", "xlsx"], label_visibility="collapsed")
        if st.button(tt("📘 Örnek Veri Kullan", "📘 Use Sample Data"), use_container_width=True):
            st.session_state["raw_data"] = sample_dataset()
            st.session_state["data_source_id"] = "sample_data"
            st.session_state.prep_done = False
            st.session_state["step1_done"] = False
            st.rerun()

        if uploaded is not None and st.session_state.get("data_source_id") != uploaded.name:
            st.session_state["raw_data"] = load_uploaded_file(uploaded)
            st.session_state["data_source_id"] = uploaded.name
            st.session_state.prep_done = False
            st.session_state["step1_done"] = False
            st.rerun()

    # ── Veri Ön İşleme ──
    if is_data_loaded:
        st.markdown(f"<div class='sb-section-label'>🧹 {tt('Veri Ön İşleme', 'Data Preprocessing')}</div>", unsafe_allow_html=True)

        # Session state başlat
        if "impute_mode_open" not in st.session_state:
            st.session_state["impute_mode_open"] = False
        if "missing_strategy_saved" not in st.session_state:
            st.session_state["missing_strategy_saved"] = "Sil"

        impute_checked = st.checkbox(
            tt("Eksik Veri Tamamla", "Impute Missing Values"),
            value=st.session_state["impute_mode_open"],
            key="cb_impute_mode",
        )
        if impute_checked != st.session_state["impute_mode_open"]:
            st.session_state["impute_mode_open"] = impute_checked
            st.rerun()

        if impute_checked:
            _impute_method = st.selectbox(
                tt("Tamamlama yöntemi", "Imputation method"),
                [tt("Medyan", "Median"), tt("Ortalama", "Mean"), tt("Interpolasyon", "Interpolation"), tt("Sıfır", "Zero")],
                key="impute_method_select",
            )
            if st.button(tt("✅ Uygula", "✅ Apply"), use_container_width=True, key="btn_impute_apply"):
                st.session_state["missing_strategy_saved"] = _impute_method
                st.rerun()
            _strat_label = st.session_state["missing_strategy_saved"]
            st.caption(tt(f"Aktif yöntem: {_strat_label}", f"Active method: {_strat_label}"))

        missing_strategy = st.session_state["missing_strategy_saved"] if impute_checked else "Sil"
        clip_outliers = st.checkbox(tt("Aykırı Değerleri (Outlier) Temizle", "Clean Outliers"), value=False)
        st.caption(tt("Yalnız sayısal sütunlara uygulanır.", "Applied to numeric columns only."))
    else:
        missing_strategy, clip_outliers = "Sil", False

    # ── Aşama Göstergesi ──
    st.markdown("<div style='margin-top:0.5rem;'></div>", unsafe_allow_html=True)
    _step_data_done = is_data_loaded
    _step_prep_done = bool(st.session_state.get("prep_done"))
    _step_result_done = st.session_state.get("analysis_result") is not None
    _step_method_active = _step_data_done and _step_prep_done and not _step_result_done
    st.markdown(
        f"""
        <div class="stepper-wrap">
            <div class="step-item"><span class="step-dot {'step-done' if _step_data_done else 'step-pending'}"></span> {tt('1. Adım', 'Step 1')} · {tt('Veri girişi', 'Data input')}{' ✅' if _step_data_done else ''}</div>
            <div class="step-item"><span class="step-dot {'step-done' if _step_prep_done else ('step-active' if _step_data_done else 'step-pending')}"></span> {tt('2. Adım', 'Step 2')} · {tt('Kriter doğrulama', 'Criteria validation')}{' ✅' if _step_prep_done else ''}</div>
            <div class="step-item"><span class="step-dot {'step-done' if _step_result_done else ('step-active' if _step_method_active else 'step-pending')}"></span> {tt('3. Adım', 'Step 3')} · {tt('Yöntem ve analiz', 'Method and analysis')}{' ✅' if _step_result_done else ''}</div>
            <div class="step-item"><span class="step-dot {'step-done' if _step_result_done else 'step-pending'}"></span> {tt('4. Adım', 'Step 4')} · {tt('Sonuç ve rapor', 'Results and report')}{' ✅' if _step_result_done else ''}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Sidebar Footer ──
    st.markdown(
        "<div class='sidebar-footer'>Prof. Dr. Ömer Faruk Rençber</div>",
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------
# ANA GÖVDE
# ---------------------------------------------------------
raw_data = st.session_state.get("raw_data")
if raw_data is None:
    st.markdown(
        f"""
        <div style="text-align:center;padding:4.5rem 1rem 3rem 1rem;">
            <h3 style="color:#1B365D;font-weight:700;font-size:1.6rem;margin-bottom:0.45rem;">{tt("Karar Destek Sistemine Hoş Geldiniz", "Welcome to the Decision Support System")}</h3>
            <p style="color:#64748B;font-size:1.02rem;max-width:760px;margin:0 auto;line-height:1.65;">{tt("Başlamak için sol panelden veri setinizi yükleyin veya örnek veriyi kullanın. Veri yüklendiğinde analiz amacı, kriter doğrulama ve yöntem seçimi adımları burada görünür olacaktır.", "To begin, upload your dataset from the left panel or use the sample data. Once data is loaded, the analysis objective, criteria validation, and method selection steps will appear here.")}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()
_needs_ranking_default = bool(st.session_state.get("needs_ranking", True))
_purpose_options = [
    tt("⚖️ Kriterleri önem düzeyine göre sıralamak (Yalnızca Ağırlık)", "⚖️ Rank criteria by importance level (Weights Only)"),
    tt("🏆 Alternatifleri kriterler ile birlikte sıralamak (Ağırlık + Sıralama)", "🏆 Rank alternatives together with criteria (Weights + Ranking)"),
]
_purpose_default_idx = 1 if _needs_ranking_default else 0
_step1_done = bool(st.session_state.get("step1_done"))
_step1_label = f"🎯 {tt('1. Adım: Analizi Amacınız Ne?', 'Step 1: What Is Your Analysis Objective?')}{' ✅' if _step1_done else ''}"

with st.expander(_step1_label, expanded=(raw_data is not None) and not _step1_done):
    st.caption(tt("Bu adımda analiz hedefinizi ve veri yapınızı belirleyin.", "In this step, define your analysis objective and data structure."))
    st.caption(tt("Tek yıl veya panel veri seçimi, sonraki hesaplama ve çıktı akışını doğrudan değiştirir.", "Your single-year or panel-data choice directly changes the next calculation and output flow."))
    st.markdown(
        f"<p style='font-size:0.82rem; color:#5C5650; margin:0 0 0.5rem 0;'>"
        f"{tt('Lütfen bu analizde yapmak istediğiniz işlemi seçin:', 'Please select what you want to accomplish in this analysis:')}"
        f"</p>",
        unsafe_allow_html=True,
    )
    _purpose_choice = st.radio(
        tt("Amacınız nedir?", "What is your objective?"),
        _purpose_options,
        index=_purpose_default_idx,
        label_visibility="collapsed",
        horizontal=True,
    )
    _scope_options = [
        tt("📅 Tek Yıl", "📅 Single Year"),
        tt("🗂️ Panel Veri", "🗂️ Panel Data"),
    ]
    _scope_default = 1 if st.session_state.get("analysis_scope") == "panel" else 0
    _scope_choice = st.radio(
        tt("Veri yapınız nedir?", "What is your data structure?"),
        _scope_options,
        index=_scope_default,
        horizontal=True,
    )

    # ── Panel veri ayarları (sadece panel seçiliyse) ──
    if _scope_choice == _scope_options[1] and isinstance(raw_data, pd.DataFrame):
        _yr_candidates = _guess_year_columns(raw_data)
        _yr_col_opts   = [str(c) for c in raw_data.columns]
        if _yr_col_opts:
            _yr_col_default = st.session_state.get("panel_year_column")
            if _yr_col_default not in _yr_col_opts:
                _yr_col_default = _yr_candidates[0] if _yr_candidates else _yr_col_opts[0]
            # Başlık + yıl sütunu etiketi tek HTML satırında
            st.markdown(
                f'<div style="background:#EEF5FB; border:1px solid #C0D8EE; border-radius:7px;'
                f'padding:0.25rem 0.55rem; margin:0.3rem 0 0 0; display:flex; align-items:center; gap:0.5rem;">'
                f'<span style="font-size:0.67rem; font-weight:700; color:#1F4A73; white-space:nowrap;">'
                f'🗂️ {tt("Panel Ayarları", "Panel Settings")}</span>'
                f'<span style="font-size:0.64rem; color:#5A7A9A;">— {tt("Yıl sütunu:", "Year column:")}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            # Selectbox etiket gizli, hemen altında kompakt
            st.markdown('<div style="margin-top:-0.4rem;"></div>', unsafe_allow_html=True)
            _panel_col_inner = st.selectbox(
                tt("Yıl sütunu", "Year column"),
                _yr_col_opts,
                index=_yr_col_opts.index(_yr_col_default),
                key="panel_year_col_select",
                label_visibility="collapsed",
            )
            st.session_state["panel_year_column"] = _panel_col_inner
            _det_years = _sorted_panel_years(raw_data[_panel_col_inner])
            if _det_years:
                if "panel_selected_years" not in st.session_state or \
                        set(st.session_state.get("panel_selected_years_all", [])) != set(_det_years):
                    st.session_state["panel_selected_years"]     = list(_det_years)
                    st.session_state["panel_selected_years_all"] = list(_det_years)
                _sel_yrs_now = []
                st.markdown(
                    f'<p style="font-size:0.64rem; color:#4A6070; margin:0.1rem 0 0.05rem 0; line-height:1.2;">'
                    f'{tt("Dönemler:", "Periods:")}</p>',
                    unsafe_allow_html=True,
                )
                _yr_sel_cols = st.columns(min(10, len(_det_years)), gap="small")
                for _yi, _yr in enumerate(_det_years):
                    with _yr_sel_cols[_yi % min(10, len(_det_years))]:
                        _yr_key = f"panel_yr_cb_{_yr}"
                        _yr_def = _yr in st.session_state.get("panel_selected_years", _det_years)
                        if st.checkbox(str(_yr), value=_yr_def, key=_yr_key):
                            _sel_yrs_now.append(_yr)
                st.session_state["panel_selected_years"] = _sel_yrs_now
                if not _sel_yrs_now:
                    st.warning(tt("En az bir dönem seçin.", "Select at least one period."))
                elif len(_sel_yrs_now) < 2:
                    st.markdown(
                        f'<p style="font-size:0.63rem; color:#7A5018; margin:0.05rem 0 0 0;">'
                        f'⚠ {tt("Karşılaştırma için ≥2 dönem önerilir.", "≥2 periods recommended.")}</p>',
                        unsafe_allow_html=True,
                    )

    _flow_steps = [
        (tt("Amaç belirleme",        "Define Objective"),        "1",   _step1_done, "#2E7D52", "#D4EFE2"),
        (tt("Veri ön hazırlık",      "Data Preparation"),        "2",   False,       "#1F5F9A", "#DAEAF7"),
        (tt("Ağırlık metodu belirle","Set Weighting Method"),    "3.1", False,       "#B5681E", "#FDE8CC"),
    ]
    if _purpose_choice == _purpose_options[1]:
        _flow_steps.insert(3, (tt("Sıralama metodu belirle", "Set Ranking Method"), "3.2", False, "#9A3030", "#FAD9D9"))
    _flow_steps.append((tt("Analiz yap ve sonuçları yorumla", "Run Analysis & Interpret"), "4", False, "#6B4FA0", "#EAE0F7"))

    _step_nodes = []
    for _idx, (_label, _num, _done, _clr, _bg) in enumerate(_flow_steps):
        _is_last = _idx == len(_flow_steps) - 1
        if _done:
            _box_bg  = _clr
            _num_clr = "rgba(255,255,255,0.80)"
            _lbl_clr = "#FFFFFF"
            _bdr_clr = _clr
            _num_txt = "✓"
        else:
            _box_bg  = _bg
            _num_clr = _clr
            _lbl_clr = _clr
            _bdr_clr = _clr + "99"
            _num_txt = _num

        _node_html = (
            f'<div style="flex:1; min-width:90px; max-width:200px;">'
            f'<div style="background:{_box_bg}; border:1.5px solid {_bdr_clr}; border-radius:10px;'
            f'padding:0.35rem 0.55rem 0.4rem 0.55rem; text-align:center;'
            f'box-shadow:0 2px 8px rgba(0,0,0,0.07);">'
            f'<div style="font-size:0.60rem; font-weight:800; color:{_num_clr}; letter-spacing:0.4px; line-height:1.2; margin-bottom:2px;">{_num_txt}</div>'
            f'<div style="font-size:0.67rem; font-weight:600; color:{_lbl_clr}; line-height:1.35;">{_label}</div>'
            f'</div>'
            f'</div>'
        )
        _step_nodes.append(_node_html)

        if not _is_last:
            _step_nodes.append(
                '<div style="width:14px; flex-shrink:0; display:flex; align-items:center; justify-content:center;">'
                '<div style="width:100%; height:2px; background:linear-gradient(90deg,#B0C8DC,#C8D8E8);"></div>'
                '</div>'
            )

    st.markdown(
        f'<div style="display:flex; align-items:stretch; gap:0; width:100%;'
        f'background:linear-gradient(135deg,#F6FAFE,#EEF4FA);'
        f'border:1px solid #C8D8E8; border-radius:12px;'
        f'padding:0.6rem 0.75rem; margin:0.5rem 0 0.25rem 0;">'
        f'{"".join(_step_nodes)}</div>',
        unsafe_allow_html=True,
    )
needs_ranking = _purpose_choice == _purpose_options[1]
st.session_state["needs_ranking"] = needs_ranking
panel_mode = _scope_choice == _scope_options[1]
st.session_state["analysis_scope"] = "panel" if panel_mode else "single"

panel_year_col = None
if panel_mode and isinstance(raw_data, pd.DataFrame):
    _stored_yr_col = st.session_state.get("panel_year_column")
    _all_cols = [str(c) for c in raw_data.columns]
    if _stored_yr_col in _all_cols:
        panel_year_col = _stored_yr_col
    elif _all_cols:
        panel_year_col = _guess_year_columns(raw_data)[0] if _guess_year_columns(raw_data) else _all_cols[0]
        st.session_state["panel_year_column"] = panel_year_col

if raw_data is not None and not st.session_state.get("step1_done"):
    st.caption(tt("Seçiminizi yaptıysanız aşağıdaki butona tıklayın.", "If you made your selection, click the button below."))
    if st.button(
        tt("✨ Şimdi verilerimizi hazırlama zamanı.. Hazırsanız devam edelim", "✨ It is time to prepare our data.. If you are ready, let us continue"),
        use_container_width=True,
        key="step1_continue_btn",
    ):
        st.session_state["step1_done"] = True
        st.rerun()
    st.stop()

# ── Veri Hazırlığı ──
working = raw_data.copy()
working = clean_dataframe(working, missing_strategy, clip_outliers)
if panel_mode and panel_year_col and panel_year_col in raw_data.columns:
    working[panel_year_col] = raw_data[panel_year_col]
working.index = [f"A{idx+1}" for idx in range(len(working))]
numeric_cols = working.select_dtypes(include=[np.number]).columns.tolist()
if panel_mode and panel_year_col in numeric_cols:
    numeric_cols = [c for c in numeric_cols if c != panel_year_col]

if len(numeric_cols) < 2:
    st.error(tt("Hata: Analiz için en az iki sayısal sütun (kriter) gereklidir.", "Error: At least two numeric criterion columns are required for analysis."))
    st.stop()

_existing_crits = set(st.session_state.get("crit_dir", {}).keys())
if _existing_crits != set(numeric_cols):
    st.session_state["crit_dir"]     = {c: (guess_direction(c) == "Max (Fayda)") for c in numeric_cols}
    st.session_state["crit_include"] = {c: True for c in numeric_cols}


# ── TEK BİRLEŞİK HAZIRLIK PANELİ ──
# Diagnostics hesabı (render öncesi)
_sel_temp = [c for c in numeric_cols if st.session_state["crit_include"].get(c, True)]
_has_diag = len(_sel_temp) >= 2
if _has_diag:
    _diag = me.generate_data_diagnostics(working, _sel_temp, {c: ("max" if st.session_state["crit_dir"].get(c, True) else "min") for c in _sel_temp})
    _score, _label = _diag_score_and_label(_diag)
    _badge_cls = "diag-badge-good" if _score >= 80 else ("diag-badge-mid" if _score >= 60 else "diag-badge-bad")
    _rec_items = _diag.get("recommendations", [])[:3]
    _sugg_weight = method_display_name(_diag.get("suggested_weight") or "Entropy")
    _sugg_rank_methods = _diag.get("suggested_ranking_methods") or ["TOPSIS", "VIKOR", "EDAS"]
    _sugg_rank = ", ".join(_sugg_rank_methods[:3])
    _weight_reason = _diag.get("weight_rationale_en") if st.session_state.get("ui_lang") == "EN" else _diag.get("weight_rationale_tr")
    _rank_reason = _diag.get("ranking_rationale_en") if st.session_state.get("ui_lang") == "EN" else _diag.get("ranking_rationale_tr")
    st.session_state["_sugg_rank"] = _sugg_rank
    st.session_state["_sugg_rank_methods"] = _sugg_rank_methods[:3]
    while len(_rec_items) < 3:
        _rec_items.append({"icon": "•",
            "text": tt("Ek kritik bulgu yok.", "No additional critical finding."),
            "action": tt("Mevcut akışla devam edebilirsiniz.", "You may proceed with the current workflow.")})
else:
    _sugg_weight = method_display_name("Entropy")
    _sugg_rank_methods = ["TOPSIS", "VIKOR", "EDAS"]
    _sugg_rank = ", ".join(_sugg_rank_methods)
    _weight_reason = tt("Varsayılan olarak Entropy başlangıç önerisi kullanılıyor.", "Entropy is used as the default initial suggestion.")
    _rank_reason = tt("Varsayılan olarak TOPSIS, VIKOR ve EDAS başlangıç önerileri kullanılıyor.", "TOPSIS, VIKOR, and EDAS are used as the default initial suggestions.")

_prep_label = f"🔍 {tt('2. Adım: Ön inceleme - Veri Ön Hazırlığı', 'Step 2: Preliminary Review - Data Preparation')}{' ✅' if st.session_state.get('prep_done') else ''}"
with st.expander(_prep_label, expanded=not st.session_state.get("prep_done")):
    st.caption(tt("Bu adımda veri yapısını inceleyin, ilk görünümü kontrol edin ve kullanılacak kriterleri netleştirin.", "In this step, inspect the data structure, review the first preview, and clarify the criteria to be used."))
    st.caption(tt("Eksik, uygunsuz veya yanlış yönlü kriter bırakmamaya dikkat edin.", "Be careful not to leave missing, unsuitable, or wrongly directed criteria in the analysis."))

    # ── 1) Ön İnceleme Sonuçları ──
    if _has_diag:
        with st.expander(
            f"🧭 {tt(f'Ön İnceleme Sonuçları  —  %{_score} Veri Uygunluğu Hesaplanmıştır', f'Preliminary Review  —  {_score}% Data Suitability Calculated')}",
            expanded=False,
        ):
            st.markdown(
                f"""<div class="assistant-grid">
                    <div class="assistant-card2">
                        <div class="assistant-title2">📊 {tt("Veri Profili", "Data Profile")}</div>
                        <div class="assistant-body2">
                            {_diag.get('n_alt', 0)} {tt("alternatif", "alternatives")} · {_diag.get('n_crit', 0)} {tt("kriter", "criteria")}<br>
                            {tt("Ort. Varyasyon Katsayısı", "Avg. Coeff. of Variation")}: <strong>{_diag.get('mean_cv', 0.0):.2f}</strong> &nbsp;·&nbsp;
                            {tt("Maks. |ρ|", "Max |ρ|")}: <strong>{_diag.get('max_corr', 0.0):.2f}</strong>
                        </div>
                    </div>
                    <div class="assistant-card2">
                        <div class="assistant-title2">⚖️ {tt("Ağırlık Önerisi", "Weighting Suggestion")}</div>
                        <div class="assistant-body2">
                            {tt("Önerilen:", "Suggested:")} <strong>{_sugg_weight}</strong><br>
                            <span style="font-size:0.76rem; color:#556070;">{_weight_reason}</span>
                        </div>
                    </div>
                    <div class="assistant-card2">
                        <div class="assistant-title2">🏆 {tt("Sıralama Önerisi", "Ranking Suggestion")}</div>
                        <div class="assistant-body2">
                            {tt("Önerilen:", "Suggested:")} <strong>{_sugg_rank}</strong><br>
                            <span style="font-size:0.76rem; color:#556070;">{_rank_reason}</span>
                        </div>
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )
            with st.expander(f"📋 {tt('Detaylı Bulgular ve Öneriler', 'Detailed Findings & Recommendations')}", expanded=False):
                for idx, rec in enumerate(_rec_items, start=1):
                    _r_text, _r_action = diag_rec_text(rec)
                    st.markdown(
                        f"""<div class="assistant-card2" style="margin-bottom:0.4rem;">
                            <div class="assistant-title2">{rec.get("icon","•")} {tt("Bulgu","Finding")} {idx}</div>
                            <div class="assistant-body2">{_r_text}<br>
                            <strong>{tt("Öneri:","Recommendation:")}</strong> {_r_action}</div>
                        </div>""",
                        unsafe_allow_html=True,
                    )
                st.markdown(
                    f"""<div class="assistant-box" style="background:#F0F6FF; margin-top:0.5rem;">
                        <div class="assistant-title2">💡 {tt("Önerilen Yöntem Felsefesi","Recommended Method Philosophy")}</div>
                        <div class="assistant-body2">
                            <strong>{tt("Ağırlık:","Weighting:")}</strong> {_sugg_weight} &nbsp;·&nbsp;
                            <strong>{tt("Sıralama:","Ranking:")}</strong> {_sugg_rank}
                        </div>
                    </div>""",
                    unsafe_allow_html=True,
                )

    # ── 2) Veri Ön İzleme ──
    with st.expander(f"📋 {tt('Veri Ön İzleme', 'Data Preview')} — {tt('ilk 3 satır', 'first 3 rows')}", expanded=False):
        render_table(working.head(3))

    # ── 3) Kriter Yapılandırması ──
    with st.expander(f"⚙️ {tt('Kriter Yapılandırması', 'Criteria Configuration')}", expanded=False):
        st.markdown("""<style>
        .ct-wrap { border:1px solid #C8D8E8; border-radius:8px; overflow:hidden; margin-top:0.1rem; }
        .ct-head {
            display:grid; grid-template-columns:32px 1fr 170px;
            background:#E8F0F8; padding:0.18rem 0.6rem;
            font-size:0.67rem; font-weight:700; color:#3A5A78;
            text-transform:uppercase; letter-spacing:0.3px;
            border-bottom:2px solid #C8D8E8;
        }
        /* Row borders */
        .ct-wrap .stHorizontalBlock {
            border-bottom:1px solid #E0EAF2 !important;
            margin:0 !important; padding:0 0.35rem !important;
            align-items:center !important;
        }
        .ct-wrap .stHorizontalBlock:last-of-type { border-bottom:none !important; }
        .ct-wrap .stHorizontalBlock:nth-of-type(even) { background:#F6FAFD; }
        /* Compact columns */
        .ct-wrap [data-testid="column"] { padding:0.05rem 0.2rem !important; }
        .ct-wrap .element-container { margin:0 !important; padding:0 !important; }
        /* Checkbox */
        .ct-wrap .stCheckbox { margin:0 !important; min-height:1.6rem !important; }
        .ct-wrap .stCheckbox > label { padding:0 !important; }
        /* Criterion name */
        .ct-crit-name { font-size:0.79rem; font-weight:600; color:#1C1C1E; line-height:1.8; }
        /* Direction radio — pill style */
        .ct-wrap .stRadio > label { display:none !important; }
        .ct-wrap .stRadio > div[role="radiogroup"] {
            flex-direction:row !important; gap:4px !important;
            margin:0 !important; padding:0 !important; flex-wrap:nowrap !important;
        }
        .ct-wrap .stRadio label {
            padding:0.08rem 0.55rem !important;
            border-radius:12px !important;
            font-size:0.70rem !important; font-weight:600 !important;
            border:1.5px solid #ccc !important;
            cursor:pointer !important; white-space:nowrap !important;
            transition:background 0.15s, color 0.15s !important;
            line-height:1.7 !important;
        }
        /* Fayda pill — green tones */
        .ct-wrap .stRadio label:first-of-type {
            border-color:#27704A !important; color:#27704A !important; background:#fff !important;
        }
        .ct-wrap .stRadio label:first-of-type:has(input:checked) {
            background:#27704A !important; color:#fff !important;
        }
        /* Maliyet pill — red tones */
        .ct-wrap .stRadio label:last-of-type {
            border-color:#B02A2A !important; color:#B02A2A !important; background:#fff !important;
        }
        .ct-wrap .stRadio label:last-of-type:has(input:checked) {
            background:#B02A2A !important; color:#fff !important;
        }
        /* Hide radio circle dot */
        .ct-wrap .stRadio input[type="radio"] {
            position:absolute !important; opacity:0 !important; width:0 !important; height:0 !important;
        }
        </style>""", unsafe_allow_html=True)

        _dir_benefit = tt("⬆ Fayda", "⬆ Benefit")
        _dir_cost    = tt("⬇ Maliyet", "⬇ Cost")

        # Header
        st.markdown(
            f'<div class="ct-wrap"><div class="ct-head">'
            f'<span></span>'
            f'<span>{tt("Kriter", "Criterion")}</span>'
            f'<span>{tt("Yön", "Direction")}</span>'
            f'</div></div>',
            unsafe_allow_html=True,
        )
        st.markdown('<div class="ct-wrap">', unsafe_allow_html=True)
        for _c in numeric_cols:
            _rc = st.columns([0.35, 3.5, 2.1], gap="small")
            with _rc[0]:
                st.session_state["crit_include"][_c] = st.checkbox(
                    "", value=st.session_state["crit_include"].get(_c, True), key=f"inc_{_c}"
                )
            with _rc[1]:
                st.markdown(f'<div class="ct-crit-name">{_c}</div>', unsafe_allow_html=True)
            with _rc[2]:
                _is_benefit = st.session_state["crit_dir"].get(_c, True)
                _choice = st.radio(
                    "", [_dir_benefit, _dir_cost],
                    index=0 if _is_benefit else 1,
                    key=f"dir_{_c}", horizontal=True, label_visibility="collapsed",
                )
                st.session_state["crit_dir"][_c] = (_choice == _dir_benefit)
        st.markdown('</div>', unsafe_allow_html=True)

    if not st.session_state.get("prep_done"):
        st.divider()

if not st.session_state.get("prep_done"):
    st.markdown(f"**{tt('Harika.. şimdi de yöntemlerimizi seçelim', 'Great.. now let us choose our methods')}**")
    if st.button(tt("✅ Veri Ön İşleme Bitti (Yöntem Seçimine Geç)", "✅ Preprocessing Complete (Proceed to Method Selection)"), use_container_width=True):
        st.session_state.prep_done = True
        st.rerun()

criteria = [c for c in numeric_cols if st.session_state["crit_include"].get(c, True)]
criteria_types = {c: ("max" if st.session_state["crit_dir"].get(c, True) else "min") for c in criteria}

if len(criteria) < 2:
    st.error(tt("En az 2 kriter seçmelisiniz.", "You must select at least 2 criteria."))
    st.stop()

if not st.session_state.prep_done:
    st.stop()
else:
    if st.button(tt("🔄 Kriter Ayarlarına Geri Dön", "🔄 Back to Criteria Settings"), use_container_width=True):
        st.session_state.prep_done = False
        st.session_state["analysis_result"] = None
        st.rerun()

# ---------------------------------------------------------
# YÖNTEM SEÇİMİ (ALT PANEL)
# ---------------------------------------------------------
st.divider()
with st.expander(tt("⚙️ 3. Adım: Yöntem Seçimi ve Karşılaştırma", "⚙️ Step 3: Method Selection and Comparison"), expanded=True):
    st.caption(tt("Bu adımda analizin omurgasını kuran yöntem tercihlerini yapın.", "In this step, make the method choices that build the backbone of the analysis."))
    st.caption(tt("Seçimleriniz, sonuçların nasıl okunacağını ve hangi tablolarda raporlanacağını belirler.", "Your selections determine how results are interpreted and which tables are reported."))

    weight_mode = ""
    weight_method = "Entropy"
    weight_mode_key = "objective"
    manual_weights: Dict[str, float] | None = None
    manual_weights_valid = True
    with st.expander(tt("🧩 3.1. Adım: Ağırlıklandırma Modu ve Yöntemi", "🧩 Step 3.1: Weighting Mode and Method"), expanded=False):
        st.caption(tt("Bu bölümde kriter önemini hangi mantıkla hesaplayacağınızı belirleyin.", "In this section, choose how criterion importance will be calculated."))
        st.caption(tt("Objektif, eşit veya manuel yaklaşım seçiminin sonuç yorumunu doğrudan etkilediğini unutmayın.", "Remember that objective, equal, or manual weighting directly affects result interpretation."))

        methods_internal = me.OBJECTIVE_WEIGHT_METHODS
        methods_display = [method_display_name(m) for m in methods_internal]
        _default_weight = st.session_state.get("weight_method_pref")
        _default_weight_display = method_display_name(_default_weight) if _default_weight else methods_display[0]
        _weight_groups = weight_method_groups()
        _default_weight_internal = method_internal_name(_default_weight_display)

        weight_mode = st.radio(
            f"**{tt('Ağırlıklandırma Modu', 'Weighting Mode')}**",
            [tt("🎯 Objektif Ağırlık", "🎯 Objective Weights"), tt("⚖️ Eşit Ağırlık", "⚖️ Equal Weights"), tt("✍️ Manuel Ağırlık", "✍️ Manual Weights")],
            horizontal=True,
            help=tt(
                "Bu seçim, ağırlıkların veriden mi, eşit dağılımdan mı yoksa sizin önceliklerinizden mi üretileceğini belirler.\nKullandığınız ağırlık mantığı, bulguların yorum çerçevesini doğrudan değiştirir.",
                "This choice determines whether weights come from data, equal distribution, or your own priorities.\nThe weighting logic directly changes the interpretation frame of the findings.",
            ),
        )

        if "Objektif" in weight_mode or "Objective" in weight_mode:
            weight_method = _default_weight_internal if _default_weight_internal in methods_internal else methods_internal[0]
            weight_mode_key = "objective"
            st.caption(tt("Tüm yöntemler aşağıda kümelenmiş halde gösterilir. Tek seçim yapılır.", "All methods are grouped below. A single selection is used."))
            # Checkbox durumlarını ilk açılışta ayarla
            for _m in methods_internal:
                if f"weight_cb_{_m}" not in st.session_state:
                    st.session_state[f"weight_cb_{_m}"] = (_m == weight_method)
            # Eğer hiç seçili yoksa varsayılanı koru
            if not any(st.session_state.get(f"weight_cb_{_m}", False) for _m in methods_internal):
                st.session_state[f"weight_cb_{weight_method}"] = True
            for group_label, group_methods in _weight_groups:
                _filtered_methods = [m for m in group_methods if m in methods_internal]
                if not _filtered_methods:
                    continue
                st.markdown(f"**{group_label}**")
                method_cols = st.columns(3)
                for i, method_name in enumerate(_filtered_methods):
                    with method_cols[i % len(method_cols)]:
                        _label = method_display_name(method_name)
                        _cb_key = f"weight_cb_{method_name}"
                        _w_help = tt(
                            f"{_method_help_text(method_name)}\nBu yöntemi seçerek devam edebilirsiniz.",
                            f"{_method_help_text(method_name)}\nYou can proceed by selecting this method.",
                        )
                        st.checkbox(
                            _label, key=_cb_key, help=_w_help,
                            on_change=_wm_single_select_cb, args=(method_name, methods_internal),
                        )
            # Seçili yöntemi güncelle
            _all_weight_checked = [m for m in methods_internal if st.session_state.get(f"weight_cb_{m}", False)]
            if _all_weight_checked:
                weight_method = _all_weight_checked[0]
                st.session_state["weight_method_pref"] = weight_method
            # Seçili yöntem felsefesi kutusu
            _wh = _method_help_text(weight_method).split("\n", 1)
            _wh_simple   = _wh[0].strip()
            _wh_academic = _wh[1].strip() if len(_wh) > 1 else ""
            _academic_html = (
                f'<div style="font-size:0.70rem; color:#4A6070; line-height:1.5; margin-top:0.15rem;">{_wh_academic}</div>'
                if _wh_academic else ""
            )
            st.markdown(
                f'<div style="background:#EAF3FB; border-left:3px solid #5A9CC5; border-radius:0 8px 8px 0; '
                f'padding:0.45rem 0.75rem; margin-top:0.5rem;">'
                f'<div style="font-size:0.72rem; font-weight:700; color:#1F4A73; margin-bottom:0.18rem;">'
                f'💡 {method_display_name(weight_method)}</div>'
                f'<div style="font-size:0.71rem; color:#2A3E54; line-height:1.55;">{_wh_simple}</div>'
                f'{_academic_html}'
                f'</div>',
                unsafe_allow_html=True,
            )
        elif "Eşit" in weight_mode or "Equal" in weight_mode:
            weight_method = "Eşit Ağırlık"
            weight_mode_key = "equal"
            st.caption(tt("Tüm kriterlere eşit önem atanır.", "All criteria are assigned equal importance."))
            _weight_help_lines = _method_help_text(weight_method).split("\n", 1)
            st.caption(_weight_help_lines[0])
            if len(_weight_help_lines) > 1:
                st.caption(_weight_help_lines[1])
        else:
            weight_method = "Manuel Ağırlık"
            weight_mode_key = "manual"
            _weight_help_lines = _method_help_text(weight_method).split("\n", 1)
            st.caption(_weight_help_lines[0])
            if len(_weight_help_lines) > 1:
                st.caption(_weight_help_lines[1])

            _mcols = st.columns(min(3, max(1, len(criteria))))
            manual_raw_text: Dict[str, str] = {}
            for i, c in enumerate(criteria):
                with _mcols[i % len(_mcols)]:
                    manual_raw_text[c] = st.text_input(
                        c,
                        value=str(st.session_state.get(f"manual_w_{c}", "")),
                        placeholder=tt("örn. 7.5", "e.g. 7.5"),
                        key=f"manual_w_{c}",
                    ).strip()

            st.caption(
                tt(
                    "Seçili kriterler için ham önem puanı girin. Değerlerin toplamının 1 olması gerekmez; sistem analiz anında oranları otomatik normalize eder. Bu alan, AHP veya benzeri bir yaklaşımla dışarıda belirlediğiniz göreli önemleri esnek biçimde girmeniz için tasarlanmıştır.",
                    "Enter raw importance values for the selected criteria. Their total does not need to be 1; the system automatically normalizes the ratios during analysis. This area is designed for flexibly entering relative priorities produced externally through AHP or a similar approach.",
                )
            )

            missing_manual: List[str] = []
            invalid_manual: List[str] = []
            parsed_manual: Dict[str, float] = {}
            for c, raw_val in manual_raw_text.items():
                if not raw_val:
                    missing_manual.append(c)
                    continue
                try:
                    val = float(raw_val.replace(",", "."))
                    if not np.isfinite(val) or val <= 0:
                        raise ValueError
                    parsed_manual[c] = val
                except Exception:
                    invalid_manual.append(c)

            if missing_manual:
                st.info(
                    tt(
                        "Analize geçmeden önce tüm seçili kriterler için değer girin.",
                        "Enter values for all selected criteria before running the analysis.",
                    )
                )
                manual_weights_valid = False
            if invalid_manual:
                st.warning(
                    tt(
                        f"Geçerli pozitif sayı beklenen kriterler: {', '.join(invalid_manual)}.",
                        f"Valid positive numbers are required for: {', '.join(invalid_manual)}.",
                    )
                )
                manual_weights_valid = False

            if parsed_manual:
                _manual_total = float(sum(parsed_manual.values()))
                manual_weights = {c: parsed_manual[c] for c in criteria if c in parsed_manual}
                _manual_preview = pd.DataFrame(
                    {
                        "Kriter": criteria,
                        tt("Girilen Değer", "Entered Value"): [parsed_manual.get(c, np.nan) for c in criteria],
                        tt("Normalize Ağırlık", "Normalized Weight"): [
                            (parsed_manual[c] / _manual_total) if c in parsed_manual and _manual_total > 0 else np.nan
                            for c in criteria
                        ],
                    }
                )
                render_table(_manual_preview)
            else:
                manual_weights = None
                manual_weights_valid = False

    ranking_methods_selected = []
    primary_rank_method: str | None = None
    vikor_v = float(st.session_state.get("vikor_v", 0.5))
    waspas_lambda = float(st.session_state.get("waspas_lambda", 0.5))
    codas_tau = float(st.session_state.get("codas_tau", 0.02))
    cocoso_lambda = float(st.session_state.get("cocoso_lambda", 0.5))
    gra_rho = float(st.session_state.get("gra_rho", 0.5))
    promethee_pref_func = st.session_state.get("promethee_pref_func", "linear")
    promethee_q = float(st.session_state.get("promethee_q", 0.05))
    promethee_p = float(st.session_state.get("promethee_p", 0.30))
    promethee_s = float(st.session_state.get("promethee_s", 0.20))
    fuzzy_spread = float(st.session_state.get("fuzzy_spread", 0.10))
    sensitivity_iterations = int(st.session_state.get("sensitivity_iterations", 400))
    sensitivity_sigma = float(st.session_state.get("sensitivity_sigma", 0.12))

    if needs_ranking:
        with st.expander(tt("🧩 3.2. Adım: Sıralama Yöntemleri", "🧩 Step 3.2: Ranking Methods"), expanded=False):
            st.caption(tt("Bu bölümde alternatifleri hangi karar mantığıyla sıralayacağınızı belirleyin.", "In this section, choose the decision logic used to rank alternatives."))
            st.caption(tt("Birden fazla yöntem seçerseniz yöntem uyumunu da karşılaştırabilirsiniz.", "If you select more than one method, you can also compare method agreement."))
            _layer_options = [
                tt("Temel Yöntemler (Klasik Mantık)", "Core Methods (Classical Logic)"),
                tt("İleri Düzey Yöntemler (Fuzzy)", "Advanced Methods (Fuzzy)"),
            ]
            layer_choice = st.radio(
                f"**{tt('Analiz Katmanı', 'Analysis Layer')}**",
                _layer_options,
                horizontal=True,
                help=tt("Klasik katman kesin sayısal değerlerle, Fuzzy katman ise belirsizlik toleransı (spread) ile çalışır.", "Classical uses crisp numeric values; fuzzy uses uncertainty tolerance (spread)."),
            )
            if layer_choice == _layer_options[0]:
                all_ranks = me.CLASSICAL_MCDM_METHODS
                def_method = "TOPSIS"
                _layer_key = "classical"
            else:
                all_ranks = me.FUZZY_MCDM_METHODS
                def_method = "Fuzzy TOPSIS"
                _layer_key = "fuzzy"

            _sugg_rank_raw = st.session_state.get("_sugg_rank", "TOPSIS, VIKOR, EDAS")
            _sugg_rank_list = st.session_state.get("_sugg_rank_methods") or ["TOPSIS", "VIKOR", "EDAS"]
            _fallback_recommendations = (
                ["TOPSIS", "VIKOR", "EDAS"]
                if _layer_key == "classical"
                else ["Fuzzy TOPSIS", "Fuzzy VIKOR", "Fuzzy EDAS"]
            )

            def _resolve_rank_method_name(candidate: str) -> str | None:
                cand = candidate.strip()
                if not cand:
                    return None
                cand_upper = cand.upper().replace("FUZZY ", "").strip()
                for method_name in all_ranks:
                    method_upper = method_name.upper().replace("FUZZY ", "").strip()
                    if method_upper == cand_upper:
                        return method_name
                return None

            recommended_rank_methods: List[str] = []
            for candidate in [*_sugg_rank_list, *_sugg_rank_raw.replace("/", ",").split(","), *_fallback_recommendations]:
                resolved_method = _resolve_rank_method_name(candidate)
                if resolved_method and resolved_method not in recommended_rank_methods:
                    recommended_rank_methods.append(resolved_method)
                if len(recommended_rank_methods) >= 3:
                    break
            quick_pick_methods = recommended_rank_methods[:3] if recommended_rank_methods else [def_method]

            st.caption(
                tt(
                    f"Bu katmanda toplam {len(all_ranks)} yöntem kullanılabilir.",
                    f"A total of {len(all_ranks)} methods are available in this layer.",
                )
            )
            _mcol1, _mcol2 = st.columns(2)
            with _mcol1:
                if st.button(tt("✅ Önerilen Yöntemleri Çalıştır", "✅ Run Recommended Methods"), use_container_width=True, key=f"btn_select_recommended_{_layer_key}"):
                    st.session_state["ranking_prefs"] = list(quick_pick_methods)
                    for rank_method in all_ranks:
                        st.session_state[f"rank_cb_{rank_method}"] = rank_method in quick_pick_methods
                    st.rerun()
            with _mcol2:
                if st.button(tt("🗑️ Temizle", "🗑️ Clear"), use_container_width=True, key=f"btn_clear_all_{_layer_key}"):
                    st.session_state["ranking_prefs"] = []
                    for rank_method in all_ranks:
                        st.session_state[f"rank_cb_{rank_method}"] = False
                    st.rerun()
            st.caption(
                tt(
                    f"Önerilen başlangıç yöntemleri: {', '.join(recommended_rank_methods[:3])}.",
                    f"Recommended starter methods: {', '.join(recommended_rank_methods[:3])}.",
                )
            )
            st.caption(
                tt(
                    f"Neden bu öneri? {_rank_reason}",
                    f"Why this recommendation? {_rank_reason}",
                )
            )

            _pref_ranks = st.session_state.get("ranking_prefs") or []
            _has_pref_in_layer = any(m in all_ranks for m in _pref_ranks)
            for group_label, group_methods in ranking_method_groups(_layer_key):
                _filtered_methods = [m for m in group_methods if m in all_ranks]
                if not _filtered_methods:
                    continue
                st.markdown(f"**{group_label}**")
                rank_cols = st.columns(3)
                for i, rank_method in enumerate(_filtered_methods):
                    with rank_cols[i % len(rank_cols)]:
                        if _pref_ranks and _has_pref_in_layer:
                            default_val = rank_method in _pref_ranks
                        else:
                            default_val = rank_method == def_method
                        _rank_widget_key = f"rank_cb_{rank_method}"
                        _rank_help = tt(
                            f"{_method_help_text(rank_method)}\nBu yöntemi seçerseniz karşılaştırma çıktılarında da birlikte değerlendirilir.",
                            f"{_method_help_text(rank_method)}\nIf you select this method, it will also be evaluated in the comparison outputs.",
                        )
                        if _rank_widget_key in st.session_state:
                            _rank_checked = st.checkbox(rank_method, key=_rank_widget_key, help=_rank_help)
                        else:
                            _rank_checked = st.checkbox(rank_method, value=default_val, key=_rank_widget_key, help=_rank_help)
                        if _rank_checked:
                            ranking_methods_selected.append(rank_method)
            st.session_state["ranking_prefs"] = ranking_methods_selected

            if not ranking_methods_selected:
                st.warning(tt("Lütfen en az bir sıralama yöntemi seçiniz.", "Please select at least one ranking method."))
                st.stop()

            _sugg_primary = next(
                (m for m in ranking_methods_selected
                 if any(m == _resolve_rank_method_name(part) for part in [*_sugg_rank_list, *_sugg_rank_raw.replace("/", ",").split(",")])),
                ranking_methods_selected[0] if ranking_methods_selected else None,
            )
            primary_rank_method = _sugg_primary or (ranking_methods_selected[0] if ranking_methods_selected else None)

            uses_vikor = ("VIKOR" in ranking_methods_selected) or ("Fuzzy VIKOR" in ranking_methods_selected)
            uses_waspas = ("WASPAS" in ranking_methods_selected) or ("Fuzzy WASPAS" in ranking_methods_selected)
            uses_codas = ("CODAS" in ranking_methods_selected) or ("Fuzzy CODAS" in ranking_methods_selected)
            uses_cocoso = ("CoCoSo" in ranking_methods_selected) or ("Fuzzy CoCoSo" in ranking_methods_selected)
            uses_gra = ("GRA" in ranking_methods_selected) or ("Fuzzy GRA" in ranking_methods_selected)
            uses_promethee = ("PROMETHEE" in ranking_methods_selected) or ("Fuzzy PROMETHEE" in ranking_methods_selected)
            uses_fuzzy = any(m.startswith("Fuzzy") for m in ranking_methods_selected)
            rec = recommend_parameter_defaults(ranking_methods_selected, len(working), len(criteria), uses_fuzzy)

            with st.expander(tt("🧩 3.3. Adım: Parametreler", "🧩 Step 3.3: Parameters"), expanded=False):
                param_mode = st.radio(
                    tt("Parametre Modu", "Parameter Mode"),
                    [tt("🎯 Önerilen Varsayılan", "🎯 Recommended Default"), tt("✍️ Manuel Değiştir", "✍️ Manual Override")],
                    horizontal=True,
                    key="param_mode_choice",
                )
                if not any([uses_vikor, uses_waspas, uses_codas, uses_cocoso, uses_gra, uses_promethee, uses_fuzzy]):
                    st.caption(tt("Seçili sıralama yöntemleri ek parametre gerektirmiyor.", "Selected ranking methods do not require extra parameters."))
                if ("Önerilen" in param_mode) or ("Recommended" in param_mode):
                    vikor_v = float(rec["vikor_v"])
                    waspas_lambda = float(rec["waspas_lambda"])
                    codas_tau = float(rec["codas_tau"])
                    cocoso_lambda = float(rec["cocoso_lambda"])
                    gra_rho = float(rec["gra_rho"])
                    promethee_pref_func = str(rec.get("promethee_pref_func", "linear"))
                    promethee_q = float(rec.get("promethee_q", 0.05))
                    promethee_p = float(rec.get("promethee_p", 0.30))
                    promethee_s = float(rec.get("promethee_s", 0.20))
                    fuzzy_spread = float(rec["fuzzy_spread"])
                    sensitivity_iterations = int(rec["sensitivity_iterations"])
                    sensitivity_sigma = float(rec["sensitivity_sigma"])
                    st.caption(tt("Önerilen değerler otomatik uygulanır. İsterseniz Manuel Değiştir moduna geçebilirsiniz.", "Recommended values are applied automatically. Switch to Manual Override if needed."))
                    rec_rows = []
                    if uses_vikor: rec_rows.append({tt("Parametre", "Parameter"): "VIKOR v", tt("Değer", "Value"): vikor_v})
                    if uses_waspas: rec_rows.append({tt("Parametre", "Parameter"): "WASPAS λ", tt("Değer", "Value"): waspas_lambda})
                    if uses_codas: rec_rows.append({tt("Parametre", "Parameter"): "CODAS τ", tt("Değer", "Value"): codas_tau})
                    if uses_cocoso: rec_rows.append({tt("Parametre", "Parameter"): "CoCoSo λ", tt("Değer", "Value"): cocoso_lambda})
                    if uses_gra: rec_rows.append({tt("Parametre", "Parameter"): "GRA ρ", tt("Değer", "Value"): gra_rho})
                    if uses_promethee:
                        rec_rows.extend([
                            {tt("Parametre", "Parameter"): "PROMETHEE pref_func", tt("Değer", "Value"): promethee_pref_func},
                            {tt("Parametre", "Parameter"): "PROMETHEE q", tt("Değer", "Value"): promethee_q},
                            {tt("Parametre", "Parameter"): "PROMETHEE p", tt("Değer", "Value"): promethee_p},
                            {tt("Parametre", "Parameter"): "PROMETHEE s", tt("Değer", "Value"): promethee_s},
                        ])
                    if uses_fuzzy: rec_rows.append({tt("Parametre", "Parameter"): "Fuzzy spread", tt("Değer", "Value"): fuzzy_spread})
                    rec_rows.append({tt("Parametre", "Parameter"): tt("Monte Carlo iterasyon", "Monte Carlo iterations"), tt("Değer", "Value"): sensitivity_iterations})
                    rec_rows.append({tt("Parametre", "Parameter"): "Monte Carlo sigma", tt("Değer", "Value"): sensitivity_sigma})
                    rec_df = pd.DataFrame(rec_rows)
                    render_table(rec_df)
                else:
                    if uses_fuzzy:
                        fuzzy_spread = st.slider(tt("Fuzzy — spread (belirsizlik bandı)", "Fuzzy — spread (uncertainty band)"), 0.01, 0.50, float(fuzzy_spread), 0.01)
                    if uses_vikor:
                        vikor_v = st.slider(tt("VIKOR — v (uzlaşı katsayısı)", "VIKOR — v (compromise factor)"), 0.0, 1.0, float(vikor_v), 0.01)
                    if uses_waspas:
                        waspas_lambda = st.slider(tt("WASPAS — λ (hibrit katsayı)", "WASPAS — λ (hybrid coefficient)"), 0.0, 1.0, float(waspas_lambda), 0.01)
                    if uses_codas:
                        codas_tau = st.slider(tt("CODAS — τ (eşik değeri)", "CODAS — τ (threshold)"), 0.0, 0.20, float(codas_tau), 0.005)
                    if uses_cocoso:
                        cocoso_lambda = st.slider(tt("CoCoSo — λ (birleşim katsayısı)", "CoCoSo — λ (aggregation coefficient)"), 0.0, 1.0, float(cocoso_lambda), 0.01)
                    if uses_gra:
                        gra_rho = st.slider(tt("GRA — ρ (ayırt edici katsayı)", "GRA — ρ (distinguishing coefficient)"), 0.10, 0.90, float(gra_rho), 0.01)
                    if uses_promethee:
                        promethee_pref_func = st.selectbox(
                            tt("PROMETHEE tercih fonksiyonu", "PROMETHEE preference function"),
                            ["linear", "usual", "u_shape", "v_shape", "level", "gaussian"],
                            index=["linear", "usual", "u_shape", "v_shape", "level", "gaussian"].index(str(promethee_pref_func) if str(promethee_pref_func) in ["linear", "usual", "u_shape", "v_shape", "level", "gaussian"] else "linear"),
                        )
                        if promethee_pref_func in {"u_shape", "level", "linear"}:
                            promethee_q = st.slider("PROMETHEE q", 0.0, 1.0, float(promethee_q), 0.01)
                        if promethee_pref_func in {"v_shape", "level", "linear"}:
                            promethee_p = st.slider("PROMETHEE p", 0.0, 1.0, float(promethee_p), 0.01)
                        if promethee_pref_func == "gaussian":
                            promethee_s = st.slider("PROMETHEE s", 0.01, 1.0, float(promethee_s), 0.01)
                    st.markdown("---")
                    sensitivity_iterations = st.slider(tt("Monte Carlo iterasyon", "Monte Carlo iterations"), 100, 5000, int(sensitivity_iterations), 100)
                    sensitivity_sigma = st.slider(tt("Monte Carlo sigma", "Monte Carlo sigma"), 0.01, 0.50, float(sensitivity_sigma), 0.01)

            st.session_state["vikor_v"] = float(vikor_v)
            st.session_state["waspas_lambda"] = float(waspas_lambda)
            st.session_state["codas_tau"] = float(codas_tau)
            st.session_state["cocoso_lambda"] = float(cocoso_lambda)
            st.session_state["gra_rho"] = float(gra_rho)
            st.session_state["promethee_pref_func"] = str(promethee_pref_func)
            st.session_state["promethee_q"] = float(promethee_q)
            st.session_state["promethee_p"] = float(promethee_p)
            st.session_state["promethee_s"] = float(promethee_s)
            st.session_state["fuzzy_spread"] = float(fuzzy_spread)
            st.session_state["sensitivity_iterations"] = int(sensitivity_iterations)
            st.session_state["sensitivity_sigma"] = float(sensitivity_sigma)

_weight_step_done = bool(weight_mode)
_ranking_step_done = bool(ranking_methods_selected) if needs_ranking else False
_method_step_done = _weight_step_done and ((not needs_ranking) or _ranking_step_done)
if _method_step_done:
    st.markdown(f"**{tt('Harika.. Artık analiz yapabiliriz! Hazırsanız başlayalım', 'Great.. We can now run the analysis! If you are ready, let us begin.') }**")

if st.button(tt("🚀 Analiz Zamanı", "🚀 Run Analysis"), use_container_width=True):
    if ("Manuel" in weight_mode or "Manual" in weight_mode) and (not manual_weights_valid):
        st.error(tt("Manuel ağırlık için her seçili kriterde geçerli bir pozitif ham değer girin.", "Enter a valid positive raw value for every selected criterion in manual weighting."))
        st.stop()
    if panel_mode and (not panel_year_col or panel_year_col not in working.columns):
        st.error(tt("Panel veri için geçerli bir yıl sütunu seçmelisiniz.", "You must select a valid year column for panel data."))
        st.stop()
    with st.spinner(tt("Laboratuvar çalışıyor, veriler işleniyor...", "Running analysis... processing data...")):
        st.session_state.docx_buffer = None
        main_rank = primary_rank_method if primary_rank_method else (ranking_methods_selected[0] if ranking_methods_selected else None)
        comp_ranks = ranking_methods_selected if len(ranking_methods_selected) > 1 else []
        if main_rank and main_rank not in ranking_methods_selected:
            st.error(tt("Ana sıralama yöntemi seçimi geçersiz. Lütfen yöntemleri yeniden seçin.", "Invalid primary ranking selection. Please reselect methods."))
            st.stop()
        st.session_state["weight_method_pref"] = weight_method
        st.session_state["ranking_prefs"] = ranking_methods_selected

        config = me.AnalysisConfig(
            criteria=criteria, criteria_types=criteria_types,
            weight_method=weight_method,
            weight_mode=weight_mode_key,
            manual_weights=manual_weights,
            ranking_method=main_rank,
            compare_methods=comp_ranks,
            vikor_v=vikor_v,
            waspas_lambda=waspas_lambda,
            codas_tau=codas_tau,
            cocoso_lambda=cocoso_lambda,
            gra_rho=gra_rho,
            promethee_pref_func=promethee_pref_func,
            promethee_q=promethee_q,
            promethee_p=promethee_p,
            promethee_s=promethee_s,
            fuzzy_spread=fuzzy_spread,
            sensitivity_iterations=sensitivity_iterations,
            sensitivity_sigma=sensitivity_sigma,
        )

        start = time.time()
        try:
            if panel_mode:
                _all_year_labels = _sorted_panel_years(working[panel_year_col])
                _user_sel_years  = st.session_state.get("panel_selected_years") or _all_year_labels
                year_labels = [y for y in _all_year_labels if y in _user_sel_years]
                if not year_labels:
                    year_labels = _all_year_labels
                if len(year_labels) < 2:
                    raise ValueError(tt("Panel veri seçildi ancak yıl sütununda en az iki farklı dönem bulunamadı.", "Panel data was selected, but the year column does not contain at least two distinct periods."))
                panel_results: Dict[str, Dict[str, Any]] = {}
                for year_label in year_labels:
                    year_slice = working.loc[_panel_mask(working, panel_year_col, year_label), criteria].copy()
                    if year_slice.empty:
                        continue
                    year_slice.index = [f"A{idx+1}" for idx in range(len(year_slice))]
                    year_result = _run_single_analysis_bundle(
                        data_slice=year_slice,
                        criteria=criteria,
                        criteria_types=criteria_types,
                        config=config,
                        weight_mode_key=weight_mode_key,
                        weight_method=weight_method,
                        main_rank=main_rank,
                    )
                    year_result["panel_year"] = year_label
                    year_result["panel_year_column"] = panel_year_col
                    panel_results[year_label] = year_result
                if not panel_results:
                    raise ValueError(tt("Yıl bazında analiz üretilemedi. Seçilen yıl sütununu kontrol edin.", "Year-based analysis could not be produced. Please verify the selected year column."))
                default_year = list(panel_results.keys())[-1]
                st.session_state["panel_results"] = panel_results
                st.session_state["panel_years"] = list(panel_results.keys())
                st.session_state["panel_view_choice"] = default_year
                result = panel_results[default_year]
            else:
                result = _run_single_analysis_bundle(
                    data_slice=working,
                    criteria=criteria,
                    criteria_types=criteria_types,
                    config=config,
                    weight_mode_key=weight_mode_key,
                    weight_method=weight_method,
                    main_rank=main_rank,
                )
                st.session_state["panel_results"] = None
                st.session_state["panel_years"] = []
                st.session_state["panel_view_choice"] = None
        except Exception as exc:
            st.error(f"{tt('Analiz Hatası', 'Analysis Error')}: {exc}")
            st.stop()

        result["analysis_time"] = time.time() - start
        if panel_mode and st.session_state.get("panel_results"):
            for year_result in st.session_state["panel_results"].values():
                year_result["analysis_time"] = result["analysis_time"]
        st.session_state["analysis_result"] = result
        if DOCX_AVAILABLE:
            if panel_mode and st.session_state.get("panel_results"):
                st.session_state["report_docx"] = generate_panel_apa_docx(
                    st.session_state["panel_results"],
                    lang=st.session_state.get("ui_lang", "TR"),
                )
            else:
                st.session_state["report_docx"] = generate_apa_docx(
                    result,
                    working[criteria].copy(),
                    lang=st.session_state.get("ui_lang", "TR"),
                )

# ---------------------------------------------------------
# SONUÇLARIN GÖSTERİMİ
# ---------------------------------------------------------
result = st.session_state.get("analysis_result")
panel_results = st.session_state.get("panel_results")
_panel_all_choice = "__ALL_YEARS__"
panel_all_requested = False
panel_active_year = None
if isinstance(panel_results, dict) and panel_results:
    panel_years = st.session_state.get("panel_years") or list(panel_results.keys())
    view_choice = st.session_state.get("panel_view_choice")
    if view_choice == _panel_all_choice:
        panel_all_requested = True
        view_choice = panel_years[-1]
    elif view_choice not in panel_results:
        view_choice = panel_years[-1]
        st.session_state["panel_view_choice"] = view_choice
    panel_active_year = view_choice
    result = panel_results[view_choice]
if result is None:
    st.stop()

st.markdown(f"<h3 style='font-size: 1.2rem; margin-top:1rem;'>📥 {tt('Analiz Sonuçları ve Raporlar', 'Analysis Results and Reports')}</h3>", unsafe_allow_html=True)

weight_df     = result["weights"]["table"]
ranking_table = result.get("ranking", {}).get("table")
alt_names     = st.session_state.get("alt_names", {})

_lead_alt = "—"
if isinstance(ranking_table, pd.DataFrame) and not ranking_table.empty:
    _alt_col = col_key(ranking_table, "Alternatif", "Alternative")
    _raw_lead = str(ranking_table.iloc[0][_alt_col])
    _lead_alt = alt_names.get(_raw_lead, _raw_lead)
_rank_method = result.get("ranking", {}).get("method") or "—"
_sensitivity_payload = result.get("sensitivity") or {}
_stability = _sensitivity_payload.get("top_stability")
_stability_txt = f"%{float(_stability) * 100:.1f}" if _stability is not None else "—"
_runtime = result.get("analysis_time")
_runtime_txt = f"{float(_runtime):.2f} {tt('sn', 'sec')}" if _runtime is not None else "—"
_shape = result.get("selected_data", pd.DataFrame()).shape
st.markdown(
    f"""
    <div class="kpi-strip">
        <div class="kpi-grid">
            <div class="kpi-item"><div class="kpi-label">{tt('Lider Alternatif', 'Leading Alternative')}</div><div class="kpi-value">{_lead_alt}</div></div>
            <div class="kpi-item"><div class="kpi-label">{tt('Sıralama Yöntemi', 'Ranking Method')}</div><div class="kpi-value">{_rank_method}</div></div>
            <div class="kpi-item"><div class="kpi-label">{tt('Kararlılık', 'Stability')}</div><div class="kpi-value">{_stability_txt}</div></div>
            <div class="kpi-item"><div class="kpi-label">{tt('Analiz Süresi', 'Analysis Time')}</div><div class="kpi-value">{_runtime_txt}</div></div>
            <div class="kpi-item"><div class="kpi-label">{tt('Veri Boyutu', 'Data Size')}</div><div class="kpi-value">{_shape[0]}×{_shape[1]}</div></div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander(tt("🧮 Hesaplama Detaylarını Gör", "🧮 View Calculation Details"), expanded=False):
    if isinstance(panel_results, dict) and panel_results:
        _all_years_label = tt("Tüm yılları görmek istiyorum", "I want to see all years")
        _year_options = list(st.session_state.get("panel_years") or panel_results.keys()) + [_all_years_label]
        _stored_choice = st.session_state.get("panel_view_choice")
        _display_choice = _all_years_label if _stored_choice == _panel_all_choice else (_stored_choice or _year_options[0])
        if _display_choice not in _year_options:
            _display_choice = _year_options[0]
        _selected_year_choice = st.selectbox(
            tt("Sonucu görmek istediğiniz yılı seçin", "Select the year whose results you want to view"),
            _year_options,
            index=_year_options.index(_display_choice),
            key="panel_year_result_selector",
        )
        _new_panel_choice = _panel_all_choice if _selected_year_choice == _all_years_label else _selected_year_choice
        if _new_panel_choice != st.session_state.get("panel_view_choice"):
            st.session_state["panel_view_choice"] = _new_panel_choice
            st.rerun()
        if st.session_state.get("panel_view_choice") == _panel_all_choice:
            st.info(
                tt(
                    "Tüm yılların toplu sonuçları yalnızca DOCX ve Excel çıktılarında verilir. Ekranda tek yıl gösterilir; ayrıntılar için dosyayı indirip inceleyin.",
                    "Combined all-year results are provided only in the DOCX and Excel outputs. The screen shows one year at a time; download the files for the full multi-year view.",
                )
            )
        if panel_active_year is not None:
            st.caption(tt(f"Ekranda gösterilen yıl: {panel_active_year}", f"Year currently shown on screen: {panel_active_year}"))
    if "show_calc_details" not in st.session_state:
        st.session_state["show_calc_details"] = False
    if st.button(tt("🔍 Hesaplama Adımlarını Göster / Gizle", "🔍 Show / Hide Calculation Steps"), use_container_width=True, key="btn_toggle_calc_details"):
        st.session_state["show_calc_details"] = not st.session_state["show_calc_details"]
    if st.session_state["show_calc_details"]:
        dt1, dt2, dt3 = st.tabs([tt("⚖️ Ağırlık Adımları", "⚖️ Weighting Steps"), tt("🏆 Sıralama Adımları", "🏆 Ranking Steps"), tt("🎲 Duyarlılık Adımları", "🎲 Sensitivity Steps")])

        with dt1:
            wm = result["weights"]["method"]
            st.markdown(f"**{tt('Yöntem', 'Method')}:** `{method_display_name(wm)}`")
            st.code(get_method_steps(wm, "weight"), language="text")
            w_frames = _extract_detail_tables(result["weights"].get("details", {}))
            if w_frames:
                w_key = st.selectbox(tt("Ağırlık detay tablosu", "Weighting detail table"), list(w_frames.keys()), key="calc_weight_tbl")
                render_table(localize_df(w_frames[w_key].head(250)))
            else:
                st.info(tt("Ağırlık yöntemi için detay tablo bulunamadı.", "No detail table found for the weighting method."))

        with dt2:
            rm = result.get("ranking", {}).get("method")
            if rm:
                comp = result.get("comparison", {}) or {}
                method_details = dict(comp.get("method_details") or {})
                method_details.setdefault(str(rm), result.get("ranking", {}).get("details", {}) or {})
                method_names = [m for m in method_details.keys() if str(m).strip()]

                if len(method_names) > 1:
                    method_tabs = st.tabs([str(m) for m in method_names])
                    for meth_name, meth_tab in zip(method_names, method_tabs):
                        with meth_tab:
                            st.markdown(f"**{tt('Yöntem', 'Method')}:** `{meth_name}`")
                            st.code(get_method_steps(str(meth_name), "ranking"), language="text")
                            r_frames = _extract_detail_tables(method_details.get(str(meth_name), {}))
                            if r_frames:
                                r_key = st.selectbox(
                                    tt("Sıralama detay tablosu", "Ranking detail table"),
                                    list(r_frames.keys()),
                                    key=f"calc_rank_tbl_{str(meth_name)}",
                                )
                                render_table(localize_df(r_frames[r_key].head(250)))
                            else:
                                st.info(tt("Sıralama yöntemi için detay tablo bulunamadı.", "No detail table found for the ranking method."))
                else:
                    st.markdown(f"**{tt('Yöntem', 'Method')}:** `{rm}`")
                    st.code(get_method_steps(rm, "ranking"), language="text")
                    r_frames = _extract_detail_tables(result.get("ranking", {}).get("details", {}))
                    if r_frames:
                        r_key = st.selectbox(tt("Sıralama detay tablosu", "Ranking detail table"), list(r_frames.keys()), key="calc_rank_tbl")
                        render_table(localize_df(r_frames[r_key].head(250)))
                    else:
                        st.info(tt("Sıralama yöntemi için detay tablo bulunamadı.", "No detail table found for the ranking method."))
            else:
                st.info(tt("Bu analizde sıralama yöntemi seçilmedi.", "No ranking method was selected in this analysis."))

        with dt3:
            sens = result.get("sensitivity")
            if sens:
                st.markdown(f"**{tt('Duyarlılık yaklaşımı', 'Sensitivity approach')}:** {tt('Lokal senaryolar + Monte Carlo bozulma analizi', 'Local scenarios + Monte Carlo perturbation analysis')}")
                st.code(
                    tt(
                        "1) Baz ağırlıkla referans sıralama\n2) Her kriterde ±%10/±%20 lokal değişim\n3) Spearman ile sıra benzerliği\n4) Monte Carlo ile birincilik oranı",
                        "1) Baseline ranking with base weights\n2) ±10%/±20% local perturbations per criterion\n3) Rank similarity via Spearman\n4) First-place frequency via Monte Carlo",
                    ),
                    language="text",
                )
                local_df = sens.get("local_scenarios")
                mc_df = sens.get("monte_carlo_summary")
                if isinstance(local_df, pd.DataFrame) and not local_df.empty:
                    st.markdown(f"**{tt('Lokal senaryolar', 'Local scenarios')}**")
                    render_table(localize_df(local_df.head(250)))
                if isinstance(mc_df, pd.DataFrame) and not mc_df.empty:
                    st.markdown(f"**{tt('Monte Carlo özeti', 'Monte Carlo summary')}**")
                    render_table(localize_df(mc_df.head(250)))
            else:
                st.info(tt("Duyarlılık analizi bu çalışmada üretilmedi.", "Sensitivity analysis was not generated in this study."))

tabs_list = [
    tt("📊 Temel İstatistik", "📊 Basic Statistics"),
    tt("⚖️ Ağırlıklar", "⚖️ Weights"),
]
if needs_ranking:
    tabs_list.extend([
        tt("🏆 Sıralama", "🏆 Ranking"),
        tt("🔁 Karşılaştırma", "🔁 Comparison"),
        tt("🛡️ Sağlamlık", "🛡️ Robustness"),
        tt("🎲 Dayanıklılık", "🎲 Durability"),
    ])
else:
    tabs_list.append(tt("🛡️ Sağlamlık", "🛡️ Robustness"))
tabs_list.append(tt("📥 Çıktı", "📥 Output"))
tabs = st.tabs(tabs_list)

_tab_stats = 0
_tab_weights = 1
if needs_ranking:
    _tab_ranking = 2
    _tab_comparison = 3
    _tab_robustness = 4
    _tab_durability = 5
    _tab_output = 6
else:
    _tab_robustness = 2
    _tab_output = 3

with tabs[_tab_stats]:
    _stat_comment = gen_stat_commentary(result)
    render_tab_assistant(_stat_comment, key="stat")
    with st.expander(f"📋 {tt('Tablolar', 'Tables')}", expanded=False):
        st.markdown(f"##### 📊 {tt('Betimleyici İstatistikler', 'Descriptive Statistics')}")
        _stats_disp = localize_df(result["stats"]).copy()
        _num_cols = _stats_disp.select_dtypes(include=[np.number]).columns
        _stats_disp[_num_cols] = _stats_disp[_num_cols].round(3)
        render_table(_stats_disp)
        with st.expander(f"💬 {tt('Tablo Yorumu', 'Table Commentary')}", expanded=False):
            st.markdown(f'<div class="commentary-box">{_html_to_plain(_stat_comment).replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

    with st.expander(f"📊 {tt('Şekiller', 'Figures')}", expanded=False):
        s1, s2 = st.columns([1, 1])
        with s1:
            st.plotly_chart(fig_box_plots(result["selected_data"], criteria), use_container_width=True)
            with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                st.markdown(f'<div class="commentary-box">{tt("Kutu grafikleri her kriterin dağılım profilini gösterir. Geniş kutular ve aykırı noktalar, o kriterin yüksek ayırt edici gücüne işaret eder.", "Box plots show the distribution profile of each criterion. Wide boxes and outlier points indicate high discriminative power.")}</div>', unsafe_allow_html=True)
        with s2:
            corr_df = result.get("correlation_matrix")
            if isinstance(corr_df, pd.DataFrame):
                fig_corr = px.imshow(
                    corr_df, text_auto=".2f", aspect="auto",
                    color_continuous_scale="RdBu_r", zmin=-1, zmax=1,
                    title=tt("Kriterler Arası Korelasyon Haritası", "Correlation Heatmap Among Criteria"),
                )
                fig_corr.update_layout(height=420, **_THEME)
                st.plotly_chart(fig_corr, use_container_width=True)
                with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                    st.markdown(f'<div class="commentary-box">{tt("Korelasyon ısı haritası, kriterler arasındaki doğrusal ilişkilerin yoğunluğunu gösterir. Kırmızı: güçlü pozitif, mavi: güçlü negatif ilişki. Yüksek korelasyon (|ρ|>0.75) bilgi tekrarına işaret eder.", "The correlation heatmap shows the intensity of linear relationships between criteria. Red: strong positive, blue: strong negative. High correlation (|ρ|>0.75) suggests information redundancy.")}</div>', unsafe_allow_html=True)
            else:
                st.info(tt("Korelasyon matrisi bulunamadı.", "Correlation matrix not available."))

with tabs[_tab_weights]:
    _weight_comment = gen_weight_commentary(result)
    render_tab_assistant(_weight_comment, key="weight")
    weight_col = col_key(weight_df, "Ağırlık", "Weight")
    top3_weights = weight_df.sort_values(weight_col, ascending=False).head(3).reset_index(drop=True)

    with st.expander(f"📋 {tt('Tablolar', 'Tables')}", expanded=False):
        w1, w2 = st.columns(2)
        with w1:
            st.markdown(f"##### 🥇 {tt('İlk 3 Önemli Kriter', 'Top 3 Important Criteria')}")
            _top3_disp = localize_df(top3_weights)
            _top3_w_col = col_key(_top3_disp, "Ağırlık", "Weight")
            if _top3_w_col in _top3_disp.columns:
                _top3_disp[_top3_w_col] = pd.to_numeric(_top3_disp[_top3_w_col], errors="coerce").round(6)
            render_table(_top3_disp)
        with w2:
            st.markdown(f"##### ⚖️ {tt('Tüm Kriter Ağırlıkları', 'All Criterion Weights')}")
            _weight_disp = localize_df(weight_df)
            _weight_disp_col = col_key(_weight_disp, "Ağırlık", "Weight")
            if _weight_disp_col in _weight_disp.columns:
                _weight_disp[_weight_disp_col] = pd.to_numeric(_weight_disp[_weight_disp_col], errors="coerce").round(6)
            render_table(_weight_disp)
        with st.expander(f"💬 {tt('Tablo Yorumu', 'Table Commentary')}", expanded=False):
            st.markdown(f'<div class="commentary-box">{_html_to_plain(_weight_comment).replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

    with st.expander(f"📊 {tt('Şekiller', 'Figures')}", expanded=False):
        fig_col1, fig_col2 = st.columns(2)
        with fig_col1:
            st.plotly_chart(fig_weight_bar(weight_df), use_container_width=True)
            with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                st.markdown(f'<div class="commentary-box">{tt("Ağırlık çubuğu grafiği, kriterlerin göreli önem sıralamasını ve büyüklüklerini karşılaştırmalı olarak gösterir. En uzun çubuk, analiz kararını en güçlü etkileyen kriteri temsil eder.", "The weight bar chart compares the relative importance and magnitude of criteria. The longest bar represents the criterion with the highest influence on the analysis decision.")}</div>', unsafe_allow_html=True)
        with fig_col2:
            st.plotly_chart(fig_weight_radar(weight_df), use_container_width=True)
            with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                st.markdown(f'<div class="commentary-box">{tt("Radar grafiği, ağırlıkların çok boyutlu dağılımını görsel olarak sunar. Dengeli bir dağılım, kararın tek bir kritere bağımlı olmadığına işaret eder; asimetrik yapı ise belirli bir kriterin hakimiyetini gösterir.", "The radar chart visually presents the multi-dimensional distribution of weights. A balanced distribution indicates the decision is not dependent on a single criterion; asymmetric structure shows dominance of a specific criterion.")}</div>', unsafe_allow_html=True)

with tabs[_tab_robustness]:
    _wr = result.get("weight_robustness", {}) or {}
    _rob_comment = gen_weight_robustness_commentary(result)
    st.markdown(f"### 🛡️ {tt('Ağırlık Sağlamlık Testleri', 'Weight Robustness Tests')}")
    if _wr.get("info"):
        st.info(str(_wr.get("info")))
    elif _wr.get("error"):
        st.warning(tt("Ağırlık sağlamlık testleri üretilemedi.", "Weight robustness tests could not be generated."))
    else:
        _loo_mean = _wr.get("loo_mean_rho")
        _loo_min = _wr.get("loo_min_rho")
        _top_match = _wr.get("loo_top_match_rate")
        _eff_n = _wr.get("effective_criteria_n")
        _dom = _wr.get("dominance_ratio")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric(tt("Ortalama Uyum (LOO)", "Mean Agreement (LOO)"), (f"{float(_loo_mean):.3f}" if _loo_mean is not None and np.isfinite(_loo_mean) else "—"))
        m2.metric(tt("Min. Uyum (LOO)", "Min. Agreement (LOO)"), (f"{float(_loo_min):.3f}" if _loo_min is not None and np.isfinite(_loo_min) else "—"))
        m3.metric(tt("Lider Kriter Tutarlılığı", "Top-Criterion Consistency"), (f"%{float(_top_match)*100:.1f}" if _top_match is not None and np.isfinite(_top_match) else "—"))
        m4.metric(tt("Etkin Kriter Sayısı", "Effective Criterion Count"), (f"{float(_eff_n):.2f}" if _eff_n is not None and np.isfinite(_eff_n) else "—"))

        with st.expander(f"💬 {tt('Sağlamlık Özet Yorumu', 'Robustness Summary Commentary')}", expanded=True):
            st.markdown(f'<div class="commentary-box">{_rob_comment}</div>', unsafe_allow_html=True)
        st.caption(tt("Detaylar için lütfen aşağıdaki tablo ve şekilleri inceleyin.", "For details, please review the tables and figures below."))

        _felsefe_msg = tt(
            "📌 Yöntem felsefesi: Bu testler 'seçtiğimiz objektif ağırlıklandırma, veri biraz oynadığında aynı hikâyeyi anlatıyor mu?' sorusunu yanıtlar.",
            "📌 Method philosophy: These tests answer: 'Does the chosen objective weighting tell the same story when data changes slightly?'"
        )
        st.caption(_felsefe_msg)
        if _dom is not None and np.isfinite(_dom):
            st.caption(tt(
                f"⚖️ Lider/ikinci kriter ağırlık oranı: {_dom:.2f}. Oran yükseldikçe karar tek kritere daha bağımlı hale gelir.",
                f"⚖️ Top/second criterion weight ratio: {_dom:.2f}. As this ratio grows, the decision depends more on a single criterion."
            ))

        with st.expander(f"📋 {tt('Tablolar — Detay Sağlamlık Verileri', 'Tables — Detailed Robustness Data')}", expanded=False):
            _loo_df = _wr.get("leave_one_out")
            if isinstance(_loo_df, pd.DataFrame) and not _loo_df.empty:
                st.markdown(f"##### 🔍 {tt('Leave-One-Out Analizi', 'Leave-One-Out Analysis')}")
                st.caption(tt(
                    "Her alternatif birer birer çıkarılarak ağırlık vektörünün ne kadar değiştiği ölçülür. Spearman ρ = 1.0 mükemmel kararlılık, düşük değer kırılganlık işareti.",
                    "Each alternative is removed one by one and the change in weight vector is measured. Spearman ρ = 1.0 indicates perfect stability; lower values signal fragility."
                ))
                _loo_disp = localize_df(_loo_df.copy())
                _rho_col = col_key(_loo_disp, "SpearmanRho", "SpearmanRho")
                _maxdiff_col = col_key(_loo_disp, "MaksMutlakFark", "MaxAbsoluteDiff")
                if _rho_col in _loo_disp.columns:
                    _loo_disp[_rho_col] = pd.to_numeric(_loo_disp[_rho_col], errors="coerce").round(3)
                if _maxdiff_col in _loo_disp.columns:
                    _loo_disp[_maxdiff_col] = pd.to_numeric(_loo_disp[_maxdiff_col], errors="coerce").round(4)
                render_table(_loo_disp)
                with st.expander(f"💬 {tt('Tablo Yorumu', 'Table Commentary')}", expanded=False):
                    st.markdown(f'<div class="commentary-box">{tt("Leave-one-out tablosu, her alternatifin çıkarıldığında ağırlık düzeninin ne ölçüde korunduğunu gösterir. SpearmanRho sütunu 1\'e yakın değerler yüksek kararlılık anlamına gelir. MaksMutlakFark ise en fazla değişen kriter ağırlığının büyüklüğünü verir.", "The leave-one-out table shows how well the weight order is preserved when each alternative is removed. SpearmanRho values close to 1 indicate high stability. MaxAbsoluteDiff gives the magnitude of the most changed criterion weight.")}</div>', unsafe_allow_html=True)

            _boot_df = _wr.get("bootstrap_summary")
            if isinstance(_boot_df, pd.DataFrame) and not _boot_df.empty:
                st.markdown(f"##### 🎲 {tt('Bootstrap Kriter Güven Aralıkları', 'Bootstrap Criterion Confidence Intervals')}")
                st.caption(tt(
                    "Veri tekrar örneklenerek her kriterin ağırlık dağılımı simüle edilir. Dar aralık = güvenilir ağırlık, geniş aralık = belirsizlik yüksek.",
                    "Data is resampled to simulate the weight distribution for each criterion. Narrow interval = reliable weight, wide interval = high uncertainty."
                ))
                _boot_disp = localize_df(_boot_df.copy())
                _mean_col = col_key(_boot_disp, "OrtalamaAğırlık", "MeanWeight")
                _std_col = col_key(_boot_disp, "StdSapma", "StdDev")
                _l5_col = col_key(_boot_disp, "AltYuzde5", "Lower5Pct")
                _u95_col = col_key(_boot_disp, "UstYuzde95", "Upper95Pct")
                for _c in [_mean_col, _std_col, _l5_col, _u95_col]:
                    if _c in _boot_disp.columns:
                        _boot_disp[_c] = pd.to_numeric(_boot_disp[_c], errors="coerce").round(4)
                render_table(_boot_disp)
                with st.expander(f"💬 {tt('Tablo Yorumu', 'Table Commentary')}", expanded=False):
                    st.markdown(f'<div class="commentary-box">{tt("Bootstrap özet tablosu, her kriterin ağırlığının istatistiksel güven aralığını gösterir. Standart sapması düşük kriterler kararlı ağırlıklara sahiptir ve analizin yapı taşını oluşturur. Geniş aralıklı kriterler için duyarlılık analizi ile çapraz doğrulama yapılması önerilir.", "The bootstrap summary table shows the statistical confidence interval for each criterion weight. Criteria with low standard deviation have stable weights and form the backbone of the analysis. Cross-validation with sensitivity analysis is recommended for criteria with wide intervals.")}</div>', unsafe_allow_html=True)

        with st.expander(f"📊 {tt('Şekiller — Sağlamlık Görselleştirmeleri', 'Figures — Robustness Visualizations')}", expanded=False):
            _boot_df2 = _wr.get("bootstrap_summary")
            if isinstance(_boot_df2, pd.DataFrame) and not _boot_df2.empty:
                _k_col = col_key(_boot_df2, "Kriter", "Criterion") if "Kriter" in _boot_df2.columns or "Criterion" in _boot_df2.columns else _boot_df2.columns[0]
                _m_col = col_key(_boot_df2, "OrtalamaAğırlık", "MeanWeight")
                _s_col = col_key(_boot_df2, "StdSapma", "StdDev")
                _l_col = col_key(_boot_df2, "AltYuzde5", "Lower5Pct")
                _u_col = col_key(_boot_df2, "UstYuzde95", "Upper95Pct")
                _boot_plot = _boot_df2.copy()
                for _c in [_m_col, _s_col, _l_col, _u_col]:
                    if _c in _boot_plot.columns:
                        _boot_plot[_c] = pd.to_numeric(_boot_plot[_c], errors="coerce")
                if all(c in _boot_plot.columns for c in [_k_col, _m_col, _l_col, _u_col]):
                    _boot_plot = _boot_plot.sort_values(_m_col, ascending=True)
                    fig_boot = go.Figure()
                    fig_boot.add_trace(go.Scatter(
                        x=_boot_plot[_m_col], y=_boot_plot[_k_col],
                        mode='markers',
                        marker=dict(color='#3D6491', size=10, symbol='circle'),
                        name=tt("Ortalama", "Mean"),
                        error_x=dict(
                            type='data',
                            symmetric=False,
                            array=(_boot_plot[_u_col] - _boot_plot[_m_col]).clip(lower=0).tolist(),
                            arrayminus=(_boot_plot[_m_col] - _boot_plot[_l_col]).clip(lower=0).tolist(),
                            color='rgba(61,100,145,0.5)',
                            thickness=2,
                            width=6,
                        ),
                    ))
                    fig_boot.update_layout(
                        height=max(280, len(_boot_plot) * 38 + 60),
                        title=tt("Bootstrap Ağırlık Güven Aralıkları (5%–95%)", "Bootstrap Weight Confidence Intervals (5%–95%)"),
                        xaxis_title=tt("Ağırlık", "Weight"),
                        yaxis_title=tt("Kriter", "Criterion"),
                        **_THEME,
                    )
                    st.plotly_chart(fig_boot, use_container_width=True)
                    with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                        st.markdown(f'<div class="commentary-box">{tt("Bootstrap güven aralığı grafiği, her kriterin ağırlığının olası aralığını gösterir. Yatay çizgi uzunluğu belirsizlik düzeyini temsil eder: kısa çizgi = güvenilir ağırlık, uzun çizgi = veri değişimine duyarlı ağırlık.", "The bootstrap confidence interval chart shows the plausible range for each criterion weight. Bar length represents uncertainty level: short bar = reliable weight, long bar = weight sensitive to data changes.")}</div>', unsafe_allow_html=True)

if needs_ranking:
    with tabs[_tab_ranking]:
        _rank_comment = gen_ranking_commentary(result, alt_names)
        render_tab_assistant(_rank_comment, key="rank")
        if ranking_table is None or ranking_table.empty:
            st.info(tt("Sıralama tablosu oluşturulamadı.", "Ranking table could not be generated."))
        else:
            _rt_disp = ranking_table.copy()
            _rt_alt_col = col_key(_rt_disp, "Alternatif", "Alternative")
            _rt_score_col = col_key(_rt_disp, "Skor", "Score")
            if alt_names:
                _rt_disp[_rt_alt_col] = _rt_disp[_rt_alt_col].astype(str).map(
                    lambda x: alt_names.get(x, x)
                )

            _rank_alt_col = col_key(ranking_table, "Alternatif", "Alternative")
            top_disp = alt_names.get(str(ranking_table.iloc[0][_rank_alt_col]), str(ranking_table.iloc[0][_rank_alt_col]))
            st.info(f"🏆 {tt('Lider Alternatif', 'Leading Alternative')}: **{top_disp}** — {tt('Yöntem', 'Method')}: {result.get('ranking', {}).get('method')}")

            with st.expander(f"📋 {tt('Tablolar', 'Tables')}", expanded=False):
                r1, r2 = st.columns(2)
                with r1:
                    st.markdown(f"##### 🥇 {tt('İlk 5 Alternatif', 'Top 5 Alternatives')}")
                    _rt_top5 = localize_df(_rt_disp.head(5))
                    _rt_top5_score_col = col_key(_rt_top5, "Skor", "Score")
                    if _rt_top5_score_col in _rt_top5.columns:
                        _rt_top5[_rt_top5_score_col] = pd.to_numeric(_rt_top5[_rt_top5_score_col], errors="coerce").round(4)
                    render_table(_rt_top5)
                with r2:
                    st.markdown(f"##### 📊 {tt('Nihai Sıralama', 'Final Ranking')}")
                    _rt_table_disp = localize_df(_rt_disp)
                    _rt_table_score_col = col_key(_rt_table_disp, "Skor", "Score")
                    if _rt_table_score_col in _rt_table_disp.columns:
                        _rt_table_disp[_rt_table_score_col] = pd.to_numeric(_rt_table_disp[_rt_table_score_col], errors="coerce").round(4)
                    render_table(_rt_table_disp)
                with st.expander(f"💬 {tt('Tablo Yorumu', 'Table Commentary')}", expanded=False):
                    st.markdown(f'<div class="commentary-box">{_html_to_plain(_rank_comment).replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

            with st.expander(f"📊 {tt('Şekiller', 'Figures')}", expanded=False):
                fig_r1, fig_r2 = st.columns(2)
                with fig_r1:
                    st.plotly_chart(fig_rank_bar(_rt_disp), use_container_width=True)
                    with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                        st.markdown(f'<div class="commentary-box">{tt("Yatay çubuk grafik, her alternatifin sıralama skorunu karşılaştırmalı olarak gösterir. Uzun çubuk yüksek performansı simgeler. Barlar arasındaki mesafe, alternatiflerin birbirinden ne kadar ayrıştığını ortaya koyar.", "The horizontal bar chart comparatively shows the ranking score of each alternative. A longer bar symbolizes higher performance. The distance between bars reveals how much alternatives are differentiated from each other.")}</div>', unsafe_allow_html=True)
                with fig_r2:
                    st.plotly_chart(fig_network_alternatives(_rt_disp), use_container_width=True)
                    with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                        st.markdown(f'<div class="commentary-box">{tt("Ağ diyagramı, alternatiflerin merkezi referansa göre göreli konumunu dairesel yerleşimde gösterir. Merkeze yakın düğümler daha yüksek skor anlamına gelir. Düğüm boyutu ve rengi performans düzeyini yansıtır.", "The network diagram shows the relative position of alternatives in circular layout against the central reference. Nodes closer to the center mean higher scores. Node size and color reflect performance level.")}</div>', unsafe_allow_html=True)

                contrib = result.get("contribution_table")
                if isinstance(contrib, pd.DataFrame):
                    _contrib_disp = contrib.copy()
                    _contrib_alt_col = col_key(_contrib_disp, "Alternatif", "Alternative")
                    if alt_names:
                        _contrib_disp[_contrib_alt_col] = _contrib_disp[_contrib_alt_col].astype(str).map(
                            lambda x: alt_names.get(x, x)
                        )
                    st.plotly_chart(fig_parallel_coords(_rt_disp, _contrib_disp, criteria), use_container_width=True)
                    with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                        st.markdown(f'<div class="commentary-box">{tt("Paralel koordinat grafiği, her alternatifin tüm kriterler boyunca performans profilini tek bir görünümde sunar. Çizgilerin rengi ve yoğunluğu skor büyüklüğünü gösterir. Çizgilerin yoğun kesiştiği bölgeler, alternatiflerin benzer kriter değerlerine sahip olduğu alanları işaret eder.", "The parallel coordinates chart presents each alternative\'s performance profile across all criteria in a single view. Line color and intensity show score magnitude. Regions where lines densely intersect indicate areas where alternatives have similar criterion values.")}</div>', unsafe_allow_html=True)

            with st.expander(f"🧠 {tt('Yöntem-Özel İçgörüler', 'Method-Specific Insights')}", expanded=False):
                render_method_specific_insights(result, alt_names)

    with tabs[_tab_comparison]:
        _comp_comment = gen_comparison_commentary(result)
        render_tab_assistant(_comp_comment, key="comp")
        comp = result.get("comparison", {})
        if comp and "rank_table" in comp:
            with st.expander(f"📋 {tt('Tablolar', 'Tables')}", expanded=False):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"##### 🔁 {tt('Yöntem Sıralama Karşılaştırması', 'Method Ranking Comparison')}")
                    _rank_comp = comp["rank_table"].copy()
                    _rank_comp_alt_col = col_key(_rank_comp, "Alternatif", "Alternative")
                    if alt_names:
                        if _rank_comp_alt_col in _rank_comp.columns:
                            _rank_comp[_rank_comp_alt_col] = _rank_comp[_rank_comp_alt_col].astype(str).map(
                                lambda x: alt_names.get(x, x)
                            )
                    render_table(localize_df(_rank_comp))
                with c2:
                    _method_rob_df = build_ranking_robustness_table(result)
                    if isinstance(_method_rob_df, pd.DataFrame) and not _method_rob_df.empty:
                        st.markdown(f"##### 🛡️ {tt('Yöntem Bazlı Sağlamlık', 'Method-Based Robustness')}")
                        st.caption(tt(
                            "Her yöntemin ana yöntemle ne kadar benzer düşündüğünü basit dille açıklar.",
                            "Explains in plain language how similarly each method thinks compared with the primary method."
                        ))
                        render_table(_method_rob_df)
                with st.expander(f"💬 {tt('Tablo Yorumu', 'Table Commentary')}", expanded=False):
                    st.markdown(f'<div class="commentary-box">{_html_to_plain(_comp_comment).replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

            with st.expander(f"📊 {tt('Şekiller', 'Figures')}", expanded=False):
                if isinstance(comp.get("spearman_matrix"), pd.DataFrame):
                    spdf = comp["spearman_matrix"]
                    _method_col = col_key(spdf, "Yöntem", "Method")
                    fig_sp = px.imshow(
                        spdf.set_index(_method_col), text_auto=".2f", aspect="auto",
                        color_continuous_scale="RdYlGn", zmin=-1, zmax=1,
                        title=tt("Spearman Uyum Matrisi", "Spearman Agreement Matrix"),
                    )
                    fig_sp.update_layout(height=420, **_THEME)
                    st.plotly_chart(fig_sp, use_container_width=True)
                    with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                        st.markdown(f'<div class="commentary-box">{tt("Spearman ısı haritası, seçili yöntemlerin birbiriyle ne kadar tutarlı sıralama ürettiğini gösterir. Yeşil hücreler (ρ ≥ 0.85) yüksek yöntem uyumuna, sarı-kırmızı hücreler ise yöntemler arası görüş ayrılığına işaret eder. Üçgen matristeki tüm hücreler yeşilse sonuçlar metodolojiden bağımsız kararlı kabul edilebilir.", "The Spearman heatmap shows how consistently the selected methods produce rankings. Green cells (ρ ≥ 0.85) indicate high method agreement, yellow-red cells indicate disagreement between methods. If all cells in the triangular matrix are green, results can be considered stable regardless of methodology.")}</div>', unsafe_allow_html=True)
        else:
            st.info(tt("Karşılaştırma için birden fazla yöntem seçilmelidir.", "Select more than one method for comparison."))

    with tabs[_tab_durability]:
        render_tab_assistant(gen_mc_commentary(result), key="mc")
        sens = result.get("sensitivity")
        if not sens:
            st.info(tt("Sağlamlık analizi verisi oluşturulamadı. Lütfen sıralama yöntemi seçerek analizi çalıştırın.", "Robustness data could not be generated. Please run analysis with a ranking method."))
        else:
            mc_df  = sens.get("monte_carlo_summary")
            loc_df = sens.get("local_scenarios")
            stab   = float(sens.get("top_stability", 0.0))
            n_iter = int(sens.get("n_iterations", 0)) if "n_iterations" in sens else "—"

            mc1, mc2, mc3 = st.columns(3)
            mc1.metric(tt("Lider Kararlılığı", "Leader Stability"), f"%{stab*100:.1f}", delta="Monte Carlo")
            mc2.metric(tt("Simülasyon Sayısı", "Simulation Count"), f"{n_iter:,}" if isinstance(n_iter, int) else n_iter)
            if isinstance(mc_df, pd.DataFrame) and not mc_df.empty:
                _mc_alt_col = col_key(mc_df, "Alternatif", "Alternative")
                runner_up = mc_df.iloc[1][_mc_alt_col] if len(mc_df) > 1 else "—"
                runner_disp = alt_names.get(str(runner_up), str(runner_up))
                mc3.metric(tt("2. Sıra Alternatif", "Runner-up Alternative"), runner_disp)

            if isinstance(mc_df, pd.DataFrame) and not mc_df.empty:
                _mc_disp = mc_df.copy()
                _mc_disp_alt_col = col_key(_mc_disp, "Alternatif", "Alternative")
                if alt_names:
                    _mc_disp[_mc_disp_alt_col] = _mc_disp[_mc_disp_alt_col].astype(str).map(
                        lambda x: alt_names.get(x, x)
                    )

                _mc_comment = gen_mc_commentary(result)

                with st.expander(f"📋 {tt('Tablolar', 'Tables')}", expanded=False):
                    st.markdown(f"##### 🎲 {tt('Simülasyon Özet Tablosu', 'Simulation Summary Table')}")
                    _mc_table_disp = localize_df(_mc_disp)
                    _fp_col = col_key(_mc_table_disp, "BirincilikOranı", "FirstPlaceRate")
                    _mr_col = col_key(_mc_table_disp, "OrtalamaSıra", "MeanRank")
                    if _fp_col in _mc_table_disp.columns:
                        _mc_table_disp[_fp_col] = (pd.to_numeric(_mc_table_disp[_fp_col], errors="coerce") * 100.0).round(2)
                        _mc_table_disp = _mc_table_disp.rename(columns={_fp_col: f"{_fp_col} (%)"})
                    if _mr_col in _mc_table_disp.columns:
                        _mc_table_disp[_mr_col] = pd.to_numeric(_mc_table_disp[_mr_col], errors="coerce").round(2)
                    render_table(_mc_table_disp)
                    with st.expander(f"💬 {tt('Tablo Yorumu', 'Table Commentary')}", expanded=False):
                        st.markdown(f'<div class="commentary-box">{_html_to_plain(_mc_comment).replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

                with st.expander(f"📊 {tt('Şekiller', 'Figures')}", expanded=False):
                    mg1, mg2 = st.columns(2)
                    with mg1:
                        st.plotly_chart(fig_mc_rank_bar(_mc_disp), use_container_width=True)
                        with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                            st.markdown(f'<div class="commentary-box">{tt("Birincilik oranı çubuğu, her alternatifin Monte Carlo simülasyonları boyunca kaç kez birinci sıraya oturduğunu yüzde olarak gösterir. %50 üzeri oran güçlü kararlılığı ifade eder; %20 altı oran alternatifin liderliğinin rastlantısal olduğuna işaret edebilir.", "The first-place rate bar shows what percentage of Monte Carlo simulations each alternative ranked first. A rate above 50% indicates strong stability; a rate below 20% may indicate the alternative\'s leadership is coincidental.")}</div>', unsafe_allow_html=True)
                    with mg2:
                        st.plotly_chart(fig_mc_stability_bubble(_mc_disp), use_container_width=True)
                        with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                            st.markdown(f'<div class="commentary-box">{tt("Balon grafiği, birincilik oranı ile ortalama sıra arasındaki ilişkiyi gösterir. Sol üst köşeye yakın (düşük ortalama sıra, yüksek birincilik oranı) alternatifler en sağlam adaylardır.", "The bubble chart shows the relationship between first-place rate and mean rank. Alternatives close to the upper-left corner (low mean rank, high first-place rate) are the most robust candidates.")}</div>', unsafe_allow_html=True)

                    if isinstance(loc_df, pd.DataFrame) and not loc_df.empty:
                        st.plotly_chart(fig_local_sensitivity(loc_df), use_container_width=True)
                        with st.expander(f"💬 {tt('Şekil Yorumu', 'Figure Commentary')}", expanded=False):
                            st.markdown(f'<div class="commentary-box">{tt("Lokal duyarlılık grafiği, her kriterin ağırlığı ±%10 ve ±%20 değiştirildiğinde sıralamanın ne kadar etkilendiğini Spearman korelasyonu ile ölçer. ρ = 1.0\'a yakın çizgiler, o kriterin ağırlık değişimine karşı sıralamanın kararlı kaldığını gösterir; aşağı düşen çizgiler kritik duyarlılık noktalarına işaret eder.", "The local sensitivity chart measures how much the ranking is affected when each criterion weight changes by ±10% and ±20%, using Spearman correlation. Lines close to ρ = 1.0 show the ranking remains stable against that criterion weight change; lines dropping down indicate critical sensitivity points.")}</div>', unsafe_allow_html=True)

with tabs[_tab_output]:
    _out_lang = st.session_state.get("ui_lang", "TR")
    _doc_sections = _build_doc_sections(result, _out_lang)
    _ref_heading = tt("Kaynakça", "References")
    _refs = _doc_sections.get(_ref_heading, [])
    _is_panel_download = isinstance(panel_results, dict) and bool(panel_results)
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            tt("📊 Tüm Sonuçları İndir (Excel)", "📊 Download All Results (Excel)"),
            data=(generate_panel_excel(panel_results, lang=_out_lang) if _is_panel_download else generate_excel(result, result["selected_data"], lang=_out_lang)),
            file_name=tt("MCDM_Sonuclari.xlsx", "MCDM_Results.xlsx"),
            use_container_width=True,
        )
    with dl2:
        if DOCX_AVAILABLE:
            _doc_bytes = (
                generate_panel_apa_docx(panel_results, lang=_out_lang)
                if _is_panel_download
                else generate_apa_docx(result, result["selected_data"], lang=_out_lang)
            )
            st.session_state["report_docx"] = _doc_bytes
            st.download_button(
                tt("📄 Akademik Raporu İndir — APA Word", "📄 Download Academic Report — APA Word"),
                data=_doc_bytes,
                file_name=tt("MCDM_Akademik_Rapor.docx", "MCDM_Academic_Report.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            st.warning(tt("Word çıktısı için python-docx kurulu olmalıdır.", "python-docx must be installed for Word output."))

    st.markdown(_reference_notice_html(_out_lang), unsafe_allow_html=True)
    st.markdown(f"### 📚 {_ref_heading}")
    if _refs:
        st.markdown("\n".join(f"{idx}. {ref}" for idx, ref in enumerate(_refs, start=1)))
    else:
        st.info(tt("Kaynakça bilgisi üretilemedi.", "Reference information could not be generated."))

st.markdown(
    f"""
    <div style="text-align:right; padding: 16px 20px 6px 20px; font-family:'Georgia',serif;
                color:#1d3557; font-size:15px; font-weight:600; font-style:italic;">
        Prof. Dr. Ömer Faruk Rençber
    </div>
    <div style="text-align:center; padding: 4px 0 24px 0; font-size:0.80rem; color:#718096;">
        <a href="https://www.ofrencber.com" target="_blank"
           style="color:#2E7D9E; text-decoration:none; font-weight:600;">
            🌐 www.ofrencber.com
        </a>
        &nbsp;·&nbsp; MCDM Toolbox Professional &nbsp;·&nbsp; {tt("Akademik Karar Destek Sistemi", "Academic Decision Support System")}
    </div>
    """,
    unsafe_allow_html=True,
)
