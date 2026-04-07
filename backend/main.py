"""
KBR RDL Data Harmonizer v3.0 — Streamlit Application
=====================================================

Entry point:  streamlit run backend/main.py
Standalone:   python launcher.py  (or RDL.exe after running build_exe.py)

Architecture overview
---------------------
This file is the complete Streamlit UI layer.  All pure business logic lives
in core.py so it can be reused independently of Streamlit.

    core.py          — framework-agnostic engine (readers, fuzzy match, harmonization)
    main.py (this)   — Streamlit UI + thin wrapper functions that call core.py

Note on local function copies
------------------------------
Several helper functions (_tokenize, _has_word_overlap, _expand_compound_names,
fuzzy_match, the file readers) exist both here and in core.py.  The copies in
this file are used by local Excel-building and enrichment functions that were
written before the core module was extracted.  The canonical source of truth
for those algorithms is core.py; the two sets are functionally identical.

Tab layout (in render order)
-----------------------------
  Tab 1 — Upload & Configure    upload files, select masters, set threshold
  Tab 2 — Dashboard             KPI cards, status charts, discipline breakdown
  Tab 3 — Gap Analysis          per-class gap table, filter, download
  Tab 4 — Enrichment Suggestions  reverse-gap additions, CFIHOS lookup, export
  Tab 5 — Attributes Explorer   side-by-side Aramco vs LTC attribute browsing
  Tab 6 — Connection Map        Graphviz diagrams per class or all-classes overview
  Tab 7 — Search                single-class lookup across all sources
  Tab 8 — Batch Process         multi-name CSV/text batch lookup
  Tab 9 — Logs                  session-level processing log
"""

import uuid
import io
import re
import pandas as pd
import streamlit as st
import altair as alt
from rapidfuzz import fuzz, process

import core as _core  # Framework-agnostic engine (for FastAPI migration)
from core import (
    FILE_TYPES, SOURCE_COLORS, ABBREVIATIONS,
    _safe_str, _normalize_abbreviations,
)

# ──────────────────────────────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="KBR RDL Data Harmonizer",
    page_icon="⚙️",
    layout="wide",
)

# ──────────────────────────────────────────────────────────────────────
# Global CSS — KBR Professional Theme
# ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Import font ── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ── Root variables ── */
:root {
    --kbr-primary: #003087;
    --kbr-primary-dark: #001d54;
    --kbr-primary-light: #c7d6f0;
    --kbr-navy: #1B3A5C;
    --kbr-blue: #2563eb;
    --kbr-green: #059669;
    --kbr-orange: #d97706;
    --kbr-purple: #7c3aed;
    --kbr-gray: #64748b;
    --kbr-accent: #00A3E0;
    --card-bg: #ffffff;
    --card-border: #e2e8f0;
    --card-shadow: 0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.06);
    --card-shadow-hover: 0 4px 12px rgba(0,0,0,0.12);
    --radius: 12px;
}

/* ── Global font ── */
html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* ── Hide default Streamlit branding ── */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* ── Main container padding ── */
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 2rem !important;
    max-width: 1400px;
}

/* ── Tab styling ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
    border-radius: 12px 12px 0 0;
    padding: 4px 4px 0 4px;
    border-bottom: 2px solid var(--kbr-primary);
}
.stTabs [data-baseweb="tab"] {
    font-weight: 500;
    font-size: 0.85rem;
    color: var(--kbr-gray);
    padding: 0.6rem 1.2rem;
    border-radius: 8px 8px 0 0;
    transition: all 0.2s ease;
}
.stTabs [aria-selected="true"] {
    background: white !important;
    color: var(--kbr-primary) !important;
    font-weight: 700;
    border-top: 3px solid var(--kbr-primary);
}
.stTabs [data-baseweb="tab"]:hover {
    color: var(--kbr-primary);
    background: rgba(0, 48, 135, 0.05);
}
.stTabs [data-baseweb="tab-panel"] {
    padding: 1.5rem 0.5rem;
}

/* ── Metric card override ── */
[data-testid="stMetric"] {
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    border-left: 4px solid var(--kbr-primary);
    border-radius: var(--radius);
    padding: 1rem 1.2rem;
    box-shadow: var(--card-shadow);
    transition: all 0.2s ease;
}
[data-testid="stMetric"]:hover {
    box-shadow: var(--card-shadow-hover);
    transform: translateY(-1px);
}
[data-testid="stMetric"] label {
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--kbr-gray) !important;
}
[data-testid="stMetric"] [data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 800 !important;
    color: var(--kbr-navy) !important;
}

/* ── Buttons ── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--kbr-primary) 0%, var(--kbr-primary-dark) 100%) !important;
    border: none !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em;
    padding: 0.6rem 1.5rem !important;
    border-radius: 8px !important;
    box-shadow: 0 2px 8px rgba(0, 48, 135, 0.3) !important;
    transition: all 0.2s ease !important;
}
.stButton > button[kind="primary"]:hover {
    box-shadow: 0 4px 16px rgba(0, 48, 135, 0.4) !important;
    transform: translateY(-1px);
}
.stButton > button[kind="secondary"], .stButton > button:not([kind]) {
    border: 1.5px solid var(--card-border) !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
    transition: all 0.2s ease !important;
}
.stButton > button[kind="secondary"]:hover, .stButton > button:not([kind]):hover {
    border-color: var(--kbr-primary) !important;
    color: var(--kbr-primary) !important;
}

/* ── Download button ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, var(--kbr-green) 0%, #047857 100%) !important;
    color: white !important;
    border: none !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    box-shadow: 0 2px 8px rgba(5, 150, 105, 0.3) !important;
}
.stDownloadButton > button:hover {
    box-shadow: 0 4px 16px rgba(5, 150, 105, 0.4) !important;
}

/* ── Expander ── */
.streamlit-expanderHeader {
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    border-radius: 8px !important;
    background: #f8fafc !important;
}

/* ── Dataframes ── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--card-border);
    border-radius: var(--radius);
    overflow: hidden;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    border: 2px dashed var(--card-border);
    border-radius: var(--radius);
    padding: 0.5rem;
    transition: border-color 0.2s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--kbr-primary);
}

/* ── Selectbox / multiselect ── */
.stSelectbox > div > div, .stMultiSelect > div > div {
    border-radius: 8px !important;
}

/* ── Slider ── */
.stSlider > div > div > div > div {
    background: var(--kbr-primary) !important;
}

/* ── Dividers ── */
hr {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--card-border), transparent);
    margin: 1.5rem 0;
}

/* ── Section card helper ── */
.section-card {
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    border-radius: var(--radius);
    padding: 1.5rem;
    box-shadow: var(--card-shadow);
    margin-bottom: 1rem;
}
.section-card h3 {
    margin-top: 0;
    color: var(--kbr-navy);
    font-weight: 700;
    font-size: 1.1rem;
}

/* ── Stat pill ── */
.stat-pill {
    display: inline-block;
    padding: 0.25rem 0.75rem;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 600;
}
.stat-pill.green { background: #d1fae5; color: #065f46; }
.stat-pill.primary { background: #c7d6f0; color: #003087; }
.stat-pill.blue { background: #dbeafe; color: #1e40af; }
.stat-pill.orange { background: #ffedd5; color: #9a3412; }
.stat-pill.purple { background: #ede9fe; color: #5b21b6; }

/* ── KPI row ── */
.kpi-row {
    display: flex; gap: 1rem; flex-wrap: wrap; margin: 1rem 0;
}
.kpi-card {
    flex: 1; min-width: 160px;
    background: white;
    border: 1px solid var(--card-border);
    border-radius: var(--radius);
    padding: 1.2rem;
    text-align: center;
    box-shadow: var(--card-shadow);
    transition: all 0.2s ease;
}
.kpi-card:hover {
    box-shadow: var(--card-shadow-hover);
    transform: translateY(-2px);
}
.kpi-card .kpi-value {
    font-size: 2rem;
    font-weight: 800;
    line-height: 1.1;
}
.kpi-card .kpi-label {
    font-size: 0.75rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--kbr-gray);
    margin-top: 0.3rem;
}
.kpi-card .kpi-sub {
    font-size: 0.8rem;
    color: var(--kbr-gray);
    margin-top: 0.2rem;
}

/* ── Chart container ── */
.chart-container {
    background: white;
    border: 1px solid var(--card-border);
    border-radius: var(--radius);
    padding: 1rem 1.5rem;
    box-shadow: var(--card-shadow);
    margin: 0.75rem 0;
}
.chart-container h4 {
    margin: 0 0 0.75rem 0;
    font-size: 0.95rem;
    font-weight: 700;
    color: var(--kbr-navy);
}

/* ── Status badge ── */
.status-badge {
    display: inline-flex; align-items: center; gap: 0.35rem;
    padding: 0.2rem 0.7rem;
    border-radius: 6px;
    font-size: 0.78rem;
    font-weight: 600;
}
.status-badge.match { background: #d1fae5; color: #065f46; }
.status-badge.gap { background: #fee2e2; color: #991b1b; }
.status-badge.partial { background: #ffedd5; color: #9a3412; }

/* ── Empty state ── */
.empty-state {
    text-align: center;
    padding: 4rem 2rem;
    color: var(--kbr-gray);
}
.empty-state .empty-icon {
    font-size: 3rem;
    margin-bottom: 1rem;
    opacity: 0.5;
}
.empty-state h3 {
    color: var(--kbr-navy);
    font-weight: 700;
    margin-bottom: 0.5rem;
}
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────
# Session state defaults
# ──────────────────────────────────────────────────────────────────────
DEFAULTS = {
    "files": {},
    "masters": [],
    "classes": [],
    "logs": [],
    "match_threshold": 75,
    "aramco_attrs": None,
    "ltc_attrs": None,
    "reverse_gaps": {},
    "search_matched_class": None,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

def add_log(msg, level="INFO"):
    """Append a timestamped log entry to the session log list."""
    st.session_state.logs.append(f"[{level}] {msg}")


# ──────────────────────────────────────────────────────────────────────
# File readers  (local copies — see core.py for full docstrings)
# ──────────────────────────────────────────────────────────────────────
def read_aramco(file) -> list[dict]:
    """Parse Saudi Aramco ISM Functional Classes Excel. See core.read_aramco."""
    try:
        df = pd.read_excel(file, sheet_name="ISM Functional Classes")
    except (ValueError, KeyError):
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for row in df.to_dict("records"):
        name = _safe_str(row.get("Name", row.get("name", "")))
        if not name:
            continue
        records.append({
            "id": _safe_str(row.get("Id", row.get("id", ""))),
            "name": name,
            "cfihos_ref": _safe_str(row.get("nmcltr:CFIHOS_1.5", "")),
            "source": "aramco",
        })
    return records


def read_cfihos(file) -> list[dict]:
    """Parse CFIHOS standard equipment class Excel. See core.read_cfihos."""
    try:
        df = pd.read_excel(file, sheet_name="equipment class")
    except (ValueError, KeyError):
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for row in df.to_dict("records"):
        name = _safe_str(row.get("equipment class name", row.get("name", "")))
        if not name:
            continue
        records.append({
            "id": _safe_str(row.get("CFIHOS unique id", row.get("id", ""))),
            "name": name,
            "source": "cfihos",
        })
    return records


def read_kbr(file) -> list[dict]:
    """Parse KBR FEED class library Excel. See core.read_kbr."""
    df = pd.read_excel(file, sheet_name=0)
    records = []
    for row in df.to_dict("records"):
        name = _safe_str(row.get("Class Name (855)", row.get("name", "")))
        if not name:
            continue
        records.append({
            "id": _safe_str(row.get("Class Id", row.get("id", ""))),
            "name": name,
            "discipline": _safe_str(row.get("Discipline", "")),
            "source": "kbr",
        })
    return records


def read_ltc(file) -> list[dict]:
    """Parse LTC ISM class Excel, stripping bracket annotations. See core.read_ltc."""
    # Prefer "ISM Physical Classes" (PCL IDs) over "ISM Functional Classes" (FCL IDs)
    df = None
    for sheet in ("ISM Physical Classes", "ISM Functional Classes"):
        try:
            df = pd.read_excel(file, sheet_name=sheet)
            break
        except (ValueError, KeyError):
            continue
    if df is None:
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for row in df.to_dict("records"):
        raw_name = _safe_str(row.get("Name", row.get("name", "")))
        if not raw_name:
            continue
        # Strip [OBSOLETE] or similar bracketed prefixes so fuzzy matching works
        clean_name = re.sub(r'\[.*?\]\s*', '', raw_name).strip()
        records.append({
            "id": _safe_str(row.get("Id", row.get("id", ""))),
            "name": clean_name if clean_name else raw_name,
            "source": "ltc",
        })
    return records


def read_sa_doc(file) -> list[dict]:
    """Parse SA Document attributes Excel. See core.read_sa_doc."""
    try:
        df = pd.read_excel(file, sheet_name="SA_DOC_attributes")
    except (ValueError, KeyError):
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for row in df.to_dict("records"):
        name = _safe_str(row.get("Attribute", row.get("name", "")))
        if not name:
            continue
        records.append({
            "id": _safe_str(row.get("ID (CFIHOS_1.5)", row.get("id", ""))),
            "name": name,
            "cfihos_name": _safe_str(row.get("Name (CFIHOS_1.5)", "")),
            "source": "sa_doc",
        })
    return records


READERS = {
    "aramco": read_aramco,
    "cfihos": read_cfihos,
    "kbr": read_kbr,
    "ltc": read_ltc,
    "sa_doc": read_sa_doc,
}


# ──────────────────────────────────────────────────────────────────────
# Attribute readers
# ──────────────────────────────────────────────────────────────────────
def read_aramco_attributes(file) -> pd.DataFrame:
    """Read Aramco 'ISM Functional Class Attributes' sheet."""
    try:
        df = pd.read_excel(file, sheet_name="ISM Functional Class Attributes")
    except (ValueError, KeyError):
        return pd.DataFrame()
    col_map = {}
    for c in df.columns:
        cl = str(c).strip()
        if cl.lower() in ("class_id", "class id"):
            col_map[c] = "Class_Id"
        elif cl.lower() == "id":
            col_map[c] = "Attribute_Id"
        elif cl.lower().startswith("lookup fc desc"):
            col_map[c] = "Class_Desc"
        elif cl.lower().startswith("lookup att desc"):
            col_map[c] = "Attribute_Desc"
        elif cl.lower().startswith("lookup group"):
            col_map[c] = "Group_Id"
        elif cl.lower() == "name":
            col_map[c] = "Name"
        elif cl.lower() == "description":
            col_map[c] = "Description"
        elif cl.lower() == "size":
            col_map[c] = "Size"
        elif cl.lower() == "presence":
            col_map[c] = "Presence"
        elif cl.lower() == "discipline":
            col_map[c] = "Discipline"
        elif cl.lower() == "uomclassid":
            col_map[c] = "UomClassId"
        elif cl.lower() == "uomrequire":
            col_map[c] = "UomRequire"
        elif cl.lower() == "validationrule":
            col_map[c] = "ValidationRule"
        elif cl.lower() == "action":
            col_map[c] = "Action"
    df = df.rename(columns=col_map)
    # Back-fill Name from Attribute_Desc when Name is empty/NaN
    if "Attribute_Desc" in df.columns and "Name" in df.columns:
        df["Name"] = df["Name"].astype(object)
        mask = df["Name"].isna() | (df["Name"].astype(str).str.strip() == "") | (df["Name"].astype(str).str.lower() == "nan")
        df.loc[mask, "Name"] = df.loc[mask, "Attribute_Desc"]
    elif "Attribute_Desc" in df.columns and "Name" not in df.columns:
        df["Name"] = df["Attribute_Desc"]
    df["_source"] = "aramco"
    return df


def read_ltc_attributes(file) -> pd.DataFrame:
    """Read LTC 'ISM Physical Class Attributes' sheet."""
    try:
        df = pd.read_excel(file, sheet_name="ISM Physical Class Attributes")
    except (ValueError, KeyError):
        return pd.DataFrame()
    col_map = {}
    for c in df.columns:
        cl = str(c).strip()
        if cl.lower() in ("class_id", "class id"):
            col_map[c] = "Class_Id"
        elif cl.lower() == "id":
            col_map[c] = "Attribute_Id"
        elif cl.lower() == "name":
            col_map[c] = "Name"
        elif cl.lower() == "description":
            col_map[c] = "Description"
        elif cl.lower() == "size":
            col_map[c] = "Size"
        elif cl.lower() == "presence":
            col_map[c] = "Presence"
        elif cl.lower() == "discipline":
            col_map[c] = "Discipline"
        elif cl.lower() == "uomclassid":
            col_map[c] = "UomClassId"
        elif cl.lower() == "uomrequire":
            col_map[c] = "UomRequire"
        elif cl.lower() == "validationrule":
            col_map[c] = "ValidationRule"
        elif cl.lower() == "validationtype":
            col_map[c] = "ValidationType"
        elif cl.lower() == "maxoccur":
            col_map[c] = "MaxOccur"
        elif cl.lower() == "minoccurs":
            col_map[c] = "MinOccurs"
        elif cl.lower() == "sortorder":
            col_map[c] = "SortOrder"
        elif cl.lower() == "obsolete":
            col_map[c] = "Obsolete"
        elif cl.lower() == "aspect":
            col_map[c] = "Aspect"
        elif cl.lower() == "_action":
            col_map[c] = "Action"
    df = df.rename(columns=col_map)
    df["_source"] = "ltc"
    return df


# ──────────────────────────────────────────────────────────────────────
# Harmonization engine
# NOTE: These local copies (_tokenize, _has_word_overlap, _expand_compound_names,
# fuzzy_match) mirror the equivalents in core.py. They exist here because
# build_excel_bytes and enrichment helpers were written before core.py was
# extracted. The canonical documented versions live in core.py.
# ──────────────────────────────────────────────────────────────────────
def _tokenize(text):
    """Extract meaningful word tokens (see core.py for full docstring)."""
    words = set(re.findall(r'[a-z]{2,}', text.lower()))
    noise = {"the", "and", "for", "with", "from", "that", "this", "its",
             "type", "class", "system", "item", "general", "other", "misc"}
    return words - noise


def _has_word_overlap(name_a, name_b):
    """Return True if two equipment names share a meaningful word (see core.py)."""
    tokens_a = _tokenize(name_a)
    tokens_b = _tokenize(name_b)
    if not tokens_a or not tokens_b:
        return True
    if tokens_a & tokens_b:
        return True
    # Check abbreviation-expanded overlap
    exp_a = set(_tokenize(_normalize_abbreviations(name_a)))
    exp_b = set(_tokenize(_normalize_abbreviations(name_b)))
    if exp_a & exp_b:
        return True
    # Substring containment check
    all_a = tokens_a | exp_a
    all_b = tokens_b | exp_b
    for a in all_a:
        for b in all_b:
            if len(a) >= 3 and len(b) >= 3 and (a in b or b in a):
                return True
    return False


def _expand_compound_names(candidates):
    """Expand slash-separated names into matchable parts (see core.py)."""
    expanded_names = []
    index_map = []
    for i, c in enumerate(candidates):
        expanded_names.append(c["name"])
        index_map.append(i)
        if "/" in c["name"]:
            for part in [p.strip() for p in c["name"].split("/") if p.strip()]:
                expanded_names.append(part)
                index_map.append(i)
    return expanded_names, index_map


def fuzzy_match(name, candidates, threshold):
    """Find best-matching record using three-pass fuzzy scoring (see core.py)."""
    if not candidates:
        return None
    expanded_names, index_map = _expand_compound_names(candidates)
    query_variants = [name]
    if "/" in name:
        query_variants += [p.strip() for p in name.split("/") if p.strip()]

    best_result = None
    best_score = 0

    # Try matching with both original and abbreviation-expanded forms
    expanded_candidates = [_normalize_abbreviations(n) for n in expanded_names]

    for q in query_variants:
        q_exp = _normalize_abbreviations(q)
        # Match against original names
        result = process.extractOne(
            q, expanded_names, scorer=fuzz.token_sort_ratio, score_cutoff=threshold,
        )
        if result and result[1] > best_score:
            best_result = result
            best_score = result[1]
        # Match against abbreviation-expanded names
        result_exp = process.extractOne(
            q_exp, expanded_candidates, scorer=fuzz.token_sort_ratio, score_cutoff=threshold,
        )
        if result_exp and result_exp[1] > best_score:
            best_result = (expanded_names[result_exp[2]], result_exp[1], result_exp[2])
            best_score = result_exp[1]
        # Also try token_set_ratio for better partial matching
        result_set = process.extractOne(
            q, expanded_names, scorer=fuzz.token_set_ratio, score_cutoff=threshold,
        )
        if result_set and result_set[1] > best_score:
            best_result = result_set
            best_score = result_set[1]

    if best_result is None:
        return None
    matched_name, score, exp_idx = best_result
    orig_idx = index_map[exp_idx]
    if not _has_word_overlap(name, matched_name) and not _has_word_overlap(name, candidates[orig_idx]["name"]):
        return None
    return {**candidates[orig_idx], "score": int(round(score))}


def classify_match(score):
    if score >= 95:
        return "Exact Match", "High"
    elif score >= 85:
        return "Strong Match", "High"
    elif score >= 75:
        return "Partial Match", "Medium"
    else:
        return "Weak Match", "Low"


def run_harmonization():
    """Wrapper: reads session state, calls core engine, stores results."""
    masters = st.session_state.masters
    files = st.session_state.files
    threshold = st.session_state.match_threshold

    harmonized, reverse_gaps, logs = _core.run_harmonization(masters, files, threshold)
    for msg in logs:
        add_log(msg)
    st.session_state.reverse_gaps = reverse_gaps
    return harmonized


# ──────────────────────────────────────────────────────────────────────
# Demo data
# ──────────────────────────────────────────────────────────────────────
def load_demo_data():
    demo = _core.get_demo_data()
    # Add sa_doc which is specific to the full demo
    demo["sa_doc"] = [
        {"id": "SA-A01", "name": "Design Pressure", "cfihos_name": "Design Pressure", "source": "sa_doc"},
        {"id": "SA-A02", "name": "Design Temperature", "cfihos_name": "Design Temperature", "source": "sa_doc"},
        {"id": "SA-A03", "name": "Material of Construction", "cfihos_name": "Material", "source": "sa_doc"},
    ]
    for key, records in demo.items():
        st.session_state.files[key] = {"filename": f"demo_{key}.xlsx", "records": records}
    add_log("Demo data loaded for all 5 sources")


# ──────────────────────────────────────────────────────────────────────
# Export — clean organized Excel
# ──────────────────────────────────────────────────────────────────────
def build_export_df(classes):
    """Wrapper: calls core build_export_df with session state masters."""
    return _core.build_export_df(classes, st.session_state.masters)


def build_excel_bytes(classes):
    """Build the main harmonization export with professional styling."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import openpyxl

    masters = st.session_state.masters

    # Style constants — KBR corporate navy blue
    kbr_primary = "003087"
    kbr_dark = "001D54"
    kbr_accent = "00A3E0"
    navy = "1B3A5C"
    white = "FFFFFF"
    light_gray = "F8FAFC"
    border_gray = "E2E8F0"
    green_bg = "D1FAE5"
    green_fg = "065F46"
    red_bg = "FEE2E2"
    red_fg = "991B1B"

    thin_border = Border(
        left=Side(style="thin", color=border_gray),
        right=Side(style="thin", color=border_gray),
        top=Side(style="thin", color=border_gray),
        bottom=Side(style="thin", color=border_gray),
    )
    title_font = Font(name="Calibri", size=16, bold=True, color=white)
    title_fill = PatternFill("solid", fgColor=kbr_primary)
    sub_font = Font(name="Calibri", size=10, color=white)
    sub_fill = PatternFill("solid", fgColor=kbr_dark)
    header_font = Font(name="Calibri", size=10, bold=True, color=white)
    header_fill = PatternFill("solid", fgColor=navy)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="Calibri", size=10, color=navy)
    data_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    alt_fill = PatternFill("solid", fgColor=light_gray)
    white_fill = PatternFill("solid", fgColor=white)
    gap_fill = PatternFill("solid", fgColor=red_bg)
    gap_font = Font(name="Calibri", size=10, bold=True, color=red_fg)
    nogap_fill = PatternFill("solid", fgColor=green_bg)
    nogap_font = Font(name="Calibri", size=10, bold=True, color=green_fg)

    logo_font = Font(name="Calibri", size=14, bold=True, color=kbr_accent)
    logo_sub_font = Font(name="Calibri", size=9, color="94A3B8")
    logo_fill = PatternFill("solid", fgColor=kbr_dark)

    def _style_title_rows(ws, title_text, subtitle_text, num_cols):
        """Add KBR-AMCDE logo + branded title + subtitle rows (rows 1-3)."""
        # Row 1: KBR-AMCDE logo bar
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        logo_cell = ws.cell(row=1, column=1, value="  KBR-AMCDE  |  RDL Data Harmonizer v3.0")
        logo_cell.font = logo_font
        logo_cell.fill = logo_fill
        logo_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30
        # Row 2: Title
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
        t = ws.cell(row=2, column=1, value=f"  {title_text}")
        t.font = title_font
        t.fill = title_fill
        t.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 36
        # Row 3: Subtitle
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=num_cols)
        s = ws.cell(row=3, column=1, value=f"  {subtitle_text}")
        s.font = sub_font
        s.fill = sub_fill
        s.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[3].height = 24

    def _write_headers(ws, row, headers):
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=ci, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        ws.row_dimensions[row].height = 28

    # ── Build data ──
    df = build_export_df(classes)
    total = len(classes)
    non_master_sources = set()
    for c in classes:
        non_master_sources.update(c["matches"].keys())
        non_master_sources.update(c["gaps"])
    non_master_sources = sorted(non_master_sources)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Sheet 1: Harmonization Results ──
    ws1 = wb.create_sheet("Harmonization Results")
    cols = list(df.columns)
    master_label = " + ".join(FILE_TYPES.get(m, m) for m in masters)
    compared = ", ".join(FILE_TYPES.get(s, s) for s in non_master_sources)
    _style_title_rows(ws1, "KBR RDL — Harmonization Results",
                      f"Master: {master_label}  |  Compared: {compared}  |  Threshold: {st.session_state.match_threshold}%  |  Total: {total} classes",
                      len(cols))
    _write_headers(ws1, 4, cols)
    ws1.auto_filter.ref = f"A4:{get_column_letter(len(cols))}4"
    ws1.freeze_panes = "A5"

    for ri, (_, row) in enumerate(df.iterrows(), 5):
        is_alt = (ri % 2 == 0)
        for ci, col_name in enumerate(cols, 1):
            val = row[col_name]
            cell = ws1.cell(row=ri, column=ci, value=val)
            cell.font = data_font
            cell.alignment = center_align if col_name in ("#", "Gap Count", "Masters Cross-Match%") or "Match%" in str(col_name) else data_align
            cell.border = thin_border
            # Color the Gap In / Status columns
            if col_name == "Gap In":
                if val == "No Gaps":
                    cell.fill = nogap_fill
                    cell.font = nogap_font
                elif val:
                    cell.fill = gap_fill
                    cell.font = gap_font
                else:
                    cell.fill = alt_fill if is_alt else white_fill
            elif "Status" in str(col_name):
                if val == "GAP":
                    cell.fill = gap_fill
                    cell.font = gap_font
                elif val and "Exact" in str(val):
                    cell.fill = nogap_fill
                    cell.font = nogap_font
                else:
                    cell.fill = alt_fill if is_alt else white_fill
            else:
                cell.fill = alt_fill if is_alt else white_fill

    # Auto-fit
    for ci, col_name in enumerate(cols, 1):
        max_len = max(len(str(col_name)), max((len(str(row[col_name] or "")) for _, row in df.iterrows()), default=5))
        ws1.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 45)

    # ── Sheet 2: Gap Analysis ──
    ws2 = wb.create_sheet("Gap Analysis")
    gap_rows = []
    for c in classes:
        for g in c["gaps"]:
            present_in = [FILE_TYPES.get(m, m) for m in masters if m in c["master_entries"]]
            present_in += [FILE_TYPES.get(s, s) for s in c["matches"]]
            gap_rows.append({
                "#": c["index"],
                "Equipment Class": c["canonical_name"],
                "Found In": ", ".join(present_in),
                "Gap In": FILE_TYPES.get(g, g),
                "Action": f"Add '{c['canonical_name']}' to {FILE_TYPES.get(g, g)} or confirm exclusion",
            })
    gap_cols = ["#", "Equipment Class", "Found In", "Gap In", "Action"]
    no_gaps_cnt = sum(1 for c in classes if not c["gaps"])
    _style_title_rows(ws2, "KBR RDL — Gap Analysis",
                      f"{len(gap_rows)} gaps identified  |  {no_gaps_cnt}/{total} classes fully matched",
                      len(gap_cols))
    _write_headers(ws2, 4, gap_cols)
    ws2.auto_filter.ref = f"A4:{get_column_letter(len(gap_cols))}4"
    ws2.freeze_panes = "A5"

    for ri, grow in enumerate(gap_rows, 5):
        is_alt = (ri % 2 == 0)
        for ci, col_name in enumerate(gap_cols, 1):
            val = grow[col_name]
            cell = ws2.cell(row=ri, column=ci, value=val)
            cell.font = data_font
            cell.alignment = center_align if col_name == "#" else data_align
            cell.border = thin_border
            if col_name == "Gap In":
                cell.fill = gap_fill
                cell.font = gap_font
            else:
                cell.fill = alt_fill if is_alt else white_fill
    gap_widths = [6, 40, 40, 25, 55]
    for ci, w in enumerate(gap_widths, 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w

    # ── Sheet 3: Summary ──
    ws3 = wb.create_sheet("Summary")
    _style_title_rows(ws3, "KBR RDL — Summary Report", f"Generated from harmonization of {total} equipment classes", 2)
    _write_headers(ws3, 4, ["Metric", "Value"])

    no_gaps = sum(1 for c in classes if not c["gaps"])
    summary_data = [
        ("Total Equipment Classes", total),
        ("Master Reference", master_label),
        ("Compared Against", compared),
        ("Match Threshold", f"{st.session_state.match_threshold}%"),
        ("Classes with No Gaps", f"{no_gaps} / {total}"),
        ("Total Gap Entries", sum(len(c["gaps"]) for c in classes)),
    ]
    for src in non_master_sources:
        cnt = sum(1 for c in classes if src in c["matches"])
        pct = int(round(cnt / total * 100)) if total else 0
        gap_cnt = total - cnt
        summary_data.append((FILE_TYPES.get(src, src), f"Matched: {cnt}/{total} ({pct}%) — Gaps: {gap_cnt}"))

    for ri, (metric, value) in enumerate(summary_data, 5):
        is_alt = (ri % 2 == 0)
        mc = ws3.cell(row=ri, column=1, value=metric)
        mc.font = Font(name="Calibri", size=10, bold=True, color=navy)
        mc.fill = alt_fill if is_alt else white_fill
        mc.border = thin_border
        vc = ws3.cell(row=ri, column=2, value=value)
        vc.font = data_font
        vc.fill = alt_fill if is_alt else white_fill
        vc.border = thin_border
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 50

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _lookup_cfihos(name, cfihos_records, threshold):
    """Try to fuzzy-match a name against CFIHOS records. Returns (id, name, match_label)."""
    if not cfihos_records:
        return "", "", ""
    match = fuzzy_match(name, cfihos_records, threshold)
    if match:
        score = match["score"]
        label = "Exact" if score == 100 else f"{score}%"
        return match["id"], match["name"], label
    return "", "", ""


def build_enriched_master_excel(selected_additions):
    """Build a professionally styled Excel file: original master records + selected additions."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    masters = st.session_state.masters
    files = st.session_state.files
    threshold = st.session_state.match_threshold
    buf = io.BytesIO()

    # ── Style definitions — KBR corporate navy blue ──
    kbr_primary = "003087"
    kbr_dark = "001D54"
    kbr_accent = "00A3E0"
    navy = "1B3A5C"
    green_bg = "D1FAE5"
    green_fg = "065F46"
    blue_bg = "DBEAFE"
    blue_fg = "1E40AF"
    orange_bg = "FFEDD5"
    orange_fg = "9A3412"
    light_gray = "F8FAFC"
    border_gray = "E2E8F0"
    white = "FFFFFF"

    thin_border = Border(
        left=Side(style="thin", color=border_gray),
        right=Side(style="thin", color=border_gray),
        top=Side(style="thin", color=border_gray),
        bottom=Side(style="thin", color=border_gray),
    )

    # Title row style
    title_font = Font(name="Calibri", size=16, bold=True, color=white)
    title_fill = PatternFill("solid", fgColor=kbr_primary)
    title_align = Alignment(horizontal="left", vertical="center")

    # Subtitle row
    sub_font = Font(name="Calibri", size=10, color=white)
    sub_fill = PatternFill("solid", fgColor=kbr_dark)

    # Header row style
    header_font = Font(name="Calibri", size=10, bold=True, color=white)
    header_fill = PatternFill("solid", fgColor=navy)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Data row styles
    orig_fill = PatternFill("solid", fgColor=white)
    sugg_fill = PatternFill("solid", fgColor=blue_bg)
    data_font = Font(name="Calibri", size=10, color=navy)
    data_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")

    # Match column styles
    exact_fill = PatternFill("solid", fgColor=green_bg)
    exact_font = Font(name="Calibri", size=10, bold=True, color=green_fg)
    partial_fill = PatternFill("solid", fgColor=orange_bg)
    partial_font = Font(name="Calibri", size=10, bold=True, color=orange_fg)
    status_orig_fill = PatternFill("solid", fgColor=green_bg)
    status_orig_font = Font(name="Calibri", size=10, bold=True, color=green_fg)
    status_sugg_fill = PatternFill("solid", fgColor=blue_bg)
    status_sugg_font = Font(name="Calibri", size=10, bold=True, color=blue_fg)

    # Alt row
    alt_fill = PatternFill("solid", fgColor=light_gray)

    # Get CFIHOS records if available
    cfihos_records = []
    if "cfihos" in files:
        cfihos_records = files["cfihos"]["records"]

    wb = __import__("openpyxl").Workbook()
    wb.remove(wb.active)  # remove default sheet

    columns = ["#", "Status", "Id", "Name", "CFIHOS Code", "CFIHOS Name", "CFIHOS Match"]
    col_widths = [6, 28, 18, 40, 16, 40, 14]

    for m in masters:
        if m not in files:
            continue
        original = files[m]["records"]
        label = FILE_TYPES.get(m, m)
        ws = wb.create_sheet(title=f"{label} Enriched")

        # ── Row 1: KBR-AMCDE logo bar ──
        nc = len(columns)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
        logo_c = ws.cell(row=1, column=1, value="  KBR-AMCDE  |  RDL Data Harmonizer v3.0")
        logo_c.font = Font(name="Calibri", size=14, bold=True, color=kbr_accent)
        logo_c.fill = PatternFill("solid", fgColor=kbr_dark)
        logo_c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30

        # ── Row 2: Title bar ──
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
        title_cell = ws.cell(row=2, column=1, value=f"  {label} — Enriched Master")
        title_cell.font = title_font
        title_cell.fill = title_fill
        title_cell.alignment = title_align
        ws.row_dimensions[2].height = 36

        # ── Row 3: Subtitle ──
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=nc)
        orig_cnt = len(original)
        add_cnt = len(selected_additions)
        sub_cell = ws.cell(row=3, column=1,
                           value=f"  Original: {orig_cnt} classes  |  Suggested Additions: {add_cnt}  |  New Total: {orig_cnt + add_cnt}  |  Threshold: {threshold}%")
        sub_cell.font = Font(name="Calibri", size=10, color=white)
        sub_cell.fill = sub_fill
        sub_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[3].height = 24

        # ── Row 4: Column headers ──
        for ci, col_name in enumerate(columns, 1):
            cell = ws.cell(row=4, column=ci, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        ws.row_dimensions[4].height = 28
        ws.auto_filter.ref = f"A4:{get_column_letter(nc)}4"

        # ── Data rows ──
        row_num = 5
        counter = 1

        # Original records
        for r in original:
            cfihos_code = r.get("cfihos_ref", "")
            cfihos_name = ""
            cfihos_match = ""
            if cfihos_code and cfihos_code != "nan" and cfihos_records:
                for cr in cfihos_records:
                    if cr["id"] == cfihos_code:
                        cfihos_name = cr["name"]
                        cfihos_match = "Exact"
                        break
                if not cfihos_name:
                    cfihos_match = "Ref Only"
            if not cfihos_code or cfihos_code == "nan":
                cfihos_code, cfihos_name, cfihos_match = _lookup_cfihos(r["name"], cfihos_records, threshold)

            values = [counter, "Original", r["id"], r["name"],
                      cfihos_code if cfihos_code != "nan" else "", cfihos_name, cfihos_match]
            is_alt = (counter % 2 == 0)
            for ci, val in enumerate(values, 1):
                cell = ws.cell(row=row_num, column=ci, value=val)
                cell.font = data_font
                cell.alignment = center_align if ci in (1, 7) else data_align
                cell.border = thin_border
                # Status column styling
                if ci == 2:
                    cell.fill = status_orig_fill
                    cell.font = status_orig_font
                    cell.alignment = center_align
                # Match column styling
                elif ci == 7 and val:
                    if val == "Exact":
                        cell.fill = exact_fill
                        cell.font = exact_font
                    else:
                        cell.fill = partial_fill
                        cell.font = partial_font
                else:
                    cell.fill = alt_fill if is_alt else orig_fill

            counter += 1
            row_num += 1

        # Separator row
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=len(columns))
        sep_cell = ws.cell(row=row_num, column=1, value="  SUGGESTED ADDITIONS")
        sep_cell.font = Font(name="Calibri", size=10, bold=True, color=white)
        sep_cell.fill = PatternFill("solid", fgColor="2563EB")
        sep_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row_num].height = 26
        row_num += 1

        # Suggested additions
        for rec in selected_additions:
            cfihos_code, cfihos_name, cfihos_match = _lookup_cfihos(rec["name"], cfihos_records, threshold)
            src_label = FILE_TYPES.get(rec["source"], rec["source"])
            values = [counter, f"Suggested from {src_label}", f"NEW-{rec['id']}", rec["name"],
                      cfihos_code, cfihos_name, cfihos_match]
            is_alt = (counter % 2 == 0)
            for ci, val in enumerate(values, 1):
                cell = ws.cell(row=row_num, column=ci, value=val)
                cell.font = data_font
                cell.alignment = center_align if ci in (1, 7) else data_align
                cell.border = thin_border
                if ci == 2:
                    cell.fill = status_sugg_fill
                    cell.font = status_sugg_font
                    cell.alignment = center_align
                elif ci == 7 and val:
                    if val == "Exact":
                        cell.fill = exact_fill
                        cell.font = exact_font
                    else:
                        cell.fill = partial_fill
                        cell.font = partial_font
                else:
                    cell.fill = sugg_fill if not is_alt else PatternFill("solid", fgColor="C7D2FE")

            counter += 1
            row_num += 1

        # Set column widths
        for ci, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        # Freeze panes below header
        ws.freeze_panes = "A5"

    # ── Summary sheet ──
    ws_sum = wb.create_sheet(title="Enrichment Summary")
    orig_total = sum(len(files[m]["records"]) for m in masters if m in files)

    # Logo row
    ws_sum.merge_cells("A1:B1")
    lc = ws_sum.cell(row=1, column=1, value="  KBR-AMCDE  |  RDL Data Harmonizer v3.0")
    lc.font = Font(name="Calibri", size=14, bold=True, color=kbr_accent)
    lc.fill = PatternFill("solid", fgColor=kbr_dark)
    lc.alignment = Alignment(horizontal="left", vertical="center")
    ws_sum.row_dimensions[1].height = 30
    # Title
    ws_sum.merge_cells("A2:B2")
    t = ws_sum.cell(row=2, column=1, value="  Enrichment Summary")
    t.font = title_font
    t.fill = title_fill
    t.alignment = title_align
    ws_sum.row_dimensions[2].height = 36

    summary_data = [
        ("Original Master", " + ".join(FILE_TYPES.get(m, m) for m in masters)),
        ("Original Classes", orig_total),
        ("Suggested Additions", len(selected_additions)),
        ("New Total", orig_total + len(selected_additions)),
        ("Match Threshold", f"{threshold}%"),
    ]
    for src in sorted(set(r["source"] for r in selected_additions)):
        cnt = sum(1 for r in selected_additions if r["source"] == src)
        summary_data.append((f"Added from {FILE_TYPES.get(src, src)}", cnt))

    # Headers
    for ci, h in enumerate(["Metric", "Value"], 1):
        cell = ws_sum.cell(row=4, column=ci, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    for ri, (metric, value) in enumerate(summary_data, 5):
        mc = ws_sum.cell(row=ri, column=1, value=metric)
        mc.font = Font(name="Calibri", size=10, bold=True, color=navy)
        mc.fill = alt_fill if ri % 2 == 0 else orig_fill
        mc.border = thin_border
        vc = ws_sum.cell(row=ri, column=2, value=value)
        vc.font = data_font
        vc.fill = alt_fill if ri % 2 == 0 else orig_fill
        vc.border = thin_border

    ws_sum.column_dimensions["A"].width = 30
    ws_sum.column_dimensions["B"].width = 35

    # ── Attributes sheet — ISM Functional Class Attributes by class ──
    aramco_attrs = st.session_state.get("aramco_attrs")
    if aramco_attrs is not None and not aramco_attrs.empty and "Class_Id" in aramco_attrs.columns:
        ws_attr = wb.create_sheet(title="Class Attributes")

        # Logo + title
        attr_cols_list = ["Class_Id", "Name", "Description", "Presence", "Size",
                          "Discipline", "UomClassId", "UomRequire", "ValidationRule", "Group_Id"]
        attr_cols_list = [c for c in attr_cols_list if c in aramco_attrs.columns]
        nc = len(attr_cols_list)
        ws_attr.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(nc, 2))
        lc = ws_attr.cell(row=1, column=1, value="  KBR-AMCDE  |  RDL Data Harmonizer v3.0")
        lc.font = Font(name="Calibri", size=14, bold=True, color=kbr_accent)
        lc.fill = PatternFill("solid", fgColor=kbr_dark)
        lc.alignment = Alignment(horizontal="left", vertical="center")
        ws_attr.row_dimensions[1].height = 30

        ws_attr.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(nc, 2))
        tc = ws_attr.cell(row=2, column=1, value="  ISM Functional Class Attributes — Organized by Class")
        tc.font = title_font
        tc.fill = title_fill
        tc.alignment = Alignment(horizontal="left", vertical="center")
        ws_attr.row_dimensions[2].height = 36

        # Build all original + suggested class names with their IDs
        all_class_ids = {}
        for m_key in masters:
            if m_key in files:
                for r in files[m_key]["records"]:
                    all_class_ids.setdefault(r["name"], set()).add(str(r["id"]).strip())
        for rec in selected_additions:
            all_class_ids.setdefault(rec["name"], set()).add(str(rec["id"]).strip())

        row_num = 3
        class_separator_fill = PatternFill("solid", fgColor=kbr_primary)
        class_separator_font = Font(name="Calibri", size=11, bold=True, color=white)

        for class_name in sorted(all_class_ids.keys()):
            ids = all_class_ids[class_name]
            mask = aramco_attrs["Class_Id"].astype(str).str.strip().isin(ids)
            subset = aramco_attrs[mask]
            if subset.empty:
                continue

            # Class separator row
            ws_attr.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=nc)
            sep = ws_attr.cell(row=row_num, column=1,
                               value=f"  {class_name}  ({', '.join(sorted(ids))})  —  {len(subset)} attributes")
            sep.font = class_separator_font
            sep.fill = class_separator_fill
            sep.alignment = Alignment(horizontal="left", vertical="center")
            ws_attr.row_dimensions[row_num].height = 26
            row_num += 1

            # Column headers for this class
            for ci, col_name in enumerate(attr_cols_list, 1):
                cell = ws_attr.cell(row=row_num, column=ci, value=col_name)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border
            ws_attr.row_dimensions[row_num].height = 24
            row_num += 1

            # Data rows
            for _, attr_row in subset.iterrows():
                is_alt = (row_num % 2 == 0)
                for ci, col_name in enumerate(attr_cols_list, 1):
                    val = attr_row.get(col_name, "")
                    if pd.isna(val):
                        val = ""
                    cell = ws_attr.cell(row=row_num, column=ci, value=str(val))
                    cell.font = data_font
                    cell.alignment = data_align
                    cell.border = thin_border
                    # Color presence column
                    if col_name == "Presence":
                        sv = str(val).strip().lower()
                        if sv in ("mandatory", "required", "m"):
                            cell.fill = PatternFill("solid", fgColor=green_bg)
                            cell.font = Font(name="Calibri", size=10, bold=True, color=green_fg)
                        elif sv in ("optional", "o"):
                            cell.fill = PatternFill("solid", fgColor=blue_bg)
                            cell.font = Font(name="Calibri", size=10, color=blue_fg)
                        else:
                            cell.fill = alt_fill if is_alt else orig_fill
                    else:
                        cell.fill = alt_fill if is_alt else orig_fill
                row_num += 1

            # Blank separator row
            row_num += 1

        # Auto-fit attribute columns
        attr_widths = {"Class_Id": 14, "Name": 35, "Description": 40, "Presence": 12,
                       "Size": 8, "Discipline": 15, "UomClassId": 14, "UomRequire": 12,
                       "ValidationRule": 20, "Group_Id": 14}
        for ci, col_name in enumerate(attr_cols_list, 1):
            ws_attr.column_dimensions[get_column_letter(ci)].width = attr_widths.get(col_name, 18)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────
# Graphviz connection diagram
# ──────────────────────────────────────────────────────────────────────
def build_connection_graph(entry):
    masters = st.session_state.masters
    lines = ["digraph G {"]
    lines.append('  rankdir=LR;')
    lines.append('  node [shape=box, style="filled,rounded", fontname="Arial", fontsize=11];')
    lines.append('  edge [fontname="Arial", fontsize=9];')

    # Master reference node (single box for the combined master)
    master_parts = []
    for m in masters:
        me = entry["master_entries"].get(m)
        if me:
            master_parts.append(f'{FILE_TYPES[m]}: {me["name"]} ({me["id"]})')
    master_label = "MASTER REFERENCE\\n" + "\\n".join(
        p.replace('"', '\\"') for p in master_parts
    )
    lines.append(f'  master [label="{master_label}", fillcolor="#dbeafe", color="#2563eb", penwidth=2, fontcolor="#1e3a5f"];')

    node_id = 0
    # Non-master matches
    for src, match in entry["matches"].items():
        color = SOURCE_COLORS.get(src, "#6b7280")
        name_esc = match["name"].replace('"', '\\"')
        label = f'{FILE_TYPES.get(src, src)}\\n{name_esc}\\nID: {match["id"]}'
        nid = f"n{node_id}"
        lines.append(f'  {nid} [label="{label}", fillcolor="{color}20", color="{color}", fontcolor="#1f2937"];')
        lines.append(f'  master -> {nid} [label="{match["score"]}%", color="{color}", fontcolor="{color}"];')
        node_id += 1

    # Gap nodes
    for g in entry["gaps"]:
        label = f'{FILE_TYPES.get(g, g)}\\nGAP'
        nid = f"g{node_id}"
        lines.append(f'  {nid} [label="{label}", fillcolor="#fee2e2", color="#ef4444", fontcolor="#ef4444", style="filled,rounded,dashed"];')
        lines.append(f'  master -> {nid} [style=dashed, color="#ef4444", label="NO MATCH", fontcolor="#ef4444"];')
        node_id += 1

    lines.append("}")
    return "\n".join(lines)


def build_overview_graph(classes, max_rows=20):
    masters = st.session_state.masters
    all_sources = set()
    for c in classes:
        all_sources.update(c["matches"].keys())
        all_sources.update(c["gaps"])
    all_sources = sorted(all_sources)

    lines = ["digraph G {"]
    lines.append('  rankdir=LR;')
    lines.append('  node [shape=box, style="filled,rounded", fontname="Arial", fontsize=10];')
    lines.append('  edge [fontname="Arial", fontsize=8];')
    lines.append('  nodesep=0.3; ranksep=1.2;')

    for idx, c in enumerate(classes[:max_rows]):
        # Master node (combined)
        master_label_parts = []
        for m in masters:
            me = c["master_entries"].get(m)
            if me:
                master_label_parts.append(me["name"][:25].replace('"', '\\"'))
        master_label = "\\n".join(master_label_parts) if master_label_parts else "?"
        mid = f"r{idx}_master"
        lines.append(f'  {mid} [label="{master_label}", fillcolor="#dbeafe", color="#2563eb"];')

        for src in all_sources:
            match = c["matches"].get(src)
            color = SOURCE_COLORS.get(src, "#6b7280")
            nid = f"r{idx}_{src}"
            if match:
                name_esc = match["name"][:25].replace('"', '\\"')
                lines.append(f'  {nid} [label="{name_esc}", fillcolor="{color}15", color="{color}"];')
                lines.append(f'  {mid} -> {nid} [label="{match["score"]}%", color="{color}"];')
            else:
                lines.append(f'  {nid} [label="GAP", fillcolor="#fee2e2", color="#ef4444", style="filled,rounded,dashed"];')
                lines.append(f'  {mid} -> {nid} [style=dashed, color="#ef4444"];')

    lines.append("}")
    return "\n".join(lines)


def _build_discipline_summary_excel(
    all_disciplines, discipline_classes, disc_aramco_data, disc_ltc_data,
    discipline_class_ids, aramco_attrs, ltc_attrs,
):
    """Build a professional Excel workbook with one sheet per discipline."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import openpyxl

    # Style constants
    kbr_primary = "003087"
    kbr_dark = "001D54"
    kbr_accent = "00A3E0"
    navy = "1B3A5C"
    white = "FFFFFF"
    light_gray = "F8FAFC"
    border_gray = "E2E8F0"
    green_bg = "D1FAE5"
    green_fg = "065F46"
    blue_bg = "DBEAFE"
    blue_fg = "1E40AF"
    orange_bg = "FFEDD5"
    orange_fg = "9A3412"

    thin_border = Border(
        left=Side(style="thin", color=border_gray),
        right=Side(style="thin", color=border_gray),
        top=Side(style="thin", color=border_gray),
        bottom=Side(style="thin", color=border_gray),
    )
    title_font = Font(name="Calibri", size=16, bold=True, color=white)
    title_fill = PatternFill("solid", fgColor=kbr_primary)
    sub_font = Font(name="Calibri", size=10, color=white)
    sub_fill = PatternFill("solid", fgColor=kbr_dark)
    header_font = Font(name="Calibri", size=10, bold=True, color=white)
    header_fill = PatternFill("solid", fgColor=navy)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="Calibri", size=10, color=navy)
    data_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    alt_fill = PatternFill("solid", fgColor=light_gray)
    white_fill = PatternFill("solid", fgColor=white)

    logo_font = Font(name="Calibri", size=14, bold=True, color=kbr_accent)
    logo_fill = PatternFill("solid", fgColor=kbr_dark)

    class_sep_fill = PatternFill("solid", fgColor=kbr_primary)
    class_sep_font = Font(name="Calibri", size=11, bold=True, color=white)
    source_sep_fill = PatternFill("solid", fgColor="2563EB")
    source_sep_font = Font(name="Calibri", size=10, bold=True, color=white)

    def _style_title_rows(ws, title_text, subtitle_text, num_cols):
        nc = max(num_cols, 2)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
        lc = ws.cell(row=1, column=1, value="  KBR-AMCDE  |  RDL Data Harmonizer v3.0")
        lc.font = logo_font
        lc.fill = logo_fill
        lc.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
        t = ws.cell(row=2, column=1, value=f"  {title_text}")
        t.font = title_font
        t.fill = title_fill
        t.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 36
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=nc)
        s = ws.cell(row=3, column=1, value=f"  {subtitle_text}")
        s.font = sub_font
        s.fill = sub_fill
        s.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[3].height = 24

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Overview sheet ──
    ws_ov = wb.create_sheet("Overview")
    _style_title_rows(ws_ov, "Discipline Attribute Summary — Expert Work Document",
                      f"{len(all_disciplines)} disciplines  |  Generated by KBR RDL Data Harmonizer", 7)

    ov_headers = ["Discipline", "Equipment Classes", "Aramco Attr Rows",
                  "Aramco Unique Attrs", "LTC Attr Rows", "LTC Unique Attrs", "Total Attributes"]
    for ci, h in enumerate(ov_headers, 1):
        cell = ws_ov.cell(row=4, column=ci, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws_ov.row_dimensions[4].height = 28
    ws_ov.auto_filter.ref = f"A4:{get_column_letter(len(ov_headers))}4"
    ws_ov.freeze_panes = "A5"

    for ri, disc in enumerate(all_disciplines, 5):
        is_alt = (ri % 2 == 0)
        n_cls = len(discipline_classes.get(disc, []))
        n_ar = len(disc_aramco_data[disc]) if disc in disc_aramco_data else 0
        u_ar = len(disc_aramco_data[disc]["Name"].dropna().unique()) if disc in disc_aramco_data and "Name" in disc_aramco_data[disc].columns else 0
        n_lt = len(disc_ltc_data[disc]) if disc in disc_ltc_data else 0
        u_lt = len(disc_ltc_data[disc]["Name"].dropna().unique()) if disc in disc_ltc_data and "Name" in disc_ltc_data[disc].columns else 0
        vals = [disc, n_cls, n_ar, u_ar, n_lt, u_lt, n_ar + n_lt]
        for ci, val in enumerate(vals, 1):
            cell = ws_ov.cell(row=ri, column=ci, value=val)
            cell.font = data_font
            cell.alignment = center_align if ci > 1 else data_align
            cell.border = thin_border
            cell.fill = alt_fill if is_alt else white_fill

    ov_widths = [22, 18, 18, 18, 15, 15, 16]
    for ci, w in enumerate(ov_widths, 1):
        ws_ov.column_dimensions[get_column_letter(ci)].width = w

    # ── Per-discipline sheets ──
    aramco_cols = ["Class_Id", "Class_Desc", "Attribute_Id", "Name", "Attribute_Desc",
                   "Presence", "Size", "Discipline", "UomClassId", "UomRequire",
                   "ValidationRule", "Group_Id"]
    ltc_cols = ["Class_Id", "Name", "Description", "Presence", "Size",
                "Discipline", "UomClassId", "UomRequire", "ValidationRule",
                "ValidationType", "MaxOccur", "Aspect"]

    for disc in all_disciplines:
        # Sanitize sheet name (max 31 chars, no special chars)
        sheet_name = disc[:28].replace("/", "-").replace("\\", "-").replace("*", "").replace("?", "").replace("[", "").replace("]", "")
        ws = wb.create_sheet(sheet_name)

        disc_cls = discipline_classes.get(disc, [])
        n_aramco = len(disc_aramco_data[disc]) if disc in disc_aramco_data else 0
        n_ltc = len(disc_ltc_data[disc]) if disc in disc_ltc_data else 0

        _style_title_rows(
            ws, f"Discipline: {disc}",
            f"{len(disc_cls)} equipment classes  |  Aramco: {n_aramco} attrs  |  LTC: {n_ltc} attrs",
            10,
        )

        row_num = 4

        # ── Equipment classes in this discipline ──
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=10)
        sep = ws.cell(row=row_num, column=1, value=f"  EQUIPMENT CLASSES ({len(disc_cls)})")
        sep.font = class_sep_font
        sep.fill = class_sep_fill
        sep.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row_num].height = 26
        row_num += 1

        eq_headers = ["#", "Class Name", "Master Sources", "Matched Sources", "Gaps"]
        for ci, h in enumerate(eq_headers, 1):
            cell = ws.cell(row=row_num, column=ci, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        ws.row_dimensions[row_num].height = 24
        row_num += 1

        for idx, c_entry in enumerate(sorted(disc_cls, key=lambda x: x["canonical_name"]), 1):
            is_alt = (idx % 2 == 0)
            master_str = ", ".join(
                f"{FILE_TYPES.get(m, m)}: {c_entry['master_entries'][m]['name']}"
                for m in st.session_state.masters if m in c_entry["master_entries"]
            )
            match_str = ", ".join(
                f"{FILE_TYPES.get(s, s)} ({match['score']}%)"
                for s, match in c_entry["matches"].items()
            )
            gap_str = ", ".join(FILE_TYPES.get(g, g) for g in c_entry["gaps"]) or "No Gaps"
            vals = [idx, c_entry["canonical_name"], master_str, match_str, gap_str]
            for ci, val in enumerate(vals, 1):
                cell = ws.cell(row=row_num, column=ci, value=val)
                cell.font = data_font
                cell.alignment = center_align if ci == 1 else data_align
                cell.border = thin_border
                if ci == 5 and val != "No Gaps":
                    cell.fill = PatternFill("solid", fgColor="FEE2E2")
                    cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B")
                elif ci == 5:
                    cell.fill = PatternFill("solid", fgColor=green_bg)
                    cell.font = Font(name="Calibri", size=10, bold=True, color=green_fg)
                else:
                    cell.fill = alt_fill if is_alt else white_fill
            row_num += 1

        row_num += 1  # blank row

        # ── Aramco attributes for this discipline ──
        if disc in disc_aramco_data:
            subset = disc_aramco_data[disc]
            avail_cols = [c for c in aramco_cols if c in subset.columns]
            nc = len(avail_cols)

            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=max(nc, 5))
            sep = ws.cell(row=row_num, column=1, value=f"  ARAMCO FUNCTIONAL ATTRIBUTES ({len(subset)} rows)")
            sep.font = source_sep_font
            sep.fill = source_sep_fill
            sep.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_num].height = 26
            row_num += 1

            for ci, h in enumerate(avail_cols, 1):
                cell = ws.cell(row=row_num, column=ci, value=h)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border
            ws.row_dimensions[row_num].height = 24
            row_num += 1

            for _, attr_row in subset.iterrows():
                is_alt = (row_num % 2 == 0)
                for ci, col_name in enumerate(avail_cols, 1):
                    val = attr_row.get(col_name, "")
                    if pd.isna(val):
                        val = ""
                    cell = ws.cell(row=row_num, column=ci, value=str(val))
                    cell.font = data_font
                    cell.alignment = data_align
                    cell.border = thin_border
                    if col_name == "Presence":
                        sv = str(val).strip().lower()
                        if sv in ("mandatory", "required", "m"):
                            cell.fill = PatternFill("solid", fgColor=green_bg)
                            cell.font = Font(name="Calibri", size=10, bold=True, color=green_fg)
                        elif sv in ("optional", "o"):
                            cell.fill = PatternFill("solid", fgColor=blue_bg)
                            cell.font = Font(name="Calibri", size=10, color=blue_fg)
                        else:
                            cell.fill = alt_fill if is_alt else white_fill
                    else:
                        cell.fill = alt_fill if is_alt else white_fill
                row_num += 1

            row_num += 1  # blank row

        # ── LTC attributes for this discipline ──
        if disc in disc_ltc_data:
            subset = disc_ltc_data[disc]
            avail_cols = [c for c in ltc_cols if c in subset.columns]
            nc = len(avail_cols)

            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=max(nc, 5))
            sep = ws.cell(row=row_num, column=1, value=f"  LTC PHYSICAL ATTRIBUTES ({len(subset)} rows)")
            sep.font = source_sep_font
            sep.fill = PatternFill("solid", fgColor="7C3AED")
            sep.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_num].height = 26
            row_num += 1

            for ci, h in enumerate(avail_cols, 1):
                cell = ws.cell(row=row_num, column=ci, value=h)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border
            ws.row_dimensions[row_num].height = 24
            row_num += 1

            for _, attr_row in subset.iterrows():
                is_alt = (row_num % 2 == 0)
                for ci, col_name in enumerate(avail_cols, 1):
                    val = attr_row.get(col_name, "")
                    if pd.isna(val):
                        val = ""
                    cell = ws.cell(row=row_num, column=ci, value=str(val))
                    cell.font = data_font
                    cell.alignment = data_align
                    cell.border = thin_border
                    if col_name == "Presence":
                        sv = str(val).strip().lower()
                        if sv in ("mandatory", "required", "m"):
                            cell.fill = PatternFill("solid", fgColor=green_bg)
                            cell.font = Font(name="Calibri", size=10, bold=True, color=green_fg)
                        elif sv in ("optional", "o"):
                            cell.fill = PatternFill("solid", fgColor=blue_bg)
                            cell.font = Font(name="Calibri", size=10, color=blue_fg)
                        else:
                            cell.fill = alt_fill if is_alt else white_fill
                    else:
                        cell.fill = alt_fill if is_alt else white_fill
                row_num += 1

        # ── Attribute gap summary between sources ──
        if disc in disc_aramco_data and disc in disc_ltc_data:
            row_num += 1
            a_names = set(disc_aramco_data[disc]["Name"].dropna().str.strip().str.lower()) if "Name" in disc_aramco_data[disc].columns else set()
            l_names = set(disc_ltc_data[disc]["Name"].dropna().str.strip().str.lower()) if "Name" in disc_ltc_data[disc].columns else set()
            common = a_names & l_names
            only_a = a_names - l_names
            only_l = l_names - a_names

            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=5)
            sep = ws.cell(row=row_num, column=1, value=f"  ATTRIBUTE GAP ANALYSIS")
            sep.font = Font(name="Calibri", size=10, bold=True, color=white)
            sep.fill = PatternFill("solid", fgColor=orange_fg)
            sep.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_num].height = 26
            row_num += 1

            gap_summary = [
                ("Common Attributes", len(common)),
                ("Only in Aramco", len(only_a)),
                ("Only in LTC", len(only_l)),
            ]
            for metric, value in gap_summary:
                mc = ws.cell(row=row_num, column=1, value=metric)
                mc.font = Font(name="Calibri", size=10, bold=True, color=navy)
                mc.border = thin_border
                mc.fill = white_fill
                vc = ws.cell(row=row_num, column=2, value=value)
                vc.font = data_font
                vc.border = thin_border
                vc.alignment = center_align
                vc.fill = white_fill
                row_num += 1

            if only_a:
                row_num += 1
                ws.cell(row=row_num, column=1, value="Attributes Only in Aramco:").font = Font(name="Calibri", size=10, bold=True, color=orange_fg)
                row_num += 1
                for name in sorted(only_a):
                    ws.cell(row=row_num, column=1, value=name).font = data_font
                    row_num += 1

            if only_l:
                row_num += 1
                ws.cell(row=row_num, column=1, value="Attributes Only in LTC:").font = Font(name="Calibri", size=10, bold=True, color=blue_fg)
                row_num += 1
                for name in sorted(only_l):
                    ws.cell(row=row_num, column=1, value=name).font = data_font
                    row_num += 1

        # Auto-fit columns
        col_widths = {1: 18, 2: 35, 3: 40, 4: 35, 5: 30, 6: 15, 7: 12, 8: 15, 9: 14, 10: 20}
        for ci, w in col_widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _build_attribute_analysis_excel(all_disc, aa_by_disc, la_by_disc):
    """Build a standalone attribute analysis Excel — one sheet per discipline, no harmonization needed."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import openpyxl

    kbr_primary = "003087"
    kbr_dark = "001D54"
    kbr_accent = "00A3E0"
    navy = "1B3A5C"
    white = "FFFFFF"
    light_gray = "F8FAFC"
    border_gray = "E2E8F0"
    green_bg = "D1FAE5"
    green_fg = "065F46"
    blue_bg = "DBEAFE"
    blue_fg = "1E40AF"
    orange_bg = "FFEDD5"
    orange_fg = "9A3412"

    thin_border = Border(
        left=Side(style="thin", color=border_gray),
        right=Side(style="thin", color=border_gray),
        top=Side(style="thin", color=border_gray),
        bottom=Side(style="thin", color=border_gray),
    )
    header_font = Font(name="Calibri", size=10, bold=True, color=white)
    header_fill = PatternFill("solid", fgColor=navy)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="Calibri", size=10, color=navy)
    data_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    alt_fill = PatternFill("solid", fgColor=light_gray)
    white_fill = PatternFill("solid", fgColor=white)
    logo_font = Font(name="Calibri", size=14, bold=True, color=kbr_accent)
    logo_fill = PatternFill("solid", fgColor=kbr_dark)
    title_font = Font(name="Calibri", size=16, bold=True, color=white)
    title_fill = PatternFill("solid", fgColor=kbr_primary)
    sub_font = Font(name="Calibri", size=10, color=white)
    sub_fill = PatternFill("solid", fgColor=kbr_dark)
    section_fill = PatternFill("solid", fgColor="2563EB")
    section_font = Font(name="Calibri", size=10, bold=True, color=white)

    def _title_rows(ws, title, subtitle, nc):
        nc = max(nc, 2)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
        lc = ws.cell(row=1, column=1, value="  KBR-AMCDE  |  RDL Data Harmonizer v3.0")
        lc.font = logo_font; lc.fill = logo_fill
        lc.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
        t = ws.cell(row=2, column=1, value=f"  {title}")
        t.font = title_font; t.fill = title_fill
        t.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 36
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=nc)
        s = ws.cell(row=3, column=1, value=f"  {subtitle}")
        s.font = sub_font; s.fill = sub_fill
        s.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[3].height = 24

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Overview sheet ──
    ws_ov = wb.create_sheet("Overview")
    total_a = sum(len(g) for g in aa_by_disc.values())
    total_l = sum(len(g) for g in la_by_disc.values())
    _title_rows(ws_ov, "Attribute Analysis by Discipline",
                f"{len(all_disc)} disciplines  |  Aramco: {total_a} attributes  |  LTC: {total_l} attributes", 12)

    ov_headers = ["Discipline", "Aramco Rows", "Aramco Unique", "Aramco Mandatory", "Aramco Optional",
                  "LTC Rows", "LTC Unique", "LTC Mandatory", "LTC Optional",
                  "Common Attrs", "Only Aramco", "Only LTC"]
    for ci, h in enumerate(ov_headers, 1):
        cell = ws_ov.cell(row=4, column=ci, value=h)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = header_align; cell.border = thin_border
    ws_ov.row_dimensions[4].height = 28
    ws_ov.auto_filter.ref = f"A4:{get_column_letter(len(ov_headers))}4"
    ws_ov.freeze_panes = "A5"

    for ri, disc in enumerate(all_disc, 5):
        is_alt = (ri % 2 == 0)
        aa_sub = aa_by_disc.get(disc)
        la_sub = la_by_disc.get(disc)
        n_aa = len(aa_sub) if aa_sub is not None else 0
        n_la = len(la_sub) if la_sub is not None else 0
        u_aa = len(aa_sub["Name"].dropna().unique()) if aa_sub is not None and "Name" in aa_sub.columns else 0
        u_la = len(la_sub["Name"].dropna().unique()) if la_sub is not None and "Name" in la_sub.columns else 0
        mand_aa = opt_aa = mand_la = opt_la = 0
        if aa_sub is not None and "Presence" in aa_sub.columns:
            pv = aa_sub["Presence"].fillna("").astype(str).str.strip().str.lower()
            mand_aa = int(pv.isin(["mandatory", "required", "m"]).sum())
            opt_aa = int(pv.isin(["optional", "o"]).sum())
        if la_sub is not None and "Presence" in la_sub.columns:
            pv = la_sub["Presence"].fillna("").astype(str).str.strip().str.lower()
            mand_la = int(pv.isin(["mandatory", "required", "m"]).sum())
            opt_la = int(pv.isin(["optional", "o"]).sum())
        a_names = set(aa_sub["Name"].dropna().str.strip().str.lower()) if aa_sub is not None and "Name" in aa_sub.columns else set()
        l_names = set(la_sub["Name"].dropna().str.strip().str.lower()) if la_sub is not None and "Name" in la_sub.columns else set()
        vals = [disc, n_aa, u_aa, mand_aa, opt_aa, n_la, u_la, mand_la, opt_la,
                len(a_names & l_names), len(a_names - l_names), len(l_names - a_names)]
        for ci, val in enumerate(vals, 1):
            cell = ws_ov.cell(row=ri, column=ci, value=val)
            cell.font = data_font
            cell.alignment = center_align if ci > 1 else data_align
            cell.border = thin_border
            cell.fill = alt_fill if is_alt else white_fill

    ov_widths = [22, 12, 13, 16, 14, 10, 11, 14, 12, 13, 12, 10]
    for ci, w in enumerate(ov_widths, 1):
        ws_ov.column_dimensions[get_column_letter(ci)].width = w

    # ── Per-discipline sheets ──
    aramco_cols = ["Class_Id", "Class_Desc", "Attribute_Id", "Name", "Attribute_Desc",
                   "Presence", "Size", "UomClassId", "UomRequire", "ValidationRule", "Group_Id"]
    ltc_cols = ["Class_Id", "Name", "Description", "Presence", "Size",
                "UomClassId", "UomRequire", "ValidationRule", "ValidationType", "MaxOccur", "Aspect"]

    for disc in all_disc:
        sheet_name = disc[:28].replace("/", "-").replace("\\", "-").replace("*", "").replace("?", "").replace("[", "").replace("]", "")
        # Avoid duplicate sheet names
        if sheet_name in [s.title for s in wb.worksheets]:
            sheet_name = sheet_name[:25] + "..."
        ws = wb.create_sheet(sheet_name)

        aa_sub = aa_by_disc.get(disc)
        la_sub = la_by_disc.get(disc)
        n_aa = len(aa_sub) if aa_sub is not None else 0
        n_la = len(la_sub) if la_sub is not None else 0

        _title_rows(ws, f"Discipline: {disc}",
                    f"Aramco: {n_aa} attributes  |  LTC: {n_la} attributes", 11)

        row_num = 4

        # ── Aramco attributes ──
        if aa_sub is not None and not aa_sub.empty:
            avail = [c for c in aramco_cols if c in aa_sub.columns]
            nc = len(avail)

            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=max(nc, 5))
            sep = ws.cell(row=row_num, column=1, value=f"  ARAMCO FUNCTIONAL ATTRIBUTES ({n_aa} rows)")
            sep.font = section_font; sep.fill = section_fill
            sep.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_num].height = 26
            row_num += 1

            for ci, h in enumerate(avail, 1):
                cell = ws.cell(row=row_num, column=ci, value=h)
                cell.font = header_font; cell.fill = header_fill
                cell.alignment = header_align; cell.border = thin_border
            ws.row_dimensions[row_num].height = 24
            row_num += 1

            for _, attr_row in aa_sub.iterrows():
                is_alt = (row_num % 2 == 0)
                for ci, col_name in enumerate(avail, 1):
                    val = attr_row.get(col_name, "")
                    if pd.isna(val):
                        val = ""
                    cell = ws.cell(row=row_num, column=ci, value=str(val))
                    cell.font = data_font; cell.alignment = data_align; cell.border = thin_border
                    if col_name == "Presence":
                        sv = str(val).strip().lower()
                        if sv in ("mandatory", "required", "m"):
                            cell.fill = PatternFill("solid", fgColor=green_bg)
                            cell.font = Font(name="Calibri", size=10, bold=True, color=green_fg)
                        elif sv in ("optional", "o"):
                            cell.fill = PatternFill("solid", fgColor=blue_bg)
                            cell.font = Font(name="Calibri", size=10, color=blue_fg)
                        else:
                            cell.fill = alt_fill if is_alt else white_fill
                    else:
                        cell.fill = alt_fill if is_alt else white_fill
                row_num += 1

            # Presence summary row
            row_num += 1
            if "Presence" in aa_sub.columns:
                pv = aa_sub["Presence"].fillna("N/A").astype(str).str.strip().value_counts()
                ws.cell(row=row_num, column=1, value="Presence Summary:").font = Font(name="Calibri", size=10, bold=True, color=navy)
                for pi, (pk, pcount) in enumerate(pv.items(), 2):
                    ws.cell(row=row_num, column=pi, value=f"{pk}: {pcount}").font = data_font
                row_num += 1

            row_num += 1

        # ── LTC attributes ──
        if la_sub is not None and not la_sub.empty:
            avail = [c for c in ltc_cols if c in la_sub.columns]
            nc = len(avail)

            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=max(nc, 5))
            sep = ws.cell(row=row_num, column=1, value=f"  LTC PHYSICAL ATTRIBUTES ({n_la} rows)")
            sep.font = section_font
            sep.fill = PatternFill("solid", fgColor="7C3AED")
            sep.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_num].height = 26
            row_num += 1

            for ci, h in enumerate(avail, 1):
                cell = ws.cell(row=row_num, column=ci, value=h)
                cell.font = header_font; cell.fill = header_fill
                cell.alignment = header_align; cell.border = thin_border
            ws.row_dimensions[row_num].height = 24
            row_num += 1

            for _, attr_row in la_sub.iterrows():
                is_alt = (row_num % 2 == 0)
                for ci, col_name in enumerate(avail, 1):
                    val = attr_row.get(col_name, "")
                    if pd.isna(val):
                        val = ""
                    cell = ws.cell(row=row_num, column=ci, value=str(val))
                    cell.font = data_font; cell.alignment = data_align; cell.border = thin_border
                    if col_name == "Presence":
                        sv = str(val).strip().lower()
                        if sv in ("mandatory", "required", "m"):
                            cell.fill = PatternFill("solid", fgColor=green_bg)
                            cell.font = Font(name="Calibri", size=10, bold=True, color=green_fg)
                        elif sv in ("optional", "o"):
                            cell.fill = PatternFill("solid", fgColor=blue_bg)
                            cell.font = Font(name="Calibri", size=10, color=blue_fg)
                        else:
                            cell.fill = alt_fill if is_alt else white_fill
                    else:
                        cell.fill = alt_fill if is_alt else white_fill
                row_num += 1

            row_num += 1
            if "Presence" in la_sub.columns:
                pv = la_sub["Presence"].fillna("N/A").astype(str).str.strip().value_counts()
                ws.cell(row=row_num, column=1, value="Presence Summary:").font = Font(name="Calibri", size=10, bold=True, color=navy)
                for pi, (pk, pcount) in enumerate(pv.items(), 2):
                    ws.cell(row=row_num, column=pi, value=f"{pk}: {pcount}").font = data_font
                row_num += 1

            row_num += 1

        # ── Attribute gap analysis ──
        a_names = set(aa_sub["Name"].dropna().str.strip().str.lower()) if aa_sub is not None and "Name" in aa_sub.columns else set()
        l_names = set(la_sub["Name"].dropna().str.strip().str.lower()) if la_sub is not None and "Name" in la_sub.columns else set()
        if a_names or l_names:
            common = a_names & l_names
            only_a = a_names - l_names
            only_l = l_names - a_names

            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=5)
            sep = ws.cell(row=row_num, column=1, value="  ATTRIBUTE GAP ANALYSIS")
            sep.font = Font(name="Calibri", size=10, bold=True, color=white)
            sep.fill = PatternFill("solid", fgColor=orange_fg)
            sep.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row_num].height = 26
            row_num += 1

            for metric, value in [("Common Attributes", len(common)), ("Only in Aramco", len(only_a)), ("Only in LTC", len(only_l))]:
                mc = ws.cell(row=row_num, column=1, value=metric)
                mc.font = Font(name="Calibri", size=10, bold=True, color=navy)
                mc.border = thin_border; mc.fill = white_fill
                vc = ws.cell(row=row_num, column=2, value=value)
                vc.font = data_font; vc.border = thin_border
                vc.alignment = center_align; vc.fill = white_fill
                row_num += 1

            if only_a:
                row_num += 1
                ws.cell(row=row_num, column=1, value="Attributes Only in Aramco:").font = Font(name="Calibri", size=10, bold=True, color=orange_fg)
                row_num += 1
                for name in sorted(only_a):
                    ws.cell(row=row_num, column=1, value=name).font = data_font
                    row_num += 1

            if only_l:
                row_num += 1
                ws.cell(row=row_num, column=1, value="Attributes Only in LTC:").font = Font(name="Calibri", size=10, bold=True, color=blue_fg)
                row_num += 1
                for name in sorted(only_l):
                    ws.cell(row=row_num, column=1, value=name).font = data_font
                    row_num += 1

        # Column widths
        col_widths = {1: 18, 2: 35, 3: 18, 4: 35, 5: 35, 6: 12, 7: 10, 8: 14, 9: 12, 10: 20, 11: 14}
        for ci, w in col_widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════

# ── Header ──
st.markdown("""
<div style="background: linear-gradient(135deg, #003087 0%, #001d54 50%, #1B3A5C 100%);
            padding: 1.8rem 2.5rem; border-radius: 16px; margin-bottom: 1.5rem;
            box-shadow: 0 8px 32px rgba(0,48,135,0.3); position: relative; overflow: hidden;">
    <div style="position: absolute; top: -20px; right: -20px; width: 200px; height: 200px;
                background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 70%);
                border-radius: 50%;"></div>
    <div style="position: absolute; bottom: -30px; left: 30%; width: 150px; height: 150px;
                background: radial-gradient(circle, rgba(0,163,224,0.15) 0%, transparent 70%);
                border-radius: 50%;"></div>
    <div style="display: flex; align-items: center; justify-content: space-between; position: relative; z-index: 1;">
        <div>
            <div style="display: flex; align-items: center; gap: 1rem; margin-bottom: 0.4rem;">
                <div style="background: rgba(255,255,255,0.12); padding: 0.5rem 1rem; border-radius: 10px;
                            backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.15);">
                    <span style="font-size: 1.1rem; font-weight: 800; color: #00A3E0;
                                 letter-spacing: 0.15em; font-family: 'Inter', sans-serif;">KBR</span><span
                          style="font-size: 0.7rem; font-weight: 600; color: rgba(255,255,255,0.7);
                                 letter-spacing: 0.08em; margin-left: 0.15rem;">-AMCDE</span>
                </div>
                <h1 style="color: white; margin: 0; font-size: 1.9rem; font-weight: 800;
                           letter-spacing: -0.02em; font-family: 'Inter', sans-serif;">
                    RDL Data Harmonizer
                </h1>
            </div>
            <p style="color: rgba(0,163,224,0.85); margin: 0; font-size: 0.88rem; font-weight: 400;
                      letter-spacing: 0.02em; padding-left: 7.5rem;">
                Equipment Class Harmonization &amp; Gap Analysis Platform
            </p>
        </div>
        <div style="text-align: right;">
            <div style="background: rgba(0,163,224,0.15); padding: 0.35rem 1rem; border-radius: 20px;
                        backdrop-filter: blur(10px); margin-bottom: 0.4rem;
                        border: 1px solid rgba(0,163,224,0.25);">
                <span style="color: #00A3E0; font-size: 0.75rem; font-weight: 600;
                             letter-spacing: 0.08em; text-transform: uppercase;">Version 3.0</span>
            </div>
            <p style="color: rgba(255,255,255,0.5); font-size: 0.72rem; margin: 0;">
                Built by <span style="color: #00A3E0;">KBR AMCDE</span> Team
            </p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

tab_upload, tab_dashboard, tab_gaps, tab_attrs, tab_visual, tab_search, tab_batch, tab_logs = st.tabs([
    "Upload & Configure",
    "Dashboard",
    "Gap Analysis",
    "Attributes",
    "Connection Map",
    "Search",
    "Batch Process",
    "Logs",
])

# ══════════════════════════════════════════════════════════════════════
# TAB: Upload & Configure
# ══════════════════════════════════════════════════════════════════════
with tab_upload:
    # Step 1
    st.markdown("""
    <div class="section-card">
        <h3>&#9312; Choose Master File(s)</h3>
        <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
            Select <strong>1 or 2</strong> file types as the master reference. They are treated as
            <strong>one combined object</strong>. All other uploaded files are compared against this master.
        </p>
    </div>
    """, unsafe_allow_html=True)

    cfg_left, cfg_right = st.columns([2, 1])
    with cfg_left:
        master_options = list(FILE_TYPES.keys())
        selected_masters = st.multiselect(
            "Master file(s)",
            options=master_options,
            default=st.session_state.masters or [],
            format_func=lambda k: FILE_TYPES[k],
            max_selections=2,
        )
        if selected_masters != st.session_state.masters:
            st.session_state.masters = selected_masters

        if len(selected_masters) == 0:
            st.warning("Select at least one master.")
        elif len(selected_masters) == 1:
            st.info(f"**Master**: {FILE_TYPES[selected_masters[0]]}")
        else:
            st.info(
                f"**Combined Master**: {FILE_TYPES[selected_masters[0]]} + "
                f"{FILE_TYPES[selected_masters[1]]}"
            )
    with cfg_right:
        threshold = st.slider(
            "Match Threshold (%)", min_value=50, max_value=100,
            value=st.session_state.match_threshold,
            help="Minimum fuzzy match score to count as a match. Below this = GAP.",
        )
        if threshold != st.session_state.match_threshold:
            st.session_state.match_threshold = threshold

    st.markdown("")

    # Step 2
    st.markdown("""
    <div class="section-card">
        <h3>&#9313; Upload Data Files</h3>
        <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
            Upload Excel files for each data source. Supported formats: <code>.xlsx</code>, <code>.xls</code>, <code>.csv</code>
        </p>
    </div>
    """, unsafe_allow_html=True)

    cols = st.columns(3)
    for i, (key, label) in enumerate(FILE_TYPES.items()):
        with cols[i % 3]:
            is_master = key in st.session_state.masters
            badge = " (MASTER)" if is_master else ""
            color = SOURCE_COLORS.get(key, "#6b7280")
            st.markdown(f'<p style="font-weight:600;font-size:0.85rem;color:{color};margin-bottom:0.2rem;">{label}{badge}</p>', unsafe_allow_html=True)
            uploaded = st.file_uploader(f"Upload {label}", type=["xlsx", "xls", "csv"], key=f"upload_{key}", label_visibility="collapsed")
            if uploaded is not None:
                if key not in st.session_state.files or st.session_state.files[key]["filename"] != uploaded.name:
                    try:
                        records = READERS[key](uploaded)
                        st.session_state.files[key] = {"filename": uploaded.name, "records": records}
                        add_log(f"Uploaded {key}: {uploaded.name} ({len(records)} records)")
                        st.success(f"{len(records)} records loaded")
                        if key == "aramco":
                            uploaded.seek(0)
                            attr_df = read_aramco_attributes(uploaded)
                            if not attr_df.empty:
                                st.session_state.aramco_attrs = attr_df
                                add_log(f"Aramco attributes loaded: {len(attr_df)} rows")
                        elif key == "ltc":
                            uploaded.seek(0)
                            attr_df = read_ltc_attributes(uploaded)
                            if not attr_df.empty:
                                st.session_state.ltc_attrs = attr_df
                                add_log(f"LTC attributes loaded: {len(attr_df)} rows")
                    except Exception as e:
                        st.error(f"Error: {e}")
                else:
                    st.success(f"{len(st.session_state.files[key]['records'])} records loaded")
            if key in st.session_state.files:
                st.caption(f"_{st.session_state.files[key]['filename']}_")

    st.markdown("")

    # Step 3
    st.markdown("""
    <div class="section-card">
        <h3>&#9314; Run Harmonization</h3>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Run Harmonization", type="primary", use_container_width=True):
            if not st.session_state.masters:
                st.error("Select at least one master.")
            elif not all(m in st.session_state.files for m in st.session_state.masters):
                st.error("Upload all master files first.")
            else:
                with st.spinner("Running harmonization engine..."):
                    st.session_state.classes = run_harmonization()
                st.rerun()
    with col2:
        if st.button("Load Demo Data", use_container_width=True):
            load_demo_data()
            if not st.session_state.masters:
                st.session_state.masters = ["aramco"]
            st.session_state.classes = run_harmonization()
            st.rerun()
    with col3:
        if st.session_state.classes:
            st.download_button(
                "Export Full Report", data=build_excel_bytes(st.session_state.classes),
                file_name="harmonization_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # Source status cards
    if st.session_state.files:
        st.markdown("")
        file_keys = list(st.session_state.files.keys())
        cards_html = '<div class="kpi-row">'
        for k in file_keys:
            v = st.session_state.files[k]
            color = SOURCE_COLORS.get(k, "#6b7280")
            tag = "MASTER" if k in st.session_state.masters else ""
            badge_html = f'<span class="stat-pill primary" style="font-size:0.65rem;">{tag}</span> ' if tag else ""
            cards_html += f'''
            <div class="kpi-card" style="border-top: 3px solid {color};">
                <div>{badge_html}</div>
                <div class="kpi-value" style="color: {color};">{len(v["records"])}</div>
                <div class="kpi-label">{FILE_TYPES[k]}</div>
                <div class="kpi-sub">{v["filename"]}</div>
            </div>'''
        cards_html += '</div>'
        st.markdown(cards_html, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# TAB: Dashboard
# ══════════════════════════════════════════════════════════════════════
with tab_dashboard:
    classes = st.session_state.classes
    if not classes:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-icon">&#128202;</div>
            <h3>No Results Yet</h3>
            <p>Go to <strong>Upload &amp; Configure</strong> tab to load data and run harmonization.</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        masters = st.session_state.masters
        total = len(classes)

        # Collect non-master sources
        non_master_keys = set()
        for c in classes:
            non_master_keys.update(c["matches"].keys())
            non_master_keys.update(c["gaps"])
        non_master_keys = sorted(non_master_keys)

        # Per-source stats
        source_stats = {}
        for src in non_master_keys:
            matched = sum(1 for c in classes if src in c["matches"])
            gaps = total - matched
            source_stats[src] = {"matched": matched, "gaps": gaps}

        no_gaps_total = sum(1 for c in classes if not c["gaps"])
        has_gaps_total = total - no_gaps_total

        master_label = " + ".join(FILE_TYPES.get(m, m) for m in masters)
        compared_label = ", ".join(FILE_TYPES.get(s, s) for s in non_master_keys)

        # ── Context bar ──
        st.markdown(f"""
        <div style="background: linear-gradient(90deg, #f8fafc, #f1f5f9); border: 1px solid #e2e8f0;
                    border-radius: 10px; padding: 0.75rem 1.5rem; margin-bottom: 1rem;
                    display: flex; gap: 2rem; flex-wrap: wrap; align-items: center;">
            <span style="font-size: 0.82rem; color: #64748b;">
                <strong style="color: #1e293b;">Master:</strong> {master_label}
            </span>
            <span style="font-size: 0.82rem; color: #64748b;">
                <strong style="color: #1e293b;">Compared Against:</strong> {compared_label}
            </span>
            <span style="font-size: 0.82rem; color: #64748b;">
                <strong style="color: #1e293b;">Threshold:</strong> {st.session_state.match_threshold}%
            </span>
        </div>
        """, unsafe_allow_html=True)

        # ── KPI cards row ──
        overall_pct = int(round(no_gaps_total / total * 100)) if total else 0
        kpi_html = '<div class="kpi-row">'
        kpi_html += f'''
        <div class="kpi-card" style="border-top: 3px solid var(--kbr-navy);">
            <div class="kpi-value" style="color: var(--kbr-navy);">{total}</div>
            <div class="kpi-label">Total Classes</div>
        </div>
        <div class="kpi-card" style="border-top: 3px solid var(--kbr-green);">
            <div class="kpi-value" style="color: var(--kbr-green);">{no_gaps_total}</div>
            <div class="kpi-label">Fully Matched</div>
            <div class="kpi-sub">{overall_pct}% coverage</div>
        </div>
        <div class="kpi-card" style="border-top: 3px solid var(--kbr-red);">
            <div class="kpi-value" style="color: var(--kbr-red);">{has_gaps_total}</div>
            <div class="kpi-label">Has Gaps</div>
        </div>'''
        for src in non_master_keys:
            s = source_stats[src]
            pct = int(round(s["matched"] / total * 100))
            color = SOURCE_COLORS.get(src, "#6b7280")
            kpi_html += f'''
            <div class="kpi-card" style="border-top: 3px solid {color};">
                <div class="kpi-value" style="color: {color};">{pct}%</div>
                <div class="kpi-label">{FILE_TYPES.get(src, src)}</div>
                <div class="kpi-sub">{s["matched"]}/{total} matched &middot; {s["gaps"]} gaps</div>
            </div>'''
        kpi_html += '</div>'
        st.markdown(kpi_html, unsafe_allow_html=True)

        # ── Charts row ──
        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:
            st.markdown('<div class="chart-container"><h4>Coverage Overview</h4></div>', unsafe_allow_html=True)
            donut_data = pd.DataFrame({
                "Status": ["Fully Matched", "Has Gaps"],
                "Count": [no_gaps_total, has_gaps_total],
                "Color": ["#059669", "#b91c1c"],
            })
            donut = alt.Chart(donut_data).mark_arc(innerRadius=60, outerRadius=100, cornerRadius=4).encode(
                theta=alt.Theta("Count:Q"),
                color=alt.Color("Status:N", scale=alt.Scale(
                    domain=["Fully Matched", "Has Gaps"],
                    range=["#059669", "#b91c1c"],
                ), legend=alt.Legend(orient="bottom", title=None, labelFontSize=12, symbolSize=120)),
                tooltip=["Status:N", "Count:Q"],
            ).properties(height=280)
            center_text = alt.Chart(pd.DataFrame({"text": [f"{overall_pct}%"]})).mark_text(
                size=28, fontWeight="bold", color="#1e293b", font="Inter",
            ).encode(text="text:N")
            st.altair_chart(donut + center_text, use_container_width=True)

        with chart_col2:
            st.markdown('<div class="chart-container"><h4>Match Rate by Source</h4></div>', unsafe_allow_html=True)
            bar_data = pd.DataFrame([
                {
                    "Source": FILE_TYPES.get(src, src),
                    "Matched": int(round(source_stats[src]["matched"] / total * 100)),
                    "Gaps": int(round(source_stats[src]["gaps"] / total * 100)),
                }
                for src in non_master_keys
            ])
            bar_melted = bar_data.melt(id_vars=["Source"], var_name="Type", value_name="Percent")
            bar_chart = alt.Chart(bar_melted).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
                x=alt.X("Source:N", axis=alt.Axis(labelAngle=0, labelFontSize=12, title=None)),
                y=alt.Y("Percent:Q", axis=alt.Axis(title="Percentage", labelFontSize=11), scale=alt.Scale(domain=[0, 100])),
                color=alt.Color("Type:N", scale=alt.Scale(
                    domain=["Matched", "Gaps"],
                    range=["#059669", "#ef4444"],
                ), legend=alt.Legend(orient="bottom", title=None, labelFontSize=12, symbolSize=120)),
                tooltip=["Source:N", "Type:N", "Percent:Q"],
            ).properties(height=260)
            st.altair_chart(bar_chart, use_container_width=True)

        # ── Score distribution chart ──
        all_scores = []
        for c in classes:
            for src, m in c["matches"].items():
                all_scores.append({"Score": m["score"], "Source": FILE_TYPES.get(src, src)})
        if all_scores:
            st.markdown('<div class="chart-container"><h4>Match Score Distribution</h4></div>', unsafe_allow_html=True)
            score_df = pd.DataFrame(all_scores)
            hist = alt.Chart(score_df).mark_bar(
                cornerRadiusTopLeft=4, cornerRadiusTopRight=4, opacity=0.85,
            ).encode(
                x=alt.X("Score:Q", bin=alt.Bin(step=5), axis=alt.Axis(title="Match Score %", labelFontSize=11)),
                y=alt.Y("count()", axis=alt.Axis(title="Count", labelFontSize=11)),
                color=alt.Color("Source:N", scale=alt.Scale(
                    domain=[FILE_TYPES.get(s, s) for s in non_master_keys],
                    range=[SOURCE_COLORS.get(s, "#6b7280") for s in non_master_keys],
                ), legend=alt.Legend(orient="bottom", title=None, labelFontSize=11, symbolSize=120)),
                tooltip=["Source:N", "count()"],
            ).properties(height=220)
            st.altair_chart(hist, use_container_width=True)

        # ── Match quality breakdown ──
        exact_cnt = sum(1 for c in classes for m in c["matches"].values() if m["score"] >= 95)
        strong_cnt = sum(1 for c in classes for m in c["matches"].values() if 85 <= m["score"] < 95)
        partial_cnt = sum(1 for c in classes for m in c["matches"].values() if 75 <= m["score"] < 85)
        weak_cnt = sum(1 for c in classes for m in c["matches"].values() if m["score"] < 75)
        total_matches = exact_cnt + strong_cnt + partial_cnt + weak_cnt

        st.markdown(f"""
        <div class="section-card">
            <h3>Match Quality Breakdown</h3>
            <div style="display:flex; gap:1rem; flex-wrap:wrap; margin-top:0.75rem;">
                <div style="flex:1; min-width: 140px; text-align:center; padding:0.75rem; background:#d1fae5; border-radius:10px;">
                    <div style="font-size:1.5rem; font-weight:800; color:#065f46;">{exact_cnt}</div>
                    <div style="font-size:0.78rem; font-weight:600; color:#065f46;">Exact (95%+)</div>
                </div>
                <div style="flex:1; min-width: 140px; text-align:center; padding:0.75rem; background:#dbeafe; border-radius:10px;">
                    <div style="font-size:1.5rem; font-weight:800; color:#1e40af;">{strong_cnt}</div>
                    <div style="font-size:0.78rem; font-weight:600; color:#1e40af;">Strong (85-94%)</div>
                </div>
                <div style="flex:1; min-width: 140px; text-align:center; padding:0.75rem; background:#ffedd5; border-radius:10px;">
                    <div style="font-size:1.5rem; font-weight:800; color:#9a3412;">{partial_cnt}</div>
                    <div style="font-size:0.78rem; font-weight:600; color:#9a3412;">Partial (75-84%)</div>
                </div>
                <div style="flex:1; min-width: 140px; text-align:center; padding:0.75rem; background:#fee2e2; border-radius:10px;">
                    <div style="font-size:1.5rem; font-weight:800; color:#991b1b;">{weak_cnt}</div>
                    <div style="font-size:0.78rem; font-weight:600; color:#991b1b;">Weak (&lt;75%)</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        # ── Class detail browser ──
        st.markdown("### Class Detail Browser")
        filter_opt = st.radio("Filter", ["All", "No Gaps only", "Has Gaps only"], horizontal=True, key="dash_filter")

        for c in classes:
            has_gaps = bool(c["gaps"])
            if filter_opt == "No Gaps only" and has_gaps:
                continue
            if filter_opt == "Has Gaps only" and not has_gaps:
                continue

            if has_gaps:
                gap_label = ":red[GAP IN: " + ", ".join(FILE_TYPES.get(g, g) for g in c["gaps"]) + "]"
            else:
                gap_label = ":green[No Gaps]"

            with st.expander(f"**{c['index']}. {c['canonical_name']}** — {gap_label}"):
                detail_left, detail_right = st.columns(2)
                with detail_left:
                    st.markdown("**Master Reference:**")
                    for m in masters:
                        me = c["master_entries"].get(m)
                        if me:
                            st.markdown(f"- {FILE_TYPES[m]}: **{me['name']}** (`{me['id']}`)")
                        else:
                            st.markdown(f"- {FILE_TYPES[m]}: :red[not in this master]")
                    if len(masters) == 2 and c.get("cross_score") is not None:
                        st.caption(f"Masters cross-match: {c['cross_score']}%")

                with detail_right:
                    st.markdown("**Comparison Results:**")
                    for src in non_master_keys:
                        match = c["matches"].get(src)
                        label = FILE_TYPES.get(src, src)
                        if match:
                            status, _ = classify_match(match["score"])
                            color = "green" if match["score"] >= 90 else "orange" if match["score"] >= 75 else "red"
                            st.markdown(f"- :{color}[**{label}**]: {match['name']} (`{match['id']}`) — {match['score']}% {status}")
                        else:
                            st.markdown(f"- :red[**{label}**: GAP — not found]")

        # Table
        st.divider()
        st.markdown("### Full Data Table")
        st.dataframe(build_export_df(classes), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════
# TAB: Gap Analysis
# ══════════════════════════════════════════════════════════════════════
with tab_gaps:
    classes = st.session_state.classes
    if not classes:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-icon">&#128269;</div>
            <h3>No Gap Analysis Available</h3>
            <p>Run harmonization first to see gap analysis results.</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        masters = st.session_state.masters
        files = st.session_state.files
        all_gaps = []
        for c in classes:
            for g in c["gaps"]:
                found_in = [FILE_TYPES.get(m, m) for m in masters if m in c["master_entries"]]
                found_in += [FILE_TYPES.get(s, s) for s in c["matches"]]
                all_gaps.append({
                    "Equipment Class": c["canonical_name"],
                    "Found In": ", ".join(found_in),
                    "Gap In": FILE_TYPES.get(g, g),
                    "Action": f"Add to {FILE_TYPES.get(g, g)} or confirm exclusion",
                })

        st.markdown(f"""
        <div class="section-card">
            <h3>Gap Analysis &mdash; {len(all_gaps)} gaps identified</h3>
            <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
                Equipment classes found in the master reference but missing from one or more sources.
            </p>
        </div>
        """, unsafe_allow_html=True)

        if not all_gaps:
            st.success("No gaps! All sources are fully matched across all datasets.")
        else:
            gap_by_source = {}
            for g in all_gaps:
                gap_by_source[g["Gap In"]] = gap_by_source.get(g["Gap In"], 0) + 1

            # Gap chart + metrics
            gap_chart_col, gap_metric_col = st.columns([2, 1])
            with gap_chart_col:
                gap_chart_data = pd.DataFrame([
                    {"Source": src, "Gaps": cnt}
                    for src, cnt in sorted(gap_by_source.items(), key=lambda x: -x[1])
                ])
                gap_bar = alt.Chart(gap_chart_data).mark_bar(
                    cornerRadiusTopLeft=6, cornerRadiusTopRight=6, color="#ef4444",
                ).encode(
                    x=alt.X("Source:N", axis=alt.Axis(labelAngle=0, labelFontSize=12, title=None), sort="-y"),
                    y=alt.Y("Gaps:Q", axis=alt.Axis(title="Gap Count", labelFontSize=11)),
                    tooltip=["Source:N", "Gaps:Q"],
                    color=alt.Color("Source:N", legend=None, scale=alt.Scale(
                        domain=list(gap_by_source.keys()),
                        range=["#ef4444", "#f97316", "#eab308", "#8b5cf6"][:len(gap_by_source)],
                    )),
                ).properties(height=220)
                st.altair_chart(gap_bar, use_container_width=True)

            with gap_metric_col:
                for src, cnt in sorted(gap_by_source.items(), key=lambda x: -x[1]):
                    st.metric(src, f"{cnt} gaps")

            st.divider()
            source_filter = st.selectbox("Filter by source", ["All"] + sorted(gap_by_source.keys()))
            filtered = all_gaps if source_filter == "All" else [g for g in all_gaps if g["Gap In"] == source_filter]
            st.dataframe(pd.DataFrame(filtered), use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("### Action Items")
            for i, g in enumerate(filtered):
                st.markdown(f"{i+1}. **{g['Equipment Class']}** — add to :red[**{g['Gap In']}**] (found in: {g['Found In']})")

        # ── Enrichment Suggestions (reverse gaps) ──────────────────
        st.divider()
        reverse_gaps = st.session_state.get("reverse_gaps", {})
        master_label = " + ".join(FILE_TYPES.get(m, m) for m in masters)

        if reverse_gaps:
            total_extra = sum(len(recs) for recs in reverse_gaps.values())
            master_total = sum(len(files[m]["records"]) for m in masters if m in files)

            st.markdown(f"""
            <div class="section-card" style="border-left: 4px solid var(--kbr-blue);">
                <h3 style="color: var(--kbr-blue);">Enrichment Suggestions &mdash; {total_extra} extra classes found</h3>
                <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
                    These classes exist in non-master sources but <strong>not in your master</strong> ({master_label}).
                    Master currently has <strong>{master_total}</strong> classes.
                    Review the suggestions below and download an enriched master file.
                </p>
            </div>
            """, unsafe_allow_html=True)

            # Metrics per source
            rev_cols = st.columns(len(reverse_gaps))
            for i, (src, recs) in enumerate(sorted(reverse_gaps.items())):
                with rev_cols[i]:
                    st.metric(f"Extra in {FILE_TYPES.get(src, src)}", len(recs))

            st.divider()

            # Build a combined preview list with CFIHOS matching
            cfihos_records = files["cfihos"]["records"] if "cfihos" in files else []
            threshold = st.session_state.match_threshold
            preview_rows = []
            for src, recs in sorted(reverse_gaps.items()):
                for rec in recs:
                    cfihos_code, cfihos_name, cfihos_match = _lookup_cfihos(rec["name"], cfihos_records, threshold)
                    row = {
                        "Source": FILE_TYPES.get(src, src),
                        "ID": rec["id"],
                        "Name": rec["name"],
                        "CFIHOS Code": cfihos_code,
                        "CFIHOS Name": cfihos_name,
                        "CFIHOS Match": cfihos_match,
                        "_source_key": src,
                    }
                    if "discipline" in rec and rec["discipline"]:
                        row["Discipline"] = rec["discipline"]
                    preview_rows.append(row)

            preview_df = pd.DataFrame(preview_rows)
            display_cols = [c for c in ["Source", "ID", "Name", "CFIHOS Code", "CFIHOS Name", "CFIHOS Match", "Discipline"] if c in preview_df.columns]

            # Source filter for enrichment
            enrich_sources = sorted(reverse_gaps.keys())
            enrich_filter = st.selectbox(
                "Filter by source",
                ["All"] + [FILE_TYPES.get(s, s) for s in enrich_sources],
                key="enrich_filter",
            )

            if enrich_filter == "All":
                filtered_preview = preview_df
            else:
                filtered_preview = preview_df[preview_df["Source"] == enrich_filter]

            st.markdown("**Preview of suggested additions:**")
            st.dataframe(
                filtered_preview[display_cols].reset_index(drop=True),
                use_container_width=True,
                hide_index=True,
            )
            st.caption(f"Showing {len(filtered_preview)} of {len(preview_df)} suggestions")

            # Select all / download
            st.divider()
            st.markdown(
                f"**Download enriched master** — your original {master_label} "
                f"({master_total} classes) + the {len(filtered_preview)} suggested additions below."
            )

            # Build selected additions from filtered preview
            selected_additions = []
            for _, row in filtered_preview.iterrows():
                selected_additions.append({
                    "id": row["ID"],
                    "name": row["Name"],
                    "source": row["_source_key"],
                })

            if selected_additions:
                enriched_bytes = build_enriched_master_excel(selected_additions)
                st.download_button(
                    f"Download Enriched {master_label} ({master_total + len(selected_additions)} classes)",
                    data=enriched_bytes,
                    file_name="enriched_master.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
                st.caption(
                    f"Original: {master_total} classes → "
                    f"Enriched: {master_total + len(selected_additions)} classes "
                    f"(+{len(selected_additions)} from suggestions)"
                )
        else:
            st.subheader("Enrichment Suggestions")
            st.success(
                f"No extra classes found. All non-master records matched something in {master_label}."
            )


# ══════════════════════════════════════════════════════════════════════
# TAB: Attributes
# ══════════════════════════════════════════════════════════════════════
with tab_attrs:
    aramco_attrs = st.session_state.aramco_attrs
    ltc_attrs = st.session_state.ltc_attrs
    classes = st.session_state.classes

    if aramco_attrs is None and ltc_attrs is None:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-icon">&#128196;</div>
            <h3>No Attribute Data</h3>
            <p>Upload <strong>Aramco</strong> (needs &lsquo;ISM Functional Class Attributes&rsquo; sheet)
            or <strong>LTC</strong> (needs &lsquo;ISM Physical Class Attributes&rsquo; sheet) to explore attributes.</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="section-card">
            <h3>Attributes Explorer</h3>
            <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
                Browse, compare, and analyze equipment class attributes across data sources.
            </p>
        </div>
        """, unsafe_allow_html=True)

        # Search filter: if a search match is active, filter attributes to that class
        # Search filter: if a search match is active, show attributes for that class directly
        search_class = st.session_state.search_matched_class
        if search_class:
            # Collect all Class_Ids for the matched class
            _sf_ids = set()
            for m in st.session_state.masters:
                me = search_class["master_entries"].get(m)
                if me:
                    _sf_ids.add(str(me["id"]).strip())
            for src, match in search_class["matches"].items():
                _sf_ids.add(str(match["id"]).strip())
            search_filter_ids = list(_sf_ids)
            _fc1, _fc2 = st.columns([5, 1])
            with _fc1:
                st.info(f"Filtered by search result: **{search_class['canonical_name']}**")
            with _fc2:
                if st.button("Clear Filter", key="clear_filter_attrs", type="secondary", use_container_width=True):
                    st.session_state.search_matched_class = None
                    st.rerun()

            # Filter attribute dataframes
            filtered_aramco = None
            filtered_ltc = None
            if aramco_attrs is not None and not aramco_attrs.empty and "Class_Id" in aramco_attrs.columns:
                filtered_aramco = aramco_attrs[aramco_attrs["Class_Id"].astype(str).str.strip().isin(search_filter_ids)]
            if ltc_attrs is not None and not ltc_attrs.empty and "Class_Id" in ltc_attrs.columns:
                filtered_ltc = ltc_attrs[ltc_attrs["Class_Id"].astype(str).str.strip().isin(search_filter_ids)]

            # Show counts
            attr_cols = st.columns(2)
            with attr_cols[0]:
                cnt = len(filtered_aramco) if filtered_aramco is not None else 0
                st.metric("Aramco Functional Attributes", cnt)
            with attr_cols[1]:
                cnt = len(filtered_ltc) if filtered_ltc is not None else 0
                st.metric("LTC Physical Attributes", cnt)

            st.divider()
            st.caption(f"Class IDs: {', '.join(search_filter_ids)}")

            # Show Aramco attributes
            st.markdown("#### Aramco Functional Attributes")
            if filtered_aramco is not None and not filtered_aramco.empty:
                display_cols = [c for c in ["Attribute_Id", "Name", "Attribute_Desc", "Class_Desc",
                                             "Description", "Presence", "Size",
                                             "Discipline", "UomClassId", "UomRequire",
                                             "ValidationRule", "Group_Id"] if c in filtered_aramco.columns]
                st.dataframe(filtered_aramco[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                st.caption(f"{len(filtered_aramco)} attributes")
            else:
                st.warning("No Aramco attributes found for this class.")

            # Show LTC attributes
            st.markdown("#### LTC Physical Attributes")
            if filtered_ltc is not None and not filtered_ltc.empty:
                display_cols = [c for c in ["Name", "Description", "Presence", "Size",
                                             "Discipline", "UomClassId", "UomRequire",
                                             "ValidationRule", "ValidationType", "MaxOccur",
                                             "MinOccurs", "Aspect"] if c in filtered_ltc.columns]
                st.dataframe(filtered_ltc[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                st.caption(f"{len(filtered_ltc)} attributes")
            else:
                st.warning("No LTC attributes found for this class.")

            # Attribute comparison between sources
            if filtered_aramco is not None and not filtered_aramco.empty and filtered_ltc is not None and not filtered_ltc.empty:
                st.divider()
                st.markdown("#### Attribute Comparison")
                a_names = set(filtered_aramco["Name"].dropna().str.strip().str.lower()) if "Name" in filtered_aramco.columns else set()
                l_names = set(filtered_ltc["Name"].dropna().str.strip().str.lower()) if "Name" in filtered_ltc.columns else set()
                common = a_names & l_names
                only_aramco = a_names - l_names
                only_ltc = l_names - a_names
                gap_cols = st.columns(3)
                with gap_cols[0]:
                    st.metric("Common", len(common))
                with gap_cols[1]:
                    st.metric("Only in Aramco", len(only_aramco))
                with gap_cols[2]:
                    st.metric("Only in LTC", len(only_ltc))
                if only_aramco:
                    with st.expander(f"Attributes only in Aramco ({len(only_aramco)})"):
                        for name in sorted(only_aramco):
                            st.markdown(f"- {name}")
                if only_ltc:
                    with st.expander(f"Attributes only in LTC ({len(only_ltc)})"):
                        for name in sorted(only_ltc):
                            st.markdown(f"- {name}")

        else:
            # No search filter — show normal view modes
            # Show counts
            attr_cols = st.columns(2)
            with attr_cols[0]:
                cnt = len(aramco_attrs) if aramco_attrs is not None else 0
                st.metric("Aramco Functional Attributes", cnt)
            with attr_cols[1]:
                cnt = len(ltc_attrs) if ltc_attrs is not None else 0
                st.metric("LTC Physical Attributes", cnt)

            st.divider()

            # Build class list for picker
            # If harmonization has been run, use the harmonized classes; otherwise derive from attribute data
            class_options = []
            class_id_map = {}  # display_name -> list of class_ids
            if classes:
                for c in classes:
                    display = c["canonical_name"]
                    ids = set()
                    for m in st.session_state.masters:
                        me = c["master_entries"].get(m)
                        if me:
                            ids.add(str(me["id"]).strip())
                    for src, match in c["matches"].items():
                        ids.add(str(match["id"]).strip())
                    class_options.append(display)
                    class_id_map[display] = list(ids)
            else:
                # Derive from attribute data directly
                seen = set()
                if aramco_attrs is not None and "Class_Id" in aramco_attrs.columns:
                    for cid in aramco_attrs["Class_Id"].dropna().unique():
                        key = str(cid).strip()
                        if key and key not in seen:
                            seen.add(key)
                            desc = ""
                            if "Class_Desc" in aramco_attrs.columns:
                                filtered = aramco_attrs[aramco_attrs["Class_Id"] == cid]
                                if not filtered.empty:
                                    desc = str(filtered.iloc[0].get("Class_Desc", ""))
                            label = f"{desc} ({key})" if desc and desc != "nan" else key
                            class_options.append(label)
                            class_id_map[label] = [key]
                if ltc_attrs is not None and "Class_Id" in ltc_attrs.columns:
                    for cid in ltc_attrs["Class_Id"].dropna().unique():
                        key = str(cid).strip()
                        if key and key not in seen:
                            seen.add(key)
                            class_options.append(key)
                            class_id_map[key] = [key]

            attr_view = st.radio(
                "View mode",
                ["Browse by Class", "Discipline-wise Attributes", "Discipline File Summary", "Attribute Analysis", "Full Aramco Attributes", "Full LTC Attributes", "Attribute Comparison"],
                horizontal=True,
            )

            if attr_view == "Browse by Class":
                if not class_options:
                    st.warning("No class data available. Upload files and/or run harmonization first.")
                else:
                    selected_class = st.selectbox("Select equipment class", class_options)
                    if selected_class:
                        ids = class_id_map.get(selected_class, [])
                        st.caption(f"Looking up Class IDs: {', '.join(ids)}")

                        # Aramco attributes for this class
                        st.markdown("#### Aramco Functional Attributes")
                        if aramco_attrs is not None and "Class_Id" in aramco_attrs.columns:
                            mask = aramco_attrs["Class_Id"].astype(str).str.strip().isin(ids)
                            subset = aramco_attrs[mask]
                            if subset.empty:
                                st.warning("No Aramco attributes found for this class.")
                            else:
                                display_cols = [c for c in ["Attribute_Id", "Name", "Attribute_Desc", "Class_Desc",
                                                             "Description", "Presence", "Size",
                                                             "Discipline", "UomClassId", "UomRequire",
                                                             "ValidationRule", "Group_Id"] if c in subset.columns]
                                st.dataframe(subset[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                                st.caption(f"{len(subset)} attributes")
                        else:
                            st.info("Aramco attribute data not loaded.")

                        # LTC attributes for this class
                        st.markdown("#### LTC Physical Attributes")
                        if ltc_attrs is not None and "Class_Id" in ltc_attrs.columns:
                            mask = ltc_attrs["Class_Id"].astype(str).str.strip().isin(ids)
                            subset = ltc_attrs[mask]
                            if subset.empty:
                                st.warning("No LTC attributes found for this class.")
                            else:
                                display_cols = [c for c in ["Name", "Description", "Presence", "Size",
                                                             "Discipline", "UomClassId", "UomRequire",
                                                             "ValidationRule", "ValidationType", "MaxOccur",
                                                             "MinOccurs", "Aspect"] if c in subset.columns]
                                st.dataframe(subset[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                                st.caption(f"{len(subset)} attributes")
                        else:
                            st.info("LTC attribute data not loaded.")

            elif attr_view == "Discipline-wise Attributes":
                # Collect discipline info from harmonized classes
                # Each class gets a discipline from KBR master or match sources
                discipline_classes = {}  # discipline -> list of class names
                discipline_class_ids = {}  # discipline -> set of class_ids (for attribute lookup)
                if classes:
                    for c in classes:
                        disc = None
                        # Check KBR master entry for discipline
                        for m in st.session_state.masters:
                            me = c["master_entries"].get(m)
                            if me and me.get("discipline"):
                                disc = me["discipline"]
                                break
                        # Check matches for KBR discipline
                        if not disc:
                            for src, match in c["matches"].items():
                                if match.get("discipline"):
                                    disc = match["discipline"]
                                    break
                        if not disc:
                            disc = "Unclassified"
                        discipline_classes.setdefault(disc, []).append(c["canonical_name"])
                        # Collect all Class_Ids for this class entry
                        _ids = discipline_class_ids.setdefault(disc, set())
                        for m in st.session_state.masters:
                            me = c["master_entries"].get(m)
                            if me:
                                _ids.add(str(me["id"]).strip())
                        for src, match in c["matches"].items():
                            _ids.add(str(match["id"]).strip())

                # Build discipline -> attributes mapping via Class_Id linkage
                disc_aramco = {}
                disc_ltc = {}

                for disc, ids in discipline_class_ids.items():
                    if aramco_attrs is not None and not aramco_attrs.empty and "Class_Id" in aramco_attrs.columns:
                        mask = aramco_attrs["Class_Id"].astype(str).str.strip().isin(ids)
                        subset = aramco_attrs[mask]
                        if not subset.empty:
                            disc_aramco[disc] = subset
                    if ltc_attrs is not None and not ltc_attrs.empty and "Class_Id" in ltc_attrs.columns:
                        mask = ltc_attrs["Class_Id"].astype(str).str.strip().isin(ids)
                        subset = ltc_attrs[mask]
                        if not subset.empty:
                            disc_ltc[disc] = subset

                # Merge all discipline names
                all_disciplines = sorted(set(list(discipline_classes.keys()) + list(disc_aramco.keys()) + list(disc_ltc.keys())))

                if not all_disciplines:
                    st.warning("No discipline information found. Upload files with discipline data (KBR FEED or attribute sheets with Discipline column).")
                else:
                    # Summary metrics
                    st.markdown("#### Discipline Summary")
                    summary_data = []
                    for disc in all_disciplines:
                        n_classes = len(discipline_classes.get(disc, []))
                        n_aramco_attrs = len(disc_aramco[disc]) if disc in disc_aramco else 0
                        n_ltc_attrs = len(disc_ltc[disc]) if disc in disc_ltc else 0
                        summary_data.append({
                            "Discipline": disc,
                            "Equipment Classes": n_classes,
                            "Aramco Attributes": n_aramco_attrs,
                            "LTC Attributes": n_ltc_attrs,
                            "Total Attributes": n_aramco_attrs + n_ltc_attrs,
                        })
                    summary_df = pd.DataFrame(summary_data)
                    st.dataframe(summary_df, use_container_width=True, hide_index=True)

                    st.divider()

                    # Discipline selector
                    selected_disc = st.selectbox("Select Discipline", all_disciplines, key="disc_select")
                    if selected_disc:
                        # Show equipment classes in this discipline
                        disc_class_list = discipline_classes.get(selected_disc, [])
                        if disc_class_list:
                            st.markdown(f"**Equipment Classes in {selected_disc}** ({len(disc_class_list)}):")
                            for cn in sorted(disc_class_list):
                                st.markdown(f"- {cn}")
                        else:
                            st.info(f"No harmonized equipment classes tagged under '{selected_disc}'.")

                        st.divider()

                        # Show Aramco attributes for this discipline
                        col_a, col_l = st.columns(2)
                        with col_a:
                            st.markdown(f"**Aramco Attributes — {selected_disc}**")
                            if selected_disc in disc_aramco:
                                subset = disc_aramco[selected_disc]
                                display_cols = [c for c in ["Class_Id", "Class_Desc", "Attribute_Id", "Name",
                                                             "Attribute_Desc", "Presence", "Size",
                                                             "UomClassId", "ValidationRule"] if c in subset.columns]
                                st.dataframe(subset[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                                st.caption(f"{len(subset)} attributes")

                                # Unique attribute names in this discipline
                                if "Name" in subset.columns:
                                    unique_attrs = sorted(subset["Name"].dropna().unique())
                                    with st.expander(f"Unique Attribute Names ({len(unique_attrs)})"):
                                        for a in unique_attrs:
                                            st.markdown(f"- {a}")
                            else:
                                st.info("No Aramco attributes for this discipline.")

                        with col_l:
                            st.markdown(f"**LTC Attributes — {selected_disc}**")
                            if selected_disc in disc_ltc:
                                subset = disc_ltc[selected_disc]
                                display_cols = [c for c in ["Class_Id", "Name", "Description",
                                                             "Presence", "Size", "UomClassId",
                                                             "ValidationRule", "Aspect"] if c in subset.columns]
                                st.dataframe(subset[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                                st.caption(f"{len(subset)} attributes")

                                if "Name" in subset.columns:
                                    unique_attrs = sorted(subset["Name"].dropna().unique())
                                    with st.expander(f"Unique Attribute Names ({len(unique_attrs)})"):
                                        for a in unique_attrs:
                                            st.markdown(f"- {a}")
                            else:
                                st.info("No LTC attributes for this discipline.")

                        # Cross-source attribute gap for this discipline
                        if selected_disc in disc_aramco and selected_disc in disc_ltc:
                            st.divider()
                            st.markdown(f"**Attribute Gap Analysis — {selected_disc}**")
                            a_names = set(disc_aramco[selected_disc]["Name"].dropna().str.strip().str.lower()) if "Name" in disc_aramco[selected_disc].columns else set()
                            l_names = set(disc_ltc[selected_disc]["Name"].dropna().str.strip().str.lower()) if "Name" in disc_ltc[selected_disc].columns else set()
                            common = a_names & l_names
                            only_a = a_names - l_names
                            only_l = l_names - a_names
                            gap_cols = st.columns(3)
                            with gap_cols[0]:
                                st.metric("Common", len(common))
                            with gap_cols[1]:
                                st.metric("Only in Aramco", len(only_a))
                            with gap_cols[2]:
                                st.metric("Only in LTC", len(only_l))
                            if only_a:
                                with st.expander(f"Attributes only in Aramco ({len(only_a)})"):
                                    for n in sorted(only_a):
                                        st.markdown(f"- {n}")
                            if only_l:
                                with st.expander(f"Attributes only in LTC ({len(only_l)})"):
                                    for n in sorted(only_l):
                                        st.markdown(f"- {n}")

            elif attr_view == "Discipline File Summary":
                # Expert work document: each discipline gets its own file summary
                discipline_classes = {}
                discipline_class_ids = {}
                if classes:
                    for c in classes:
                        disc = None
                        for m in st.session_state.masters:
                            me = c["master_entries"].get(m)
                            if me and me.get("discipline"):
                                disc = me["discipline"]
                                break
                        if not disc:
                            for src, match in c["matches"].items():
                                if match.get("discipline"):
                                    disc = match["discipline"]
                                    break
                        if not disc:
                            disc = "Unclassified"
                        discipline_classes.setdefault(disc, []).append(c)
                        _ids = discipline_class_ids.setdefault(disc, set())
                        for m in st.session_state.masters:
                            me = c["master_entries"].get(m)
                            if me:
                                _ids.add(str(me["id"]).strip())
                        for src, match in c["matches"].items():
                            _ids.add(str(match["id"]).strip())

                # Build per-discipline attribute data
                disc_aramco_data = {}
                disc_ltc_data = {}
                for disc, ids in discipline_class_ids.items():
                    if aramco_attrs is not None and not aramco_attrs.empty and "Class_Id" in aramco_attrs.columns:
                        mask = aramco_attrs["Class_Id"].astype(str).str.strip().isin(ids)
                        subset = aramco_attrs[mask]
                        if not subset.empty:
                            disc_aramco_data[disc] = subset
                    if ltc_attrs is not None and not ltc_attrs.empty and "Class_Id" in ltc_attrs.columns:
                        mask = ltc_attrs["Class_Id"].astype(str).str.strip().isin(ids)
                        subset = ltc_attrs[mask]
                        if not subset.empty:
                            disc_ltc_data[disc] = subset

                all_disciplines = sorted(set(list(discipline_classes.keys()) + list(disc_aramco_data.keys()) + list(disc_ltc_data.keys())))

                if not all_disciplines:
                    st.warning("No discipline information found. Upload files with discipline data (KBR FEED) and run harmonization.")
                else:
                    st.markdown("""
                    <div class="section-card">
                        <h3>Discipline File Summary — Expert Work Document</h3>
                        <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
                            Each discipline has its own attribute file summary. Export generates a professional
                            Excel workbook with one sheet per discipline.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)

                    # Export button
                    disc_export_data = _build_discipline_summary_excel(
                        all_disciplines, discipline_classes, disc_aramco_data, disc_ltc_data,
                        discipline_class_ids, aramco_attrs, ltc_attrs,
                    )
                    st.download_button(
                        "Export Discipline Reports (Excel)",
                        data=disc_export_data,
                        file_name="discipline_attribute_summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                    st.divider()

                    # Overview table
                    overview_rows = []
                    for disc in all_disciplines:
                        n_classes = len(discipline_classes.get(disc, []))
                        n_aramco = len(disc_aramco_data[disc]) if disc in disc_aramco_data else 0
                        n_ltc = len(disc_ltc_data[disc]) if disc in disc_ltc_data else 0
                        unique_aramco = len(disc_aramco_data[disc]["Name"].dropna().unique()) if disc in disc_aramco_data and "Name" in disc_aramco_data[disc].columns else 0
                        unique_ltc = len(disc_ltc_data[disc]["Name"].dropna().unique()) if disc in disc_ltc_data and "Name" in disc_ltc_data[disc].columns else 0
                        overview_rows.append({
                            "Discipline": disc,
                            "Equipment Classes": n_classes,
                            "Aramco Attr Rows": n_aramco,
                            "Aramco Unique Attrs": unique_aramco,
                            "LTC Attr Rows": n_ltc,
                            "LTC Unique Attrs": unique_ltc,
                            "Total Attributes": n_aramco + n_ltc,
                        })
                    st.dataframe(pd.DataFrame(overview_rows), use_container_width=True, hide_index=True)

                    st.divider()

                    # Per-discipline expandable sections
                    for disc in all_disciplines:
                        disc_cls = discipline_classes.get(disc, [])
                        n_aramco = len(disc_aramco_data[disc]) if disc in disc_aramco_data else 0
                        n_ltc = len(disc_ltc_data[disc]) if disc in disc_ltc_data else 0
                        with st.expander(f"{disc} — {len(disc_cls)} classes, {n_aramco + n_ltc} attributes"):
                            # Equipment classes list
                            if disc_cls:
                                st.markdown(f"**Equipment Classes ({len(disc_cls)}):**")
                                class_names = sorted([c["canonical_name"] for c in disc_cls])
                                # Show as comma-separated for compactness
                                st.markdown(", ".join(f"`{n}`" for n in class_names))

                            # Aramco attributes summary
                            if disc in disc_aramco_data:
                                st.markdown(f"**Aramco Attributes ({n_aramco} rows)**")
                                subset = disc_aramco_data[disc]
                                display_cols = [c for c in ["Class_Id", "Class_Desc", "Name",
                                                             "Presence", "Size", "UomClassId",
                                                             "ValidationRule"] if c in subset.columns]
                                st.dataframe(subset[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)

                            # LTC attributes summary
                            if disc in disc_ltc_data:
                                st.markdown(f"**LTC Attributes ({n_ltc} rows)**")
                                subset = disc_ltc_data[disc]
                                display_cols = [c for c in ["Class_Id", "Name", "Description",
                                                             "Presence", "Size", "UomClassId",
                                                             "ValidationRule"] if c in subset.columns]
                                st.dataframe(subset[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)

                            # Attribute gap between sources for this discipline
                            if disc in disc_aramco_data and disc in disc_ltc_data:
                                a_names = set(disc_aramco_data[disc]["Name"].dropna().str.strip().str.lower()) if "Name" in disc_aramco_data[disc].columns else set()
                                l_names = set(disc_ltc_data[disc]["Name"].dropna().str.strip().str.lower()) if "Name" in disc_ltc_data[disc].columns else set()
                                common = a_names & l_names
                                only_a = a_names - l_names
                                only_l = l_names - a_names
                                g1, g2, g3 = st.columns(3)
                                with g1:
                                    st.metric("Common Attrs", len(common))
                                with g2:
                                    st.metric("Only Aramco", len(only_a))
                                with g3:
                                    st.metric("Only LTC", len(only_l))

            elif attr_view == "Attribute Analysis":
                # Standalone analysis — only Aramco + LTC attributes, no harmonization needed
                # Groups by Discipline column directly from attribute sheets
                if aramco_attrs is None and ltc_attrs is None:
                    st.warning("Upload **Aramco** or **LTC** files first. This view reads the Discipline column directly from their attribute sheets.")
                else:
                    st.markdown("""
                    <div class="section-card">
                        <h3>Attribute Analysis by Discipline</h3>
                        <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
                            Standalone analysis of Aramco &amp; LTC attributes grouped by discipline.
                            No harmonization required — reads the Discipline column directly from attribute sheets.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)

                    # Group Aramco attributes by Discipline
                    aa_by_disc = {}
                    if aramco_attrs is not None and not aramco_attrs.empty and "Discipline" in aramco_attrs.columns:
                        for disc_val, grp in aramco_attrs.groupby(aramco_attrs["Discipline"].fillna("Unclassified").astype(str).str.strip()):
                            key = disc_val if disc_val and disc_val != "nan" else "Unclassified"
                            aa_by_disc[key] = grp

                    # Group LTC attributes by Discipline
                    la_by_disc = {}
                    if ltc_attrs is not None and not ltc_attrs.empty and "Discipline" in ltc_attrs.columns:
                        for disc_val, grp in ltc_attrs.groupby(ltc_attrs["Discipline"].fillna("Unclassified").astype(str).str.strip()):
                            key = disc_val if disc_val and disc_val != "nan" else "Unclassified"
                            la_by_disc[key] = grp

                    all_disc = sorted(set(list(aa_by_disc.keys()) + list(la_by_disc.keys())))

                    if not all_disc:
                        st.warning("No Discipline column found in the uploaded attribute sheets.")
                    else:
                        # KPI row
                        total_aramco = len(aramco_attrs) if aramco_attrs is not None else 0
                        total_ltc = len(ltc_attrs) if ltc_attrs is not None else 0
                        kpi_c1, kpi_c2, kpi_c3 = st.columns(3)
                        with kpi_c1:
                            st.metric("Disciplines", len(all_disc))
                        with kpi_c2:
                            st.metric("Total Aramco Attributes", total_aramco)
                        with kpi_c3:
                            st.metric("Total LTC Attributes", total_ltc)

                        st.divider()

                        # Overview table
                        ov_rows = []
                        for disc in all_disc:
                            aa_sub = aa_by_disc.get(disc)
                            la_sub = la_by_disc.get(disc)
                            n_aa = len(aa_sub) if aa_sub is not None else 0
                            n_la = len(la_sub) if la_sub is not None else 0
                            u_aa = len(aa_sub["Name"].dropna().unique()) if aa_sub is not None and "Name" in aa_sub.columns else 0
                            u_la = len(la_sub["Name"].dropna().unique()) if la_sub is not None and "Name" in la_sub.columns else 0
                            # Presence breakdown
                            mand_aa = 0
                            opt_aa = 0
                            if aa_sub is not None and "Presence" in aa_sub.columns:
                                p_vals = aa_sub["Presence"].fillna("").astype(str).str.strip().str.lower()
                                mand_aa = int(p_vals.isin(["mandatory", "required", "m"]).sum())
                                opt_aa = int(p_vals.isin(["optional", "o"]).sum())
                            mand_la = 0
                            opt_la = 0
                            if la_sub is not None and "Presence" in la_sub.columns:
                                p_vals = la_sub["Presence"].fillna("").astype(str).str.strip().str.lower()
                                mand_la = int(p_vals.isin(["mandatory", "required", "m"]).sum())
                                opt_la = int(p_vals.isin(["optional", "o"]).sum())
                            # Common / gap
                            a_names = set(aa_sub["Name"].dropna().str.strip().str.lower()) if aa_sub is not None and "Name" in aa_sub.columns else set()
                            l_names = set(la_sub["Name"].dropna().str.strip().str.lower()) if la_sub is not None and "Name" in la_sub.columns else set()
                            common = len(a_names & l_names)
                            only_a = len(a_names - l_names)
                            only_l = len(l_names - a_names)
                            ov_rows.append({
                                "Discipline": disc,
                                "Aramco Rows": n_aa,
                                "Aramco Unique": u_aa,
                                "Aramco Mandatory": mand_aa,
                                "Aramco Optional": opt_aa,
                                "LTC Rows": n_la,
                                "LTC Unique": u_la,
                                "LTC Mandatory": mand_la,
                                "LTC Optional": opt_la,
                                "Common Attrs": common,
                                "Only Aramco": only_a,
                                "Only LTC": only_l,
                            })
                        ov_df = pd.DataFrame(ov_rows)
                        st.markdown("#### Overview by Discipline")
                        st.dataframe(ov_df, use_container_width=True, hide_index=True)

                        # Export button
                        aa_export = _build_attribute_analysis_excel(all_disc, aa_by_disc, la_by_disc)
                        st.download_button(
                            "Export Attribute Analysis (Excel)",
                            data=aa_export,
                            file_name="attribute_analysis_by_discipline.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )

                        st.divider()

                        # Per-discipline detail
                        for disc in all_disc:
                            aa_sub = aa_by_disc.get(disc)
                            la_sub = la_by_disc.get(disc)
                            n_aa = len(aa_sub) if aa_sub is not None else 0
                            n_la = len(la_sub) if la_sub is not None else 0

                            with st.expander(f"{disc} — Aramco: {n_aa} / LTC: {n_la} attributes"):
                                # Aramco
                                if aa_sub is not None and not aa_sub.empty:
                                    st.markdown(f"**Aramco Functional Attributes ({n_aa})**")
                                    display_cols = [c for c in ["Class_Id", "Class_Desc", "Name",
                                                                 "Attribute_Desc", "Presence", "Size",
                                                                 "UomClassId", "ValidationRule"] if c in aa_sub.columns]
                                    st.dataframe(aa_sub[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)

                                    # Presence breakdown
                                    if "Presence" in aa_sub.columns:
                                        p_counts = aa_sub["Presence"].fillna("N/A").astype(str).str.strip().value_counts()
                                        st.markdown("**Presence breakdown:**  " + "  |  ".join(f"`{k}`: {v}" for k, v in p_counts.items()))
                                else:
                                    st.info("No Aramco attributes for this discipline.")

                                st.divider()

                                # LTC
                                if la_sub is not None and not la_sub.empty:
                                    st.markdown(f"**LTC Physical Attributes ({n_la})**")
                                    display_cols = [c for c in ["Class_Id", "Name", "Description",
                                                                 "Presence", "Size", "UomClassId",
                                                                 "ValidationRule", "Aspect"] if c in la_sub.columns]
                                    st.dataframe(la_sub[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)

                                    if "Presence" in la_sub.columns:
                                        p_counts = la_sub["Presence"].fillna("N/A").astype(str).str.strip().value_counts()
                                        st.markdown("**Presence breakdown:**  " + "  |  ".join(f"`{k}`: {v}" for k, v in p_counts.items()))
                                else:
                                    st.info("No LTC attributes for this discipline.")

                                # Attribute gap
                                if aa_sub is not None and la_sub is not None:
                                    st.divider()
                                    a_names = set(aa_sub["Name"].dropna().str.strip().str.lower()) if "Name" in aa_sub.columns else set()
                                    l_names = set(la_sub["Name"].dropna().str.strip().str.lower()) if "Name" in la_sub.columns else set()
                                    common = a_names & l_names
                                    only_a = a_names - l_names
                                    only_l = l_names - a_names
                                    g1, g2, g3 = st.columns(3)
                                    with g1:
                                        st.metric("Common", len(common))
                                    with g2:
                                        st.metric("Only in Aramco", len(only_a))
                                    with g3:
                                        st.metric("Only in LTC", len(only_l))
                                    if only_a:
                                        with st.expander(f"Attributes only in Aramco ({len(only_a)})"):
                                            for n in sorted(only_a):
                                                st.markdown(f"- {n}")
                                    if only_l:
                                        with st.expander(f"Attributes only in LTC ({len(only_l)})"):
                                            for n in sorted(only_l):
                                                st.markdown(f"- {n}")

            elif attr_view == "Full Aramco Attributes":
                if aramco_attrs is not None:
                    st.dataframe(aramco_attrs.drop(columns=["_source"], errors="ignore"), use_container_width=True, hide_index=True)
                    st.caption(f"Total: {len(aramco_attrs)} rows")
                else:
                    st.info("Aramco attribute data not loaded.")

            elif attr_view == "Full LTC Attributes":
                if ltc_attrs is not None:
                    st.dataframe(ltc_attrs.drop(columns=["_source"], errors="ignore"), use_container_width=True, hide_index=True)
                    st.caption(f"Total: {len(ltc_attrs)} rows")
                else:
                    st.info("LTC attribute data not loaded.")

            elif attr_view == "Attribute Comparison":
                if aramco_attrs is None or ltc_attrs is None:
                    st.warning("Need both Aramco and LTC attribute data for comparison. Upload both files.")
                elif not class_options:
                    st.warning("No class data available. Upload files and/or run harmonization.")
                else:
                    selected_class = st.selectbox("Select equipment class for comparison", class_options, key="attr_cmp_class")
                    if selected_class:
                        ids = class_id_map.get(selected_class, [])

                        aramco_sub = aramco_attrs[aramco_attrs["Class_Id"].astype(str).str.strip().isin(ids)]
                        ltc_sub = ltc_attrs[ltc_attrs["Class_Id"].astype(str).str.strip().isin(ids)]

                        aramco_names = set(aramco_sub["Name"].dropna().str.strip().str.lower()) if "Name" in aramco_sub.columns else set()
                        ltc_names = set(ltc_sub["Name"].dropna().str.strip().str.lower()) if "Name" in ltc_sub.columns else set()

                        common = aramco_names & ltc_names
                        only_aramco = aramco_names - ltc_names
                        only_ltc = ltc_names - aramco_names

                        cmp_cols = st.columns(3)
                        with cmp_cols[0]:
                            st.metric("Common Attributes", len(common))
                        with cmp_cols[1]:
                            st.metric("Only in Aramco", len(only_aramco))
                        with cmp_cols[2]:
                            st.metric("Only in LTC", len(only_ltc))

                        st.divider()
                        col_left, col_right = st.columns(2)
                        with col_left:
                            st.markdown("**Aramco Attributes**")
                            if not aramco_sub.empty:
                                display_cols = [c for c in ["Name", "Attribute_Desc", "Presence", "Size", "Description"] if c in aramco_sub.columns]
                                st.dataframe(aramco_sub[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                            else:
                                st.info("No Aramco attributes for this class.")
                        with col_right:
                            st.markdown("**LTC Attributes**")
                            if not ltc_sub.empty:
                                display_cols = [c for c in ["Name", "Presence", "Size", "Description"] if c in ltc_sub.columns]
                                st.dataframe(ltc_sub[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)
                            else:
                                st.info("No LTC attributes for this class.")

                        if only_aramco:
                            st.markdown("**:orange[Attributes only in Aramco (gap in LTC):]**")
                            for name in sorted(only_aramco):
                                st.markdown(f"- {name}")
                        if only_ltc:
                            st.markdown("**:blue[Attributes only in LTC (gap in Aramco):]**")
                            for name in sorted(only_ltc):
                                st.markdown(f"- {name}")
                        if not only_aramco and not only_ltc:
                            st.success("Both sources have the same attributes for this class.")


# ══════════════════════════════════════════════════════════════════════
# TAB: Connection Map
# ══════════════════════════════════════════════════════════════════════
with tab_visual:
    classes = st.session_state.classes
    if not classes:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-icon">&#128279;</div>
            <h3>No Connection Data</h3>
            <p>Run harmonization first to visualize connections between data sources.</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="section-card">
            <h3>Connection Map</h3>
            <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
                Visual representation of how equipment classes connect across data sources. Solid lines = matched, dashed red = gap.
            </p>
        </div>
        """, unsafe_allow_html=True)

        # Search filter: if a search match is active, show only that class
        search_class = st.session_state.search_matched_class
        if search_class:
            _fc3, _fc4 = st.columns([5, 1])
            with _fc3:
                st.info(f"Filtered by search result: **{search_class['canonical_name']}**")
            with _fc4:
                if st.button("Clear Filter", key="clear_filter_connmap", type="secondary", use_container_width=True):
                    st.session_state.search_matched_class = None
                    st.rerun()
            st.graphviz_chart(build_connection_graph(search_class), use_container_width=True)
            st.caption("Solid = matched. Dashed red = GAP.")
        else:
            view_mode = st.radio("View", ["Per-class detail", "Overview (first 20)"], horizontal=True)
            if view_mode == "Per-class detail":
                options = [f"{c['index']}. {c['canonical_name']}" for c in classes]
                selected = st.selectbox("Select equipment class", options)
                if selected:
                    idx = int(selected.split(".")[0]) - 1
                    st.graphviz_chart(build_connection_graph(classes[idx]), use_container_width=True)
                    st.caption("Solid = matched. Dashed red = GAP.")
            else:
                st.graphviz_chart(build_overview_graph(classes, 20), use_container_width=True)
                if len(classes) > 20:
                    st.caption(f"First 20 of {len(classes)}.")


# ══════════════════════════════════════════════════════════════════════
# TAB: Search
# ══════════════════════════════════════════════════════════════════════
with tab_search:
    st.markdown("""
    <div class="section-card">
        <h3>Search Equipment Classes</h3>
        <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
            Type an equipment name to find the best match in the harmonized dataset.
        </p>
    </div>
    """, unsafe_allow_html=True)
    query = st.text_input("Equipment name", placeholder="e.g. centrifugal pump...")

    if query and st.session_state.classes:
        # Build expanded search index: include compound "/" parts so
        # searching "Junction Box" finds "Junction Box / Splice Case"
        search_names = []
        search_idx_map = []  # maps expanded index -> original class index
        for ci, c in enumerate(st.session_state.classes):
            search_names.append(c["canonical_name"])
            search_idx_map.append(ci)
            if "/" in c["canonical_name"]:
                for part in [p.strip() for p in c["canonical_name"].split("/") if p.strip()]:
                    search_names.append(part)
                    search_idx_map.append(ci)

        # Match with dual scorer + abbreviation expansion
        expanded_candidates = [_normalize_abbreviations(n) for n in search_names]
        query_variants = [query]
        if "/" in query:
            query_variants += [p.strip() for p in query.split("/") if p.strip()]

        best_result = None
        best_score = 0
        for q in query_variants:
            q_exp = _normalize_abbreviations(q)
            for scorer in (fuzz.token_set_ratio, fuzz.token_sort_ratio):
                result = process.extractOne(q, search_names, scorer=scorer)
                if result and result[1] > best_score:
                    best_result = result
                    best_score = result[1]
                result_exp = process.extractOne(q_exp, expanded_candidates, scorer=scorer)
                if result_exp and result_exp[1] > best_score:
                    best_result = (search_names[result_exp[2]], result_exp[1], result_exp[2])
                    best_score = result_exp[1]

        if best_result:
            matched_name, score, exp_idx = best_result
            orig_idx = search_idx_map[exp_idx]
            score = int(round(score))
            c = st.session_state.classes[orig_idx]
            # Store matched class so Attributes & Connection tabs can filter
            st.session_state.search_matched_class = c
            status, _ = classify_match(score)
            color = "green" if score >= 90 else "orange" if score >= 75 else "red"

            name = c["canonical_name"]
            st.markdown(f"### :{color}[**{name}** — {score}% ({status})]")

            st.info(f"The **Attributes** and **Connection Map** tabs are now filtered to show only data for **{name}**.")

            col_l, col_r = st.columns(2)
            with col_l:
                st.markdown("**Master Reference:**")
                for m in st.session_state.masters:
                    me = c["master_entries"].get(m)
                    if me:
                        st.markdown(f"- {FILE_TYPES[m]}: {me['name']} (`{me['id']}`)")
            with col_r:
                if c["matches"]:
                    st.markdown("**Matches:**")
                    for src, match in c["matches"].items():
                        st.markdown(f"- {FILE_TYPES.get(src, src)}: {match['name']} — {match['score']}%")
                if c["gaps"]:
                    st.markdown("**:red[Gap In:]** " + ", ".join(FILE_TYPES.get(g, g) for g in c["gaps"]))

            st.divider()
            st.graphviz_chart(build_connection_graph(c), use_container_width=True)
        else:
            st.session_state.search_matched_class = None
    elif query:
        st.warning("Run harmonization first.")
    else:
        # No query entered — clear search filter
        st.session_state.search_matched_class = None


# ══════════════════════════════════════════════════════════════════════
# TAB: Batch
# ══════════════════════════════════════════════════════════════════════
with tab_batch:
    st.markdown("""
    <div class="section-card">
        <h3>Batch Process</h3>
        <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
            Process multiple equipment names at once. Enter manually or upload a CSV file.
        </p>
    </div>
    """, unsafe_allow_html=True)
    batch_input = st.text_area("Equipment names (one per line)", height=150)
    batch_file = st.file_uploader("Or upload CSV", type=["csv"], key="batch_csv")

    if st.button("Process Batch", type="primary"):
        items = []
        if batch_file:
            try:
                bdf = pd.read_csv(batch_file)
                if bdf.empty:
                    st.error("CSV file is empty.")
                else:
                    col = "name" if "name" in bdf.columns else bdf.columns[0]
                    items = bdf[col].dropna().astype(str).tolist()
            except Exception as e:
                st.error(f"Error reading CSV: {e}")
        elif batch_input.strip():
            items = [l.strip() for l in batch_input.strip().split("\n") if l.strip()]

        if not items:
            st.warning("Enter at least one item.")
        elif not st.session_state.classes:
            st.error("Run harmonization first.")
        else:
            # Build expanded search index with compound name parts
            search_names = []
            search_idx_map = []
            for ci, c in enumerate(st.session_state.classes):
                search_names.append(c["canonical_name"])
                search_idx_map.append(ci)
                if "/" in c["canonical_name"]:
                    for part in [p.strip() for p in c["canonical_name"].split("/") if p.strip()]:
                        search_names.append(part)
                        search_idx_map.append(ci)

            expanded_candidates = [_normalize_abbreviations(n) for n in search_names]
            results = []
            progress = st.progress(0, text="Processing batch...")
            for item_idx, item in enumerate(items):
                query_variants = [item]
                if "/" in item:
                    query_variants += [p.strip() for p in item.split("/") if p.strip()]

                best_result = None
                best_score = 0
                for q in query_variants:
                    q_exp = _normalize_abbreviations(q)
                    for scorer in (fuzz.token_sort_ratio, fuzz.token_set_ratio):
                        result = process.extractOne(q, search_names, scorer=scorer)
                        if result and result[1] > best_score:
                            best_result = result
                            best_score = result[1]
                        result_exp = process.extractOne(q_exp, expanded_candidates, scorer=scorer)
                        if result_exp and result_exp[1] > best_score:
                            best_result = (search_names[result_exp[2]], result_exp[1], result_exp[2])
                            best_score = result_exp[1]

                if best_result and best_score >= 50:
                    matched_name, score, exp_idx = best_result
                    orig_idx = search_idx_map[exp_idx]
                    c = st.session_state.classes[orig_idx]
                    score = int(round(score))
                    status, _ = classify_match(score)
                    gaps = ", ".join(FILE_TYPES.get(g, g) for g in c["gaps"]) or "No Gaps"
                    results.append({"Input": item, "Best Match": c["canonical_name"], "Score%": score, "Status": status, "Gap In": gaps})
                else:
                    results.append({"Input": item, "Best Match": "—", "Score%": 0, "Status": "No Match", "Gap In": "—"})
                progress.progress((item_idx + 1) / len(items), text=f"Processing {item_idx + 1}/{len(items)}...")
            progress.empty()
            result_df = pd.DataFrame(results)
            matched_count = sum(1 for r in results if r["Score%"] > 0)
            st.success(f"Processed {len(results)} items: {matched_count} matched, {len(results) - matched_count} unmatched")
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            st.download_button("Download Results", data=result_df.to_csv(index=False).encode(), file_name="batch_results.csv", mime="text/csv")


# ══════════════════════════════════════════════════════════════════════
# TAB: Logs
# ══════════════════════════════════════════════════════════════════════
with tab_logs:
    st.markdown("""
    <div class="section-card">
        <h3>System Logs</h3>
        <p style="color: var(--kbr-gray); font-size: 0.85rem; margin: 0;">
            Operational log of all harmonization activities.
        </p>
    </div>
    """, unsafe_allow_html=True)
    if st.session_state.logs:
        st.code("\n".join(reversed(st.session_state.logs)), language="log")
    else:
        st.info("No logs yet.")

# ── Footer ──
st.markdown("""
<div style="margin-top: 3rem; padding: 1.5rem 2rem; background: linear-gradient(90deg, #f8fafc, #f1f5f9);
            border-top: 2px solid #e2e8f0; border-radius: 12px;">
    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;">
        <div>
            <span style="font-weight: 700; color: #1e293b; font-size: 0.9rem;">KBR RDL Data Harmonizer</span>
            <span style="color: #94a3b8; font-size: 0.82rem;"> &mdash; v3.0</span>
        </div>
        <div style="color: #94a3b8; font-size: 0.78rem;">
            Built by <strong style="color: #003087;">KBR AMCDE Team</strong> &bull;
            Equipment Class Harmonization &amp; Gap Analysis Platform
        </div>
    </div>
</div>
""", unsafe_allow_html=True)
