"""
KBR RDL Data Harmonizer v3.0 — Streamlit App
Run:  streamlit run main.py
"""

import uuid
import io
import re
import pandas as pd
import streamlit as st
import altair as alt
from rapidfuzz import fuzz, process

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
    --kbr-red: #b91c1c;
    --kbr-red-dark: #7f1d1d;
    --kbr-red-light: #fecaca;
    --kbr-navy: #1e293b;
    --kbr-blue: #2563eb;
    --kbr-green: #059669;
    --kbr-orange: #d97706;
    --kbr-purple: #7c3aed;
    --kbr-gray: #64748b;
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
    border-bottom: 2px solid var(--kbr-red);
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
    color: var(--kbr-red) !important;
    font-weight: 700;
    border-top: 3px solid var(--kbr-red);
}
.stTabs [data-baseweb="tab"]:hover {
    color: var(--kbr-red);
    background: rgba(185, 28, 28, 0.05);
}
.stTabs [data-baseweb="tab-panel"] {
    padding: 1.5rem 0.5rem;
}

/* ── Metric card override ── */
[data-testid="stMetric"] {
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    border-left: 4px solid var(--kbr-red);
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
    background: linear-gradient(135deg, var(--kbr-red) 0%, var(--kbr-red-dark) 100%) !important;
    border: none !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em;
    padding: 0.6rem 1.5rem !important;
    border-radius: 8px !important;
    box-shadow: 0 2px 8px rgba(185, 28, 28, 0.3) !important;
    transition: all 0.2s ease !important;
}
.stButton > button[kind="primary"]:hover {
    box-shadow: 0 4px 16px rgba(185, 28, 28, 0.4) !important;
    transform: translateY(-1px);
}
.stButton > button[kind="secondary"], .stButton > button:not([kind]) {
    border: 1.5px solid var(--card-border) !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
    transition: all 0.2s ease !important;
}
.stButton > button[kind="secondary"]:hover, .stButton > button:not([kind]):hover {
    border-color: var(--kbr-red) !important;
    color: var(--kbr-red) !important;
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
    border-color: var(--kbr-red);
}

/* ── Selectbox / multiselect ── */
.stSelectbox > div > div, .stMultiSelect > div > div {
    border-radius: 8px !important;
}

/* ── Slider ── */
.stSlider > div > div > div > div {
    background: var(--kbr-red) !important;
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
.stat-pill.red { background: #fee2e2; color: #991b1b; }
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
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

FILE_TYPES = {
    "aramco":  "Saudi Aramco 9COM",
    "cfihos":  "CFIHOS Standard",
    "kbr":     "KBR FEED",
    "ltc":     "LTC Contractor",
    "sa_doc":  "SA Document",
}

SOURCE_COLORS = {
    "aramco": "#16a34a",
    "cfihos": "#2563eb",
    "kbr":    "#dc2626",
    "ltc":    "#9333ea",
    "sa_doc": "#d97706",
}


def add_log(msg, level="INFO"):
    st.session_state.logs.append(f"[{level}] {msg}")


# ──────────────────────────────────────────────────────────────────────
# File readers
# ──────────────────────────────────────────────────────────────────────
def read_aramco(file) -> list[dict]:
    try:
        df = pd.read_excel(file, sheet_name="ISM Functional Classes")
    except Exception:
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("Id", row.get("id", ""))),
            "name": str(row.get("Name", row.get("name", ""))),
            "cfihos_ref": str(row.get("nmcltr:CFIHOS_1.5", "")),
            "source": "aramco",
        })
    return [r for r in records if r["name"].strip()]


def read_cfihos(file) -> list[dict]:
    try:
        df = pd.read_excel(file, sheet_name="equipment class")
    except Exception:
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("CFIHOS unique id", row.get("id", ""))),
            "name": str(row.get("equipment class name", row.get("name", ""))),
            "source": "cfihos",
        })
    return [r for r in records if r["name"].strip()]


def read_kbr(file) -> list[dict]:
    df = pd.read_excel(file, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("Class Id", row.get("id", ""))),
            "name": str(row.get("Class Name (855)", row.get("name", ""))),
            "discipline": str(row.get("Discipline", "")),
            "source": "kbr",
        })
    return [r for r in records if r["name"].strip()]


def read_ltc(file) -> list[dict]:
    # Prefer "ISM Physical Classes" (PCL IDs) over "ISM Functional Classes" (FCL IDs)
    df = None
    for sheet in ("ISM Physical Classes", "ISM Functional Classes"):
        try:
            df = pd.read_excel(file, sheet_name=sheet)
            break
        except Exception:
            continue
    if df is None:
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        raw_name = str(row.get("Name", row.get("name", "")))
        # Strip [OBSOLETE] or similar bracketed prefixes so fuzzy matching works
        clean_name = re.sub(r'\[.*?\]\s*', '', raw_name).strip()
        records.append({
            "id": str(row.get("Id", row.get("id", ""))),
            "name": clean_name if clean_name else raw_name,
            "source": "ltc",
        })
    return [r for r in records if r["name"].strip()]


def read_sa_doc(file) -> list[dict]:
    try:
        df = pd.read_excel(file, sheet_name="SA_DOC_attributes")
    except Exception:
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("ID (CFIHOS_1.5)", row.get("id", ""))),
            "name": str(row.get("Attribute", row.get("name", ""))),
            "cfihos_name": str(row.get("Name (CFIHOS_1.5)", "")),
            "source": "sa_doc",
        })
    return [r for r in records if r["name"].strip()]


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
    except Exception:
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
    except Exception:
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
# ──────────────────────────────────────────────────────────────────────
def _tokenize(text):
    words = set(re.findall(r'[a-z]{2,}', text.lower()))
    noise = {"the", "and", "for", "with", "from", "that", "this", "its"}
    return words - noise


def _has_word_overlap(name_a, name_b):
    tokens_a = _tokenize(name_a)
    tokens_b = _tokenize(name_b)
    if not tokens_a or not tokens_b:
        return True
    if tokens_a & tokens_b:
        return True
    for a in tokens_a:
        for b in tokens_b:
            if len(a) >= 3 and len(b) >= 3 and (a in b or b in a):
                return True
    return False


def _expand_compound_names(candidates):
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
    if not candidates:
        return None
    expanded_names, index_map = _expand_compound_names(candidates)
    query_variants = [name]
    if "/" in name:
        query_variants += [p.strip() for p in name.split("/") if p.strip()]

    best_result = None
    best_score = 0
    for q in query_variants:
        result = process.extractOne(
            q, expanded_names, scorer=fuzz.token_sort_ratio, score_cutoff=threshold,
        )
        if result and result[1] > best_score:
            best_result = result
            best_score = result[1]

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
    """Masters = ONE unified reference. Then compare each non-master source."""
    masters = st.session_state.masters
    files = st.session_state.files
    threshold = st.session_state.match_threshold

    # Step 1: Build unified master rows
    if len(masters) == 1:
        master_records = files[masters[0]]["records"]
        merged = []
        seen = set()
        for r in master_records:
            key = r["name"].strip().lower()
            if key in seen:
                continue
            seen.add(key)
            merged.append({
                "master_entries": {masters[0]: r},
                "canonical_name": r["name"],
            })
    else:
        m1, m2 = masters[0], masters[1]
        m1_records = files[m1]["records"]
        m2_records = list(files[m2]["records"])
        m2_used = set()
        seen = set()
        merged = []

        for r1 in m1_records:
            key = r1["name"].strip().lower()
            if key in seen:
                continue
            seen.add(key)
            match = fuzzy_match(r1["name"], m2_records, threshold)
            entry = {"master_entries": {m1: r1}, "canonical_name": r1["name"]}
            if match:
                for i, r2 in enumerate(m2_records):
                    if r2["id"] == match["id"] and r2["name"] == match["name"]:
                        m2_used.add(i)
                        break
                match_clean = {k: v for k, v in match.items() if k != "score"}
                entry["master_entries"][m2] = match_clean
                entry["cross_score"] = match["score"]
            merged.append(entry)

        for i, r2 in enumerate(m2_records):
            if i not in m2_used:
                key = r2["name"].strip().lower()
                if key not in seen:
                    seen.add(key)
                    merged.append({
                        "master_entries": {m2: r2},
                        "canonical_name": r2["name"],
                    })

    add_log(f"Master reference built from {masters}: {len(merged)} unique classes (threshold={threshold}%)")

    # Step 2: Compare each non-master source against the unified master
    non_masters = {k: v["records"] for k, v in files.items() if k not in masters}

    harmonized = []
    for idx, row in enumerate(merged):
        entry = {
            "uid": str(uuid.uuid4()),
            "index": idx + 1,
            "canonical_name": row["canonical_name"],
            "master_entries": row["master_entries"],
            "cross_score": row.get("cross_score"),
            "matches": {},
            "gaps": [],
        }
        for src, candidates in non_masters.items():
            match = fuzzy_match(row["canonical_name"], candidates, threshold)
            if match:
                entry["matches"][src] = match
            else:
                entry["gaps"].append(src)
        harmonized.append(entry)

    add_log(f"Harmonization complete: {len(harmonized)} classes, compared against {list(non_masters.keys())}")

    # Step 3: Compute reverse gaps — non-master records that matched nothing in the master
    reverse_gaps = {}
    for src, candidates in non_masters.items():
        # Collect IDs of candidates that were matched by at least one master entry
        matched_ids = set()
        for entry in harmonized:
            m = entry["matches"].get(src)
            if m:
                matched_ids.add((m["id"], m["name"]))
        # Find candidates that were never matched
        unmatched = []
        for rec in candidates:
            if (rec["id"], rec["name"]) not in matched_ids:
                unmatched.append(rec)
        if unmatched:
            reverse_gaps[src] = unmatched
    if reverse_gaps:
        parts = [f"{FILE_TYPES.get(s, s)}: {len(recs)} extra" for s, recs in reverse_gaps.items()]
        add_log(f"Reverse gaps (extra classes not in master): {', '.join(parts)}")
    st.session_state.reverse_gaps = reverse_gaps

    return harmonized


# ──────────────────────────────────────────────────────────────────────
# Demo data
# ──────────────────────────────────────────────────────────────────────
def load_demo_data():
    demo = {
        "aramco": [
            {"id": "AC-001", "name": "Centrifugal Pump", "cfihos_ref": "CF-101", "source": "aramco"},
            {"id": "AC-002", "name": "Shell and Tube Heat Exchanger", "cfihos_ref": "CF-201", "source": "aramco"},
            {"id": "AC-003", "name": "Pressure Vessel", "cfihos_ref": "CF-301", "source": "aramco"},
            {"id": "AC-004", "name": "Control Valve", "cfihos_ref": "CF-401", "source": "aramco"},
            {"id": "AC-005", "name": "Electric Motor", "cfihos_ref": "CF-501", "source": "aramco"},
            {"id": "AC-006", "name": "Air Cooled Heat Exchanger", "cfihos_ref": "CF-202", "source": "aramco"},
            {"id": "AC-007", "name": "Reciprocating Compressor", "cfihos_ref": "CF-601", "source": "aramco"},
            {"id": "AC-008", "name": "Storage Tank", "cfihos_ref": "CF-701", "source": "aramco"},
        ],
        "cfihos": [
            {"id": "CF-101", "name": "Centrifugal Pump", "source": "cfihos"},
            {"id": "CF-201", "name": "Shell and Tube Heat Exchanger", "source": "cfihos"},
            {"id": "CF-301", "name": "Pressure Vessel", "source": "cfihos"},
            {"id": "CF-401", "name": "Control Valve", "source": "cfihos"},
            {"id": "CF-501", "name": "Electric Motor", "source": "cfihos"},
            {"id": "CF-202", "name": "Air Cooled Exchanger", "source": "cfihos"},
            {"id": "CF-601", "name": "Reciprocating Compressor", "source": "cfihos"},
            {"id": "CF-701", "name": "Atmospheric Storage Tank", "source": "cfihos"},
        ],
        "kbr": [
            {"id": "KBR-P01", "name": "Centrifugal Pump", "discipline": "Mechanical", "source": "kbr"},
            {"id": "KBR-H01", "name": "Shell & Tube Heat Exchanger", "discipline": "Mechanical", "source": "kbr"},
            {"id": "KBR-V01", "name": "Pressure Vessel", "discipline": "Mechanical", "source": "kbr"},
            {"id": "KBR-CV1", "name": "Control Valve Assembly", "discipline": "Instrumentation", "source": "kbr"},
            {"id": "KBR-E01", "name": "Electric Motor Driver", "discipline": "Electrical", "source": "kbr"},
            {"id": "KBR-H02", "name": "Air Cooled Heat Exchanger", "discipline": "Mechanical", "source": "kbr"},
            {"id": "KBR-C01", "name": "Reciprocating Compressor", "discipline": "Mechanical", "source": "kbr"},
            # Extra KBR classes not in Aramco (for enrichment suggestion demo)
            {"id": "KBR-T01", "name": "Cooling Tower", "discipline": "Mechanical", "source": "kbr"},
            {"id": "KBR-F01", "name": "Flare Stack", "discipline": "Process", "source": "kbr"},
            {"id": "KBR-D01", "name": "Drum Separator", "discipline": "Process", "source": "kbr"},
        ],
        "ltc": [
            {"id": "LTC-001", "name": "Centrifugal Pump Unit", "source": "ltc"},
            {"id": "LTC-002", "name": "Shell and Tube HX", "source": "ltc"},
            {"id": "LTC-003", "name": "Pressure Vessel", "source": "ltc"},
            {"id": "LTC-004", "name": "Control Valve", "source": "ltc"},
            {"id": "LTC-005", "name": "Electric Motor", "source": "ltc"},
            {"id": "LTC-006", "name": "Air Cooled Exchanger", "source": "ltc"},
        ],
        "sa_doc": [
            {"id": "SA-A01", "name": "Design Pressure", "cfihos_name": "Design Pressure", "source": "sa_doc"},
            {"id": "SA-A02", "name": "Design Temperature", "cfihos_name": "Design Temperature", "source": "sa_doc"},
            {"id": "SA-A03", "name": "Material of Construction", "cfihos_name": "Material", "source": "sa_doc"},
        ],
    }
    for key, records in demo.items():
        st.session_state.files[key] = {"filename": f"demo_{key}.xlsx", "records": records}
    add_log("Demo data loaded for all 5 sources")


# ──────────────────────────────────────────────────────────────────────
# Export — clean organized Excel
# ──────────────────────────────────────────────────────────────────────
def build_export_df(classes):
    """Build export DataFrame.

    Layout:
      # | MASTER REFERENCE (one block) | Source1 columns | Source2 columns | Gap In
      Masters are shown together as one reference block.
      Each non-master source gets: ID | Name | Match% | Status
      Last column: "Gap In" lists which sources are missing, or "No Gaps".
    """
    masters = st.session_state.masters
    non_master_sources = set()
    for c in classes:
        non_master_sources.update(c["matches"].keys())
        non_master_sources.update(c["gaps"])
    non_master_sources = sorted(non_master_sources)

    rows = []
    for c in classes:
        gap_count = len(c["gaps"])

        row = {"#": c["index"]}

        # --- MASTER REFERENCE block ---
        for m in masters:
            label = FILE_TYPES.get(m, m)
            me = c["master_entries"].get(m)
            row[f"Master ({label}) ID"] = me["id"] if me else ""
            row[f"Master ({label}) Name"] = me["name"] if me else ""

        if len(masters) == 2:
            cs = c.get("cross_score")
            row["Masters Cross-Match%"] = cs if cs is not None else ""

        # --- Each non-master source ---
        for s in non_master_sources:
            label = FILE_TYPES.get(s, s)
            m = c["matches"].get(s)
            if m:
                status, _ = classify_match(m["score"])
                row[f"{label} ID"] = m["id"]
                row[f"{label} Name"] = m["name"]
                row[f"{label} Match%"] = m["score"]
                row[f"{label} Status"] = status
            else:
                row[f"{label} ID"] = ""
                row[f"{label} Name"] = ""
                row[f"{label} Match%"] = ""
                row[f"{label} Status"] = "GAP"

        # --- Gap summary ---
        if gap_count == 0:
            row["Gap In"] = "No Gaps"
        else:
            row["Gap In"] = " | ".join(FILE_TYPES.get(g, g) for g in c["gaps"])
        row["Gap Count"] = gap_count

        rows.append(row)

    return pd.DataFrame(rows)


def build_excel_bytes(classes):
    """Build the main harmonization export with professional styling."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import openpyxl

    masters = st.session_state.masters

    # Style constants
    kbr_red = "B91C1C"
    kbr_dark = "7F1D1D"
    navy = "1E293B"
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
    title_fill = PatternFill("solid", fgColor=kbr_red)
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

    def _style_title_rows(ws, title_text, subtitle_text, num_cols):
        """Add branded title + subtitle rows to a worksheet."""
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        t = ws.cell(row=1, column=1, value=f"  {title_text}")
        t.font = title_font
        t.fill = title_fill
        t.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 36
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
        s = ws.cell(row=2, column=1, value=f"  {subtitle_text}")
        s.font = sub_font
        s.fill = sub_fill
        s.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 24

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
    _write_headers(ws1, 3, cols)
    ws1.auto_filter.ref = f"A3:{get_column_letter(len(cols))}3"
    ws1.freeze_panes = "A4"

    for ri, (_, row) in enumerate(df.iterrows(), 4):
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
    _write_headers(ws2, 3, gap_cols)
    ws2.auto_filter.ref = f"A3:{get_column_letter(len(gap_cols))}3"
    ws2.freeze_panes = "A4"

    for ri, grow in enumerate(gap_rows, 4):
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
    _write_headers(ws3, 3, ["Metric", "Value"])

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

    for ri, (metric, value) in enumerate(summary_data, 4):
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

    # ── Style definitions ──
    kbr_red = "B91C1C"
    kbr_dark = "7F1D1D"
    navy = "1E293B"
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
    title_fill = PatternFill("solid", fgColor=kbr_red)
    title_align = Alignment(horizontal="left", vertical="center")

    # Subtitle row
    sub_font = Font(name="Calibri", size=10, color="94A3B8")
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

        # ── Row 1: Title bar ──
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
        title_cell = ws.cell(row=1, column=1, value=f"  KBR RDL Data Harmonizer — {label} Enriched Master")
        title_cell.font = title_font
        title_cell.fill = title_fill
        title_cell.alignment = title_align
        ws.row_dimensions[1].height = 36

        # ── Row 2: Subtitle ──
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(columns))
        orig_cnt = len(original)
        add_cnt = len(selected_additions)
        sub_cell = ws.cell(row=2, column=1,
                           value=f"  Original: {orig_cnt} classes  |  Suggested Additions: {add_cnt}  |  New Total: {orig_cnt + add_cnt}  |  Threshold: {threshold}%")
        sub_cell.font = Font(name="Calibri", size=10, color=white)
        sub_cell.fill = sub_fill
        sub_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 24

        # ── Row 3: Column headers ──
        for ci, col_name in enumerate(columns, 1):
            cell = ws.cell(row=3, column=ci, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        ws.row_dimensions[3].height = 28
        ws.auto_filter.ref = f"A3:{get_column_letter(len(columns))}3"

        # ── Data rows ──
        row_num = 4
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
        ws.freeze_panes = "A4"

    # ── Summary sheet ──
    ws_sum = wb.create_sheet(title="Enrichment Summary")
    orig_total = sum(len(files[m]["records"]) for m in masters if m in files)

    # Title
    ws_sum.merge_cells("A1:B1")
    t = ws_sum.cell(row=1, column=1, value="  Enrichment Summary")
    t.font = title_font
    t.fill = title_fill
    t.alignment = title_align
    ws_sum.row_dimensions[1].height = 36

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
        cell = ws_sum.cell(row=3, column=ci, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    for ri, (metric, value) in enumerate(summary_data, 4):
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


# ══════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════

# ── Header ──
st.markdown("""
<div style="background: linear-gradient(135deg, #b91c1c 0%, #7f1d1d 50%, #1e293b 100%);
            padding: 1.8rem 2.5rem; border-radius: 16px; margin-bottom: 1.5rem;
            box-shadow: 0 8px 32px rgba(127,29,29,0.3); position: relative; overflow: hidden;">
    <div style="position: absolute; top: -20px; right: -20px; width: 200px; height: 200px;
                background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 70%);
                border-radius: 50%;"></div>
    <div style="position: absolute; bottom: -30px; left: 30%; width: 150px; height: 150px;
                background: radial-gradient(circle, rgba(255,255,255,0.05) 0%, transparent 70%);
                border-radius: 50%;"></div>
    <div style="display: flex; align-items: center; justify-content: space-between; position: relative; z-index: 1;">
        <div>
            <div style="display: flex; align-items: center; gap: 0.75rem; margin-bottom: 0.3rem;">
                <div style="background: rgba(255,255,255,0.15); padding: 0.5rem 0.75rem; border-radius: 10px;
                            backdrop-filter: blur(10px); font-size: 1.5rem;">&#9881;</div>
                <h1 style="color: white; margin: 0; font-size: 1.9rem; font-weight: 800;
                           letter-spacing: -0.02em; font-family: 'Inter', sans-serif;">
                    KBR RDL Data Harmonizer
                </h1>
            </div>
            <p style="color: rgba(252,165,165,0.9); margin: 0; font-size: 0.88rem; font-weight: 400;
                      letter-spacing: 0.02em; padding-left: 3.5rem;">
                Equipment Class Harmonization &amp; Gap Analysis Platform
            </p>
        </div>
        <div style="text-align: right;">
            <div style="background: rgba(255,255,255,0.12); padding: 0.35rem 1rem; border-radius: 20px;
                        backdrop-filter: blur(10px); margin-bottom: 0.4rem;">
                <span style="color: #fca5a5; font-size: 0.75rem; font-weight: 600;
                             letter-spacing: 0.08em; text-transform: uppercase;">Version 3.0</span>
            </div>
            <p style="color: rgba(255,255,255,0.5); font-size: 0.72rem; margin: 0;">
                Built by KBR AMCDE Team
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
            badge_html = f'<span class="stat-pill red" style="font-size:0.65rem;">{tag}</span> ' if tag else ""
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
                            row = aramco_attrs[aramco_attrs["Class_Id"] == cid].iloc[0]
                            desc = str(row.get("Class_Desc", ""))
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
            ["Browse by Class", "Full Aramco Attributes", "Full LTC Attributes", "Attribute Comparison"],
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

        # Try token_set_ratio first (better for subset queries like "Junction Box")
        # then fall back to token_sort_ratio
        best_result = None
        best_score = 0
        for scorer in (fuzz.token_set_ratio, fuzz.token_sort_ratio):
            result = process.extractOne(query, search_names, scorer=scorer)
            if result and result[1] > best_score:
                best_result = result
                best_score = result[1]

        if best_result:
            matched_name, score, exp_idx = best_result
            orig_idx = search_idx_map[exp_idx]
            score = int(round(score))
            c = st.session_state.classes[orig_idx]
            status, _ = classify_match(score)
            color = "green" if score >= 90 else "orange" if score >= 75 else "red"

            name = c["canonical_name"]
            st.markdown(f"### :{color}[**{name}** — {score}% ({status})]")

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
    elif query:
        st.warning("Run harmonization first.")


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
            bdf = pd.read_csv(batch_file)
            col = "name" if "name" in bdf.columns else bdf.columns[0]
            items = bdf[col].dropna().astype(str).tolist()
        elif batch_input.strip():
            items = [l.strip() for l in batch_input.strip().split("\n") if l.strip()]

        if not items:
            st.warning("Enter at least one item.")
        elif not st.session_state.classes:
            st.error("Run harmonization first.")
        else:
            names = [c["canonical_name"] for c in st.session_state.classes]
            results = []
            for item in items:
                match = process.extractOne(item, names, scorer=fuzz.token_sort_ratio)
                if match:
                    name, score, idx = match
                    c = st.session_state.classes[idx]
                    status, _ = classify_match(int(round(score)))
                    gaps = ", ".join(FILE_TYPES.get(g, g) for g in c["gaps"]) or "No Gaps"
                    results.append({"Input": item, "Best Match": name, "Score%": int(round(score)), "Status": status, "Gap In": gaps})
                else:
                    results.append({"Input": item, "Best Match": "—", "Score%": 0, "Status": "No Match", "Gap In": "—"})
            result_df = pd.DataFrame(results)
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            st.download_button("Download", data=result_df.to_csv(index=False).encode(), file_name="batch_results.csv", mime="text/csv")


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
            Built by <strong style="color: #b91c1c;">KBR AMCDE Team</strong> &bull;
            Equipment Class Harmonization &amp; Gap Analysis Platform
        </div>
    </div>
</div>
""", unsafe_allow_html=True)
