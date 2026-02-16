"""
KBR RDL Data Harmonizer v2.0 — Streamlit App
Run:  streamlit run main.py
"""

import uuid
import io
import re
import pandas as pd
import streamlit as st
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
# Session state defaults
# ──────────────────────────────────────────────────────────────────────
DEFAULTS = {
    "files": {},
    "masters": [],
    "classes": [],
    "logs": [],
    "match_threshold": 75,
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
    try:
        df = pd.read_excel(file, sheet_name="ISM Functional Classes")
    except Exception:
        df = pd.read_excel(file, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("Id", row.get("id", ""))),
            "name": str(row.get("Name", row.get("name", ""))),
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
    buf = io.BytesIO()
    masters = st.session_state.masters

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Sheet 1: Harmonization
        df = build_export_df(classes)
        df.to_excel(writer, sheet_name="Harmonization Results", index=False)

        # Sheet 2: Gap Analysis
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
        gap_df = pd.DataFrame(gap_rows) if gap_rows else pd.DataFrame(
            columns=["#", "Equipment Class", "Found In", "Gap In", "Action"]
        )
        gap_df.to_excel(writer, sheet_name="Gap Analysis", index=False)

        # Sheet 3: Summary
        total = len(classes)
        non_master_sources = set()
        for c in classes:
            non_master_sources.update(c["matches"].keys())
            non_master_sources.update(c["gaps"])

        no_gaps = sum(1 for c in classes if not c["gaps"])
        summary_rows = [
            {"Metric": "Total Equipment Classes", "Value": total},
            {"Metric": "Master Reference", "Value": " + ".join(FILE_TYPES.get(m, m) for m in masters)},
            {"Metric": "Compared Against", "Value": ", ".join(FILE_TYPES.get(s, s) for s in sorted(non_master_sources))},
            {"Metric": "Match Threshold", "Value": f"{st.session_state.match_threshold}%"},
            {"Metric": "Classes with No Gaps", "Value": f"{no_gaps} / {total}"},
            {"Metric": "Total Gap Entries", "Value": sum(len(c["gaps"]) for c in classes)},
        ]
        for src in sorted(non_master_sources):
            cnt = sum(1 for c in classes if src in c["matches"])
            pct = int(round(cnt / total * 100)) if total else 0
            gap_cnt = total - cnt
            summary_rows.append({
                "Metric": f"{FILE_TYPES.get(src, src)}",
                "Value": f"Matched: {cnt}/{total} ({pct}%) — Gaps: {gap_cnt}",
            })
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

        # Auto-fit columns
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col_cells in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 3, 50)

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
st.markdown("""
<div style="background: linear-gradient(90deg, #b91c1c, #7f1d1d); padding: 1.5rem 2rem; border-radius: 0.75rem; margin-bottom: 1rem;">
    <h1 style="color: white; margin: 0; font-size: 1.8rem;">KBR RDL Data Harmonizer</h1>
    <p style="color: #fca5a5; margin: 0; font-size: 0.9rem;">Equipment Class Harmonization v2.0</p>
</div>
""", unsafe_allow_html=True)

tab_upload, tab_dashboard, tab_gaps, tab_visual, tab_search, tab_batch, tab_logs = st.tabs([
    "Upload & Configure",
    "Dashboard",
    "Gap Analysis",
    "Connection Map",
    "Search",
    "Batch Process",
    "Logs",
])

# ══════════════════════════════════════════════════════════════════════
# TAB: Upload & Configure
# ══════════════════════════════════════════════════════════════════════
with tab_upload:
    st.subheader("Step 1 — Choose Master File(s)")
    st.caption(
        "Select **1 or 2** file types as the master reference. "
        "They are treated as **one combined object**. "
        "All other uploaded files are compared against this master."
    )

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

    st.divider()
    threshold = st.slider(
        "Match Threshold (%)", min_value=50, max_value=100,
        value=st.session_state.match_threshold,
        help="Minimum score to count as a match. Below this = GAP.",
    )
    if threshold != st.session_state.match_threshold:
        st.session_state.match_threshold = threshold

    st.divider()
    st.subheader("Step 2 — Upload Data Files")
    cols = st.columns(3)
    for i, (key, label) in enumerate(FILE_TYPES.items()):
        with cols[i % 3]:
            is_master = key in st.session_state.masters
            badge = " MASTER" if is_master else ""
            uploaded = st.file_uploader(f"{label}{badge}", type=["xlsx", "xls", "csv"], key=f"upload_{key}")
            if uploaded is not None:
                if key not in st.session_state.files or st.session_state.files[key]["filename"] != uploaded.name:
                    try:
                        records = READERS[key](uploaded)
                        st.session_state.files[key] = {"filename": uploaded.name, "records": records}
                        add_log(f"Uploaded {key}: {uploaded.name} ({len(records)} records)")
                        st.success(f"{len(records)} records loaded")
                    except Exception as e:
                        st.error(f"Error: {e}")
                else:
                    st.success(f"{len(st.session_state.files[key]['records'])} records loaded")
            if key in st.session_state.files:
                st.caption(f"File: {st.session_state.files[key]['filename']}")

    st.divider()
    st.subheader("Step 3 — Run")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Run Harmonization", type="primary", use_container_width=True):
            if not st.session_state.masters:
                st.error("Select at least one master.")
            elif not all(m in st.session_state.files for m in st.session_state.masters):
                st.error("Upload all master files first.")
            else:
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
                "Export Excel", data=build_excel_bytes(st.session_state.classes),
                file_name="harmonization_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if st.session_state.files:
        st.divider()
        status_cols = st.columns(len(st.session_state.files))
        for i, (k, v) in enumerate(st.session_state.files.items()):
            with status_cols[i]:
                tag = "MASTER " if k in st.session_state.masters else ""
                st.metric(f"{tag}{FILE_TYPES[k]}", f"{len(v['records'])} records")


# ══════════════════════════════════════════════════════════════════════
# TAB: Dashboard
# ══════════════════════════════════════════════════════════════════════
with tab_dashboard:
    classes = st.session_state.classes
    if not classes:
        st.info("No results yet. Go to **Upload & Configure**.")
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

        # Metrics
        mc = st.columns(2 + len(non_master_keys))
        with mc[0]:
            st.metric("Total Classes", total)
        with mc[1]:
            st.metric("No Gaps", f"{no_gaps_total}/{total}")
        for i, src in enumerate(non_master_keys):
            with mc[2 + i]:
                s = source_stats[src]
                pct = int(round(s["matched"] / total * 100))
                st.metric(FILE_TYPES.get(src, src), f"{pct}% match ({s['gaps']} gaps)")

        master_label = " + ".join(FILE_TYPES.get(m, m) for m in masters)
        compared_label = ", ".join(FILE_TYPES.get(s, s) for s in non_master_keys)
        st.caption(f"**Master:** {master_label}  |  **Compared against:** {compared_label}  |  **Threshold:** {st.session_state.match_threshold}%")
        st.divider()

        # Filter
        filter_opt = st.radio("Show", ["All", "No Gaps only", "Has Gaps only"], horizontal=True)

        for c in classes:
            has_gaps = bool(c["gaps"])
            if filter_opt == "No Gaps only" and has_gaps:
                continue
            if filter_opt == "Has Gaps only" and not has_gaps:
                continue

            # Header
            gap_label = ":red[GAP IN: " + ", ".join(FILE_TYPES.get(g, g) for g in c["gaps"]) + "]" if has_gaps else ":green[No Gaps]"

            with st.expander(f"**{c['index']}. {c['canonical_name']}** — {gap_label}"):
                # Master reference
                st.markdown("**Master Reference:**")
                for m in masters:
                    me = c["master_entries"].get(m)
                    if me:
                        st.markdown(f"- {FILE_TYPES[m]}: **{me['name']}** (`{me['id']}`)")
                    else:
                        st.markdown(f"- {FILE_TYPES[m]}: :red[not in this master]")

                if len(masters) == 2 and c.get("cross_score") is not None:
                    st.caption(f"Masters cross-match: {c['cross_score']}%")

                # Per-source comparison
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
        st.subheader("Table View")
        st.dataframe(build_export_df(classes), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════
# TAB: Gap Analysis
# ══════════════════════════════════════════════════════════════════════
with tab_gaps:
    classes = st.session_state.classes
    if not classes:
        st.info("Run harmonization first.")
    else:
        masters = st.session_state.masters
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

        st.subheader(f"Gap Analysis — {len(all_gaps)} gaps found")

        if not all_gaps:
            st.success("No gaps! All sources are fully matched.")
        else:
            gap_by_source = {}
            for g in all_gaps:
                gap_by_source[g["Gap In"]] = gap_by_source.get(g["Gap In"], 0) + 1

            gap_cols = st.columns(len(gap_by_source))
            for i, (src, cnt) in enumerate(sorted(gap_by_source.items(), key=lambda x: -x[1])):
                with gap_cols[i]:
                    st.metric(src, f"{cnt} gaps")

            st.divider()
            source_filter = st.selectbox("Filter by source", ["All"] + sorted(gap_by_source.keys()))
            filtered = all_gaps if source_filter == "All" else [g for g in all_gaps if g["Gap In"] == source_filter]
            st.dataframe(pd.DataFrame(filtered), use_container_width=True, hide_index=True)

            st.divider()
            st.subheader("Action Items")
            for i, g in enumerate(filtered):
                st.markdown(f"{i+1}. **{g['Equipment Class']}** — add to :red[**{g['Gap In']}**] (found in: {g['Found In']})")


# ══════════════════════════════════════════════════════════════════════
# TAB: Connection Map
# ══════════════════════════════════════════════════════════════════════
with tab_visual:
    classes = st.session_state.classes
    if not classes:
        st.info("Run harmonization first.")
    else:
        st.subheader("Connection Map")
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
    st.subheader("Search Equipment Classes")
    query = st.text_input("Equipment name", placeholder="e.g. centrifugal pump...")

    if query and st.session_state.classes:
        names = [c["canonical_name"] for c in st.session_state.classes]
        result = process.extractOne(query, names, scorer=fuzz.token_sort_ratio)
        if result:
            name, score, idx = result
            score = int(round(score))
            c = st.session_state.classes[idx]
            status, _ = classify_match(score)
            color = "green" if score >= 90 else "orange" if score >= 75 else "red"

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
    st.subheader("Batch Process")
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
    st.subheader("System Logs")
    if st.session_state.logs:
        st.code("\n".join(reversed(st.session_state.logs)), language="log")
    else:
        st.info("No logs yet.")

st.divider()
st.caption("Built by KBR AMCDE Team — RDL Data Harmonizer v2.0")
