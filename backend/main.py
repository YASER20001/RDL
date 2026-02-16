"""
KBR RDL Data Harmonizer v2.0 — Streamlit App
Run:  streamlit run main.py
"""

import uuid
import io
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
def fuzzy_match(name, candidates, threshold):
    """Return the single best match above threshold, or None."""
    if not candidates:
        return None
    names = [c["name"] for c in candidates]
    result = process.extractOne(
        name, names, scorer=fuzz.token_sort_ratio, score_cutoff=threshold,
    )
    if result is None:
        return None
    matched_name, score, idx = result
    return {**candidates[idx], "score": int(round(score))}


def classify_match(score):
    """Return (status, confidence) for a given score."""
    if score >= 95:
        return "Exact Match", "High"
    elif score >= 85:
        return "Strong Match", "High"
    elif score >= 75:
        return "Partial Match", "Medium"
    else:
        return "Weak Match", "Low"


def run_harmonization():
    """Harmonize: merge masters into unified rows, match others against them.

    Dual master: Master-1 rows are matched 1:1 against Master-2. Unmatched
    Master-2 rows are appended. Each row has both master IDs/names merged.
    Non-master sources are then matched against the canonical name.
    """
    masters = st.session_state.masters
    files = st.session_state.files
    threshold = st.session_state.match_threshold

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

        # Deduplicate master-1
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

        # Append unmatched + deduplicated master-2 records
        for i, r2 in enumerate(m2_records):
            if i not in m2_used:
                key = r2["name"].strip().lower()
                if key not in seen:
                    seen.add(key)
                    merged.append({
                        "master_entries": {m2: r2},
                        "canonical_name": r2["name"],
                    })

    add_log(f"Master set from {masters}: {len(merged)} unique rows (threshold={threshold}%)")

    # Non-master sources
    other = {k: v["records"] for k, v in files.items() if k not in masters}

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
        for src, candidates in other.items():
            match = fuzzy_match(row["canonical_name"], candidates, threshold)
            if match:
                entry["matches"][src] = match
            else:
                entry["gaps"].append(src)
        harmonized.append(entry)

    add_log(f"Harmonization complete: {len(harmonized)} classes")
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
# Export helper — clean organized Excel
# ──────────────────────────────────────────────────────────────────────
def build_export_df(classes):
    """Build a clean, well-organized DataFrame for export.

    Layout for dual-master (e.g. KBR + CFIHOS):
        # | Status | Confidence | KBR ID | KBR Name | CFIHOS ID | CFIHOS Name |
        Master Match% | ARAMCO ID | ARAMCO Name | ARAMCO Match% | ARAMCO Status | ...
        | Gap Sources | Gap Count

    Gaps are clearly marked as "GAP - NO MATCH" in the Name cells.
    """
    masters = st.session_state.masters
    # Collect ALL non-master sources (even if only in gaps)
    non_master_sources = set()
    for c in classes:
        non_master_sources.update(c["matches"].keys())
        non_master_sources.update(c["gaps"])
    non_master_sources = sorted(non_master_sources)

    rows = []
    for c in classes:
        # --- Overall status ---
        total_other = len(non_master_sources)
        matched_count = len(c["matches"])
        gap_count = len(c["gaps"])

        if gap_count == 0 and (len(masters) == 1 or c.get("cross_score") is not None):
            overall_status = "FULLY MATCHED"
        elif gap_count == total_other:
            overall_status = "NO MATCHES"
        elif gap_count > 0:
            overall_status = "PARTIAL - HAS GAPS"
        else:
            overall_status = "MATCHED"

        # Overall confidence from worst match
        all_scores = [m["score"] for m in c["matches"].values()]
        if c.get("cross_score"):
            all_scores.append(c["cross_score"])
        if all_scores:
            min_score = min(all_scores)
            _, overall_conf = classify_match(min_score)
        else:
            overall_conf = "None"

        row = {
            "#": c["index"],
            "Overall Status": overall_status,
            "Confidence": overall_conf,
        }

        # --- Master columns ---
        for m in masters:
            label = FILE_TYPES.get(m, m)
            me = c["master_entries"].get(m)
            row[f"{label} ID"] = me["id"] if me else ""
            row[f"{label} Name"] = me["name"] if me else "GAP - NOT IN THIS MASTER"

        if len(masters) == 2:
            cs = c.get("cross_score")
            if cs is not None:
                status, conf = classify_match(cs)
                row["Masters Match%"] = cs
                row["Masters Status"] = status
            else:
                row["Masters Match%"] = ""
                row["Masters Status"] = "GAP - NO CROSS-MATCH"

        # --- Non-master columns ---
        for s in non_master_sources:
            label = FILE_TYPES.get(s, s)
            m = c["matches"].get(s)
            if m:
                status, conf = classify_match(m["score"])
                row[f"{label} ID"] = m["id"]
                row[f"{label} Name"] = m["name"]
                row[f"{label} Match%"] = m["score"]
                row[f"{label} Status"] = status
            else:
                row[f"{label} ID"] = ""
                row[f"{label} Name"] = "GAP - NO MATCH"
                row[f"{label} Match%"] = 0
                row[f"{label} Status"] = "GAP"

        # --- Gap summary ---
        row["Gap Sources"] = " | ".join(FILE_TYPES.get(g, g) for g in c["gaps"]) if c["gaps"] else ""
        row["Gap Count"] = gap_count

        rows.append(row)

    return pd.DataFrame(rows)


def build_excel_bytes(classes):
    """Build a styled Excel file with multiple sheets."""
    buf = io.BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Sheet 1: Harmonization Results
        df = build_export_df(classes)
        df.to_excel(writer, sheet_name="Harmonization Results", index=False)

        # Sheet 2: Gap Analysis
        gap_rows = []
        masters = st.session_state.masters
        for c in classes:
            for g in c["gaps"]:
                gap_rows.append({
                    "#": c["index"],
                    "Equipment Class": c["canonical_name"],
                    "Present In": ", ".join(
                        FILE_TYPES.get(m, m) for m in masters if m in c["master_entries"]
                    ) + (
                        ", " + ", ".join(
                            FILE_TYPES.get(s, s) for s in c["matches"]
                        ) if c["matches"] else ""
                    ),
                    "MISSING From": FILE_TYPES.get(g, g),
                    "Action Required": f"Add '{c['canonical_name']}' to {FILE_TYPES.get(g, g)} or confirm exclusion",
                })
        if gap_rows:
            gap_df = pd.DataFrame(gap_rows)
        else:
            gap_df = pd.DataFrame(columns=["#", "Equipment Class", "Present In", "MISSING From", "Action Required"])
        gap_df.to_excel(writer, sheet_name="Gap Analysis", index=False)

        # Sheet 3: Summary Statistics
        total = len(classes)
        non_master_sources = set()
        for c in classes:
            non_master_sources.update(c["matches"].keys())
            non_master_sources.update(c["gaps"])

        summary_rows = [
            {"Metric": "Total Equipment Classes", "Value": total},
            {"Metric": "Masters", "Value": ", ".join(FILE_TYPES.get(m, m) for m in masters)},
            {"Metric": "Match Threshold", "Value": f"{st.session_state.match_threshold}%"},
            {"Metric": "Total Gap Entries", "Value": sum(len(c["gaps"]) for c in classes)},
        ]
        fully = sum(1 for c in classes if not c["gaps"] and (len(masters) == 1 or c.get("cross_score") is not None))
        summary_rows.append({"Metric": "Fully Matched Classes", "Value": f"{fully} / {total}"})

        for src in sorted(non_master_sources):
            cnt = sum(1 for c in classes if src in c["matches"])
            pct = int(round(cnt / total * 100)) if total else 0
            summary_rows.append({
                "Metric": f"{FILE_TYPES.get(src, src)} Match Rate",
                "Value": f"{cnt}/{total} ({pct}%)",
            })

        summary_df = pd.DataFrame(summary_rows)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # Auto-fit column widths (approximate)
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    val = str(cell.value) if cell.value is not None else ""
                    max_len = max(max_len, len(val))
                ws.column_dimensions[col_letter].width = min(max_len + 3, 50)

    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────
# Graphviz connection diagram builder
# ──────────────────────────────────────────────────────────────────────
def build_connection_graph(entry):
    masters = st.session_state.masters
    lines = []
    lines.append("digraph G {")
    lines.append('  rankdir=LR;')
    lines.append('  node [shape=box, style="filled,rounded", fontname="Arial", fontsize=11];')
    lines.append('  edge [fontname="Arial", fontsize=9];')

    canon = entry["canonical_name"].replace('"', '\\"')
    lines.append(f'  center [label="{canon}", fillcolor="#fef3c7", color="#d97706", penwidth=2];')

    node_id = 0
    for m in masters:
        me = entry["master_entries"].get(m)
        if me:
            color = SOURCE_COLORS.get(m, "#6b7280")
            name_esc = me["name"].replace('"', '\\"')
            label = f'{FILE_TYPES[m]}\\n{name_esc}\\nID: {me["id"]}'
            nid = f"m{node_id}"
            lines.append(f'  {nid} [label="{label}", fillcolor="{color}20", color="{color}", fontcolor="#1f2937"];')
            score_label = ""
            if len(masters) == 2 and entry.get("cross_score") is not None:
                score_label = f' [label="{entry["cross_score"]}%", color="{color}", fontcolor="{color}"]'
            elif len(masters) == 1:
                score_label = f' [label="master", color="{color}", fontcolor="{color}"]'
            lines.append(f'  {nid} -> center{score_label};')
            node_id += 1
        else:
            color = "#ef4444"
            nid = f"m{node_id}"
            label = f'{FILE_TYPES[m]}\\nNOT IN THIS MASTER'
            lines.append(f'  {nid} [label="{label}", fillcolor="#fee2e2", color="{color}", fontcolor="{color}", style="filled,rounded,dashed"];')
            lines.append(f'  {nid} -> center [style=dashed, color="{color}"];')
            node_id += 1

    for src, match in entry["matches"].items():
        color = SOURCE_COLORS.get(src, "#6b7280")
        name_esc = match["name"].replace('"', '\\"')
        label = f'{FILE_TYPES.get(src, src)}\\n{name_esc}\\nID: {match["id"]}'
        nid = f"n{node_id}"
        lines.append(f'  {nid} [label="{label}", fillcolor="{color}20", color="{color}", fontcolor="#1f2937"];')
        lines.append(f'  center -> {nid} [label="{match["score"]}%", color="{color}", fontcolor="{color}"];')
        node_id += 1

    for g in entry["gaps"]:
        color = "#ef4444"
        label = f'{FILE_TYPES.get(g, g)}\\nGAP - NO MATCH'
        nid = f"g{node_id}"
        lines.append(f'  {nid} [label="{label}", fillcolor="#fee2e2", color="{color}", fontcolor="{color}", style="filled,rounded,dashed"];')
        lines.append(f'  center -> {nid} [style=dashed, color="{color}"];')
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

    lines = []
    lines.append("digraph G {")
    lines.append('  rankdir=LR;')
    lines.append('  node [shape=box, style="filled,rounded", fontname="Arial", fontsize=10];')
    lines.append('  edge [fontname="Arial", fontsize=8];')
    lines.append('  nodesep=0.3; ranksep=1.5;')

    display = classes[:max_rows]

    for idx, c in enumerate(display):
        for m in masters:
            me = c["master_entries"].get(m)
            if me:
                color = SOURCE_COLORS.get(m, "#6b7280")
                nid = f"r{idx}_{m}"
                name_esc = me["name"][:25].replace('"', '\\"')
                label = f'{me["id"]}\\n{name_esc}'
                lines.append(f'  {nid} [label="{label}", fillcolor="{color}15", color="{color}"];')

        if len(masters) == 2:
            m1, m2 = masters
            if m1 in c["master_entries"] and m2 in c["master_entries"]:
                score = c.get("cross_score", "")
                lbl = f'{score}%' if score else ""
                lines.append(f'  r{idx}_{m1} -> r{idx}_{m2} [label="{lbl}", color="#6b7280", dir=both];')

        for src in all_sources:
            match = c["matches"].get(src)
            color = SOURCE_COLORS.get(src, "#6b7280")
            nid = f"r{idx}_{src}"
            if match:
                name_esc = match["name"][:25].replace('"', '\\"')
                label = f'{match["id"]}\\n{name_esc}'
                lines.append(f'  {nid} [label="{label}", fillcolor="{color}15", color="{color}"];')
                first_m = masters[0] if masters[0] in c["master_entries"] else masters[-1]
                lines.append(f'  r{idx}_{first_m} -> {nid} [label="{match["score"]}%", color="{color}"];')
            else:
                label = f'GAP'
                lines.append(f'  {nid} [label="{label}", fillcolor="#fee2e2", color="#ef4444", style="filled,rounded,dashed"];')
                first_m = masters[0] if masters[0] in c["master_entries"] else masters[-1]
                lines.append(f'  r{idx}_{first_m} -> {nid} [style=dashed, color="#ef4444"];')

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
        "Select **1 or 2** file types as the master source. "
        "The master defines the base set of equipment classes; "
        "all other uploaded files are fuzzy-matched against it."
    )

    master_options = list(FILE_TYPES.keys())
    selected_masters = st.multiselect(
        "Master file(s)",
        options=master_options,
        default=st.session_state.masters or [],
        format_func=lambda k: FILE_TYPES[k],
        max_selections=2,
        help="Pick 1 or 2 systems to use as the master reference.",
    )

    if selected_masters != st.session_state.masters:
        st.session_state.masters = selected_masters
        add_log(f"Master config changed to: {selected_masters}")

    if len(selected_masters) == 0:
        st.warning("Select at least one master to run harmonization.")
    elif len(selected_masters) == 1:
        st.info(f"**Single master**: {FILE_TYPES[selected_masters[0]]} will be the base.")
    else:
        st.info(
            f"**Dual master**: {FILE_TYPES[selected_masters[0]]} + "
            f"{FILE_TYPES[selected_masters[1]]} will be merged row-by-row."
        )

    # Match threshold slider
    st.divider()
    threshold = st.slider(
        "Match Threshold (%)",
        min_value=50, max_value=100, value=st.session_state.match_threshold,
        help="Minimum fuzzy score to count as a match. Below this = GAP. "
             "75% is recommended — rejects weak matches like 'Junction Box' vs 'Function Block'.",
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
            uploaded = st.file_uploader(
                f"{label}{badge}",
                type=["xlsx", "xls", "csv"],
                key=f"upload_{key}",
            )
            if uploaded is not None:
                if key not in st.session_state.files or st.session_state.files[key]["filename"] != uploaded.name:
                    try:
                        reader = READERS[key]
                        records = reader(uploaded)
                        st.session_state.files[key] = {"filename": uploaded.name, "records": records}
                        add_log(f"Uploaded {key}: {uploaded.name} ({len(records)} records)")
                        st.success(f"{len(records)} records loaded")
                    except Exception as e:
                        st.error(f"Error reading file: {e}")
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
                st.error("Select at least one master file first.")
            elif not all(m in st.session_state.files for m in st.session_state.masters):
                missing = [m for m in st.session_state.masters if m not in st.session_state.files]
                st.error(f"Master file(s) not uploaded: {', '.join(missing)}")
            else:
                st.session_state.classes = run_harmonization()
                st.success(f"Harmonized {len(st.session_state.classes)} classes!")
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
            xlsx_bytes = build_excel_bytes(st.session_state.classes)
            st.download_button(
                "Export Excel",
                data=xlsx_bytes,
                file_name="harmonization_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if st.session_state.files:
        st.divider()
        st.caption("**Uploaded files:**")
        status_cols = st.columns(len(st.session_state.files))
        for i, (k, v) in enumerate(st.session_state.files.items()):
            with status_cols[i]:
                is_m = "MASTER " if k in st.session_state.masters else ""
                st.metric(f"{is_m}{FILE_TYPES[k]}", f"{len(v['records'])} records")


# ══════════════════════════════════════════════════════════════════════
# TAB: Dashboard
# ══════════════════════════════════════════════════════════════════════
with tab_dashboard:
    classes = st.session_state.classes

    if not classes:
        st.info("No harmonization results yet. Go to **Upload & Configure** to get started.")
    else:
        masters = st.session_state.masters
        total = len(classes)
        non_master_sources = set()
        gap_count = 0
        source_match_counts = {}
        fully_matched = 0
        for c in classes:
            non_master_sources.update(c["matches"].keys())
            non_master_sources.update(c["gaps"])
            for src in c["matches"]:
                source_match_counts[src] = source_match_counts.get(src, 0) + 1
            gap_count += len(c["gaps"])
            if not c["gaps"] and (len(masters) == 1 or c.get("cross_score") is not None):
                fully_matched += 1

        # Metrics row
        mc = st.columns(3 + len(source_match_counts))
        with mc[0]:
            st.metric("Total Classes", total)
        with mc[1]:
            st.metric("Fully Matched", f"{fully_matched}/{total}")
        with mc[2]:
            st.metric("Total Gaps", gap_count)
        for i, (src, cnt) in enumerate(source_match_counts.items()):
            with mc[3 + i]:
                pct = int(round(cnt / total * 100))
                st.metric(f"{FILE_TYPES.get(src, src)}", f"{pct}% matched")

        st.caption(
            f"**Masters:** {', '.join(FILE_TYPES.get(m, m) for m in masters)} "
            f" |  **Threshold:** {st.session_state.match_threshold}%"
        )
        st.divider()

        # Expandable results
        st.subheader("Harmonized Equipment Classes")

        # Filter
        filter_opt = st.radio(
            "Show", ["All", "Fully Matched only", "Has Gaps only"], horizontal=True,
        )

        for c in classes:
            has_gaps = bool(c["gaps"])
            missing_master = len(masters) == 2 and c.get("cross_score") is None and len(c["master_entries"]) < 2

            if filter_opt == "Fully Matched only" and (has_gaps or missing_master):
                continue
            if filter_opt == "Has Gaps only" and not has_gaps and not missing_master:
                continue

            # Status badge
            if has_gaps or missing_master:
                badge = ":red[HAS GAPS]"
            else:
                badge = ":green[FULLY MATCHED]"

            header_parts = []
            for m in masters:
                me = c["master_entries"].get(m)
                if me:
                    header_parts.append(f"{FILE_TYPES[m]}: {me['name']}")
                else:
                    header_parts.append(f"{FILE_TYPES[m]}: MISSING")
            header = " | ".join(header_parts) if header_parts else c["canonical_name"]

            with st.expander(f"**{c['index']}. {header}** — {badge}"):
                # Master info
                st.markdown("**Masters:**")
                for m in masters:
                    me = c["master_entries"].get(m)
                    if me:
                        st.markdown(f"- **{FILE_TYPES[m]}**: {me['name']} (ID: `{me['id']}`)")
                    else:
                        st.markdown(f"- :red[**{FILE_TYPES[m]}**: GAP - NOT IN THIS MASTER]")

                if len(masters) == 2:
                    cs = c.get("cross_score")
                    if cs is not None:
                        status, _ = classify_match(cs)
                        st.markdown(f"- Masters cross-match: **{cs}%** ({status})")
                    else:
                        st.markdown("- :red[Masters cross-match: **GAP** — only in one master]")

                # Other matches
                if c["matches"]:
                    st.markdown("**Matched sources:**")
                    for src, match in c["matches"].items():
                        score = match["score"]
                        status, conf = classify_match(score)
                        if score >= 90:
                            color = "green"
                        elif score >= 75:
                            color = "orange"
                        else:
                            color = "red"
                        st.markdown(
                            f"- :{color}[**{FILE_TYPES.get(src, src)}**]: "
                            f"{match['name']} (ID: `{match['id']}`) — **{score}%** ({status})"
                        )

                if c["gaps"]:
                    st.markdown("**:red[GAPS — Missing from these sources:]**")
                    for g in c["gaps"]:
                        st.markdown(
                            f"- :red[**{FILE_TYPES.get(g, g)}**] — No match above "
                            f"{st.session_state.match_threshold}% threshold"
                        )

        # Table view
        st.divider()
        st.subheader("Table View")
        export_df = build_export_df(classes)
        st.dataframe(export_df, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════
# TAB: Gap Analysis
# ══════════════════════════════════════════════════════════════════════
with tab_gaps:
    classes = st.session_state.classes
    if not classes:
        st.info("Run harmonization first to see gap analysis.")
    else:
        masters = st.session_state.masters

        # Collect all gaps
        all_gaps = []
        for c in classes:
            # Gap in non-master sources
            for g in c["gaps"]:
                present_in = [FILE_TYPES.get(m, m) for m in masters if m in c["master_entries"]]
                present_in += [FILE_TYPES.get(s, s) for s in c["matches"]]
                all_gaps.append({
                    "Equipment Class": c["canonical_name"],
                    "Present In": ", ".join(present_in),
                    "MISSING From": FILE_TYPES.get(g, g),
                    "Gap Type": "Not in source",
                })
            # Gap in dual-master cross-match
            if len(masters) == 2:
                for m in masters:
                    if m not in c["master_entries"]:
                        present_in = [FILE_TYPES.get(mm, mm) for mm in masters if mm in c["master_entries"]]
                        all_gaps.append({
                            "Equipment Class": c["canonical_name"],
                            "Present In": ", ".join(present_in),
                            "MISSING From": FILE_TYPES.get(m, m),
                            "Gap Type": "Not in master",
                        })

        st.subheader(f"Gap Analysis — {len(all_gaps)} gaps found")

        if not all_gaps:
            st.success("No gaps! All equipment classes are matched across all sources.")
        else:
            # Summary by source
            st.markdown("**Gaps per source:**")
            gap_by_source = {}
            for g in all_gaps:
                src = g["MISSING From"]
                gap_by_source[src] = gap_by_source.get(src, 0) + 1

            gap_cols = st.columns(len(gap_by_source))
            for i, (src, cnt) in enumerate(sorted(gap_by_source.items(), key=lambda x: -x[1])):
                with gap_cols[i]:
                    st.metric(src, f"{cnt} gaps")

            st.divider()

            # Filter by source
            source_filter = st.selectbox(
                "Filter by missing source",
                ["All"] + sorted(gap_by_source.keys()),
            )

            filtered = all_gaps if source_filter == "All" else [g for g in all_gaps if g["MISSING From"] == source_filter]

            gap_df = pd.DataFrame(filtered)
            st.dataframe(gap_df, use_container_width=True, hide_index=True)

            # Actionable list
            st.divider()
            st.subheader("Action Items")
            for i, g in enumerate(filtered):
                st.markdown(
                    f"{i+1}. **{g['Equipment Class']}** — "
                    f"Add to :red[**{g['MISSING From']}**] "
                    f"(currently in: {g['Present In']})"
                )


# ══════════════════════════════════════════════════════════════════════
# TAB: Connection Map
# ══════════════════════════════════════════════════════════════════════
with tab_visual:
    classes = st.session_state.classes
    if not classes:
        st.info("Run harmonization first to see the connection map.")
    else:
        st.subheader("Connection Map")

        view_mode = st.radio(
            "View",
            ["Per-class detail", "Overview (first 20 rows)"],
            horizontal=True,
        )

        if view_mode == "Per-class detail":
            options = [f"{c['index']}. {c['canonical_name']}" for c in classes]
            selected = st.selectbox("Select equipment class", options)
            if selected:
                idx = int(selected.split(".")[0]) - 1
                entry = classes[idx]
                dot = build_connection_graph(entry)
                st.graphviz_chart(dot, use_container_width=True)
                st.caption("Solid lines = matched. Dashed red = GAP (no match). Numbers = match score %.")
        else:
            dot = build_overview_graph(classes, max_rows=20)
            st.graphviz_chart(dot, use_container_width=True)
            if len(classes) > 20:
                st.caption(f"Showing first 20 of {len(classes)} classes.")


# ══════════════════════════════════════════════════════════════════════
# TAB: Search — best match only
# ══════════════════════════════════════════════════════════════════════
with tab_search:
    st.subheader("Search Equipment Classes")
    query = st.text_input("Equipment name", placeholder="e.g. centrifugal pump, heat exchanger...")

    if query and st.session_state.classes:
        names = [c["canonical_name"] for c in st.session_state.classes]
        result = process.extractOne(query, names, scorer=fuzz.token_sort_ratio)

        if result:
            name, score, idx = result
            score = int(round(score))
            c = st.session_state.classes[idx]
            status, conf = classify_match(score)

            if score >= 90:
                color = "green"
            elif score >= 75:
                color = "orange"
            else:
                color = "red"

            st.markdown(f"### :{color}[Best match: **{name}** — {score}% ({status})]")

            masters = st.session_state.masters
            col_l, col_r = st.columns(2)
            with col_l:
                st.markdown("**Master entries:**")
                for m in masters:
                    me = c["master_entries"].get(m)
                    if me:
                        st.markdown(f"- **{FILE_TYPES[m]}**: {me['name']} (`{me['id']}`)")
                    else:
                        st.markdown(f"- :red[**{FILE_TYPES[m]}**: GAP]")

            with col_r:
                if c["matches"]:
                    st.markdown("**Other matches:**")
                    for src, match in c["matches"].items():
                        st.markdown(f"- **{FILE_TYPES.get(src, src)}**: {match['name']} — {match['score']}%")
                if c["gaps"]:
                    st.markdown("**:red[Gaps:]** " + ", ".join(FILE_TYPES.get(g, g) for g in c["gaps"]))

            st.divider()
            dot = build_connection_graph(c)
            st.graphviz_chart(dot, use_container_width=True)
        else:
            st.warning("No match found.")
    elif query:
        st.warning("Run harmonization first to search.")


# ══════════════════════════════════════════════════════════════════════
# TAB: Batch
# ══════════════════════════════════════════════════════════════════════
with tab_batch:
    st.subheader("Batch Process")
    st.caption("Enter one equipment name per line, or upload a CSV with a `name` column.")

    batch_input = st.text_area("Equipment names", height=150, placeholder="Centrifugal Pump\nHeat Exchanger\nPressure Vessel")
    batch_file = st.file_uploader("Or upload CSV", type=["csv"], key="batch_csv")

    if st.button("Process Batch", type="primary"):
        items = []
        if batch_file:
            bdf = pd.read_csv(batch_file)
            col = "name" if "name" in bdf.columns else bdf.columns[0]
            items = bdf[col].dropna().astype(str).tolist()
        elif batch_input.strip():
            items = [line.strip() for line in batch_input.strip().split("\n") if line.strip()]

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
                    score = int(round(score))
                    c = st.session_state.classes[idx]
                    status, conf = classify_match(score)
                    results.append({
                        "Input": item,
                        "Best Match": name,
                        "Score %": score,
                        "Status": status,
                        "Gaps": ", ".join(FILE_TYPES.get(g, g) for g in c["gaps"]) or "None",
                    })
                else:
                    results.append({"Input": item, "Best Match": "—", "Score %": 0, "Status": "No Match", "Gaps": "—"})

            result_df = pd.DataFrame(results)
            st.dataframe(result_df, use_container_width=True, hide_index=True)

            csv_out = result_df.to_csv(index=False).encode()
            st.download_button("Download batch results", data=csv_out, file_name="batch_results.csv", mime="text/csv")


# ══════════════════════════════════════════════════════════════════════
# TAB: Logs
# ══════════════════════════════════════════════════════════════════════
with tab_logs:
    st.subheader("System Logs")
    if st.session_state.logs:
        log_text = "\n".join(reversed(st.session_state.logs))
        st.code(log_text, language="log")
    else:
        st.info("No logs yet.")

# ── Footer ────────────────────────────────────────────────────────────
st.divider()
st.caption("Built by KBR AMCDE Team — RDL Data Harmonizer v2.0")
