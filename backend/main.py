"""
KBR RDL Data Harmonizer v2.0 — Streamlit App
Run:  streamlit run main.py
"""

import uuid
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
    "files": {},           # key -> {filename, records: list[dict]}
    "masters": [],         # list of selected master keys
    "classes": [],         # harmonized output
    "logs": [],            # log entries
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
def fuzzy_match(name, candidates, threshold=60):
    if not candidates:
        return None
    names = [c["name"] for c in candidates]
    result = process.extractOne(name, names, scorer=fuzz.token_sort_ratio, score_cutoff=threshold)
    if result is None:
        return None
    matched_name, score, idx = result
    return {**candidates[idx], "score": round(score, 1)}


def run_harmonization():
    masters = st.session_state.masters
    files = st.session_state.files

    # Build master class list (union of 1 or 2 masters)
    master_classes = []
    seen = set()
    for m in masters:
        for r in files[m]["records"]:
            key = r["name"].strip().lower()
            if key not in seen:
                seen.add(key)
                master_classes.append(r)

    add_log(f"Master set from {masters}: {len(master_classes)} unique classes")

    # Non-master sources
    other = {k: v["records"] for k, v in files.items() if k not in masters}

    harmonized = []
    for idx, mc in enumerate(master_classes):
        entry = {
            "uid": str(uuid.uuid4()),
            "index": idx + 1,
            "master_id": mc["id"],
            "master_name": mc["name"],
            "master_source": mc.get("source", masters[0]),
            "matches": {},
            "gaps": [],
        }
        # Match against non-master files
        for src, candidates in other.items():
            match = fuzzy_match(mc["name"], candidates)
            if match:
                entry["matches"][src] = match
            else:
                entry["gaps"].append(src)

        # Cross-match between two masters
        if len(masters) == 2:
            other_master = [m for m in masters if m != mc.get("source")]
            if other_master and other_master[0] in files:
                match = fuzzy_match(mc["name"], files[other_master[0]]["records"])
                if match:
                    entry["matches"][other_master[0]] = match

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
# CSV export helper
# ──────────────────────────────────────────────────────────────────────
def build_export_csv(classes):
    sources = set()
    for c in classes:
        sources.update(c["matches"].keys())
        sources.update(c["gaps"])
    sources = sorted(sources)

    rows = []
    for c in classes:
        row = {
            "#": c["index"],
            "Master ID": c["master_id"],
            "Master Name": c["master_name"],
            "Master Source": c["master_source"],
        }
        for s in sources:
            m = c["matches"].get(s)
            row[f"{s}_id"] = m["id"] if m else ""
            row[f"{s}_name"] = m["name"] if m else ""
            row[f"{s}_score"] = m["score"] if m else ""
        row["Gaps"] = ", ".join(c["gaps"])
        rows.append(row)

    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════

# ── Header ────────────────────────────────────────────────────────────
st.markdown("""
<div style="background: linear-gradient(90deg, #b91c1c, #7f1d1d); padding: 1.5rem 2rem; border-radius: 0.75rem; margin-bottom: 1rem;">
    <h1 style="color: white; margin: 0; font-size: 1.8rem;">KBR RDL Data Harmonizer</h1>
    <p style="color: #fca5a5; margin: 0; font-size: 0.9rem;">Equipment Class Harmonization v2.0</p>
</div>
""", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────
tab_upload, tab_dashboard, tab_search, tab_batch, tab_logs = st.tabs([
    "Upload & Configure",
    "Dashboard",
    "Search",
    "Batch Process",
    "Logs",
])

# ══════════════════════════════════════════════════════════════════════
# TAB: Upload & Configure
# ══════════════════════════════════════════════════════════════════════
with tab_upload:

    # ── Master selection ──────────────────────────────────────────────
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
            f"{FILE_TYPES[selected_masters[1]]} will be merged as the base."
        )

    st.divider()

    # ── File uploads ──────────────────────────────────────────────────
    st.subheader("Step 2 — Upload Data Files")
    cols = st.columns(3)
    for i, (key, label) in enumerate(FILE_TYPES.items()):
        with cols[i % 3]:
            is_master = key in st.session_state.masters
            badge = " ⭐ MASTER" if is_master else ""
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

    # ── Actions ───────────────────────────────────────────────────────
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
            st.success("Demo loaded and harmonized!")
            st.rerun()

    with col3:
        if st.session_state.classes:
            csv_df = build_export_csv(st.session_state.classes)
            csv_bytes = csv_df.to_csv(index=False).encode()
            st.download_button(
                "Export CSV",
                data=csv_bytes,
                file_name="harmonization_export.csv",
                mime="text/csv",
                use_container_width=True,
            )

    # ── Status summary ────────────────────────────────────────────────
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
        # ── Metrics row ───────────────────────────────────────────────
        sources = set()
        gap_count = 0
        source_match_counts = {}
        for c in classes:
            sources.update(c["matches"].keys())
            sources.update(c["gaps"])
            for src in c["matches"]:
                source_match_counts[src] = source_match_counts.get(src, 0) + 1
            gap_count += len(c["gaps"])

        total = len(classes)
        metric_cols = st.columns(2 + len(source_match_counts))
        with metric_cols[0]:
            st.metric("Total Classes", total)
        with metric_cols[1]:
            st.metric("Gaps", gap_count)
        for i, (src, cnt) in enumerate(source_match_counts.items()):
            with metric_cols[2 + i]:
                pct = round(cnt / total * 100, 1)
                st.metric(f"{src.upper()} Match", f"{pct}%")

        st.caption(f"**Master:** {', '.join(st.session_state.masters)}")
        st.divider()

        # ── Results table ─────────────────────────────────────────────
        st.subheader("Harmonized Equipment Classes")

        for c in classes:
            with st.expander(f"**{c['index']}. {c['master_name']}** ({c['master_source']})"):
                if c["matches"]:
                    st.markdown("**Matches:**")
                    for src, m in c["matches"].items():
                        score = m["score"]
                        if score >= 90:
                            color = "green"
                        elif score >= 70:
                            color = "orange"
                        else:
                            color = "red"
                        st.markdown(
                            f"- :{color}[**{src.upper()}**]: {m['name']} "
                            f"(ID: `{m['id']}`) — **{score}%**"
                        )
                if c["gaps"]:
                    st.markdown("**Gaps (no match):**")
                    for g in c["gaps"]:
                        st.markdown(f"- :red[{g.upper()}] — No match found")

        # ── Full table view ───────────────────────────────────────────
        st.divider()
        st.subheader("Table View")
        export_df = build_export_csv(classes)
        st.dataframe(export_df, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════
# TAB: Search
# ══════════════════════════════════════════════════════════════════════
with tab_search:
    st.subheader("Search Equipment Classes")
    query = st.text_input("Equipment name", placeholder="e.g. centrifugal pump, heat exchanger...")

    if query and st.session_state.classes:
        names = [c["master_name"] for c in st.session_state.classes]
        matches = process.extract(query, names, scorer=fuzz.token_sort_ratio, limit=15)
        for name, score, idx in matches:
            c = st.session_state.classes[idx]
            if score >= 90:
                color = "green"
            elif score >= 70:
                color = "orange"
            else:
                color = "red"
            with st.expander(f":{color}[**{name}**] — {score}% match"):
                st.write(f"**Source:** {c['master_source']}  |  **ID:** `{c['master_id']}`")
                if c["matches"]:
                    for src, m in c["matches"].items():
                        st.write(f"- {src}: {m['name']} ({m['score']}%)")
                if c["gaps"]:
                    st.write(f"**Gaps:** {', '.join(c['gaps'])}")
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
            names = [c["master_name"] for c in st.session_state.classes]
            results = []
            for item in items:
                match = process.extractOne(item, names, scorer=fuzz.token_sort_ratio)
                if match:
                    name, score, idx = match
                    c = st.session_state.classes[idx]
                    results.append({
                        "Input": item,
                        "Best Match": name,
                        "Score": round(score, 1),
                        "Master Source": c["master_source"],
                        "Gaps": ", ".join(c["gaps"]),
                    })
                else:
                    results.append({"Input": item, "Best Match": "—", "Score": 0, "Master Source": "—", "Gaps": "—"})

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
