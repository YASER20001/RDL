"""
KBR RDL Data Harmonizer — Core Engine
Framework-agnostic business logic: file readers, fuzzy matching, harmonization, demo data.
"""

import re
import uuid
import io
import pandas as pd
from rapidfuzz import fuzz, process

# ──────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────
FILE_TYPES = {
    "aramco": "Saudi Aramco 9COM",
    "cfihos": "CFIHOS Standard",
    "kbr": "KBR FEED",
    "ltc": "LTC Contractor",
    "sa_doc": "SA Document",
}

SOURCE_COLORS = {
    "aramco": "#16a34a",
    "cfihos": "#2563eb",
    "kbr": "#dc2626",
    "ltc": "#9333ea",
    "sa_doc": "#d97706",
}

# Industry-standard abbreviation dictionary for smarter matching
ABBREVIATIONS = {
    "hx": "heat exchanger",
    "ht": "heat",
    "xchg": "exchanger",
    "vlv": "valve",
    "cmp": "compressor",
    "comp": "compressor",
    "pmp": "pump",
    "gen": "generator",
    "xfmr": "transformer",
    "sep": "separator",
    "tnk": "tank",
    "tk": "tank",
    "vsl": "vessel",
    "mtr": "motor",
    "drv": "driver",
    "ctrl": "control",
    "inst": "instrument",
    "elec": "electric",
    "mech": "mechanical",
    "recip": "reciprocating",
    "centrif": "centrifugal",
    "atm": "atmospheric",
    "ss": "stainless steel",
}


def _normalize_abbreviations(text):
    """Expand known industry abbreviations in text for better matching."""
    words = text.lower().split()
    expanded = []
    for w in words:
        clean = re.sub(r'[^a-z0-9]', '', w)
        if clean in ABBREVIATIONS:
            expanded.append(ABBREVIATIONS[clean])
        else:
            expanded.append(w)
    return " ".join(expanded)


# ──────────────────────────────────────────────────────────────────────
# File readers — use df.to_dict('records') for performance
# ──────────────────────────────────────────────────────────────────────
def _safe_str(val):
    """Convert value to string, treating NaN/None as empty string."""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


def read_aramco(file) -> list[dict]:
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
        clean_name = re.sub(r'\[.*?\]\s*', '', raw_name).strip()
        records.append({
            "id": _safe_str(row.get("Id", row.get("id", ""))),
            "name": clean_name if clean_name else raw_name,
            "source": "ltc",
        })
    return records


def read_sa_doc(file) -> list[dict]:
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
# Fuzzy matching engine
# ──────────────────────────────────────────────────────────────────────
def _tokenize(text):
    words = set(re.findall(r'[a-z]{2,}', text.lower()))
    noise = {"the", "and", "for", "with", "from", "that", "this", "its",
             "type", "class", "system", "item", "general", "other", "misc"}
    return words - noise


def _has_word_overlap(name_a, name_b):
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

    # Try matching with both original and abbreviation-expanded forms
    name_expanded = _normalize_abbreviations(name)
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
            # Map back to original index
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


# ──────────────────────────────────────────────────────────────────────
# Harmonization
# ──────────────────────────────────────────────────────────────────────
def run_harmonization(masters, files, threshold):
    """Run full harmonization. Returns (harmonized, reverse_gaps, logs)."""
    logs = []

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

    logs.append(f"Master reference built from {masters}: {len(merged)} unique classes (threshold={threshold}%)")

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

    logs.append(f"Harmonization complete: {len(harmonized)} classes, compared against {list(non_masters.keys())}")

    # Step 3: Compute reverse gaps
    reverse_gaps = {}
    for src, candidates in non_masters.items():
        matched_ids = set()
        for entry in harmonized:
            m = entry["matches"].get(src)
            if m:
                matched_ids.add((m["id"], m["name"]))
        unmatched = []
        for rec in candidates:
            if (rec["id"], rec["name"]) not in matched_ids:
                unmatched.append(rec)
        if unmatched:
            reverse_gaps[src] = unmatched
    if reverse_gaps:
        parts = [f"{FILE_TYPES.get(s, s)}: {len(recs)} extra" for s, recs in reverse_gaps.items()]
        logs.append(f"Reverse gaps (extra classes not in master): {', '.join(parts)}")

    return harmonized, reverse_gaps, logs


# ──────────────────────────────────────────────────────────────────────
# Demo data
# ──────────────────────────────────────────────────────────────────────
def get_demo_data():
    """Return demo data as a files dict (same structure as parsed uploads)."""
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
    files = {}
    for key, records in demo.items():
        files[key] = {"filename": f"demo_{key}.xlsx", "records": records}
    return files


# ──────────────────────────────────────────────────────────────────────
# CFIHOS lookup
# ──────────────────────────────────────────────────────────────────────
def lookup_cfihos(name, cfihos_records, threshold):
    """Fuzzy-match a name against CFIHOS records. Returns (id, name, match_label)."""
    if not cfihos_records:
        return "", "", ""
    match = fuzzy_match(name, cfihos_records, threshold)
    if match:
        score = match["score"]
        label = "Exact" if score == 100 else f"{score}%"
        return match["id"], match["name"], label
    return "", "", ""


# ──────────────────────────────────────────────────────────────────────
# Export helpers
# ──────────────────────────────────────────────────────────────────────
def build_export_df(classes, masters):
    """Build export DataFrame for harmonization results."""
    non_master_sources = set()
    for c in classes:
        non_master_sources.update(c["matches"].keys())
        non_master_sources.update(c["gaps"])
    non_master_sources = sorted(non_master_sources)

    rows = []
    for c in classes:
        gap_count = len(c["gaps"])
        row = {"#": c["index"]}

        for m in masters:
            label = FILE_TYPES.get(m, m)
            me = c["master_entries"].get(m)
            row[f"Master ({label}) ID"] = me["id"] if me else ""
            row[f"Master ({label}) Name"] = me["name"] if me else ""

        if len(masters) == 2:
            cs = c.get("cross_score")
            row["Masters Cross-Match%"] = cs if cs is not None else ""

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

        if gap_count == 0:
            row["Gap In"] = "No Gaps"
        else:
            row["Gap In"] = " | ".join(FILE_TYPES.get(g, g) for g in c["gaps"])
        row["Gap Count"] = gap_count
        rows.append(row)

    return pd.DataFrame(rows)


def search_classes(query, classes):
    """Search for equipment class by name. Returns (matched_class, score) or (None, 0)."""
    if not classes or not query:
        return None, 0

    search_names = []
    search_idx_map = []
    for ci, c in enumerate(classes):
        search_names.append(c["canonical_name"])
        search_idx_map.append(ci)
        if "/" in c["canonical_name"]:
            for part in [p.strip() for p in c["canonical_name"].split("/") if p.strip()]:
                search_names.append(part)
                search_idx_map.append(ci)

    best_result = None
    best_score = 0
    query_expanded = _normalize_abbreviations(query)
    expanded_search = [_normalize_abbreviations(n) for n in search_names]

    for scorer in (fuzz.token_set_ratio, fuzz.token_sort_ratio):
        # Match original query against original names
        result = process.extractOne(query, search_names, scorer=scorer)
        if result and result[1] > best_score:
            best_result = result
            best_score = result[1]
        # Match expanded query against expanded names
        result_exp = process.extractOne(query_expanded, expanded_search, scorer=scorer)
        if result_exp and result_exp[1] > best_score:
            best_result = (search_names[result_exp[2]], result_exp[1], result_exp[2])
            best_score = result_exp[1]

    if best_result:
        matched_name, score, exp_idx = best_result
        orig_idx = search_idx_map[exp_idx]
        return classes[orig_idx], int(round(score))
    return None, 0


def batch_process(items, classes):
    """Process a list of equipment names. Returns list of result dicts."""
    # Build expanded name list with compound name support (consistent with fuzzy_match)
    search_names = []
    search_idx_map = []
    for ci, c in enumerate(classes):
        search_names.append(c["canonical_name"])
        search_idx_map.append(ci)
        if "/" in c["canonical_name"]:
            for part in [p.strip() for p in c["canonical_name"].split("/") if p.strip()]:
                search_names.append(part)
                search_idx_map.append(ci)

    # Also build abbreviation-expanded versions
    expanded_search = [_normalize_abbreviations(n) for n in search_names]

    results = []
    for item in items:
        best_score = 0
        best_idx = None
        best_name = None

        item_expanded = _normalize_abbreviations(item)

        for scorer in (fuzz.token_sort_ratio, fuzz.token_set_ratio):
            match = process.extractOne(item, search_names, scorer=scorer)
            if match and match[1] > best_score:
                best_name, best_score, best_exp_idx = match
                best_idx = search_idx_map[best_exp_idx]
            # Try expanded
            match_exp = process.extractOne(item_expanded, expanded_search, scorer=scorer)
            if match_exp and match_exp[1] > best_score:
                best_name = search_names[match_exp[2]]
                best_score = match_exp[1]
                best_idx = search_idx_map[match_exp[2]]

        if best_idx is not None:
            c = classes[best_idx]
            score = int(round(best_score))
            status, _ = classify_match(score)
            gaps = ", ".join(FILE_TYPES.get(g, g) for g in c["gaps"]) or "No Gaps"
            results.append({
                "Input": item,
                "Best Match": c["canonical_name"],
                "Score%": score,
                "Status": status,
                "Gap In": gaps,
            })
        else:
            results.append({
                "Input": item,
                "Best Match": "-",
                "Score%": 0,
                "Status": "No Match",
                "Gap In": "-",
            })
    return results
