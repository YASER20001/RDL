"""
KBR RDL Data Harmonizer v2.0 - Backend API
Flexible master file selection: users choose 1 or 2 files as the master source.
"""

import os
import uuid
import logging
from datetime import datetime
from typing import Optional

import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from rapidfuzz import fuzz, process
from pydantic import BaseModel

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("harmonizer")

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
app = FastAPI(title="KBR RDL Data Harmonizer", version="2.0.0")

cors_origins = os.getenv("CORS_ORIGINS", "*").split(",")
app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
MAX_UPLOAD_SIZE = int(os.getenv("MAX_UPLOAD_SIZE_MB", "50")) * 1024 * 1024

# ---------------------------------------------------------------------------
# In-memory state
# ---------------------------------------------------------------------------
FILES: dict = {}          # type -> {path, filename, df_dict}
CLASSES: list = []        # harmonized classes
LOGS: list = []           # system logs
MASTER_CONFIG: dict = {   # which files act as master
    "masters": [],        # list of file type keys (e.g. ["aramco", "cfihos"])
}

FILE_TYPES = ["aramco", "cfihos", "kbr", "ltc", "sa_doc"]


def add_log(msg: str, level: str = "info"):
    entry = {"ts": datetime.utcnow().isoformat(), "level": level, "msg": msg}
    LOGS.append(entry)
    getattr(logger, level, logger.info)(msg)


# ---------------------------------------------------------------------------
# Pydantic models
# ---------------------------------------------------------------------------
class SearchRequest(BaseModel):
    query: str
    limit: int = 20


class BatchRequest(BaseModel):
    items: list[str]


class MasterConfigRequest(BaseModel):
    """Configure which uploaded file(s) serve as the master source.
    Accepts 1 or 2 file type keys from: aramco, cfihos, kbr, ltc, sa_doc
    """
    masters: list[str]


# ---------------------------------------------------------------------------
# File readers – each returns a list[dict] with at least {id, name}
# ---------------------------------------------------------------------------

def _read_aramco(path: str) -> list[dict]:
    """Read Aramco 9COM .xlsx – sheet 'ISM Functional Classes'."""
    try:
        df = pd.read_excel(path, sheet_name="ISM Functional Classes")
    except Exception:
        df = pd.read_excel(path, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("Id", row.get("id", ""))),
            "name": str(row.get("Name", row.get("name", ""))),
            "cfihos_ref": str(row.get("nmcltr:CFIHOS_1.5", "")),
            "source": "aramco",
        })
    return [r for r in records if r["name"].strip()]


def _read_cfihos(path: str) -> list[dict]:
    """Read CFIHOS .xlsx – sheet 'equipment class'."""
    try:
        df = pd.read_excel(path, sheet_name="equipment class")
    except Exception:
        df = pd.read_excel(path, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("CFIHOS unique id", row.get("id", ""))),
            "name": str(row.get("equipment class name", row.get("name", ""))),
            "source": "cfihos",
        })
    return [r for r in records if r["name"].strip()]


def _read_kbr(path: str) -> list[dict]:
    """Read KBR FEED .xlsx."""
    df = pd.read_excel(path, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("Class Id", row.get("id", ""))),
            "name": str(row.get("Class Name (855)", row.get("name", ""))),
            "discipline": str(row.get("Discipline", "")),
            "source": "kbr",
        })
    return [r for r in records if r["name"].strip()]


def _read_ltc(path: str) -> list[dict]:
    """Read LTC .xlsx – sheet 'ISM Functional Classes'."""
    try:
        df = pd.read_excel(path, sheet_name="ISM Functional Classes")
    except Exception:
        df = pd.read_excel(path, sheet_name=0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "id": str(row.get("Id", row.get("id", ""))),
            "name": str(row.get("Name", row.get("name", ""))),
            "source": "ltc",
        })
    return [r for r in records if r["name"].strip()]


def _read_sa_doc(path: str) -> list[dict]:
    """Read SA DOC .xlsx – sheet 'SA_DOC_attributes'."""
    try:
        df = pd.read_excel(path, sheet_name="SA_DOC_attributes")
    except Exception:
        df = pd.read_excel(path, sheet_name=0)
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
    "aramco": _read_aramco,
    "cfihos": _read_cfihos,
    "kbr": _read_kbr,
    "ltc": _read_ltc,
    "sa_doc": _read_sa_doc,
}


# ---------------------------------------------------------------------------
# Harmonization engine
# ---------------------------------------------------------------------------

def _fuzzy_match(name: str, candidates: list[dict], threshold: int = 60) -> Optional[dict]:
    """Return the best fuzzy match from candidates or None."""
    if not candidates:
        return None
    names = [c["name"] for c in candidates]
    result = process.extractOne(name, names, scorer=fuzz.token_sort_ratio, score_cutoff=threshold)
    if result is None:
        return None
    matched_name, score, idx = result
    return {**candidates[idx], "score": round(score, 1)}


def run_harmonization() -> list[dict]:
    """Run harmonization using the configured master(s).

    If one master is selected, that file's classes are used as the base and
    all other uploaded files are matched against it.

    If two masters are selected, classes from both masters are merged
    (union) to form the base, and the remaining files are matched against
    this combined set.
    """
    masters = MASTER_CONFIG.get("masters", [])
    if not masters:
        raise ValueError("No master file configured. Set at least one master.")

    # Validate masters are uploaded
    for m in masters:
        if m not in FILES:
            raise ValueError(f"Master '{m}' is configured but not uploaded.")

    # Build master class list
    master_classes: list[dict] = []
    seen_names: set = set()
    for m in masters:
        reader = READERS[m]
        records = reader(FILES[m]["path"])
        for r in records:
            key = r["name"].strip().lower()
            if key not in seen_names:
                seen_names.add(key)
                master_classes.append(r)

    add_log(f"Master set built from {masters} – {len(master_classes)} unique classes")

    # Build lookup tables for non-master files
    other_sources: dict[str, list[dict]] = {}
    for ftype, finfo in FILES.items():
        if ftype not in masters:
            reader = READERS[ftype]
            other_sources[ftype] = reader(finfo["path"])

    # Harmonize
    harmonized = []
    for idx, mc in enumerate(master_classes):
        entry: dict = {
            "uid": str(uuid.uuid4()),
            "index": idx + 1,
            "master_id": mc["id"],
            "master_name": mc["name"],
            "master_source": mc.get("source", masters[0]),
            "matches": {},
            "gaps": [],
        }
        for src, candidates in other_sources.items():
            match = _fuzzy_match(mc["name"], candidates)
            if match:
                entry["matches"][src] = match
            else:
                entry["gaps"].append(src)

        # If two masters, also cross-match between them
        if len(masters) == 2:
            other_master = [m for m in masters if m != mc.get("source")][0] if mc.get("source") in masters else None
            if other_master and other_master in FILES:
                other_master_records = READERS[other_master](FILES[other_master]["path"])
                match = _fuzzy_match(mc["name"], other_master_records)
                if match:
                    entry["matches"][other_master] = match

        harmonized.append(entry)

    add_log(f"Harmonization complete: {len(harmonized)} classes processed")
    return harmonized


# ---------------------------------------------------------------------------
# Demo data generator
# ---------------------------------------------------------------------------

def _generate_demo_data():
    """Generate small in-memory demo data for all five sources."""
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
    return demo


# ---------------------------------------------------------------------------
# API Routes
# ---------------------------------------------------------------------------

@app.get("/api/health")
def health():
    return {"status": "ok", "version": "2.0.0"}


# ---- Master configuration -------------------------------------------------

@app.get("/api/master-config")
def get_master_config():
    """Return the current master file configuration."""
    return {
        "masters": MASTER_CONFIG["masters"],
        "available": list(FILES.keys()),
        "all_types": FILE_TYPES,
    }


@app.post("/api/master-config")
def set_master_config(req: MasterConfigRequest):
    """Set which uploaded file(s) act as the master source (1 or 2)."""
    if len(req.masters) < 1 or len(req.masters) > 2:
        raise HTTPException(400, "Select exactly 1 or 2 master files.")
    for m in req.masters:
        if m not in FILE_TYPES:
            raise HTTPException(400, f"Unknown file type: {m}")
    MASTER_CONFIG["masters"] = req.masters
    add_log(f"Master config updated: {req.masters}")
    return {"masters": MASTER_CONFIG["masters"]}


# ---- File upload -----------------------------------------------------------

@app.post("/api/upload/{file_type}")
async def upload_file(file_type: str, file: UploadFile = File(...)):
    if file_type not in FILE_TYPES:
        raise HTTPException(400, f"Unknown file type: {file_type}. Use one of {FILE_TYPES}")

    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in (".xlsx", ".xls", ".csv"):
        raise HTTPException(400, "Only .xlsx, .xls, and .csv files are accepted.")

    content = await file.read()
    if len(content) > MAX_UPLOAD_SIZE:
        raise HTTPException(413, f"File exceeds {MAX_UPLOAD_SIZE // (1024*1024)}MB limit.")

    save_path = os.path.join(UPLOAD_DIR, f"{file_type}_{uuid.uuid4().hex}{ext}")
    with open(save_path, "wb") as f:
        f.write(content)

    # Quick validation – try to read it
    try:
        reader = READERS[file_type]
        records = reader(save_path)
    except Exception as e:
        os.remove(save_path)
        raise HTTPException(422, f"Could not parse file: {e}")

    FILES[file_type] = {"path": save_path, "filename": file.filename, "count": len(records)}
    add_log(f"Uploaded {file_type}: {file.filename} ({len(records)} records)")

    # Auto-set master if none configured yet
    if not MASTER_CONFIG["masters"]:
        MASTER_CONFIG["masters"] = [file_type]
        add_log(f"Auto-set master to [{file_type}] (first upload)")

    return {
        "file_type": file_type,
        "filename": file.filename,
        "records": len(records),
        "masters": MASTER_CONFIG["masters"],
    }


# ---- Harmonize -------------------------------------------------------------

@app.post("/api/harmonize")
def harmonize():
    if not FILES:
        raise HTTPException(400, "No files uploaded yet.")
    if not MASTER_CONFIG["masters"]:
        raise HTTPException(400, "No master file configured.")
    try:
        global CLASSES
        CLASSES = run_harmonization()
    except ValueError as e:
        raise HTTPException(400, str(e))
    return {"total": len(CLASSES), "masters": MASTER_CONFIG["masters"]}


# ---- Classes ---------------------------------------------------------------

@app.get("/api/classes")
def list_classes(skip: int = 0, limit: int = 100):
    return {"total": len(CLASSES), "items": CLASSES[skip : skip + limit]}


@app.get("/api/classes/{uid}")
def get_class(uid: str):
    for c in CLASSES:
        if c["uid"] == uid:
            return c
    raise HTTPException(404, "Class not found")


# ---- Search ----------------------------------------------------------------

@app.post("/api/search")
def search_classes(req: SearchRequest):
    if not CLASSES:
        return {"results": []}
    names = [c["master_name"] for c in CLASSES]
    matches = process.extract(req.query, names, scorer=fuzz.token_sort_ratio, limit=req.limit)
    results = []
    for name, score, idx in matches:
        entry = {**CLASSES[idx], "search_score": round(score, 1)}
        results.append(entry)
    return {"results": results}


# ---- Batch -----------------------------------------------------------------

@app.post("/api/batch")
def batch_process(req: BatchRequest):
    if not CLASSES:
        raise HTTPException(400, "Run harmonization first.")
    results = []
    names = [c["master_name"] for c in CLASSES]
    for item in req.items:
        match = process.extractOne(item, names, scorer=fuzz.token_sort_ratio)
        if match:
            name, score, idx = match
            results.append({"input": item, "match": CLASSES[idx], "score": round(score, 1)})
        else:
            results.append({"input": item, "match": None, "score": 0})
    return {"results": results}


# ---- Stats -----------------------------------------------------------------

@app.get("/api/stats")
def get_stats():
    total = len(CLASSES)
    if total == 0:
        return {"total": 0, "matches": {}, "gaps": 0, "masters": MASTER_CONFIG["masters"]}

    source_match_counts: dict[str, int] = {}
    gap_count = 0
    for c in CLASSES:
        for src in c["matches"]:
            source_match_counts[src] = source_match_counts.get(src, 0) + 1
        gap_count += len(c["gaps"])

    match_rates = {src: round(cnt / total * 100, 1) for src, cnt in source_match_counts.items()}

    return {
        "total": total,
        "matches": match_rates,
        "gap_count": gap_count,
        "masters": MASTER_CONFIG["masters"],
        "files_uploaded": list(FILES.keys()),
    }


# ---- Export ----------------------------------------------------------------

@app.get("/api/export/csv")
def export_csv():
    if not CLASSES:
        raise HTTPException(400, "No harmonization data to export.")

    import io, csv
    buf = io.StringIO()
    writer = csv.writer(buf)

    # Header
    sources = set()
    for c in CLASSES:
        sources.update(c["matches"].keys())
        sources.update(c["gaps"])
    sources = sorted(sources)

    header = ["#", "Master ID", "Master Name", "Master Source"]
    for s in sources:
        header += [f"{s}_id", f"{s}_name", f"{s}_score"]
    header.append("Gaps")
    writer.writerow(header)

    for c in CLASSES:
        row = [c["index"], c["master_id"], c["master_name"], c["master_source"]]
        for s in sources:
            m = c["matches"].get(s)
            if m:
                row += [m.get("id", ""), m.get("name", ""), m.get("score", "")]
            else:
                row += ["", "", ""]
        row.append(", ".join(c["gaps"]))
        writer.writerow(row)

    buf.seek(0)
    return StreamingResponse(
        iter([buf.getvalue()]),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=harmonization_export.csv"},
    )


# ---- Demo ------------------------------------------------------------------

@app.post("/api/load-demo")
def load_demo(masters: Optional[list[str]] = Query(default=None)):
    """Load demo data. Optionally specify master(s) via query param, e.g.
    ?masters=aramco&masters=cfihos  (two masters)
    ?masters=kbr                   (one master)
    Defaults to ["aramco"].
    """
    demo = _generate_demo_data()

    # Store demo data as virtual files
    for ftype, records in demo.items():
        # Write a small xlsx so readers work
        df = pd.DataFrame(records)
        path = os.path.join(UPLOAD_DIR, f"demo_{ftype}.xlsx")
        df.to_excel(path, index=False)
        FILES[ftype] = {"path": path, "filename": f"demo_{ftype}.xlsx", "count": len(records)}

    chosen_masters = masters or ["aramco"]
    for m in chosen_masters:
        if m not in FILE_TYPES:
            raise HTTPException(400, f"Unknown master type: {m}")
    if len(chosen_masters) > 2:
        raise HTTPException(400, "Maximum 2 masters allowed.")

    MASTER_CONFIG["masters"] = chosen_masters
    add_log(f"Demo data loaded. Masters: {chosen_masters}")

    global CLASSES
    CLASSES = run_harmonization()

    return {
        "message": "Demo data loaded",
        "masters": chosen_masters,
        "total_classes": len(CLASSES),
        "files": list(FILES.keys()),
    }


# ---- Logs ------------------------------------------------------------------

@app.get("/api/logs")
def get_logs():
    return {"logs": LOGS[-100:]}


# ---- Documents (SA DOC) ---------------------------------------------------

@app.get("/api/documents")
def get_documents():
    if "sa_doc" not in FILES:
        return {"documents": []}
    records = READERS["sa_doc"](FILES["sa_doc"]["path"])
    return {"documents": records}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("main:app", host=host, port=port, reload=True)
