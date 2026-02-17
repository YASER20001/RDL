# KBR RDL Data Harmonizer v3.0

**Equipment Class Harmonization & Gap Analysis Platform**

Built by the **KBR AMCDE Team**

---

## Table of Contents

1. [The Problem](#the-problem)
2. [The Solution](#the-solution)
3. [Key Benefits](#key-benefits)
4. [Architecture Overview](#architecture-overview)
5. [Supported Data Sources](#supported-data-sources)
6. [Installation & Setup](#installation--setup)
7. [Quick Start](#quick-start)
8. [Feature Guide](#feature-guide)
   - [Upload & Configure](#1-upload--configure)
   - [Dashboard](#2-dashboard)
   - [Gap Analysis](#3-gap-analysis)
   - [Enrichment Suggestions](#4-enrichment-suggestions)
   - [Attributes Explorer](#5-attributes-explorer)
   - [Connection Map](#6-connection-map)
   - [Search](#7-search)
   - [Batch Process](#8-batch-process)
   - [Logs](#9-logs)
9. [Harmonization Engine — How It Works](#harmonization-engine--how-it-works)
   - [Fuzzy Matching Algorithm](#fuzzy-matching-algorithm)
   - [Compound Name Expansion](#compound-name-expansion)
   - [Word Overlap Validation](#word-overlap-validation)
   - [Match Classification](#match-classification)
   - [Bidirectional Gap Analysis](#bidirectional-gap-analysis)
   - [CFIHOS Code Matching](#cfihos-code-matching)
10. [Excel Export Details](#excel-export-details)
    - [Harmonization Report](#harmonization-report-excel)
    - [Enriched Master File](#enriched-master-excel)
11. [Demo Mode](#demo-mode)
12. [File Format Requirements](#file-format-requirements)
13. [Configuration](#configuration)
14. [Tech Stack](#tech-stack)
15. [Project Structure](#project-structure)

---

## The Problem

In large-scale industrial projects (oil & gas, petrochemical, infrastructure), equipment classification data is maintained across **multiple independent systems** by different organizations:

- **Saudi Aramco** maintains ISM Functional Classes (the owner's reference data library)
- **CFIHOS** provides an international equipment class standard
- **KBR** maintains its own FEED (Front-End Engineering Design) class library
- **LTC (Lump-sum Turnkey Contractor)** works with ISM Physical Classes
- **SA Document** contains attribute-level specifications

Each system uses **different naming conventions, different IDs, and different scopes**. For example:

| Aramco | KBR | LTC | CFIHOS |
|--------|-----|-----|--------|
| Shell and Tube Heat Exchanger | Shell & Tube Heat Exchanger | Shell and Tube HX | Shell and Tube Heat Exchanger |
| Control Valve | Control Valve Assembly | Control Valve | Control Valve |
| Electric Motor | Electric Motor Driver | Electric Motor | Electric Motor |

This creates serious problems:

- **No single source of truth** — engineers cannot quickly determine if all systems describe the same equipment
- **Hidden gaps** — some classes exist in one system but not another, leading to procurement and design errors
- **Manual reconciliation** — teams spend weeks in spreadsheets trying to cross-reference thousands of classes
- **Inconsistent scope** — one system may have 146 classes while another has 200, and nobody knows which extras are real additions vs. naming mismatches
- **No attribute-level visibility** — even when class names match, their underlying attributes (design pressure, material, etc.) may differ

**The cost of getting this wrong**: missed equipment in procurement, incorrect specifications, rework during construction, and project delays.

---

## The Solution

The **KBR RDL Data Harmonizer** automates the entire reconciliation process:

1. **Upload** Excel files from all 5 data sources
2. **Select** which source(s) serve as the master reference
3. **Run** the fuzzy matching harmonization engine
4. **Analyze** results through interactive dashboards, charts, connection maps, and detailed gap reports
5. **Enrich** the master file with suggested additions from non-master sources
6. **Export** professional, color-coded Excel reports ready for stakeholder review

What used to take **weeks of manual effort** now takes **minutes**.

---

## Key Benefits

| Benefit | Description |
|---------|-------------|
| **Automated Matching** | Fuzzy matching engine handles naming variations (abbreviations, word order, punctuation) automatically |
| **Bidirectional Gap Analysis** | Detects classes missing from non-master sources (forward gaps) AND classes in non-master sources not in the master (reverse gaps) |
| **Enrichment Workflow** | Suggests additions to the master file with CFIHOS code matching, preview, and one-click download |
| **Visual Connection Maps** | Graphviz diagrams show how each equipment class connects across all sources |
| **Professional Excel Reports** | KBR-branded, color-coded exports with multiple sheets, ready for management review |
| **Attribute-Level Comparison** | Browse and compare ISM Functional Class Attributes (Aramco) vs. ISM Physical Class Attributes (LTC) |
| **Interactive Dashboard** | Donut charts, bar charts, score histograms, and KPI cards for instant insight |
| **Batch Processing** | Process hundreds of equipment names at once via text input or CSV upload |
| **CFIHOS Integration** | Automatic CFIHOS code lookup with exact/fuzzy match indicators |
| **Zero Infrastructure** | Single Python file, runs locally via Streamlit — no database, no server setup |

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    Streamlit Web UI (8 tabs)                 │
│  Upload | Dashboard | Gaps | Attributes | Map | Search | …  │
└──────────────────────────┬──────────────────────────────────┘
                           │
              ┌────────────┴────────────┐
              │   Harmonization Engine   │
              │  - Fuzzy matching        │
              │  - Compound expansion    │
              │  - Word overlap filter   │
              │  - Bidirectional gaps    │
              │  - CFIHOS lookup         │
              └────────────┬────────────┘
                           │
         ┌─────────────────┼─────────────────┐
         │                 │                 │
    ┌────┴────┐      ┌────┴────┐      ┌────┴────┐
    │  File   │      │ Session │      │  Excel  │
    │ Readers │      │  State  │      │ Builder │
    │ (5 src) │      │(Streamlit)│    │(openpyxl)│
    └─────────┘      └─────────┘      └─────────┘
```

The entire application is a **single-file Streamlit app** (`backend/main.py`, ~2550 lines) organized into clearly separated sections:

1. **CSS Theme** — KBR corporate branding (Inter font, navy blue palette, card layouts)
2. **Session State** — Persistent data across reruns (files, classes, attributes, gaps)
3. **File Readers** — 5 specialized parsers, one per data source
4. **Attribute Readers** — Aramco Functional + LTC Physical class attributes
5. **Harmonization Engine** — Fuzzy matching + gap detection + reverse gap computation
6. **Demo Data** — Built-in sample data for testing
7. **Excel Builders** — Professional styled exports (harmonization report + enriched master)
8. **Graphviz Builders** — Connection diagram generators
9. **UI Tabs** — 8 interactive tabs with full Streamlit components

---

## Supported Data Sources

| Source | Key | Excel Sheet Expected | ID Column | Name Column |
|--------|-----|---------------------|-----------|-------------|
| Saudi Aramco 9COM | `aramco` | `ISM Functional Classes` | `Id` | `Name` |
| CFIHOS Standard | `cfihos` | `equipment class` | `CFIHOS unique id` | `equipment class name` |
| KBR FEED | `kbr` | First sheet | `Class Id` | `Class Name (855)` |
| LTC Contractor | `ltc` | `ISM Physical Classes` (preferred) or `ISM Functional Classes` | `Id` | `Name` |
| SA Document | `sa_doc` | `SA_DOC_attributes` | `ID (CFIHOS_1.5)` | `Attribute` |

Each reader gracefully falls back to the first sheet if the expected sheet name is not found.

### Additional Sheets Parsed

- **Aramco**: `ISM Functional Class Attributes` — loaded into the Attributes tab and used in the enriched Excel Class Attributes sheet
- **LTC**: `ISM Physical Class Attributes` — loaded into the Attributes tab for comparison

---

## Installation & Setup

### Prerequisites

- Python 3.9 or higher
- pip (Python package manager)

### Install Dependencies

```bash
cd backend
pip install -r requirements.txt
```

The `requirements.txt` contains:

```
streamlit>=1.28.0
pandas>=2.1.0
openpyxl>=3.1.0
rapidfuzz>=3.5.0
altair>=5.0.0
```

### Run the Application

```bash
cd backend
streamlit run main.py
```

The app will open in your browser at `http://localhost:8501`.

---

## Quick Start

1. **Launch** the app with `streamlit run main.py`
2. On the **Upload & Configure** tab:
   - Select a master file (e.g., `aramco`)
   - Upload Excel files for each source, or click **Load Demo Data** to try with sample data
   - Adjust the **Match Threshold** slider (default: 75%)
   - Click **Run Harmonization**
3. Switch to the **Dashboard** tab to see results:
   - KPI cards showing total classes, coverage %, and per-source match rates
   - Donut chart for coverage overview
   - Bar chart for match rates by source
   - Score distribution histogram
4. Check the **Gap Analysis** tab for:
   - Forward gaps (master classes missing from non-master sources)
   - **Enrichment Suggestions** (non-master classes not in the master — scroll down)
5. Explore **Connection Map** for visual diagrams
6. Use **Search** to find specific equipment classes
7. Click **Export Full Report** to download the KBR-branded Excel report

---

## Feature Guide

### 1. Upload & Configure

The entry point for the application. Three steps:

**Step 1 — Choose Master File(s)**
- Select 1 or 2 file types as the master reference
- When 2 masters are selected, they are fused into a **single unified reference** (cross-matched against each other first, then deduplicated)
- All other uploaded files are compared against this combined master

**Step 2 — Upload Data Files**
- Upload `.xlsx`, `.xls`, or `.csv` files for each of the 5 sources
- Each file is parsed by its specialized reader
- Attributes are automatically extracted from Aramco and LTC files if the corresponding sheets exist
- Source status cards show record counts and filenames

**Step 3 — Run Harmonization**
- Click "Run Harmonization" to start the matching engine
- Or click "Load Demo Data" to use built-in sample data
- Once complete, all other tabs become active
- "Export Full Report" button appears for downloading the Excel output

### 2. Dashboard

A comprehensive analytics view with:

- **Context bar** — shows which master, compared sources, and threshold
- **KPI cards** — Total Classes, Fully Matched count (with % coverage), Has Gaps count, and per-source match percentages
- **Coverage Overview** — Donut chart showing Fully Matched vs. Has Gaps, with center percentage
- **Match Rate by Source** — Grouped bar chart comparing Matched vs. Gaps percentage for each non-master source
- **Match Score Distribution** — Histogram of all match scores (binned by 5%), color-coded by source
- **Match Quality Breakdown** — Four colored tiles: Exact (95%+), Strong (85-94%), Partial (75-84%), Weak (<75%)
- **Class Detail Browser** — Expandable list of every equipment class with master reference and comparison results, filterable by "All", "No Gaps only", or "Has Gaps only"
- **Full Data Table** — Complete harmonization results in a sortable, searchable Streamlit dataframe

### 3. Gap Analysis

Forward gap analysis — classes in the master that are missing from non-master sources:

- **Gap chart** — Bar chart showing gap count per source
- **Gap metrics** — Per-source gap counts
- **Filterable gap table** — Filter by source, shows Equipment Class, Found In, Gap In, and recommended Action
- **Action Items** — Numbered list of specific remediation steps

### 4. Enrichment Suggestions

Reverse gap analysis — classes in non-master sources that don't exist in the master:

This is the key differentiator. When a non-master source (e.g., KBR with 200 classes) has more equipment classes than the master (e.g., Aramco with 146), the system:

1. **Detects** all unmatched non-master records
2. **Shows metrics** — how many extra classes per source
3. **CFIHOS-matches** each suggestion — looks up the closest CFIHOS code with "Exact" or percentage match indicator
4. **Previews** the suggestions in a filterable table (Source, ID, Name, CFIHOS Code, CFIHOS Name, CFIHOS Match, Discipline)
5. **Downloads** an enriched master Excel file containing the original master records + all accepted suggestions

The enriched Excel includes:
- **Status column** — "Original" (green) or "Suggested from [source]" (blue)
- **CFIHOS columns** — Code, Name, and Match quality
- **Enrichment Summary sheet** — Original count, additions count, new total
- **Class Attributes sheet** — ISM Functional Class Attributes organized by equipment class with color-coded Presence column (green = mandatory, blue = optional)

### 5. Attributes Explorer

Four view modes:

- **Browse by Class** — Select an equipment class and see its Aramco Functional Attributes and LTC Physical Attributes side by side
- **Full Aramco Attributes** — Complete attribute dataset in a sortable table
- **Full LTC Attributes** — Complete attribute dataset in a sortable table
- **Attribute Comparison** — For a selected class, shows:
  - Common attributes (shared between Aramco and LTC)
  - Attributes only in Aramco (gaps in LTC)
  - Attributes only in LTC (gaps in Aramco)
  - Side-by-side attribute tables

### 6. Connection Map

Graphviz-powered visual diagrams:

- **Per-class detail** — Select any equipment class to see a directed graph:
  - Center node: Master Reference (blue, showing all master entries)
  - Connected nodes: Each matched source (colored by source, showing name + ID + match %)
  - Dashed red nodes: GAP sources with "NO MATCH" labels
- **Overview (first 20)** — Shows the first 20 classes in a single large graph with all connections

### 7. Search

Intelligent single-class lookup:

- Type an equipment name (e.g., "junction box", "centrifugal pump")
- Uses **compound name expansion** — if a class is "Junction Box / Splice Case", searching "Junction Box" will find it
- Uses **dual scorer** — tries both `token_set_ratio` and `token_sort_ratio`, takes the best result
- Shows: match score, status, master reference entries, all source matches/gaps, and a connection diagram

### 8. Batch Process

Process multiple equipment names at once:

- Enter names manually (one per line) or upload a CSV file
- Each name is matched against the harmonized dataset
- Results table shows: Input, Best Match, Score%, Status, Gap In
- Download results as CSV

### 9. Logs

Operational log of all harmonization activities:

- File uploads with record counts
- Harmonization runs with class counts and parameters
- Reverse gap detection results
- Displayed in reverse chronological order

---

## Harmonization Engine — How It Works

### Fuzzy Matching Algorithm

The core matching uses [RapidFuzz](https://github.com/maxbachmann/RapidFuzz), a high-performance fuzzy string matching library. The engine uses `token_sort_ratio` as the primary scorer, which:

1. Tokenizes both strings (splits into words)
2. Sorts tokens alphabetically
3. Computes the Levenshtein distance ratio

This means "Shell and Tube Heat Exchanger" and "Heat Exchanger Shell and Tube" score 100% — word order doesn't matter.

### Compound Name Expansion

Many industrial classes use compound names with "/" separators, such as:
- "Junction Box / Splice Case"
- "Pressure Transmitter / Gauge"

The engine expands these into individual search targets:

```
"Junction Box / Splice Case"  →  ["Junction Box / Splice Case",
                                   "Junction Box",
                                   "Splice Case"]
```

Each expanded part maps back to the original class, so a search for "Junction Box" correctly finds "Junction Box / Splice Case" at high confidence.

### Word Overlap Validation

To prevent false positives (e.g., "Junction Box" matching "Function Block" due to similar letters), every fuzzy match is validated with a **word overlap check**:

1. Both names are tokenized into meaningful words (2+ characters, noise words removed)
2. At least one word must be shared, OR one word must be a substring of another (3+ character words)
3. If no overlap exists, the match is rejected regardless of the fuzzy score

This eliminates matches where the fuzzy score is high but the actual equipment is completely different.

### Match Classification

Every match is classified into quality tiers:

| Score Range | Label | Confidence |
|-------------|-------|------------|
| 95-100% | Exact Match | High |
| 85-94% | Strong Match | High |
| 75-84% | Partial Match | Medium |
| <75% | Weak Match | Low |

The default threshold is 75%. Anything below the threshold is considered a **GAP** (not matched).

### Bidirectional Gap Analysis

The engine performs two types of gap analysis in a single harmonization pass:

**Forward Gaps** (Master → Non-master):
For each master class, check if it exists in every non-master source. If not found → forward gap.

**Reverse Gaps** (Non-master → Master):
For each non-master source, check which records were never matched by any master class. These are "extra" classes that exist in the non-master source but have no equivalent in the master.

Reverse gaps feed the **Enrichment Suggestions** feature.

### CFIHOS Code Matching

When building the enriched master, each class is matched against the CFIHOS standard through a three-tier approach:

1. **Direct reference** — If the Aramco record has a `cfihos_ref` field, look up that ID directly in the CFIHOS records → "Exact" match
2. **ID lookup** — Check if the reference ID exists but no name was found → "Ref Only"
3. **Fuzzy match** — If no direct reference exists, fuzzy-match the class name against all CFIHOS records → shows percentage (e.g., "85%")

The result populates three columns in the enriched Excel: CFIHOS Code, CFIHOS Name, and CFIHOS Match.

---

## Excel Export Details

### Harmonization Report Excel

Downloaded via "Export Full Report" on the Upload tab. Contains 3 sheets:

**Sheet 1: Harmonization Results**
- Row 1: KBR-AMCDE logo bar (navy dark background, cyan accent text)
- Row 2: Title bar (KBR primary blue)
- Row 3: Subtitle bar (master, compared sources, threshold, total count)
- Row 4: Column headers (navy fill, white text, auto-filter enabled, freeze panes)
- Data rows: Alternating white/light gray, with:
  - Master Reference columns (ID + Name per master)
  - Per-source columns (ID, Name, Match%, Status)
  - Gap In column (green for "No Gaps", red for gaps)
  - Gap Count column

**Sheet 2: Gap Analysis**
- Same branded header structure
- Columns: #, Equipment Class, Found In, Gap In, Action
- Red-highlighted Gap In cells

**Sheet 3: Summary**
- Total Equipment Classes, Master Reference, Compared Against, Match Threshold
- Classes with No Gaps ratio
- Per-source match/gap breakdown

### Enriched Master Excel

Downloaded from the Enrichment Suggestions section. Contains 3-4 sheets:

**Sheet 1: [Source] Enriched** (e.g., "Saudi Aramco 9COM Enriched")
- Same branded header structure
- Columns: #, Status, Id, Name, CFIHOS Code, CFIHOS Name, CFIHOS Match
- Original records with green "Original" status
- Blue separator bar: "SUGGESTED ADDITIONS"
- Suggested records with blue "Suggested from [source]" status
- CFIHOS Match column: green for "Exact", orange for percentage matches

**Sheet 2: Enrichment Summary**
- Original Master, Original Classes count, Suggested Additions count, New Total
- Additions breakdown by source

**Sheet 3: Class Attributes** (if Aramco attributes are loaded)
- ISM Functional Class Attributes organized by equipment class
- Navy blue separator rows for each class (showing class name, IDs, attribute count)
- Column headers repeated per class for readability
- Presence column color-coded: green = Mandatory/Required, blue = Optional
- Covers both original and suggested classes

---

## Demo Mode

Click **Load Demo Data** to instantly load sample data for all 5 sources without uploading files:

- **Aramco**: 8 classes (Centrifugal Pump, Shell and Tube Heat Exchanger, Pressure Vessel, Control Valve, Electric Motor, Air Cooled Heat Exchanger, Reciprocating Compressor, Storage Tank)
- **CFIHOS**: 8 classes (matching Aramco with slight naming variations)
- **KBR**: 10 classes (8 matching + 3 extras: Cooling Tower, Flare Stack, Drum Separator — to demonstrate enrichment suggestions)
- **LTC**: 6 classes (subset — to demonstrate forward gaps)
- **SA Document**: 3 attribute records

This demonstrates:
- Exact and fuzzy matches across sources
- Forward gaps (Storage Tank and Reciprocating Compressor missing from LTC)
- Reverse gaps / enrichment suggestions (Cooling Tower, Flare Stack, Drum Separator from KBR)
- CFIHOS code matching

---

## File Format Requirements

All uploaded files must be Excel format (`.xlsx`, `.xls`) or CSV (`.csv`).

### Aramco File
- Must contain a sheet named `ISM Functional Classes`
- Required columns: `Id`, `Name`
- Optional column: `nmcltr:CFIHOS_1.5` (for direct CFIHOS reference)
- Optional sheet: `ISM Functional Class Attributes` (for attribute analysis)

### CFIHOS File
- Must contain a sheet named `equipment class`
- Required columns: `CFIHOS unique id`, `equipment class name`

### KBR File
- Uses first sheet
- Required columns: `Class Id`, `Class Name (855)`
- Optional column: `Discipline`

### LTC File
- Prefers sheet `ISM Physical Classes`, falls back to `ISM Functional Classes`, then first sheet
- Required columns: `Id`, `Name`
- Note: `[OBSOLETE]` prefixes in names are automatically stripped for matching
- Optional sheet: `ISM Physical Class Attributes` (for attribute analysis)

### SA Document File
- Must contain a sheet named `SA_DOC_attributes`
- Required columns: `ID (CFIHOS_1.5)`, `Attribute`
- Optional column: `Name (CFIHOS_1.5)`

---

## Configuration

### Match Threshold

The **Match Threshold** slider (50-100%) controls the minimum fuzzy match score required to count as a match. Below this threshold, the comparison is recorded as a **GAP**.

| Threshold | Behavior |
|-----------|----------|
| 95-100% | Very strict — only near-exact name matches count |
| 85-94% | Strict — handles minor abbreviations (HX → Heat Exchanger) |
| 75-84% | Balanced (default) — catches most naming variations |
| 50-74% | Loose — may produce false positives, use with caution |

### Master Selection

- **Single master**: One source becomes the reference. All others are compared against it.
- **Dual master**: Two sources are cross-matched first (fused into one unified reference with deduplication), then all other sources are compared against the combined reference.

---

## Tech Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| Web Framework | [Streamlit](https://streamlit.io/) >= 1.28.0 | Interactive web UI with session state |
| Data Processing | [pandas](https://pandas.pydata.org/) >= 2.1.0 | DataFrame operations, Excel reading |
| Fuzzy Matching | [RapidFuzz](https://github.com/maxbachmann/RapidFuzz) >= 3.5.0 | High-performance string matching |
| Excel Export | [openpyxl](https://openpyxl.readthedocs.io/) >= 3.1.0 | Styled Excel generation with cell-level formatting |
| Charts | [Altair](https://altair-viz.github.io/) >= 5.0.0 | Declarative statistical visualizations |
| Diagrams | [Graphviz](https://graphviz.org/) (via Streamlit) | Connection map visualizations |
| Fonts | [Google Fonts — Inter](https://fonts.google.com/specimen/Inter) | Professional typography |

---

## Project Structure

```
RDL/
├── backend/
│   ├── main.py              # Complete application (~2550 lines)
│   ├── requirements.txt     # Python dependencies
│   └── uploads/
│       └── .gitkeep         # Upload directory placeholder
├── .gitignore
└── README.md                # This file
```

The application is intentionally a **single-file design** for simplicity of deployment — no database, no separate API server, no build step. Just install dependencies and run.

---

## Color & Branding

The application follows KBR corporate branding:

| Element | Color | Hex |
|---------|-------|-----|
| KBR Primary | Navy Blue | `#003087` |
| KBR Dark | Deep Navy | `#001D54` |
| KBR Accent | Cyan | `#00A3E0` |
| Navy | Slate Navy | `#1B3A5C` |
| Success | Green | `#059669` |
| Error/Gap | Red | `#EF4444` |
| Warning | Orange | `#D97706` |

The KBR-AMCDE logo appears in the application header and on every Excel export sheet.

---

## Version History

| Version | Description |
|---------|-------------|
| v1.0 | Initial harmonization engine with basic gap analysis |
| v2.0 | Added attributes tab, connection maps, batch processing |
| v3.0 | Complete UI overhaul, KBR corporate branding, enrichment suggestions, CFIHOS matching, professional Excel exports, Class Attributes sheet, dashboard charts, compound name search fix |

---

*KBR RDL Data Harmonizer v3.0 — Equipment Class Harmonization & Gap Analysis Platform*
*Built by KBR AMCDE Team*
