"""
KBR RDL Data Harmonizer — VIKTOR Parametrization
Defines the left-panel input fields: file uploads, master selection, threshold, search, batch.
"""

import viktor as vkt


class Parametrization(vkt.Parametrization):
    # ══════════════════════════════════════════════════════════════════
    # Tab 1: Configuration
    # ══════════════════════════════════════════════════════════════════
    config = vkt.Tab("Configuration")

    config.masters = vkt.Section("Master Selection")
    config.masters.info = vkt.Text(
        "Select 1 or 2 file types as the master reference. "
        "All other uploaded files are compared against this master."
    )
    config.masters.master_1 = vkt.OptionField(
        "Primary Master",
        options=["aramco", "cfihos", "kbr", "ltc", "sa_doc"],
        default="aramco",
        description="Main reference source for harmonization",
    )
    config.masters.master_2 = vkt.OptionField(
        "Secondary Master (optional)",
        options=["None", "aramco", "cfihos", "kbr", "ltc", "sa_doc"],
        default="None",
        description="Optional second master for combined reference",
    )

    config.settings = vkt.Section("Match Settings")
    config.settings.threshold = vkt.NumberField(
        "Match Threshold (%)",
        min=50,
        max=100,
        default=75,
        step=5,
        description="Minimum fuzzy match score to count as a match. Below = GAP.",
    )
    config.settings.use_demo = vkt.BooleanField(
        "Use Demo Data",
        default=False,
        description="Load built-in sample data instead of uploaded files",
    )

    config.uploads = vkt.Section(
        "File Uploads",
        description="Upload Excel files for each data source (.xlsx, .xls, .csv)",
    )
    config.uploads.aramco_file = vkt.FileField(
        "Saudi Aramco 9COM",
        file_types=[".xlsx", ".xls", ".csv"],
        description="Must contain 'ISM Functional Classes' sheet",
    )
    config.uploads.cfihos_file = vkt.FileField(
        "CFIHOS Standard",
        file_types=[".xlsx", ".xls", ".csv"],
        description="Must contain 'equipment class' sheet",
    )
    config.uploads.kbr_file = vkt.FileField(
        "KBR FEED",
        file_types=[".xlsx", ".xls", ".csv"],
        description="Uses first sheet with 'Class Id' and 'Class Name (855)' columns",
    )
    config.uploads.ltc_file = vkt.FileField(
        "LTC Contractor",
        file_types=[".xlsx", ".xls", ".csv"],
        description="Prefers 'ISM Physical Classes' sheet",
    )
    config.uploads.sa_doc_file = vkt.FileField(
        "SA Document",
        file_types=[".xlsx", ".xls", ".csv"],
        description="Must contain 'SA_DOC_attributes' sheet",
    )

    config.exports = vkt.Section("Export Reports")
    config.exports.download_report = vkt.DownloadButton(
        "Download Harmonization Report",
        method="download_report",
    )
    config.exports.download_enriched = vkt.DownloadButton(
        "Download Enriched Master",
        method="download_enriched",
    )

    # ══════════════════════════════════════════════════════════════════
    # Tab 2: Search & Batch
    # ══════════════════════════════════════════════════════════════════
    tools = vkt.Tab("Search & Batch")

    tools.search = vkt.Section("Search Equipment Class")
    tools.search.query = vkt.TextField(
        "Equipment Name",
        default="",
        description="Type an equipment name to find the best match (e.g., 'centrifugal pump')",
    )

    tools.batch = vkt.Section("Batch Process")
    tools.batch.names = vkt.TextAreaField(
        "Equipment Names (one per line)",
        default="",
        description="Enter multiple equipment names, one per line",
    )
    tools.batch.csv_file = vkt.FileField(
        "Or Upload CSV",
        file_types=[".csv"],
        description="CSV file with 'name' column or first column used",
    )
