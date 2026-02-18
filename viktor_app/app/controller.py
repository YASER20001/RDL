"""
KBR RDL Data Harmonizer — VIKTOR Controller
All views (Dashboard, Results, Gaps, Enrichment, Attributes, Search, Batch)
and download handlers.
"""

import io
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import viktor as vkt

from .parametrization import Parametrization
from .engine import (
    FILE_TYPES, SOURCE_COLORS, READERS,
    get_demo_data, run_harmonization, classify_match,
    build_export_df, search_classes, batch_process, lookup_cfihos,
    read_aramco_attributes, read_ltc_attributes,
)
from .excel_builder import build_excel_bytes, build_enriched_master_excel


def _parse_files(params):
    """Parse uploaded files from params. Returns (files, aramco_attrs, ltc_attrs)."""
    if params.config.settings.use_demo:
        return get_demo_data(), None, None

    file_map = {
        "aramco": params.config.uploads.aramco_file,
        "cfihos": params.config.uploads.cfihos_file,
        "kbr": params.config.uploads.kbr_file,
        "ltc": params.config.uploads.ltc_file,
        "sa_doc": params.config.uploads.sa_doc_file,
    }

    files = {}
    aramco_attrs = None
    ltc_attrs = None

    for key, file_resource in file_map.items():
        if file_resource is not None:
            file_content = file_resource.file.open_binary()
            reader = READERS[key]
            records = reader(file_content)
            files[key] = {"filename": file_resource.filename, "records": records}

            # Read attributes from Aramco and LTC
            if key == "aramco":
                file_content.seek(0)
                aramco_attrs = read_aramco_attributes(file_content)
            elif key == "ltc":
                file_content.seek(0)
                ltc_attrs = read_ltc_attributes(file_content)

    return files, aramco_attrs, ltc_attrs


def _get_masters(params):
    """Get list of master keys from params."""
    masters = [params.config.masters.master_1]
    if params.config.masters.master_2 != "None":
        masters.append(params.config.masters.master_2)
    return masters


def _run_pipeline(params):
    """Full pipeline: parse files -> harmonize. Returns dict with all results."""
    files, aramco_attrs, ltc_attrs = _parse_files(params)
    masters = _get_masters(params)
    threshold = int(params.config.settings.threshold)

    # Check we have master files
    for m in masters:
        if m not in files:
            return None

    harmonized, reverse_gaps, logs = run_harmonization(masters, files, threshold)

    return {
        "files": files,
        "masters": masters,
        "threshold": threshold,
        "classes": harmonized,
        "reverse_gaps": reverse_gaps,
        "logs": logs,
        "aramco_attrs": aramco_attrs,
        "ltc_attrs": ltc_attrs,
    }


class Controller(vkt.Controller):
    parametrization = Parametrization

    # ══════════════════════════════════════════════════════════════════
    # Dashboard — Plotly charts + KPIs
    # ══════════════════════════════════════════════════════════════════
    @vkt.PlotlyAndDataView("Dashboard", duration_guess=6)
    def dashboard(self, params, **kwargs):
        result = _run_pipeline(params)
        if result is None:
            fig = go.Figure()
            fig.add_annotation(text="Upload files and configure masters to see dashboard",
                               xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False,
                               font=dict(size=16, color="#64748b"))
            fig.update_layout(template="plotly_white", height=400)
            data = vkt.DataGroup(vkt.DataItem("Status", "No data", status=vkt.DataStatus.WARNING))
            return vkt.PlotlyAndDataResult(fig.to_json(), data)

        classes = result["classes"]
        masters = result["masters"]
        total = len(classes)

        non_master_keys = set()
        for c in classes:
            non_master_keys.update(c["matches"].keys())
            non_master_keys.update(c["gaps"])
        non_master_keys = sorted(non_master_keys)

        no_gaps = sum(1 for c in classes if not c["gaps"])
        has_gaps = total - no_gaps
        overall_pct = int(round(no_gaps / total * 100)) if total else 0

        # Build Plotly figure with subplots
        fig = make_subplots(
            rows=2, cols=2,
            specs=[[{"type": "pie"}, {"type": "bar"}],
                   [{"type": "bar", "colspan": 2}, None]],
            subplot_titles=["Coverage Overview", "Match Rate by Source", "Match Score Distribution"],
            vertical_spacing=0.15,
            horizontal_spacing=0.1,
        )

        # Donut chart
        fig.add_trace(go.Pie(
            labels=["Fully Matched", "Has Gaps"],
            values=[no_gaps, has_gaps],
            hole=0.6,
            marker=dict(colors=["#059669", "#ef4444"]),
            textinfo="label+value",
            textfont=dict(size=12),
        ), row=1, col=1)

        # Bar chart — match rate by source
        for src in non_master_keys:
            matched = sum(1 for c in classes if src in c["matches"])
            gaps = total - matched
            fig.add_trace(go.Bar(
                name=FILE_TYPES.get(src, src),
                x=[FILE_TYPES.get(src, src)],
                y=[int(round(matched / total * 100)) if total else 0],
                marker_color=SOURCE_COLORS.get(src, "#6b7280"),
                text=[f"{int(round(matched / total * 100))}%"],
                textposition="auto",
                showlegend=False,
            ), row=1, col=2)

        # Score distribution histogram
        all_scores = []
        for c in classes:
            for src, m in c["matches"].items():
                all_scores.append(m["score"])
        if all_scores:
            fig.add_trace(go.Histogram(
                x=all_scores,
                xbins=dict(start=50, end=100, size=5),
                marker_color="#003087",
                opacity=0.8,
                name="Scores",
                showlegend=False,
            ), row=2, col=1)

        fig.update_layout(
            template="plotly_white",
            height=700,
            title=dict(
                text=f"KBR RDL Harmonization Dashboard — {total} Classes",
                font=dict(size=18, color="#003087"),
            ),
            font=dict(family="Inter, Arial, sans-serif"),
        )
        fig.update_xaxes(title_text="Score %", row=2, col=1)
        fig.update_yaxes(title_text="Count", row=2, col=1)

        # DataView with KPIs
        master_label = " + ".join(FILE_TYPES.get(m, m) for m in masters)
        source_items = []
        for src in non_master_keys:
            matched = sum(1 for c in classes if src in c["matches"])
            pct = int(round(matched / total * 100)) if total else 0
            source_items.append(
                vkt.DataItem(FILE_TYPES.get(src, src), pct, suffix="%",
                             status=vkt.DataStatus.SUCCESS if pct >= 80 else vkt.DataStatus.WARNING)
            )

        exact_cnt = sum(1 for c in classes for m in c["matches"].values() if m["score"] >= 95)
        strong_cnt = sum(1 for c in classes for m in c["matches"].values() if 85 <= m["score"] < 95)
        partial_cnt = sum(1 for c in classes for m in c["matches"].values() if 75 <= m["score"] < 85)
        weak_cnt = sum(1 for c in classes for m in c["matches"].values() if m["score"] < 75)

        data = vkt.DataGroup(
            vkt.DataItem("Master Reference", master_label),
            vkt.DataItem("Threshold", result["threshold"], suffix="%"),
            vkt.DataItem("Total Classes", total),
            vkt.DataItem("Fully Matched", no_gaps, status=vkt.DataStatus.SUCCESS),
            vkt.DataItem("Has Gaps", has_gaps,
                         status=vkt.DataStatus.ERROR if has_gaps > 0 else vkt.DataStatus.SUCCESS),
            vkt.DataItem("Coverage", overall_pct, suffix="%",
                         status=vkt.DataStatus.SUCCESS if overall_pct >= 80 else vkt.DataStatus.WARNING),
            vkt.DataItem("Match Quality", "",
                         subgroup=vkt.DataGroup(
                             vkt.DataItem("Exact (95%+)", exact_cnt, status=vkt.DataStatus.SUCCESS),
                             vkt.DataItem("Strong (85-94%)", strong_cnt, status=vkt.DataStatus.SUCCESS),
                             vkt.DataItem("Partial (75-84%)", partial_cnt, status=vkt.DataStatus.WARNING),
                             vkt.DataItem("Weak (<75%)", weak_cnt, status=vkt.DataStatus.ERROR),
                         )),
            vkt.DataItem("Per-Source Match Rate", "",
                         subgroup=vkt.DataGroup(*source_items) if source_items else vkt.DataGroup(
                             vkt.DataItem("No sources", "-"))),
        )

        return vkt.PlotlyAndDataResult(fig.to_json(), data)

    # ══════════════════════════════════════════════════════════════════
    # Harmonization Results Table
    # ══════════════════════════════════════════════════════════════════
    @vkt.TableView("Results Table", duration_guess=6)
    def results_table(self, params, **kwargs):
        result = _run_pipeline(params)
        if result is None:
            return vkt.TableResult(pd.DataFrame({"Status": ["Upload files and configure masters first"]}))

        df = build_export_df(result["classes"], result["masters"])
        return vkt.TableResult(df)

    # ══════════════════════════════════════════════════════════════════
    # Gap Analysis
    # ══════════════════════════════════════════════════════════════════
    @vkt.PlotlyAndDataView("Gap Analysis", duration_guess=6)
    def gap_analysis(self, params, **kwargs):
        result = _run_pipeline(params)
        if result is None:
            fig = go.Figure()
            fig.add_annotation(text="No data available", xref="paper", yref="paper",
                               x=0.5, y=0.5, showarrow=False, font=dict(size=16, color="#64748b"))
            fig.update_layout(template="plotly_white", height=300)
            data = vkt.DataGroup(vkt.DataItem("Status", "No data", status=vkt.DataStatus.WARNING))
            return vkt.PlotlyAndDataResult(fig.to_json(), data)

        classes = result["classes"]
        masters = result["masters"]
        total = len(classes)

        # Collect gaps
        all_gaps = []
        gap_by_source = {}
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
                gap_by_source[FILE_TYPES.get(g, g)] = gap_by_source.get(FILE_TYPES.get(g, g), 0) + 1

        # Gap chart
        fig = go.Figure()
        if gap_by_source:
            sources = sorted(gap_by_source.keys(), key=lambda x: -gap_by_source[x])
            colors = ["#ef4444", "#f97316", "#eab308", "#8b5cf6"]
            fig.add_trace(go.Bar(
                x=sources,
                y=[gap_by_source[s] for s in sources],
                marker_color=colors[:len(sources)],
                text=[gap_by_source[s] for s in sources],
                textposition="auto",
            ))
            fig.update_layout(
                title="Gap Count by Source",
                template="plotly_white",
                height=350,
                font=dict(family="Inter, Arial, sans-serif"),
                xaxis_title="Source",
                yaxis_title="Gap Count",
            )
        else:
            fig.add_annotation(text="No gaps found - all sources fully matched!",
                               xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False,
                               font=dict(size=16, color="#059669"))
            fig.update_layout(template="plotly_white", height=300)

        # Data panel
        no_gaps = sum(1 for c in classes if not c["gaps"])
        items = [
            vkt.DataItem("Total Gaps", len(all_gaps),
                         status=vkt.DataStatus.ERROR if all_gaps else vkt.DataStatus.SUCCESS),
            vkt.DataItem("Fully Matched", f"{no_gaps}/{total}",
                         status=vkt.DataStatus.SUCCESS),
        ]
        for src, cnt in sorted(gap_by_source.items(), key=lambda x: -x[1]):
            items.append(vkt.DataItem(src, cnt, suffix="gaps", status=vkt.DataStatus.ERROR))

        # Reverse gaps (enrichment)
        reverse_gaps = result["reverse_gaps"]
        if reverse_gaps:
            rev_items = []
            for src, recs in sorted(reverse_gaps.items()):
                rev_items.append(vkt.DataItem(
                    f"Extra in {FILE_TYPES.get(src, src)}", len(recs),
                    status=vkt.DataStatus.INFO))
            items.append(vkt.DataItem("Enrichment Suggestions", "",
                                      subgroup=vkt.DataGroup(*rev_items)))

        data = vkt.DataGroup(*items)
        return vkt.PlotlyAndDataResult(fig.to_json(), data)

    # ══════════════════════════════════════════════════════════════════
    # Enrichment Suggestions Table
    # ══════════════════════════════════════════════════════════════════
    @vkt.TableView("Enrichment", duration_guess=6)
    def enrichment_table(self, params, **kwargs):
        result = _run_pipeline(params)
        if result is None:
            return vkt.TableResult(pd.DataFrame({"Status": ["Upload files and configure masters first"]}))

        reverse_gaps = result["reverse_gaps"]
        if not reverse_gaps:
            return vkt.TableResult(pd.DataFrame({
                "Status": ["No extra classes found - all non-master records matched the master"]}))

        files = result["files"]
        threshold = result["threshold"]
        cfihos_records = files["cfihos"]["records"] if "cfihos" in files else []

        preview_rows = []
        for src, recs in sorted(reverse_gaps.items()):
            for rec in recs:
                cfihos_code, cfihos_name, cfihos_match = lookup_cfihos(
                    rec["name"], cfihos_records, threshold)
                row = {
                    "Source": FILE_TYPES.get(src, src),
                    "ID": rec["id"],
                    "Name": rec["name"],
                    "CFIHOS Code": cfihos_code,
                    "CFIHOS Name": cfihos_name,
                    "CFIHOS Match": cfihos_match,
                }
                if "discipline" in rec and rec["discipline"]:
                    row["Discipline"] = rec["discipline"]
                preview_rows.append(row)

        df = pd.DataFrame(preview_rows)
        return vkt.TableResult(df)

    # ══════════════════════════════════════════════════════════════════
    # Attributes Table
    # ══════════════════════════════════════════════════════════════════
    @vkt.TableView("Attributes", duration_guess=6)
    def attributes_table(self, params, **kwargs):
        files, aramco_attrs, ltc_attrs = _parse_files(params)

        if aramco_attrs is not None and not aramco_attrs.empty:
            display_cols = [c for c in ["Class_Id", "Name", "Attribute_Desc", "Class_Desc",
                                        "Description", "Presence", "Size", "Discipline",
                                        "UomClassId", "UomRequire", "ValidationRule", "Group_Id"]
                           if c in aramco_attrs.columns]
            return vkt.TableResult(aramco_attrs[display_cols])

        if ltc_attrs is not None and not ltc_attrs.empty:
            display_cols = [c for c in ["Class_Id", "Name", "Description", "Presence", "Size",
                                        "Discipline", "UomClassId", "UomRequire",
                                        "ValidationRule", "ValidationType", "MaxOccur", "Aspect"]
                           if c in ltc_attrs.columns]
            return vkt.TableResult(ltc_attrs[display_cols])

        return vkt.TableResult(pd.DataFrame({
            "Status": ["Upload Aramco (ISM Functional Class Attributes) "
                       "or LTC (ISM Physical Class Attributes) to see attribute data"]}))

    # ══════════════════════════════════════════════════════════════════
    # Connection Map (Graphviz SVG via WebView)
    # ══════════════════════════════════════════════════════════════════
    @vkt.WebView("Connection Map", duration_guess=6)
    def connection_map(self, params, **kwargs):
        result = _run_pipeline(params)
        if result is None:
            return vkt.WebResult(html="<div style='text-align:center;padding:4rem;color:#64748b;'>"
                                      "<h2>No Connection Data</h2>"
                                      "<p>Upload files and configure masters to see the connection map.</p>"
                                      "</div>")

        classes = result["classes"]
        masters = result["masters"]

        # Build HTML connection map using CSS/HTML (no Graphviz dependency needed)
        html_parts = ["""
        <html><head>
        <style>
            body { font-family: Inter, Arial, sans-serif; background: #f8fafc; padding: 1rem; }
            .class-card { background: white; border: 1px solid #e2e8f0; border-radius: 12px;
                         padding: 1rem; margin: 0.75rem 0; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
            .class-name { font-size: 1.1rem; font-weight: 700; color: #003087; margin-bottom: 0.5rem; }
            .connections { display: flex; flex-wrap: wrap; gap: 0.5rem; }
            .conn-badge { padding: 0.3rem 0.8rem; border-radius: 8px; font-size: 0.82rem; font-weight: 500; }
            .conn-match { background: #d1fae5; color: #065f46; }
            .conn-gap { background: #fee2e2; color: #991b1b; }
            .master-ref { font-size: 0.85rem; color: #64748b; margin-bottom: 0.3rem; }
            h2 { color: #003087; margin-bottom: 1rem; }
        </style></head><body>
        <h2>Connection Map — {total} Equipment Classes</h2>
        """.format(total=len(classes))]

        for c in classes[:30]:  # Limit to 30 for performance
            master_parts = []
            for m in masters:
                me = c["master_entries"].get(m)
                if me:
                    master_parts.append(f"{FILE_TYPES[m]}: {me['name']} ({me['id']})")

            connections = []
            for src, match in c["matches"].items():
                connections.append(
                    f'<span class="conn-badge conn-match">'
                    f'{FILE_TYPES.get(src, src)}: {match["name"]} — {match["score"]}%</span>'
                )
            for g in c["gaps"]:
                connections.append(
                    f'<span class="conn-badge conn-gap">'
                    f'{FILE_TYPES.get(g, g)}: GAP</span>'
                )

            html_parts.append(f"""
            <div class="class-card">
                <div class="class-name">{c['index']}. {c['canonical_name']}</div>
                <div class="master-ref">Master: {' | '.join(master_parts)}</div>
                <div class="connections">{''.join(connections)}</div>
            </div>
            """)

        if len(classes) > 30:
            html_parts.append(f'<p style="color:#64748b;text-align:center;">Showing 30 of {len(classes)} classes</p>')

        html_parts.append("</body></html>")
        return vkt.WebResult(html="".join(html_parts))

    # ══════════════════════════════════════════════════════════════════
    # Search Results
    # ══════════════════════════════════════════════════════════════════
    @vkt.DataView("Search Results", duration_guess=3)
    def search_results(self, params, **kwargs):
        query = params.tools.search.query
        if not query or not query.strip():
            return vkt.DataResult(vkt.DataGroup(
                vkt.DataItem("Status", "Enter an equipment name in the Search tab to find matches")))

        result = _run_pipeline(params)
        if result is None:
            return vkt.DataResult(vkt.DataGroup(
                vkt.DataItem("Status", "Upload files first", status=vkt.DataStatus.WARNING)))

        matched_class, score = search_classes(query, result["classes"])
        if matched_class is None:
            return vkt.DataResult(vkt.DataGroup(
                vkt.DataItem("Status", "No match found", status=vkt.DataStatus.ERROR)))

        status_label, confidence = classify_match(score)
        status = vkt.DataStatus.SUCCESS if score >= 90 else vkt.DataStatus.WARNING if score >= 75 else vkt.DataStatus.ERROR

        # Build master reference items
        master_items = []
        for m in result["masters"]:
            me = matched_class["master_entries"].get(m)
            if me:
                master_items.append(vkt.DataItem(FILE_TYPES[m], f"{me['name']} ({me['id']})"))
            else:
                master_items.append(vkt.DataItem(FILE_TYPES[m], "Not in this master",
                                                 status=vkt.DataStatus.WARNING))

        # Build match items
        match_items = []
        for src, match in matched_class["matches"].items():
            ms = vkt.DataStatus.SUCCESS if match["score"] >= 90 else vkt.DataStatus.WARNING
            match_items.append(vkt.DataItem(
                FILE_TYPES.get(src, src),
                f"{match['name']} ({match['id']}) — {match['score']}%",
                status=ms))

        # Build gap items
        gap_items = []
        for g in matched_class["gaps"]:
            gap_items.append(vkt.DataItem(FILE_TYPES.get(g, g), "GAP", status=vkt.DataStatus.ERROR))

        items = [
            vkt.DataItem("Search Query", query),
            vkt.DataItem("Best Match", matched_class["canonical_name"], status=status),
            vkt.DataItem("Match Score", score, suffix="%"),
            vkt.DataItem("Match Quality", status_label),
            vkt.DataItem("Master Reference", "",
                         subgroup=vkt.DataGroup(*master_items) if master_items else vkt.DataGroup(
                             vkt.DataItem("None", "-"))),
        ]
        if match_items:
            items.append(vkt.DataItem("Source Matches", "",
                                      subgroup=vkt.DataGroup(*match_items)))
        if gap_items:
            items.append(vkt.DataItem("Gaps", "",
                                      subgroup=vkt.DataGroup(*gap_items)))

        return vkt.DataResult(vkt.DataGroup(*items))

    # ══════════════════════════════════════════════════════════════════
    # Batch Results
    # ══════════════════════════════════════════════════════════════════
    @vkt.TableView("Batch Results", duration_guess=6)
    def batch_results(self, params, **kwargs):
        result = _run_pipeline(params)
        if result is None:
            return vkt.TableResult(pd.DataFrame({"Status": ["Upload files and configure masters first"]}))

        # Get items from text area or CSV
        items = []
        if params.tools.batch.csv_file is not None:
            csv_content = params.tools.batch.csv_file.file.open_binary()
            bdf = pd.read_csv(csv_content)
            col = "name" if "name" in bdf.columns else bdf.columns[0]
            items = bdf[col].dropna().astype(str).tolist()
        elif params.tools.batch.names and params.tools.batch.names.strip():
            items = [l.strip() for l in params.tools.batch.names.strip().split("\n") if l.strip()]

        if not items:
            return vkt.TableResult(pd.DataFrame({
                "Status": ["Enter equipment names in the Batch tab (one per line) or upload a CSV"]}))

        results = batch_process(items, result["classes"])
        return vkt.TableResult(pd.DataFrame(results))

    # ══════════════════════════════════════════════════════════════════
    # System Logs
    # ══════════════════════════════════════════════════════════════════
    @vkt.DataView("Logs", duration_guess=3)
    def logs_view(self, params, **kwargs):
        result = _run_pipeline(params)
        if result is None:
            return vkt.DataResult(vkt.DataGroup(
                vkt.DataItem("Status", "No logs yet — run harmonization first")))

        log_items = []
        for i, log in enumerate(result["logs"]):
            log_items.append(vkt.DataItem(f"Log {i+1}", log))

        if not log_items:
            log_items.append(vkt.DataItem("Status", "No logs"))

        return vkt.DataResult(vkt.DataGroup(*log_items))

    # ══════════════════════════════════════════════════════════════════
    # Download handlers
    # ══════════════════════════════════════════════════════════════════
    def download_report(self, params, **kwargs):
        """Download the harmonization Excel report."""
        result = _run_pipeline(params)
        if result is None:
            return vkt.DownloadResult("No data to export. Upload files first.", "error.txt")

        excel_bytes = build_excel_bytes(
            result["classes"], result["masters"], result["threshold"])
        return vkt.DownloadResult(excel_bytes, "harmonization_report.xlsx")

    def download_enriched(self, params, **kwargs):
        """Download the enriched master Excel with suggestions."""
        result = _run_pipeline(params)
        if result is None:
            return vkt.DownloadResult("No data to export. Upload files first.", "error.txt")

        reverse_gaps = result["reverse_gaps"]
        if not reverse_gaps:
            return vkt.DownloadResult("No enrichment suggestions available.", "no_enrichment.txt")

        # Build selected additions from all reverse gaps
        selected_additions = []
        for src, recs in reverse_gaps.items():
            for rec in recs:
                selected_additions.append({
                    "id": rec["id"],
                    "name": rec["name"],
                    "source": src,
                })

        excel_bytes = build_enriched_master_excel(
            selected_additions,
            result["masters"],
            result["files"],
            result["threshold"],
            aramco_attrs=result.get("aramco_attrs"),
        )
        return vkt.DownloadResult(excel_bytes, "enriched_master.xlsx")
