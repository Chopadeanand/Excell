import streamlit as st
import pandas as pd
import os, sys, tempfile, shutil, io, time, zipfile
from pathlib import Path

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="HiRATE Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Sora', sans-serif;
}

/* Background */
.stApp {
    background: linear-gradient(135deg, #0f1923 0%, #1a2a3a 50%, #0d1f2d 100%);
    min-height: 100vh;
}

/* Hide streamlit default branding */
#MainMenu, footer, header {visibility: hidden;}

/* Hero header */
.hero-header {
    background: linear-gradient(90deg, #1F4E79 0%, #2E75B6 50%, #1F4E79 100%);
    border-radius: 16px;
    padding: 36px 40px;
    margin-bottom: 32px;
    border: 1px solid rgba(46,117,182,0.4);
    box-shadow: 0 8px 40px rgba(46,117,182,0.25), inset 0 1px 0 rgba(255,255,255,0.1);
    position: relative;
    overflow: hidden;
}
.hero-header::before {
    content: '';
    position: absolute;
    top: -50%; right: -10%;
    width: 400px; height: 400px;
    background: radial-gradient(circle, rgba(146,208,80,0.15) 0%, transparent 70%);
    border-radius: 50%;
}
.hero-title {
    font-size: 2.4rem;
    font-weight: 800;
    color: #ffffff;
    margin: 0;
    letter-spacing: -0.5px;
    line-height: 1.1;
}
.hero-title span { color: #92D050; }
.hero-sub {
    font-size: 1rem;
    color: rgba(255,255,255,0.65);
    margin: 8px 0 0 0;
    font-weight: 300;
    letter-spacing: 0.3px;
}

/* Cards */
.card {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 14px;
    padding: 24px;
    margin-bottom: 20px;
    backdrop-filter: blur(10px);
}
.card-title {
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #92D050;
    margin-bottom: 16px;
}

/* Upload zone */
.upload-zone {
    border: 2px dashed rgba(46,117,182,0.5);
    border-radius: 14px;
    background: rgba(31,78,121,0.15);
    padding: 32px;
    text-align: center;
    transition: all 0.3s;
}
.upload-zone:hover {
    border-color: #2E75B6;
    background: rgba(31,78,121,0.25);
}

/* File item */
.file-item {
    display: flex;
    align-items: center;
    gap: 12px;
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 10px;
    padding: 12px 16px;
    margin: 8px 0;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.85rem;
    color: rgba(255,255,255,0.85);
}
.file-icon { font-size: 1.2rem; }
.file-valid { border-left: 3px solid #92D050; }
.file-invalid { border-left: 3px solid #ED7D31; }
.file-error { border-left: 3px solid #FF4444; }

/* Status badges */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}
.badge-ok { background: rgba(146,208,80,0.2); color: #92D050; border: 1px solid rgba(146,208,80,0.3); }
.badge-warn { background: rgba(237,125,49,0.2); color: #ED7D31; border: 1px solid rgba(237,125,49,0.3); }
.badge-error { background: rgba(255,68,68,0.2); color: #FF6666; border: 1px solid rgba(255,68,68,0.3); }

/* Stats row */
.stat-box {
    background: rgba(31,78,121,0.3);
    border: 1px solid rgba(46,117,182,0.3);
    border-radius: 12px;
    padding: 20px;
    text-align: center;
}
.stat-num {
    font-size: 2.2rem;
    font-weight: 800;
    color: #ffffff;
    line-height: 1;
    font-family: 'JetBrains Mono', monospace;
}
.stat-num.green { color: #92D050; }
.stat-num.orange { color: #ED7D31; }
.stat-num.red { color: #FF6666; }
.stat-label {
    font-size: 0.72rem;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: rgba(255,255,255,0.5);
    margin-top: 6px;
    font-weight: 600;
}

/* Validation result */
.val-pass {
    background: rgba(146,208,80,0.1);
    border: 1px solid rgba(146,208,80,0.3);
    border-left: 4px solid #92D050;
    border-radius: 10px;
    padding: 16px 20px;
    color: #92D050;
    font-size: 0.9rem;
    font-weight: 600;
}
.val-fail {
    background: rgba(255,68,68,0.08);
    border: 1px solid rgba(255,68,68,0.25);
    border-left: 4px solid #FF4444;
    border-radius: 10px;
    padding: 16px 20px;
    color: #FF8888;
    font-size: 0.9rem;
}
.val-issue {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.8rem;
    color: rgba(255,180,180,0.85);
    margin: 6px 0 0 16px;
}

/* Generate button */
.stButton > button {
    width: 100%;
    background: linear-gradient(135deg, #1F4E79, #2E75B6) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 16px 32px !important;
    font-family: 'Sora', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.5px !important;
    cursor: pointer !important;
    transition: all 0.3s !important;
    box-shadow: 0 4px 20px rgba(46,117,182,0.4) !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 28px rgba(46,117,182,0.6) !important;
}

/* Progress bar */
.stProgress > div > div {
    background: linear-gradient(90deg, #1F4E79, #92D050) !important;
    border-radius: 4px !important;
}

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #145a32, #1e8449) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Sora', sans-serif !important;
    font-weight: 600 !important;
    width: 100% !important;
    box-shadow: 0 4px 16px rgba(30,132,73,0.35) !important;
}

/* Scrollable area */
.results-scroll {
    max-height: 420px;
    overflow-y: auto;
    padding-right: 4px;
}
.results-scroll::-webkit-scrollbar { width: 4px; }
.results-scroll::-webkit-scrollbar-track { background: rgba(255,255,255,0.05); border-radius: 4px; }
.results-scroll::-webkit-scrollbar-thumb { background: #2E75B6; border-radius: 4px; }

/* Separator */
.sep { height: 1px; background: rgba(255,255,255,0.07); margin: 24px 0; }

/* Info text */
.info-text { color: rgba(255,255,255,0.5); font-size: 0.85rem; line-height: 1.6; }

/* Rule list */
.rule-item {
    display: flex;
    gap: 10px;
    align-items: flex-start;
    padding: 8px 0;
    border-bottom: 1px solid rgba(255,255,255,0.05);
    font-size: 0.83rem;
    color: rgba(255,255,255,0.7);
}
.rule-num {
    background: rgba(46,117,182,0.3);
    color: #2E75B6;
    border-radius: 4px;
    padding: 1px 7px;
    font-weight: 700;
    font-size: 0.75rem;
    flex-shrink: 0;
    margin-top: 1px;
    font-family: 'JetBrains Mono', monospace;
}
</style>
""", unsafe_allow_html=True)


# ── Load report engine ────────────────────────────────────────────────────────
import importlib.util, sys

@st.cache_resource
def load_engine():  # v2
    engine_path = Path(__file__).parent / "sipl_report.py"
    if not engine_path.exists():
        st.error(f"sipl_report.py not found in {Path(__file__).parent}")
        return None
    spec = importlib.util.spec_from_file_location("sipl_report", engine_path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["sipl_report"] = mod
    spec.loader.exec_module(mod)
    return mod

engine = load_engine()


@st.cache_resource
def load_ppt_engine():
    ppt_path = Path(__file__).parent / "sipl_ppt.py"
    if not ppt_path.exists():
        return None, "sipl_ppt.py not found next to sipl_app.py"
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("sipl_ppt", ppt_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod, None
    except Exception as e:
        return None, str(e)


# ── Helper functions ──────────────────────────────────────────────────────────
def validate_file_ui(df):
    """Returns (is_valid, errors_list, warnings_list, cleaned_df)
    Calls clean_df first to remove NA rows for ratings 1,5,10 before validation."""
    if engine is None:
        return False, ["Engine not loaded"], [], df
    # Clean first — remove NA-remark rows for ratings 1,5,10
    cleaned_df, removed = engine.clean_df(df)
    warnings = []
    if removed > 0:
        warnings.append(f"{removed} rows with Rating in (1,5,10) and Remarks='NA' were removed before validation")
    all_issues = engine.validate_file(cleaned_df)
    # WARN: prefix = warnings (allow processing), rest = errors (block)
    warnings += [i[5:] for i in all_issues if i.startswith("WARN:")]
    errors    = [i for i in all_issues if not i.startswith("WARN:")]
    is_valid  = len(errors) == 0
    return is_valid, errors, warnings, cleaned_df

def generate_report(df, output_path, report_title=None):
    if engine is None:
        raise RuntimeError("Engine not loaded")
    import inspect
    sig = inspect.signature(engine.build_report)
    if 'report_title' in sig.parameters:
        engine.build_report(df, output_path, report_title=report_title)
    else:
        engine.build_report(df, output_path)

def read_sipl_file(uploaded_file):
    """Read uploaded file - tries ALL_SIPL_Ratings_List sheet, falls back to first sheet"""
    try:
        xl = pd.ExcelFile(uploaded_file)
        # Use ALL_SIPL_Ratings_List if present, otherwise first sheet
        if 'ALL_SIPL_Ratings_List' in xl.sheet_names:
            df = xl.parse('ALL_SIPL_Ratings_List')
        else:
            df = xl.parse(xl.sheet_names[0])
        return df, None
    except Exception as e:
        return None, str(e)

def fmt_size(n_bytes):
    if n_bytes < 1024: return f"{n_bytes} B"
    elif n_bytes < 1024**2: return f"{n_bytes/1024:.1f} KB"
    return f"{n_bytes/1024**2:.1f} MB"


# ── HERO HEADER ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-header">
    <p class="hero-title">HiRATE <span>Audit</span> Report Generator</p>
    <p class="hero-sub">Automated audit report generation · Validates data · Produces Excel reports with charts</p>
</div>
""", unsafe_allow_html=True)


# ── LAYOUT: Left column (upload + rules) | Right column (results) ─────────────
left_col, right_col = st.columns([1, 1.1], gap="large")

with left_col:

    # Upload section
    st.markdown('<div class="card-title">📂 Upload Files</div>', unsafe_allow_html=True)
    st.markdown('<p class="info-text">Upload one or more xlsx files. The <code>ALL_SIPL_Ratings_List</code> sheet will be used if present, otherwise the first sheet.</p>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Drop SIPL xlsx files here",
        type=["xlsx"],
        accept_multiple_files=True,
        key="uploader",
        label_visibility="collapsed",
    )

    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)

    # Validation rules reference
    st.markdown('<div class="card-title">📋 Validation Rules</div>', unsafe_allow_html=True)
    rules = [
        ("R1", "Asset Type = 'Median Opening' → Row must be deleted"),
        ("R2", "HO Rating = '-' with HO Remarks ≠ '-' → Invalid"),
        ("R3", "HO Rating = 10 with HO Remarks ≠ '-' → Invalid"),
        ("R4", "HO Rating = 1 or 5 with HO Remarks = '-' or 'NA' → Invalid"),
    ]
    for num, desc in rules:
        st.markdown(f"""
        <div class="rule-item">
            <span class="rule-num">{num}</span>
            <span>{desc}</span>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)

    # Report contents
    st.markdown('<div class="card-title">📊 Report Contents</div>', unsafe_allow_html=True)
    st.markdown("""
    <p class="info-text">
    Each generated report contains:<br>
    • Main table — 39 asset rows with totals &amp; % of issues<br>
    • Category summary table (6 categories + overall)<br>
    • Overall satisfactory vs observations<br>
    • Chart 1 — Category-wise clustered bar<br>
    • Chart 2 — Observations summary pie<br>
    • Chart 3 — Division-wise bar (sorted by issues, zeros removed)
    </p>
    """, unsafe_allow_html=True)


# ── RIGHT COLUMN: File status + generate ─────────────────────────────────────
with right_col:

    if not uploaded_files:
        st.markdown("""
        <div style="text-align:center; padding: 80px 20px; color: rgba(255,255,255,0.25);">
            <div style="font-size:4rem; margin-bottom:16px;">📁</div>
            <div style="font-size:1rem; font-weight:600; letter-spacing:1px;">AWAITING FILES</div>
            <div style="font-size:0.8rem; margin-top:8px;">Upload xlsx files on the left to begin</div>
        </div>
        """, unsafe_allow_html=True)

    else:
        # ── Scan all files ──
        file_results = []
        for uf in uploaded_files:
            uf.seek(0)
            df, err = read_sipl_file(uf)
            if err:
                file_results.append({
                    "file": uf, "name": uf.name, "size": uf.size,
                    "df": None, "valid": False, "errors": [f"Read error: {err}"], "warnings": [], "status": "error"
                })
            else:
                is_valid, errors, warnings, cleaned_df = validate_file_ui(df)
                status = "ok" if (is_valid and not warnings) else ("warn_ok" if is_valid else "warn")
                file_results.append({
                    "file": uf, "name": uf.name, "size": uf.size,
                    "df": cleaned_df, "valid": is_valid, "errors": errors, "warnings": warnings,
                    "status": status
                })

        # ── Summary stats ──
        n_total = len(file_results)
        n_ok = sum(1 for r in file_results if r["status"] in ("ok", "warn_ok"))
        n_warn = sum(1 for r in file_results if r["status"] == "warn")
        n_err = sum(1 for r in file_results if r["status"] == "error")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="stat-box"><div class="stat-num">{n_total}</div><div class="stat-label">Files</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="stat-box"><div class="stat-num green">{n_ok}</div><div class="stat-label">Ready</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="stat-box"><div class="stat-num orange">{n_warn}</div><div class="stat-label">Issues</div></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="stat-box"><div class="stat-num red">{n_err}</div><div class="stat-label">Errors</div></div>', unsafe_allow_html=True)

        st.markdown('<div class="sep"></div>', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📄 File Status</div>', unsafe_allow_html=True)

        # ── Per-file status ──
        for r in file_results:
            badge_map = {
                "ok":      ("badge-ok",   "✓ READY"),
                "warn_ok": ("badge-warn",  "⚠ READY WITH WARNINGS"),
                "warn":    ("badge-error", "⚠ VALIDATION FAILED"),
                "error":   ("badge-error", "✗ ERROR"),
            }
            badge_cls, badge_txt = badge_map.get(r["status"], ("badge-warn", "UNKNOWN"))
            file_cls = "ok" if r["status"] in ("ok","warn_ok") else ("warn" if r["status"]=="warn" else "error")
            st.markdown(f"""
            <div class="file-item file-{file_cls}">
                <span class="file-icon">📊</span>
                <span style="flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">{r['name']}</span>
                <span style="color:rgba(255,255,255,0.35); font-size:0.75rem;">{fmt_size(r['size'])}</span>
                <span class="badge {badge_cls}">{badge_txt}</span>
            </div>""", unsafe_allow_html=True)

            for w in r.get("warnings", []):
                st.markdown(f'<div class="val-issue" style="color:rgba(255,200,80,0.9);">⚠ {w}</div>', unsafe_allow_html=True)
            for e in r.get("errors", []):
                st.markdown(f'<div class="val-issue">⚑ {e}</div>', unsafe_allow_html=True)

        st.markdown('<div class="sep"></div>', unsafe_allow_html=True)

        # ── Only generate for valid files ──
        valid_files = [r for r in file_results if r["status"] in ("ok", "warn_ok")]

        if not valid_files:
            st.markdown("""
            <div class="val-fail">
                ⚠ No files passed validation. Fix the issues above before generating reports.
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="card-title">🚀 Generate Reports ({len(valid_files)} file{"s" if len(valid_files)>1 else ""})</div>', unsafe_allow_html=True)

            if st.button(f"⚡ Generate {len(valid_files)} Report{'s' if len(valid_files)>1 else ''}", use_container_width=True):

                progress_bar = st.progress(0)
                status_text = st.empty()
                generated = []

                for idx, r in enumerate(valid_files):
                    status_text.markdown(f'<p style="color:rgba(255,255,255,0.6); font-size:0.85rem;">Processing: <code>{r["name"]}</code>...</p>', unsafe_allow_html=True)

                    try:
                        # Write uploaded file to temp location (needed as input_path)
                        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_in:
                            r["file"].seek(0)
                            tmp_in.write(r["file"].read())
                            tmp_in_path = tmp_in.name

                        out_name = Path(r["name"]).stem + "_REPORT.xlsx"
                        tmp_out_path = tempfile.mktemp(suffix=".xlsx")

                        generate_report(r["df"], tmp_out_path, report_title=Path(r["name"]).stem)

                        with open(tmp_out_path, "rb") as f:
                            report_bytes = f.read()

                        generated.append({"name": out_name, "bytes": report_bytes, "source": r["name"]})
                        os.unlink(tmp_in_path)
                        os.unlink(tmp_out_path)

                    except Exception as e:
                        generated.append({"name": r["name"], "bytes": None, "error": str(e)})

                    progress_bar.progress((idx + 1) / len(valid_files))
                    time.sleep(0.1)

                status_text.empty()
                progress_bar.empty()

                # ── Download section ──
                n_success = sum(1 for g in generated if g.get("bytes"))
                st.markdown(f"""
                <div class="val-pass">
                    ✓ {n_success} report{"s" if n_success != 1 else ""} generated successfully!
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

                # Show individual xlsx download buttons
                for g in generated:
                    if g.get("bytes"):
                        st.download_button(
                            label=f"⬇  Download  {g['name']}",
                            data=g["bytes"],
                            file_name=g["name"],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{g['name']}",
                            use_container_width=True,
                        )
                    else:
                        st.markdown(f'<div class="val-fail">✗ Failed: {g["name"]}<div class="val-issue">{g.get("error","")}</div></div>', unsafe_allow_html=True)

                # ── Generate PPT from all successful reports ──
                success_reports = [(g["name"], g["bytes"]) for g in generated if g.get("bytes")]
                ppt_bytes = None
                html_bytes = None
                if success_reports:
                    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)
                    st.markdown('<div class="card-title">📊 PowerPoint Report</div>', unsafe_allow_html=True)
                    ppt_status = st.empty()
                    ppt_status.markdown('<p style="color:rgba(255,255,255,0.6); font-size:0.85rem;">⏳ Generating PowerPoint...</p>', unsafe_allow_html=True)
                    try:
                        ppt_engine, ppt_load_err = load_ppt_engine()
                        if ppt_engine is None:
                            ppt_status.markdown(f'<div class="val-fail">✗ {ppt_load_err}</div>', unsafe_allow_html=True)
                        else:
                            tmp_ppt = tempfile.mktemp(suffix=".pptx")
                            ok = ppt_engine.generate_ppt_from_reports(success_reports, tmp_ppt)
                            if ok and os.path.exists(tmp_ppt):
                                with open(tmp_ppt, "rb") as fp:
                                    ppt_bytes = fp.read()
                                os.unlink(tmp_ppt)
                                ppt_status.empty()
                                n_proj = len(success_reports)
                                st.markdown(f'''<div class="val-pass">✓ PowerPoint generated — {n_proj} project{"s" if n_proj!=1 else ""} · {n_proj*2} slides</div>''', unsafe_allow_html=True)
                                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                                st.download_button(
                                    label="⬇  Download  HiRATE_Report.pptx",
                                    data=ppt_bytes,
                                    file_name="HiRATE_Report.pptx",
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                    key="dl_ppt",
                                    use_container_width=True,
                                )
                            else:
                                ppt_status.markdown('<div class="val-fail">✗ PPT generation failed — check that python-pptx is installed</div>', unsafe_allow_html=True)
                    except Exception as ppt_err:
                        ppt_status.markdown(f'<div class="val-fail">✗ PPT error: {ppt_err}</div>', unsafe_allow_html=True)

                    # ── HTML Dashboard ──
                    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)
                    st.markdown('<div class="card-title">🌐 Interactive HTML Dashboard</div>', unsafe_allow_html=True)
                    html_status = st.empty()
                    html_status.markdown('<p style="color:rgba(255,255,255,0.6); font-size:0.85rem;">⏳ Generating dashboard...</p>', unsafe_allow_html=True)
                    try:
                        dash_path = Path(__file__).parent / "sipl_dashboard.py"
                        if not dash_path.exists():
                            html_status.markdown('<div class="val-fail">✗ sipl_dashboard.py not found</div>', unsafe_allow_html=True)
                        else:
                            import importlib.util
                            spec = importlib.util.spec_from_file_location("sipl_dashboard", dash_path)
                            dash_mod = importlib.util.module_from_spec(spec)
                            spec.loader.exec_module(dash_mod)
                            tmp_html = tempfile.mktemp(suffix=".html")
                            ok2 = dash_mod.generate_dashboard_from_reports(success_reports, tmp_html)
                            if ok2 and os.path.exists(tmp_html):
                                with open(tmp_html, "rb") as fp:
                                    html_bytes = fp.read()
                                os.unlink(tmp_html)
                                html_status.empty()
                                st.markdown(f'<div class="val-pass">✓ Dashboard generated — {len(success_reports)} project{"s" if len(success_reports)!=1 else ""}</div>', unsafe_allow_html=True)
                                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                                st.download_button(
                                    label="⬇  Download  HiRATE_Dashboard.html",
                                    data=html_bytes,
                                    file_name="HiRATE_Dashboard.html",
                                    mime="text/html",
                                    key="dl_html",
                                    use_container_width=True,
                                )
                            else:
                                html_status.markdown('<div class="val-fail">✗ Dashboard generation failed</div>', unsafe_allow_html=True)
                    except Exception as html_err:
                        html_status.markdown(f'<div class="val-fail">✗ Dashboard error: {html_err}</div>', unsafe_allow_html=True)

                    # ── Word Summary ──
                    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)
                    st.markdown('<div class="card-title">📝 Word Summary Report</div>', unsafe_allow_html=True)
                    summary_status = st.empty()
                    summary_status.markdown('<p style="color:rgba(255,255,255,0.6); font-size:0.85rem;">⏳ Generating Word summary...</p>', unsafe_allow_html=True)
                    summary_bytes = None
                    try:
                        summary_path = Path(__file__).parent / "sipl_summary.py"
                        if not summary_path.exists():
                            summary_status.markdown('<div class="val-fail">✗ sipl_summary.py not found</div>', unsafe_allow_html=True)
                        else:
                            import importlib.util
                            spec = importlib.util.spec_from_file_location("sipl_summary", summary_path)
                            sum_mod = importlib.util.module_from_spec(spec)
                            spec.loader.exec_module(sum_mod)
                            tmp_docx = tempfile.mktemp(suffix=".docx")
                            ok3 = sum_mod.generate_summary_from_reports(success_reports, tmp_docx)
                            if ok3 and os.path.exists(tmp_docx):
                                with open(tmp_docx, "rb") as fp:
                                    summary_bytes = fp.read()
                                os.unlink(tmp_docx)
                                summary_status.empty()
                                st.markdown(f'<div class="val-pass">✓ Word summary generated — {len(success_reports)} project{"s" if len(success_reports)!=1 else ""}</div>', unsafe_allow_html=True)
                                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                                st.download_button(
                                    label="⬇  Download  HiRATE_Summary.docx",
                                    data=summary_bytes,
                                    file_name="HiRATE_Summary.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key="dl_summary",
                                    use_container_width=True,
                                )
                            else:
                                summary_status.markdown('<div class="val-fail">✗ Summary generation failed</div>', unsafe_allow_html=True)
                    except Exception as sum_err:
                        summary_status.markdown(f'<div class="val-fail">✗ Summary error: {sum_err}</div>', unsafe_allow_html=True)

                # ── ZIP: all reports + PPT + HTML in one folder ──
                if success_reports:
                    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)
                    st.markdown('<div class="card-title">📦 Download Everything</div>', unsafe_allow_html=True)

                    zip_buf = io.BytesIO()
                    folder_name = "HiRATE_Reports"
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for g in generated:
                            if g.get("bytes"):
                                zf.writestr(f"{folder_name}/{g['name']}", g["bytes"])
                        if ppt_bytes:
                            zf.writestr(f"{folder_name}/HiRATE_Report.pptx", ppt_bytes)
                        if html_bytes:
                            zf.writestr(f"{folder_name}/HiRATE_Dashboard.html", html_bytes)
                        if summary_bytes:
                            zf.writestr(f"{folder_name}/HiRATE_Summary.docx", summary_bytes)
                    zip_buf.seek(0)

                    st.download_button(
                        label="⬇  Download All  (HiRATE_Reports.zip)",
                        data=zip_buf.getvalue(),
                        file_name="HiRATE_Reports.zip",
                        mime="application/zip",
                        key="dl_zip",
                        use_container_width=True,
                    )
