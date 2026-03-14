"""
sipl_ppt.py  —  HiRATE Report → PowerPoint generator
======================================================
Reads one or more *_REPORT.xlsx files (output of sipl_report.py),
extracts chart + summary data, and builds 2 slides per project into
a single cumulative HiRATE_Report.pptx:

  Slide 1 (per project):  Category-wise clustered bar chart
  Slide 2 (per project):  Division-wise bar chart  +  Observations pie

Run:
  python sipl_ppt.py                        # auto-finds *_REPORT.xlsx in same folder
  python sipl_ppt.py FILE1.xlsx FILE2.xlsx  # specific files
  python sipl_ppt.py --output MyDeck.pptx  # custom output name
"""

import os, sys, glob, json, subprocess, tempfile, argparse
import openpyxl

DEFAULT_OUTPUT = "HiRATE_Report.pptx"
SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))

# ══════════════════════════════════════════════════════════════════════════════
# DATA EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def extract_report_data(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb['Sheet1']

    # Project name from data sheet tab
    data_sheet = next((s for s in wb.sheetnames if s != 'Sheet1'), None)
    proj_name = "Project"
    if data_sheet:
        import re
        m = re.match(r'ALL_(.+?)_Ratings_List', data_sheet)
        proj_name = m.group(1) if m else data_sheet

    # Category-wise (rows 3-9, cols J-M)
    cat_labels, cat_total, cat_issues, cat_pct = [], [], [], []
    for r in range(3, 10):
        lbl = ws.cell(r, 10).value
        if lbl is None:
            continue
        cat_labels.append(str(lbl))
        cat_total.append(int(ws.cell(r, 11).value or 0))
        cat_issues.append(int(ws.cell(r, 12).value or 0))
        p = ws.cell(r, 13).value
        cat_pct.append(round((p or 0) * 100, 1))

    # Pie data
    satisfactory = int(ws['Q3'].value or 0)
    observations = int(ws['R3'].value or 0)

    # Division-wise staging at row 200+
    div_labels, div_total, div_issues, div_pct = [], [], [], []
    r = 201
    while True:
        asset = ws.cell(r, 1).value
        if asset is None:
            break
        div_labels.append(str(asset))
        div_total.append(int(ws.cell(r, 2).value or 0))
        div_issues.append(int(ws.cell(r, 3).value or 0))
        p = ws.cell(r, 4).value
        if isinstance(p, float) and p < 1.5:
            p = round(p * 100, 1)
        div_pct.append(float(p or 0))
        r += 1

    return {
        "proj_name":    proj_name,
        "cat_labels":   cat_labels,
        "cat_total":    cat_total,
        "cat_issues":   cat_issues,
        "cat_pct":      cat_pct,
        "satisfactory": satisfactory,
        "observations": observations,
        "div_labels":   div_labels,
        "div_total":    div_total,
        "div_issues":   div_issues,
        "div_pct":      div_pct,
    }

# ══════════════════════════════════════════════════════════════════════════════
# JS TEMPLATE
# ══════════════════════════════════════════════════════════════════════════════

JS_CODE = r"""
const pptxgen = require("pptxgenjs");
const projects = PROJECT_DATA;
const outFile  = OUTPUT_FILE;

const NAVY   = "1F4E79";
const ORANGE = "ED7D31";
const GREEN  = "92D050";
const WHITE  = "FFFFFF";
const GRAY   = "555555";
const BG_HDR = "1A2E45";
const BG_SLD = "F7F9FC";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "HiRATE Audit Report";

function addHeader(slide, title, proj) {
  slide.addShape(pres.shapes.RECTANGLE,
    { x:0, y:0, w:10, h:0.72, fill:{color:BG_HDR}, line:{color:BG_HDR} });
  slide.addText("HiRATE", {
    x:0.18, y:0, w:1.4, h:0.72,
    fontSize:20, bold:true, color:"00B0F0", fontFace:"Calibri", valign:"middle", margin:0
  });
  slide.addText(title, {
    x:1.7, y:0, w:7.0, h:0.72,
    fontSize:13, bold:true, color:WHITE, fontFace:"Calibri", valign:"middle", margin:0
  });
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE,
    { x:8.55, y:0.13, w:1.3, h:0.46, fill:{color:"0070C0"}, line:{color:"0070C0"}, rectRadius:0.07 });
  slide.addText(proj, {
    x:8.55, y:0.13, w:1.3, h:0.46,
    fontSize:10, bold:true, color:WHITE, fontFace:"Calibri", align:"center", valign:"middle", margin:0
  });
}

function addFooter(slide, lbl) {
  slide.addShape(pres.shapes.RECTANGLE,
    { x:0, y:5.42, w:10, h:0.2, fill:{color:"0D1B2A"}, line:{color:"0D1B2A"} });
  slide.addText("HiRATE Audit Report  |  Confidential", {
    x:0.2, y:5.43, w:7, h:0.18,
    fontSize:7, color:"888888", fontFace:"Calibri", valign:"middle", margin:0
  });
  slide.addText(lbl, {
    x:8.5, y:5.43, w:1.4, h:0.18,
    fontSize:7, color:"888888", fontFace:"Calibri", align:"right", valign:"middle", margin:0
  });
}

projects.forEach((p, pi) => {

  // ── SLIDE 1: Division-wise + Pie ────────────────────
  const s1 = pres.addSlide();
  s1.background = { color: BG_SLD };
  addHeader(s1, "HiRATE Observations — Division Wise & Summary", p.proj_name);

  s1.addText("HiRATE Observations Division wise — " + p.proj_name, {
    x:0.3, y:0.75, w:6.2, h:0.15,
    fontSize:8, color:"999999", italic:true, fontFace:"Calibri", margin:0
  });

  // Y-axis max: capped to ~1.5x the max No of Issues so orange bars are clearly visible.
  // Total Audited bars will go off the top (clipped) but their data labels still show.
  const divMaxIssues = Math.max(...p.div_issues, 1);
  const divYMax = Math.ceil(divMaxIssues * 1.5 / 5) * 5;

  // Single chart — Total Audited + No of Issues, Y-axis scaled to No of Issues
  s1.addChart(pres.charts.BAR, [
    { name:"Total Audited", labels:p.div_labels, values:p.div_total },
    { name:"No of Issues",  labels:p.div_labels, values:p.div_issues },
  ], {
    x:0.3, y:0.9, w:6.2, h:4.3,
    barDir:"col", barGrouping:"clustered", barGapWidthPct:80,
    chartColors:[NAVY, ORANGE],
    chartArea:{ fill:{color:WHITE}, roundedCorners:false },
    plotArea:{ fill:{color:WHITE} },
    showValue:true, dataLabelFontSize:7, dataLabelColor:"333333",
    catAxisLabelColor:GRAY, catAxisLabelFontSize:7,
    valAxisLabelColor:GRAY, valAxisLabelFontSize:8,
    valAxisMaxVal:divYMax,
    valGridLine:{color:"E0E0E0", size:0.5},
    catGridLine:{style:"none"},
    showLegend:true, legendPos:"b", legendFontSize:8,
    showTitle:false,
  });

  // Dashed divider
  s1.addShape(pres.shapes.LINE, {
    x:6.65, y:0.9, w:0, h:4.3,
    line:{ color:"CCCCCC", width:1, dashType:"dash" }
  });

  // Pie chart (right side)
  s1.addChart(pres.charts.PIE, [{
    name:"Summary",
    labels:["Satisfactory","Observations"],
    values:[p.satisfactory, p.observations]
  }], {
    x:6.75, y:0.9, w:3.0, h:2.8,
    chartColors:[GREEN, ORANGE],
    chartArea:{ fill:{color:WHITE}, roundedCorners:false },
    showPercent:true, showValue:true,
    dataLabelFontSize:10, dataLabelColor:"333333",
    showLegend:true, legendPos:"b", legendFontSize:9,
    showTitle:true, title:"OBSERVATIONS SUMMARY",
    titleFontSize:10, titleColor:"333333",
  });

  // Stat boxes below pie
  const total  = p.satisfactory + p.observations;
  const satPct = total > 0 ? ((p.satisfactory/total)*100).toFixed(1) : "0";
  const obsPct = total > 0 ? ((p.observations/total)*100).toFixed(1) : "0";

  s1.addShape(pres.shapes.RECTANGLE,
    { x:6.75, y:3.85, w:1.4, h:0.85, fill:{color:"E8F5E1"}, line:{color:GREEN, width:1.5} });
  s1.addText([
    { text:String(p.satisfactory), options:{fontSize:20, bold:true, color:"2E7D32", breakLine:true} },
    { text:satPct+"% Satisfactory",  options:{fontSize:8, color:GRAY} }
  ], { x:6.75, y:3.85, w:1.4, h:0.85, align:"center", valign:"middle", margin:0 });

  s1.addShape(pres.shapes.RECTANGLE,
    { x:8.35, y:3.85, w:1.4, h:0.85, fill:{color:"FFF3E0"}, line:{color:ORANGE, width:1.5} });
  s1.addText([
    { text:String(p.observations), options:{fontSize:20, bold:true, color:"E65100", breakLine:true} },
    { text:obsPct+"% Issues",        options:{fontSize:8, color:GRAY} }
  ], { x:8.35, y:3.85, w:1.4, h:0.85, align:"center", valign:"middle", margin:0 });

  addFooter(s1, "Slide " + (pi*2+1));

  // ── SLIDE 2: Category-wise ──────────────────────────
  const s2 = pres.addSlide();
  s2.background = { color: BG_SLD };
  addHeader(s2, "HiRATE Observations — Category Wise", p.proj_name);

  s2.addText("HiRATE Observations Category wise — " + p.proj_name, {
    x:0.3, y:0.75, w:9.4, h:0.15,
    fontSize:8, color:"999999", italic:true, fontFace:"Calibri", margin:0
  });

  // Category bar — only Total Audited + No of Issues (no % of issues)
  s2.addChart(pres.charts.BAR, [
    { name:"Total Audited", labels:p.cat_labels, values:p.cat_total },
    { name:"No of Issues",  labels:p.cat_labels, values:p.cat_issues },
  ], {
    x:0.3, y:0.9, w:9.4, h:4.3,
    barDir:"col", barGrouping:"clustered", barGapWidthPct:120,
    chartColors:[NAVY, ORANGE],
    chartArea:{ fill:{color:WHITE}, roundedCorners:false },
    plotArea:{ fill:{color:WHITE} },
    showValue:true, dataLabelFontSize:8, dataLabelColor:"333333",
    catAxisLabelColor:GRAY, catAxisLabelFontSize:9,
    valAxisLabelColor:GRAY, valAxisLabelFontSize:9,
    valGridLine:{color:"E0E0E0", size:0.5},
    catGridLine:{style:"none"},
    showLegend:true, legendPos:"b", legendFontSize:9,
    showTitle:false,
  });

  addFooter(s2, "Slide " + (pi*2+2));
});

pres.writeFile({ fileName: outFile })
  .then(() => console.log("PPT_OK:" + outFile))
  .catch(e  => { console.error("PPT_ERR:" + e); process.exit(1); });
"""

# ══════════════════════════════════════════════════════════════════════════════
# BUILD
# ══════════════════════════════════════════════════════════════════════════════

def build_ppt(report_files, output_path):
    projects = []
    for f in report_files:
        print(f"  Reading: {os.path.basename(f)}")
        try:
            d = extract_report_data(f)
            projects.append(d)
            print(f"    → {d['proj_name']}  cats={len(d['cat_labels'])}  divs={len(d['div_labels'])}")
        except Exception as e:
            print(f"    ⚠ Skipped: {e}")

    if not projects:
        print("No valid report files.")
        return False

    js = JS_CODE.replace("PROJECT_DATA", json.dumps(projects, ensure_ascii=False)) \
                .replace("OUTPUT_FILE",  json.dumps(output_path))

    with tempfile.NamedTemporaryFile(suffix=".js", mode="w", encoding="utf-8",
                                     dir=SCRIPT_DIR, delete=False) as tf:
        tf.write(js)
        js_path = tf.name

    try:
        IS_WIN   = sys.platform == "win32"
        npm_cmd  = ["npm.cmd", "install", "pptxgenjs"] if IS_WIN else ["npm", "install", "pptxgenjs"]
        node_cmd = ["node.exe"] if IS_WIN else ["node"]

        nm_path = os.path.join(SCRIPT_DIR, "node_modules", "pptxgenjs")
        if not os.path.isdir(nm_path):
            print("  Installing pptxgenjs (first run only)...")
            inst = subprocess.run(npm_cmd, capture_output=True, text=True, cwd=SCRIPT_DIR)
            if inst.returncode != 0:
                print("npm install failed:\n", inst.stdout + inst.stderr)
                return False
            print("  pptxgenjs installed successfully.")

        res = subprocess.run(node_cmd + [js_path], capture_output=True, text=True, cwd=SCRIPT_DIR)
        out = res.stdout + res.stderr
        if "PPT_OK:" in out:
            print(f"\n✓  Saved: {output_path}")
            print(f"   {len(projects)} project(s)  ×  2 slides  =  {len(projects)*2} slides total")
            return True
        else:
            print("Node error:\n", out)
            return False
    finally:
        os.unlink(js_path)


def find_report_files(folder):
    return sorted(glob.glob(os.path.join(folder, "*_REPORT.xlsx")))


def main():
    ap = argparse.ArgumentParser(description="HiRATE PPT generator")
    ap.add_argument("files", nargs="*", help="*_REPORT.xlsx files")
    ap.add_argument("--output", "-o", default=None, help="Output .pptx path")
    args = ap.parse_args()

    files = [os.path.abspath(f) for f in args.files] if args.files \
            else find_report_files(SCRIPT_DIR)

    if not files:
        print(f"ERROR: No *_REPORT.xlsx files found in {SCRIPT_DIR}")
        sys.exit(1)

    out = os.path.abspath(args.output) if args.output \
          else os.path.join(SCRIPT_DIR, DEFAULT_OUTPUT)

    print(f"\nHiRATE PPT Generator  ·  {len(files)} file(s)  →  {out}\n{'─'*50}")
    sys.exit(0 if build_ppt(files, out) else 1)


def generate_ppt_from_reports(report_bytes_list, output_path):
    """Called by sipl_app.py. report_bytes_list = [(filename, bytes), ...]"""
    tmp = []
    try:
        for fname, fbytes in report_bytes_list:
            tf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False,
                 prefix=os.path.splitext(fname)[0] + "_")
            tf.write(fbytes); tf.close()
            tmp.append(tf.name)
        return build_ppt(tmp, output_path)
    finally:
        for f in tmp:
            try: os.unlink(f)
            except: pass


if __name__ == "__main__":
    main()
