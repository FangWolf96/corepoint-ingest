import io
import re
import statistics
from datetime import datetime, timezone, date

from bs4 import BeautifulSoup
from flask import Flask, request, render_template_string, send_file
import pandas as pd

app = Flask(__name__)

# ---------- Config you can tweak ----------
EXCLUDED_COLUMNS_FOR_AGE = {"Completed", "Canceled", "New Parts Request"}
REQUESTED_LANES = [
    "Receipt Confirmed <7 days", "Aging >7 Days", "Stale >14 Days",
    "Assigned to Department", "Parts Ordered", "Arrived/In-Hand",
    "Contacted", "Customer Unreachable", "Scheduled"
]
FOCUSED_LABELS = [
    "Demand Repair", "Install", "Dispatch",
    "NOT COOLING/HEATING", "Warranty", "Comfort Shield Warranty"
]
ALL_LABELS = [
    "Backordered","Warranty","Escalation","Prepaid Service","COSTCO ESCALATION",
    "Demand Repair","1st Year Warranty","Install","Comfort Shield Warranty","Recall",
    "Schedule Return Visit","Electrical","Duct Cleaning","Senior Tech Callback",
    "Commercial","Possible Payment Issue","Replacement Opp","Manager Visit","Damage Claim",
    "NOT COOLING/HEATING","Multiple Systems","Missing Quote","READY TO SCHEDULE",
    "Missing Parts","URGENT","Dispatch","ISR - Service","Aging Card",
    "In Service Recovery - Stale","Call Customer","Check Warranty","Plumbing",
    "Truck Stock","Multiple Parts Request"
]
# -----------------------------------------

DATE_PAT = re.compile(r"Received[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})", re.IGNORECASE)
PRICE_PAT = re.compile(r"Quoted Price\s*[$:]*\s*([\d,]+)", re.IGNORECASE)

def parse_date(s: str):
    for fmt in ("%m/%d/%y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None

def extract_cards(html_bytes: bytes):
    """Return list of dicts: {column, text, date, age, price}"""
    soup = BeautifulSoup(html_bytes, "html.parser")
    today = datetime.now(timezone.utc).date()
    out = []

    for outer in soup.find_all("div", class_=lambda c: c and c.startswith("_outerWrapper_")):
        header = outer.find("div", class_=lambda c: c and c.startswith("_headerName_"))
        if not header:
            continue
        column = header.get_text(" ", strip=True)

        cards = outer.select("div.card") or outer.find_all("div", recursive=True)

        for card in cards:
            text = card.get_text(" ", strip=True)
            if not text:
                continue
            m = DATE_PAT.search(text)
            if not m:
                continue
            d = parse_date(m.group(1))
            if not d:
                continue
            age = (today - d).days

            pm = PRICE_PAT.search(text)
            price = int(pm.group(1).replace(",", "")) if pm else None

            out.append({"column": column, "text": text, "date": d, "age": age, "price": price})
    return out

def avg(xs):
    return round(statistics.mean(xs), 2) if xs else 0.0

def build_report(cards: list[dict]):
    active = [c for c in cards if c["column"] not in EXCLUDED_COLUMNS_FOR_AGE]

    scope_rows = [{
        "Scope": "All cards (excluding Completed/Canceled)",
        "Count": len(active),
        "Average Age (days)": avg([c["age"] for c in active])
    }]

    for lbl in FOCUSED_LABELS:
        ages = [c["age"] for c in active if lbl.lower() in c["text"].lower()]
        scope_rows.append({
            "Scope": lbl,
            "Count": len(ages),
            "Average Age (days)": avg(ages) if ages else 0
        })

    lane_rows = []
    for lane in REQUESTED_LANES:
        ages = [c["age"] for c in cards if c["column"] == lane]
        if ages:
            lane_rows.append({"Lane": lane, "Count": len(ages), "Average Age (days)": avg(ages)})

    active_prices = [c["price"] for c in active if c["price"] is not None]
    completed_prices = [c["price"] for c in cards if c["column"] == "Completed" and c["price"] is not None]
    canceled_prices  = [c["price"] for c in cards if c["column"] == "Canceled"  and c["price"] is not None]

    quoted_stats = [{
        "Total Value Count": len(active_prices),
        "Total Value": sum(active_prices) if active_prices else 0,
        "Average Value": avg(active_prices) if active_prices else 0,
        "Total Won Count": len(completed_prices),
        "Total Won": sum(completed_prices) if completed_prices else 0,
        "Average Won": avg(completed_prices) if completed_prices else 0,
        "Total Lost Count": len(canceled_prices),
        "Total Lost": sum(canceled_prices) if canceled_prices else 0,
        "Average Lost": avg(canceled_prices) if canceled_prices else 0
    }]

    all_label_rows = []
    if ALL_LABELS:
        for lbl in ALL_LABELS:
            ages = [c["age"] for c in active if lbl.lower() in c["text"].lower()]
            all_label_rows.append({"Label": lbl, "Count": len(ages), "Average Age (days)": avg(ages) if ages else 0})

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame(scope_rows).to_excel(writer, sheet_name="Scope", index=False)
        pd.DataFrame(lane_rows).to_excel(writer, sheet_name="Lane", index=False)
        pd.DataFrame(quoted_stats).to_excel(writer, sheet_name="Quoted Prices", index=False)
        if all_label_rows:
            pd.DataFrame(all_label_rows).to_excel(writer, sheet_name="All Labels", index=False)
    output.seek(0)

    summary = {"scope": scope_rows, "lanes": lane_rows, "quoted": quoted_stats[0]}
    return output, summary

INDEX_TMPL = """<!doctype html>
<title>CorePoint Board Analyzer</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
 body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;max-width:900px;margin:40px auto;padding:0 16px}
 h1{font-size:1.6rem;margin:0 0 12px}
 form{border:1px solid #ddd;padding:16px;border-radius:12px}
 .row{margin:8px 0}
 .btn{padding:10px 16px;border:1px solid #222;border-radius:10px;background:#fff;cursor:pointer}
 table{border-collapse:collapse;width:100%;margin:16px 0}
 th,td{border:1px solid #ddd;padding:8px;text-align:left}
 th{background:#f7f7f7}
</style>
<h1>CorePoint Board Analyzer</h1>
<form method="post" action="/analyze" enctype="multipart/form-data">
  <div class="row">Upload your CorePoint HTML export:</div>
  <div class="row"><input type="file" name="file" accept=".html,.htm" required></div>
  <div class="row"><button class="btn" type="submit">Analyze</button></div>
</form>
<p style="color:#777;margin-top:10px">Defaults: exclude Completed, Canceled, New Parts Request from age stats. Prices are read from “Quoted Price $: …”.</p>
"""

RESULT_TMPL = """<!doctype html>
<title>CorePoint Report</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
 body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;max-width:1100px;margin:40px auto;padding:0 16px}
 h1{font-size:1.6rem;margin:0 0 12px}
 table{border-collapse:collapse;width:100%;margin:16px 0}
 th,td{border:1px solid #ddd;padding:8px;text-align:left}
 th{background:#f7f7f7}
 .section{margin-top:24px}
 .btn{padding:10px 16px;border:1px solid #222;border-radius:10px;background:#fff;cursor:pointer;text-decoration:none}
</style>
<h1>Report</h1>
<a class="btn" href="/download">Download Excel</a>

<div class="section">
<h2>Scope</h2>
<table>
  <tr><th>Scope</th><th>Count</th><th>Average Age (days)</th></tr>
  {% for r in scope %}
  <tr><td>{{r['Scope']}}</td><td>{{r['Count']}}</td><td>{{"%.2f"|format(r['Average Age (days)'])}}</td></tr>
  {% endfor %}
</table>
</div>

<div class="section">
<h2>Lane</h2>
<table>
  <tr><th>Lane</th><th>Count</th><th>Average Age (days)</th></tr>
  {% for r in lanes %}
  <tr><td>{{r['Lane']}}</td><td>{{r['Count']}}</td><td>{{"%.2f"|format(r['Average Age (days)'])}}</td></tr>
  {% endfor %}
</table>
</div>

<div class="section">
<h2>Quoted Prices</h2>
<table>
  <tr><th>Metric</th><th>Count</th><th>Value</th></tr>
  <tr><td>Total Value</td><td>{{quoted['Total Value Count']}}</td><td>${{quoted['Total Value']}}</td></tr>
  <tr><td>Average Value</td><td></td><td>${{"%.2f"|format(quoted['Average Value'])}}</td></tr>
  <tr><td>Total Won (Completed)</td><td>{{quoted['Total Won Count']}}</td><td>${{quoted['Total Won']}}</td></tr>
  <tr><td>Average Won</td><td></td><td>${{"%.2f"|format(quoted['Average Won'])}}</td></tr>
  <tr><td>Total Lost (Canceled)</td><td>{{quoted['Total Lost Count']}}</td><td>${{quoted['Total Lost']}}</td></tr>
  <tr><td>Average Lost</td><td></td><td>${{"%.2f"|format(quoted['Average Lost'])}}</td></tr>
</table>
</div>
"""

_last_workbook = None

@app.get("/")
def index():
    return render_template_string(INDEX_TMPL)

@app.post("/analyze")
def analyze():
    global _last_workbook
    file = request.files.get("file")
    if not file:
        return "No file uploaded", 400
    html = file.read()

    cards = extract_cards(html)
    xlsx, summary = build_report(cards)
    _last_workbook = xlsx

    return render_template_string(RESULT_TMPL, **summary)

@app.get("/download")
def download():
    global _last_workbook
    if not _last_workbook:
        return "No report available. Upload a file first.", 400
    _last_workbook.seek(0)
    return send_file(
        _last_workbook,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"planka_report_{datetime.now().date()}.xlsx",
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
