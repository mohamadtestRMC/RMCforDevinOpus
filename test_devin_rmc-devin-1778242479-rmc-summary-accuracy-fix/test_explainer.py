"""Quick test: run engine, explain a row, generate PDF."""
import io
from engine.fill_jobtrack import fill_jobtrack
from engine.explainer import explain_row, build_all_explanations
from engine.pdf_report import generate_row_pdf

BASE = "Template_Files"

with open(f"{BASE}/Jobtrack Feb Without MRR.xlsx", "rb") as f:
    jt = io.BytesIO(f.read())
with open(f"{BASE}/Stores Recordings.xlsx", "rb") as f:
    stores = io.BytesIO(f.read())
with open(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx", "rb") as f:
    pr = io.BytesIO(f.read())

output_bytes, results_log, fill_stats = fill_jobtrack(jt, stores, pr)
output_data = output_bytes.getvalue()

print(f"Engine done: {len(results_log)} log entries, {fill_stats}")

# Test row 7 (PRINTING)
expl = explain_row(output_data, results_log, 7)
print(f"\nRow 7 explanation:")
print(f"  UID: {expl['uid']}")
print(f"  Process: {expl['process']}")
print(f"  Confidence: {expl['confidence']}")
print(f"  Issues: {len(expl['issues'])}")
for i in expl['issues']:
    print(f"    - {i['title']} [{i['severity']}]")
print(f"  Rate breakdown entries: {len(expl['rate_breakdown'])}")

# Generate PDF
pdf_bytes = generate_row_pdf(expl)
with open("test_row_report.pdf", "wb") as f:
    f.write(pdf_bytes)
print(f"\nPDF saved: test_row_report.pdf ({len(pdf_bytes)} bytes)")

# Test row 15 (LAM)
expl_lam = explain_row(output_data, results_log, 15)
print(f"\nRow 15 explanation:")
print(f"  UID: {expl_lam['uid']}")
print(f"  Process: {expl_lam['process']}")
print(f"  Confidence: {expl_lam['confidence']}")
print(f"  Issues: {len(expl_lam['issues'])}")
for i in expl_lam['issues']:
    print(f"    - {i['title']} [{i['severity']}]")

pdf_lam = generate_row_pdf(expl_lam)
with open("test_row_report_lam.pdf", "wb") as f:
    f.write(pdf_lam)
print(f"LAM PDF saved: test_row_report_lam.pdf ({len(pdf_lam)} bytes)")

# Test row 9 (INH/WPE — has fallback)
expl_inh = explain_row(output_data, results_log, 9)
print(f"\nRow 9 (INH) explanation:")
print(f"  UID: {expl_inh['uid']}")
print(f"  Confidence: {expl_inh['confidence']}")
print(f"  Issues: {len(expl_inh['issues'])}")
for i in expl_inh['issues']:
    print(f"    - {i['title']} [{i['severity']}]: {i['detail'][:80]}")

pdf_inh = generate_row_pdf(expl_inh)
with open("test_row_report_inh.pdf", "wb") as f:
    f.write(pdf_inh)
print(f"INH PDF saved: test_row_report_inh.pdf ({len(pdf_inh)} bytes)")

# Build ALL explanations (for caching test)
all_expls = build_all_explanations(output_data, results_log)
print(f"\nAll explanations: {len(all_expls)} rows")

# Confidence breakdown
confs = {'HIGH': 0, 'MEDIUM': 0, 'LOW': 0}
for _, e in all_expls.items():
    confs[e['confidence']['level']] += 1
print(f"Confidence breakdown: {confs}")
