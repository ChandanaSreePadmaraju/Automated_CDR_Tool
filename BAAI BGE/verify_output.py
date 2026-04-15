"""
Full verification of the output document.
Checks every post-processing pass and structural requirement.
"""
import sys, zipfile, re
sys.path.insert(0, '.')
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

OUTPUT   = 'output/Filled_CDR_v5.docx'
TEMPLATE = 'Template_D001024021 CDR ISO 17664-2 (2021) ProductName RX.Y Rev C.docx'

PASS = "\033[92m[PASS]\033[0m"
FAIL = "\033[91m[FAIL]\033[0m"
WARN = "\033[93m[WARN]\033[0m"
INFO = "\033[94m[INFO]\033[0m"

results = []

def check(label, ok, detail=""):
    icon = PASS if ok else FAIL
    print(f"  {icon} {label}" + (f": {detail}" if detail else ""))
    results.append((label, ok))

doc = Document(OUTPUT)
body = doc.element.body

with zipfile.ZipFile(OUTPUT) as z:
    doc_xml  = z.read('word/document.xml').decode('utf-8')
    rels_xml = z.read('word/_rels/document.xml.rels').decode('utf-8')
    media    = [n for n in z.namelist() if 'media/' in n]

print("\n" + "="*65)
print(" 1. COVER PAGE")
print("="*65)

cover_tbl = list(body)[9]
rows = cover_tbl.findall('.//' + qn('w:tr'))
for ri, row in enumerate(rows):
    for ci, cell in enumerate(row.findall('.//' + qn('w:tc'))):
        txt  = ''.join(t.text or '' for t in cell.iter(qn('w:t'))).strip()
        rids = [b.get(qn('r:embed')) for b in cell.findall('.//' + qn('a:blip'))]
        if txt:
            check(f"Cover row{ri} col{ci} placeholder replaced",
                  'ProductName RX.Y' not in txt,
                  f"'{txt[:60]}'")
        if rids:
            defined = set(re.findall(r' Id="(rId\d+)"', rels_xml))
            all_ok  = all(r in defined for r in rids)
            check(f"Cover row{ri} col{ci} image rIds valid",
                  all_ok, str(rids))


print("\n" + "="*65)
print(" 2. IMAGES")
print("="*65)

embeds  = re.findall(r'r:embed="([^"]+)"', doc_xml)
defined = set(re.findall(r' Id="(rId\d+)"', rels_xml))
leftover_placeholders = [e for e in embeds if 'CDRCOPY' in e]
missing_embeds        = [e for e in embeds if e not in defined]

check("No leftover CDRCOPY placeholder rIds", not leftover_placeholders,
      str(leftover_placeholders) if leftover_placeholders else "")
check("All r:embed refs resolve to a relationship", not missing_embeds,
      str(missing_embeds) if missing_embeds else "")
print(f"  {INFO} Media files: {media}")

img_rels = re.findall(r'Id="(rId\d+)"[^/]*/image[^/]*Target="([^"]+)"', rels_xml)
for rid, target in img_rels:
    in_zip = 'word/' + target.lstrip('../') in [m for m in media] or \
             target.lstrip('../') in [m.replace('word/','') for m in media] or \
             any(target.split('/')[-1] in m for m in media)
    check(f"Image {rid} target exists in zip", in_zip, target)


print("\n" + "="*65)
print(" 3. PLACEHOLDERS REMAINING IN BODY")
print("="*65)

leftover_body = []
for p in doc.paragraphs:
    t = p.text.strip()
    if '<ProductName RX.Y>' in t:
        leftover_body.append(t[:80])
# Also check tables
for tbl in doc.tables:
    for row in tbl.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                t = p.text.strip()
                if '<ProductName RX.Y>' in t:
                    leftover_body.append(t[:80])

check("No <ProductName RX.Y> left in body", not leftover_body,
      str(leftover_body) if leftover_body else "")

# Also check raw XML for the placeholder
raw_leftover = doc_xml.count('ProductName RX.Y')
check("No ProductName RX.Y in raw XML", raw_leftover == 0,
      f"{raw_leftover} occurrences remain")


print("\n" + "="*65)
print(" 4. HEADERS AND FOOTERS")
print("="*65)

hf_issues = []
for si, section in enumerate(doc.sections):
    for part_name, part in [
        ('header', section.header),
        ('footer', section.footer),
    ]:
        try:
            for p in part.paragraphs:
                if '<ProductName RX.Y>' in p.text:
                    hf_issues.append(f"section{si} {part_name}: {p.text[:60]}")
            for tbl in part.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if '<ProductName RX.Y>' in p.text:
                                hf_issues.append(f"section{si} {part_name} table: {p.text[:60]}")
        except Exception:
            pass

check("No <ProductName RX.Y> in headers/footers", not hf_issues,
      str(hf_issues) if hf_issues else "")


print("\n" + "="*65)
print(" 5. POST-PROCESSING: GUIDANCE STYLE REMOVED")
print("="*65)

guidance_paras = [p for p in doc.paragraphs if p.style.name == 'Guidance']
check("All Guidance-style paragraphs removed", len(guidance_paras) == 0,
      f"{len(guidance_paras)} remaining" if guidance_paras else "")


print("\n" + "="*65)
print(" 6. POST-PROCESSING: TEXT COLOR")
print("="*65)

color_elems = body.findall('.//' + qn('w:color'))
non_auto = [e for e in color_elems if e.get(qn('w:val'), '').lower() not in ('auto', '', 'ffffff')]
check("No colored text in body (all text black or auto)", len(non_auto) == 0,
      f"{len(non_auto)} color runs remain" if non_auto else "")


print("\n" + "="*65)
print(" 7. POST-PROCESSING: PRE-CDR SECTION REMOVED")
print("="*65)

pre_cdr_headings = [p for p in doc.paragraphs
                    if p.style.name.startswith('Heading') and 'Pre-created CDR history' in p.text]
check("'Pre-created CDR history' heading removed", len(pre_cdr_headings) == 0,
      "still present" if pre_cdr_headings else "")


print("\n" + "="*65)
print(" 8. SECTION FILL STATUS")
print("="*65)

import json
with open('prompts.json') as f:
    mappings = json.load(f).get('heading_mappings', {})

from src.heading_extractor import extract_headings
template_headings = extract_headings(TEMPLATE)
data_headings     = extract_headings(OUTPUT)

print(f"  {INFO} Template has {len(template_headings)} headings")
print(f"  {INFO} Output has   {len(data_headings)} headings")

output_heading_texts = {h['text'] for h in data_headings}
for h in template_headings:
    txt = h['text']
    if txt == 'Pre-created CDR history':
        check(f"Heading '{txt}' correctly removed", txt not in output_heading_texts)
    else:
        check(f"Heading '{txt}' present in output", txt in output_heading_texts,
              "MISSING" if txt not in output_heading_texts else "")


print("\n" + "="*65)
print(" 9. TABLE STRUCTURE")
print("="*65)

tables = doc.tables
print(f"  {INFO} Total tables in output: {len(tables)}")
empty_data_rows = 0
for ti, tbl in enumerate(tables):
    for ri, row in enumerate(tbl.rows):
        if ri == 0:
            continue  # skip header
        txt = ''.join(c.text.strip() for c in row.cells)
        if not txt:
            empty_data_rows += 1
check("No empty data rows in tables", empty_data_rows == 0,
      f"{empty_data_rows} empty rows found" if empty_data_rows else "")


print("\n" + "="*65)
print(" 10. RISK ANALYSIS / COMPLIANCE CHECKLIST CONTENT")
print("="*65)

risk_count  = doc_xml.count('Risk analysis')
valid_count = doc_xml.count('Validation of the processes')
check("'Risk analysis' content present", risk_count > 0, f"{risk_count} occurrences")
check("'Validation of the processes' present", valid_count > 0, f"{valid_count} occurrences")


print("\n" + "="*65)
print(" SUMMARY")
print("="*65)
passed  = sum(1 for _, ok in results if ok)
failed  = sum(1 for _, ok in results if not ok)
total   = len(results)
print(f"  Passed: {passed}/{total}")
if failed:
    print(f"  Failed checks:")
    for label, ok in results:
        if not ok:
            print(f"    - {label}")
print()
