# -*- coding: utf-8 -*-
"""
Web App: AI Assistant - Phan bo chi phi
Rules: I.1 Phan bo chi phi dien, I.2 Phan bo chi phi nuoc uong
"""
import os
import re
import io
import tempfile
import shutil
from datetime import datetime
import calendar

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel

import pdfplumber
import openpyxl

try:
    import pytesseract
    from PIL import Image
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

# ==========================================
# RULES CONFIG
# ==========================================
RULES = {
    "dien": {
        "name": "Phan bo chi phi dien",
        "data_start_row": 5,
        "col_code": 8,
        "col_fill": 15,
        "sheet_index": 1,
        "excel_prefix": "Bang phan bo thanh toan chi phi dien",
        "title_format": "BANG PHAN BO CHI PHI THANG {month}/2026",
        "sheet_format": "Phan bo_{month}.26",
        "has_date_shift": True,
        "col_ky_tt": 13,
        "col_ky_tt_truoc": 14,
        "split_duplicates": False,
        "has_di_doi": True,
        "ref_month": 7,
    },
    "nuoc": {
        "name": "Phan bo chi phi nuoc uong",
        "data_start_row": 5,
        "col_code": 2,
        "col_fill": 8,
        "sheet_index": 2,
        "excel_prefix": "Bang phan bo chi phi nuoc uong",
        "title_format": "BANG PHAN BO CHI PHI NUOC UONG THANG {month}/2026",
        "sheet_format": "T{month}",
        "has_date_shift": False,
        "col_ky_tt": 13,
        "col_ky_tt_truoc": 14,
        "split_duplicates": True,
        "has_di_doi": False,
        "ref_month": 7,
    },
}

# ==========================================
# APP
# ==========================================
app = FastAPI(title="AI Assistant - Phan bo chi phi", version="2.0")

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.get("/")
async def root():
    return {
        "app": "AI Assistant - Phan bo chi phi",
        "version": "2.0",
        "rules": list(RULES.keys()),
        "endpoints": {
            "health": "/health",
            "docs": "/docs",
            "process": "/process (POST)",
            "download": "/process/download (POST)",
        }
    }

# ==========================================
# HELPER FUNCTIONS
# ==========================================
def add_months(date_str, n):
    try:
        d = datetime.strptime(date_str.strip(), '%d/%m/%Y')
        total = d.year * 12 + d.month - 1 + n
        y, m = total // 12, total % 12 + 1
        return d.replace(year=y, month=m, day=min(d.day, calendar.monthrange(y, m)[1])).strftime('%d/%m/%Y')
    except:
        return None

def shift_date_range(s, n):
    if not s: return None
    m = re.match(r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', str(s).strip())
    if not m: return None
    a, b = add_months(m.group(1), n), add_months(m.group(2), n)
    return f"{a} - {b}" if a and b else None

def _code(text):
    if not text: return ""
    t = str(text).strip()
    m = re.match(r'^([A-Z]{2}[0-9A-Z]+)', t) or re.search(r'\b([A-Z]{2}\d{8,})\b', t)
    return m.group(1) if m else ""

def _amt(d, c):
    s = ""
    if d and d != '0': s = d.replace(',', '').replace('.00', '').strip()
    elif c and c != '0': s = c.replace(',', '').replace('.00', '').strip()
    return float(s) if s else 0

# ==========================================
# DATA EXTRACTION
# ==========================================
def extract_pdf_dien(fp):
    """Extract from electricity PDF (bank statement format)"""
    data = []
    with pdfplumber.open(fp) as pdf:
        for p in pdf.pages:
            for t in p.extract_tables():
                for r in t:
                    if not r or not r[0]: continue
                    if not str(r[0]).strip().isdigit(): continue
                    code = _code(str(r[5]) if r[5] else "")
                    amt = _amt(str(r[6]) if r[6] else "", str(r[7]) if r[7] else "")
                    if code and code != 'MSB' and amt > 0:
                        data.append({'code': code, 'amount': amt})
    return data

def extract_pdf_nuoc(fp):
    """Extract from water PDF (sales detail format).
    1. Find customer code (C\d+) in table col 0 -> take col 6 (So Luong = cot thu 7)
    2. For codes split across pages -> get from 'Theo mat hang/By SKU:' text
    3. If code not found -> left empty in Excel
    """
    result = {}
    with pdfplumber.open(fp) as pdf:
        all_text = ""
        all_tables = []
        for page in pdf.pages:
            all_text += (page.extract_text() or "") + "\n"
            for table in page.extract_tables():
                all_tables.append(table)

        # STEP 1: Find codes in table (col 0 = code, col 6 = So Luong)
        for table in all_tables:
            for row in table:
                if not row or not row[0]: continue
                m = re.match(r'^(C\d+)', str(row[0]).strip())
                if not m: continue
                code = m.group(1)
                qty_str = str(row[6]).strip() if len(row) > 6 and row[6] else ""
                try: qty = float(qty_str.replace(',', ''))
                except: qty = 0
                if qty > 0 and code not in result:
                    result[code] = qty

        # STEP 2: Codes split across pages -> 'Theo mat hang/By SKU:' fallback
        lines = all_text.split('\n')
        current_code = None
        for i, line in enumerate(lines):
            line = line.strip()
            m = re.match(r'^(C\d+)', line)
            if m: current_code = m.group(1)
            if ('Theo mat hang' in line or 'By SKU' in line) and current_code:
                if current_code not in result and i + 1 < len(lines):
                    sku_line = lines[i + 1].strip()
                    nums = re.findall(r'(\d[\d,]*)', sku_line)
                    if nums:
                        try:
                            qty = float(nums[-1].replace(',', ''))
                            if qty > 0: result[current_code] = qty
                        except: pass

    return [{'code': k, 'amount': v} for k, v in result.items()]

def extract_excel_data(fp):
    data = []
    wb = openpyxl.load_workbook(fp, keep_vba=False, data_only=True)
    for ws in wb.worksheets:
        for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=5, values_only=True):
            if not r or not r[0]: continue
            code = str(r[0]).strip()
            amt = 0
            if len(r) > 2 and r[2]:
                try: amt = float(str(r[2]).replace(',', '').replace('.00', '').strip())
                except: pass
            if code and amt > 0: data.append({'code': code, 'amount': amt})
    wb.close()
    return data

def extract_image_data(fp):
    if not HAS_OCR: raise HTTPException(status_code=500, detail="OCR not available")
    tess_path = os.environ.get('TESSERACT_CMD', '/usr/bin/tesseract')
    pytesseract.pytesseract.tesseract_cmd = tess_path
    text = pytesseract.image_to_string(Image.open(fp), lang='eng')
    data = []
    for line in text.split('\n'):
        code = _code(line)
        nums = re.findall(r'[\d,]+\.?\d*', line)
        amt = float(nums[-1].replace(',', '')) if nums else 0
        if code and code != 'MSB' and amt > 0: data.append({'code': code, 'amount': amt})
    return data

def build_lookup(data_list):
    lk = {}
    for item in data_list:
        if item['code'] not in lk: lk[item['code']] = item
    return lk

# ==========================================
# PROCESS EXCEL
# ==========================================
def process_excel(excel_path, lookup, rule_key, month_shift=0):
    rule = RULES[rule_key]
    wb = openpyxl.load_workbook(excel_path, keep_vba=False)
    ws = wb.worksheets[rule["sheet_index"]]
    results = []
    updated = skipped_dd = skipped_nm = 0

    # Count code occurrences for split_duplicates
    code_counts = {}
    if rule["split_duplicates"]:
        for ri in range(rule["data_start_row"], ws.max_row + 1):
            c = ws.cell(row=ri, column=rule["col_code"]).value
            if c: code_counts[str(c).strip()] = code_counts.get(str(c).strip(), 0) + 1

    for ri in range(rule["data_start_row"], ws.max_row + 1):
        ec = ws.cell(row=ri, column=rule["col_code"]).value
        if not ec: continue
        ec = str(ec).strip()

        # Check di dời
        if rule["has_di_doi"]:
            c13 = ws.cell(row=ri, column=rule["col_ky_tt"]).value or ""
            c14 = ws.cell(row=ri, column=rule["col_ky_tt_truoc"]).value or ""
            if any(k in str(c13).lower() for k in ['di d', 'dời', 'ngừng']) or \
               any(k in str(c14).lower() for k in ['di d', 'dời', 'ngừng']):
                ws.cell(row=ri, column=rule["col_fill"]).value = None
                skipped_dd += 1
                results.append({'row': ri, 'code': ec, 'action': 'SKIP', 'amount': 'blank'})
                continue
            if month_shift and rule["has_date_shift"]:
                n13 = shift_date_range(c13, month_shift)
                if n13: ws.cell(row=ri, column=rule["col_ky_tt"]).value = n13
                n14 = shift_date_range(c14, month_shift)
                if n14: ws.cell(row=ri, column=rule["col_ky_tt_truoc"]).value = n14

        if ec in lookup:
            a = lookup[ec]['amount']
            if rule["split_duplicates"] and code_counts.get(ec, 1) > 1:
                a = a / code_counts[ec]
            ws.cell(row=ri, column=rule["col_fill"]).value = a
            updated += 1
            results.append({'row': ri, 'code': ec, 'action': 'UPDATED', 'amount': a})
        else:
            # Not found -> leave empty
            ws.cell(row=ri, column=rule["col_fill"]).value = None
            skipped_nm += 1
            results.append({'row': ri, 'code': ec, 'action': 'NO_MATCH', 'amount': 0})

    # Update title and sheet name
    month_num = str(7 + month_shift).zfill(2)
    ws.cell(row=3, column=1).value = rule["title_format"].format(month=month_num)
    ws.title = rule["sheet_format"].format(month=month_num)

    return wb, results, updated, skipped_dd, skipped_nm

# ==========================================
# API ENDPOINTS
# ==========================================
class ProcessResponse(BaseModel):
    rule: str
    updated: int
    skipped_di_doi: int
    no_match: int
    total: int
    results: list

@app.post("/process", response_model=ProcessResponse)
async def process(
    rule: str = Form(...),
    month: int = Form(...),
    input_file: UploadFile = File(...),
    excel_template: UploadFile = File(...),
):
    """Process input file and update Excel template.
    rule: 'dien' or 'nuoc'
    month: input month number (e.g., 6, 7, 9)
    """
    if rule not in RULES:
        raise HTTPException(status_code=400, detail=f"Invalid rule: {rule}. Use 'dien' or 'nuoc'")

    tmpdir = tempfile.mkdtemp()
    try:
        # Save uploaded files
        input_path = os.path.join(tmpdir, input_file.filename)
        with open(input_path, 'wb') as f: f.write(await input_file.read())
        excel_path = os.path.join(tmpdir, excel_template.filename)
        with open(excel_path, 'wb') as f: f.write(await excel_template.read())

        # Extract data based on file type and rule
        ext = os.path.splitext(input_file.filename)[1].lower()
        if ext == '.pdf':
            if rule == 'nuoc':
                data = extract_pdf_nuoc(input_path)
            else:
                data = extract_pdf_dien(input_path)
        elif ext in ('.xlsx', '.xls'):
            data = extract_excel_data(input_path)
        elif ext in ('.jpg', '.jpeg', '.png', '.gif'):
            data = extract_image_data(input_path)
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported file type: {ext}")

        lookup = build_lookup(data)
        month_shift = month - RULES[rule]["ref_month"]
        wb, results, updated, skipped_dd, skipped_nm = process_excel(excel_path, lookup, rule, month_shift)

        # Save output
        rule_cfg = RULES[rule]
        output_name = f"{rule_cfg['excel_prefix']} thang {month}.2026.xlsx"
        output_path = os.path.join(tmpdir, output_name)
        wb.save(output_path)
        wb.close()

        return ProcessResponse(
            rule=rule,
            updated=updated,
            skipped_di_doi=skipped_dd,
            no_match=skipped_nm,
            total=updated + skipped_dd + skipped_nm,
            results=results[:50],
        )
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

@app.post("/process/download")
async def process_download(
    rule: str = Form(...),
    month: int = Form(...),
    input_file: UploadFile = File(...),
    excel_template: UploadFile = File(...),
):
    """Process and return the updated Excel file for download."""
    if rule not in RULES:
        raise HTTPException(status_code=400, detail=f"Invalid rule: {rule}")

    tmpdir = tempfile.mkdtemp()
    try:
        input_path = os.path.join(tmpdir, input_file.filename)
        with open(input_path, 'wb') as f: f.write(await input_file.read())
        excel_path = os.path.join(tmpdir, excel_template.filename)
        with open(excel_path, 'wb') as f: f.write(await excel_template.read())

        ext = os.path.splitext(input_file.filename)[1].lower()
        if ext == '.pdf':
            data = extract_pdf_nuoc(input_path) if rule == 'nuoc' else extract_pdf_dien(input_path)
        elif ext in ('.xlsx', '.xls'):
            data = extract_excel_data(input_path)
        elif ext in ('.jpg', '.jpeg', '.png', '.gif'):
            data = extract_image_data(input_path)
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported: {ext}")

        lookup = build_lookup(data)
        month_shift = month - RULES[rule]["ref_month"]
        wb, results, updated, skipped_dd, skipped_nm = process_excel(excel_path, lookup, rule, month_shift)

        rule_cfg = RULES[rule]
        output_name = f"{rule_cfg['excel_prefix']} thang {month}.2026.xlsx"
        output_path = os.path.join(tmpdir, output_name)
        wb.save(output_path)
        wb.close()

        return FileResponse(output_path, filename=output_name,
                            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
