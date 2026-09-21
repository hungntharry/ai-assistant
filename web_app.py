# -*- coding: utf-8 -*-
"""
Web App: AI Assistant - Phan bo chi phi + Doi soat du lieu
I. Doi soat du lieu (dien, nuoc uong)
II. Phan bo chi phi (dien, nuoc uong)
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
# CONFIG
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
app = FastAPI(title="AI Assistant - Phan bo + Doi soat", version="3.0")

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.get("/")
async def root():
    return {
        "app": "AI Assistant",
        "version": "3.0",
        "menu": {
            "I": "Doi soat du lieu",
            "II": "Phan bo chi phi",
        },
        "endpoints": {
            "health": "/health",
            "docs": "/docs",
            "phan-bo": "/phan-bo (POST)",
            "doi-soat": "/doi-soat (POST)",
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
    result = {}
    with pdfplumber.open(fp) as pdf:
        all_text = ""
        all_tables = []
        for page in pdf.pages:
            all_text += (page.extract_text() or "") + "\n"
            for table in page.extract_tables():
                all_tables.append(table)
        for table in all_tables:
            for row in table:
                if not row or not row[0]: continue
                m = re.match(r'^(C\d+)', str(row[0]).strip())
                if not m: continue
                code = m.group(1)
                qty_str = str(row[6]).strip() if len(row) > 6 and row[6] else ""
                try: qty = float(qty_str.replace(',', ''))
                except: qty = 0
                if qty > 0 and code not in result: result[code] = qty
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

def extract_pdf_nuoc_by_pgh(fp):
    """Extract So PGH (9-digit) -> quantity from water PDF"""
    result = {}
    with pdfplumber.open(fp) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if 'PHIẾU GIAO HÀNG' in text or 'DELIVERY TICKET' in text:
                m = re.search(r'SỐ PHIẾU[/\s]*NO[\.\s]*\s*(\d{9})', text)
                if not m:
                    m = re.search(r'NO\.\s*(\d{9})', text)
                if not m:
                    lines = text.split('\n')
                    for i, line in enumerate(lines):
                        if 'SỐ PHIẾU' in line or 'NO.' in line:
                            for j in range(i, min(i + 4, len(lines))):
                                nums = re.findall(r'\d{9}', lines[j])
                                if nums:
                                    m = re.match(r'(\d{9})', nums[0])
                                    break
                            if m: break
                if m:
                    pgh = m.group(1) if m else ""
                    tables = page.extract_tables()
                    qty = 0
                    for table in tables:
                        if len(table) >= 2:
                            header = table[0]
                            row1 = table[1]
                            qty_col = -1
                            for ci, h in enumerate(header):
                                if h and 'lượng' in str(h).lower():
                                    qty_col = ci
                                    break
                            if qty_col >= 0 and len(row1) > qty_col and row1[qty_col]:
                                try: qty = float(str(row1[qty_col]).replace(',', ''))
                                except: qty = 0
                            break
                    if pgh and qty > 0 and pgh not in result:
                        result[pgh] = qty
            else:
                for table in page.extract_tables():
                    for row in table:
                        if not row: continue
                        col1 = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                        col2 = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                        col6 = str(row[6]).strip() if len(row) > 6 and row[6] else ""
                        if re.match(r'\d{2}/\d{2}/\d{4}', col1):
                            m = re.search(r'\d{9}', col2)
                            if m:
                                pgh = m.group(0)
                                try: qty = float(col6.replace(',', ''))
                                except: qty = 0
                                if pgh not in result: result[pgh] = qty
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

def extract_xl_by_pgh(fp):
    """Extract So PGH -> quantity from Excel. Col E=PGH, Col P=qty"""
    result = {}
    wb = openpyxl.load_workbook(fp, keep_vba=False, data_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
            if not row: continue
            pgh_val = row[4] if len(row) > 4 else None
            qty_val = row[15] if len(row) > 15 else None
            if pgh_val:
                m = re.search(r'\d{9}', str(pgh_val))
                if m and qty_val:
                    pgh = m.group(0)
                    try: qty = float(str(qty_val).replace(',', ''))
                    except: qty = 0
                    if qty > 0 and pgh not in result: result[pgh] = qty
    wb.close()
    return [{'code': k, 'amount': v} for k, v in result.items()]

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
# PHAN BO CHI PHI
# ==========================================
def process_phan_bo(excel_path, lookup, rule_key, month_shift=0):
    rule = RULES[rule_key]
    wb = openpyxl.load_workbook(excel_path, keep_vba=False)
    ws = wb.worksheets[rule["sheet_index"]]
    results = []
    updated = skipped_dd = skipped_nm = 0
    merged_ranges = list(ws.merged_cells.ranges)
    code_counts = {}
    if rule["split_duplicates"]:
        for ri in range(rule["data_start_row"], ws.max_row + 1):
            c = ws.cell(row=ri, column=rule["col_code"]).value
            if c: code_counts[str(c).strip()] = code_counts.get(str(c).strip(), 0) + 1

    for ri in range(rule["data_start_row"], ws.max_row + 1):
        ec = ws.cell(row=ri, column=rule["col_code"]).value
        if not ec or not str(ec).strip(): continue
        ec = str(ec).strip()
        if "HO HCM" in ec.upper():
            results.append({'row': ri, 'code': ec, 'action': 'KEEP FORMULA', 'amount': '(HO HCM)'})
            continue
        cell = ws.cell(row=ri, column=rule["col_fill"])
        if any(cell.coordinate in mr for mr in merged_ranges): continue
        if rule["has_di_doi"]:
            c13 = ws.cell(row=ri, column=rule["col_ky_tt"]).value or ""
            c14 = ws.cell(row=ri, column=rule["col_ky_tt_truoc"]).value or ""
            if any(k in str(c13).lower() for k in ['di d', 'dời', 'ngừng']) or \
               any(k in str(c14).lower() for k in ['di d', 'dời', 'ngừng']):
                cell.value = None; skipped_dd += 1
                results.append({'row': ri, 'code': ec, 'action': 'SKIP', 'amount': 'blank'}); continue
            if month_shift and rule["has_date_shift"]:
                n13 = shift_date_range(c13, month_shift)
                if n13: ws.cell(row=ri, column=rule["col_ky_tt"]).value = n13
                n14 = shift_date_range(c14, month_shift)
                if n14: ws.cell(row=ri, column=rule["col_ky_tt_truoc"]).value = n14
        if ec in lookup:
            a = lookup[ec]['amount']
            if rule["split_duplicates"] and code_counts.get(ec, 1) > 1: a = a / code_counts[ec]
            cell.value = a; updated += 1
            results.append({'row': ri, 'code': ec, 'action': 'UPDATED', 'amount': a})
        else:
            cell.value = 0; skipped_nm += 1
            results.append({'row': ri, 'code': ec, 'action': 'NO MATCH', 'amount': 0})
    if month_shift != 0:
        mn = str(rule["ref_month"] + month_shift).zfill(2)
        ws.cell(row=3, column=1).value = rule["title_format"].format(month=mn)
        ws.title = rule["sheet_format"].format(month=mn)
    return wb, results, updated, skipped_dd, skipped_nm

# ==========================================
# DOI SOAT DU LIEU
# ==========================================
def process_doi_soat(file1_path, file2_path, rule_key):
    """Compare 2 files by So PGH (9-digit). Returns pgh_result dict."""
    ext1 = os.path.splitext(file1_path)[1].lower()
    ext2 = os.path.splitext(file2_path)[1].lower()

    # Extract by PGH from both files
    data1, data2 = [], []
    for fp, data_list in [(file1_path, 'data1'), (file2_path, 'data2')]:
        ext = os.path.splitext(fp)[1].lower()
        if ext == '.pdf':
            if rule_key == 'nuoc':
                d = extract_pdf_nuoc_by_pgh(fp)
            else:
                d = extract_pdf_dien(fp)
        elif ext in ('.xlsx', '.xls'):
            d = extract_xl_by_pgh(fp)
        elif ext in ('.jpg', '.jpeg', '.png', '.gif'):
            d = extract_image_data(fp)
        else:
            d = []
        if data_list == 'data1':
            data1 = d
        else:
            data2 = d

    lk1 = build_lk(data1)
    lk2 = build_lk(data2)

    pgh_result = {}
    for pgh in set(list(lk1.keys()) + list(lk2.keys())):
        q1 = lk1.get(pgh, {}).get('amount', None)
        q2 = lk2.get(pgh, {}).get('amount', None)
        if q1 is not None and q2 is not None:
            pgh_result[pgh] = 'Đúng' if abs(q1 - q2) < 1 else 'Sai'
        elif q1 is not None and q2 is None:
            pgh_result[pgh] = 'Không có kết quả'
        elif q2 is not None and q1 is None:
            pgh_result[pgh] = 'Không có kết quả'

    return pgh_result, len(data1), len(data2)

def save_doi_soat_excel(excel_path, pgh_result):
    """Save doi soat results to Excel - add result column, scan col E for PGH"""
    wb = openpyxl.load_workbook(excel_path, keep_vba=False)
    ws = wb.active
    max_col = ws.max_column
    result_col = max_col + 1
    ws.cell(row=1, column=result_col).value = "Kết quả đối soát"
    ws.cell(row=1, column=result_col).font = openpyxl.styles.Font(bold=True)

    written = 0
    for row_idx in range(2, ws.max_row + 1):
        pgh_val = ws.cell(row=row_idx, column=5).value  # Column E
        if pgh_val:
            m = re.search(r'\d{9}', str(pgh_val))
            if m:
                pgh = m.group(0)
                result = pgh_result.get(pgh, 'Không có kết quả')
                cell = ws.cell(row=row_idx, column=result_col)
                cell.value = result
                if result == 'Đúng':
                    cell.font = openpyxl.styles.Font(color='008000', bold=True)
                elif result == 'Sai':
                    cell.font = openpyxl.styles.Font(color='FF0000', bold=True)
                else:
                    cell.font = openpyxl.styles.Font(color='808080')
                cell.alignment = openpyxl.styles.Alignment(horizontal='center')
                written += 1

    return wb, written, result_col

# ==========================================
# API ENDPOINTS
# ==========================================
class PhanBoResponse(BaseModel):
    rule: str
    updated: int
    skipped_di_doi: int
    no_match: int
    total: int
    results: list

class DoiSoatResponse(BaseModel):
    rule: str
    file1_count: int
    file2_count: int
    matched: int
    mismatched: int
    no_result: int
    total: int

@app.post("/phan-bo", response_model=PhanBoResponse)
async def phan_bo(
    rule: str = Form(...),
    month: int = Form(...),
    input_file: UploadFile = File(...),
    excel_template: UploadFile = File(...),
):
    """Phan bo chi phi - update Excel template with input data."""
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
        wb, results, updated, skipped_dd, skipped_nm = process_phan_bo(excel_path, lookup, rule, month_shift)

        rule_cfg = RULES[rule]
        output_name = f"{rule_cfg['excel_prefix']} thang {month}.2026.xlsx"
        output_path = os.path.join(tmpdir, output_name)
        wb.save(output_path)
        wb.close()

        return PhanBoResponse(
            rule=rule, updated=updated, skipped_di_doi=skipped_dd,
            no_match=skipped_nm, total=updated+skipped_dd+skipped_nm,
            results=results[:50],
        )
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

@app.post("/phan-bo/download")
async def phan_bo_download(
    rule: str = Form(...),
    month: int = Form(...),
    input_file: UploadFile = File(...),
    excel_template: UploadFile = File(...),
):
    """Phan bo - return updated Excel file."""
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
        wb, results, updated, skipped_dd, skipped_nm = process_phan_bo(excel_path, lookup, rule, month_shift)

        rule_cfg = RULES[rule]
        output_name = f"{rule_cfg['excel_prefix']} thang {month}.2026.xlsx"
        output_path = os.path.join(tmpdir, output_name)
        wb.save(output_path)
        wb.close()

        return FileResponse(output_path, filename=output_name,
                            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

@app.post("/doi-soat", response_model=DoiSoatResponse)
async def doi_soat(
    rule: str = Form(...),
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
):
    """Doi soat du lieu - compare 2 files by So PGH."""
    if rule not in RULES:
        raise HTTPException(status_code=400, detail=f"Invalid rule: {rule}")

    tmpdir = tempfile.mkdtemp()
    try:
        f1_path = os.path.join(tmpdir, file1.filename)
        with open(f1_path, 'wb') as f: f.write(await file1.read())
        f2_path = os.path.join(tmpdir, file2.filename)
        with open(f2_path, 'wb') as f: f.write(await file2.read())

        pgh_result, count1, count2 = process_doi_soat(f1_path, f2_path, rule)

        matched = sum(1 for v in pgh_result.values() if v == 'Đúng')
        mismatched = sum(1 for v in pgh_result.values() if v == 'Sai')
        no_result = sum(1 for v in pgh_result.values() if v == 'Không có kết quả')

        return DoiSoatResponse(
            rule=rule, file1_count=count1, file2_count=count2,
            matched=matched, mismatched=mismatched, no_result=no_result,
            total=len(pgh_result),
        )
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

@app.post("/doi-soat/download")
async def doi_soat_download(
    rule: str = Form(...),
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
):
    """Doi soat - return Excel with results. If file1 or file2 is Excel, use that format."""
    if rule not in RULES:
        raise HTTPException(status_code=400, detail=f"Invalid rule: {rule}")

    tmpdir = tempfile.mkdtemp()
    try:
        f1_path = os.path.join(tmpdir, file1.filename)
        with open(f1_path, 'wb') as f: f.write(await file1.read())
        f2_path = os.path.join(tmpdir, file2.filename)
        with open(f2_path, 'wb') as f: f.write(await file2.read())

        pgh_result, count1, count2 = process_doi_soat(f1_path, f2_path, rule)

        # Check if either file is Excel
        ext1 = os.path.splitext(f1_path)[1].lower()
        ext2 = os.path.splitext(f2_path)[1].lower()
        excel_file = None
        if ext1 in ('.xlsx', '.xls'):
            excel_file = f1_path
        elif ext2 in ('.xlsx', '.xls'):
            excel_file = f2_path

        if excel_file:
            wb, written, result_col = save_doi_soat_excel(excel_file, pgh_result)
            base_name = os.path.splitext(os.path.basename(excel_file))[0]
            output_name = f"Doi soat {base_name}.xlsx"
        else:
            # PDF vs PDF - create new Excel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Doi soat"
            ws.cell(row=1, column=1).value = "Số PGH"
            ws.cell(row=1, column=2).value = "Kết quả"
            ws.cell(row=1, column=1).font = openpyxl.styles.Font(bold=True)
            ws.cell(row=1, column=2).font = openpyxl.styles.Font(bold=True)
            for ri, (pgh, result) in enumerate(sorted(pgh_result.items()), 2):
                ws.cell(row=ri, column=1).value = pgh
                ws.cell(row=ri, column=2).value = result
            output_name = "Doi soat.xlsx"

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
