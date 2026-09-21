# -*- coding: utf-8 -*-
"""
AI Assistant - Multi-function app
Menu:
  I. Phan bo chi phi
    1. Phan bo chi phi dien
    2. Phan bo chi phi nuoc uong
  II. Doi soat du lieu
    1. Doi soat du lieu dien
    2. Doi soat du lieu nuoc uong
  III. Quan ly chi phi (placeholder)
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pdfplumber
import openpyxl
import re
import os
import sys
import io
import shutil
import json
from datetime import datetime
import calendar

try:
    import pytesseract
    from PIL import Image
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# ==========================================
# CONFIG
# ==========================================
APP_NAME = "AI Assistant"
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".ai_assistant_settings.json")
HISTORY_FILE = os.path.join(os.path.expanduser("~"), ".ai_assistant_history.json")

RULES = {
    "dien": {
        "name": "Phan bo chi phi dien",
        "menu": "II.1",
        "color": "#FFF8DC",
        "dir": r"D:\File storage\AI\AI Agent\Dien",
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
        "desc": "Match: Cot 8 (Ma KH) -> Cot 15 | Shift ngay: Co",
    },
    "nuoc": {
        "name": "Phan bo chi phi nuoc uong",
        "menu": "II.2",
        "color": "#E0F0FF",
        "dir": r"D:\File storage\AI\AI Agent\Nuoc uong",
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
        "desc": "Match: Cot 2 (Code DV) -> Cot 8 (So luong) | Chia tach: Co",
    },
}

# Doi soat config (same dirs, different function)
DOI_SOAT = {
    "dien": {
        "name": "Doi soat du lieu dien",
        "menu": "I.1",
        "dir": r"D:\File storage\AI\AI Agent\Dien",
        "col_code": 8,
        "col_amount": 15,
        "sheet_index": 1,
        "data_start_row": 5,
        "desc": "Doi soat ma KH + so tien giua PDF va Excel",
    },
    "nuoc": {
        "name": "Doi soat du lieu nuoc uong",
        "menu": "I.2",
        "dir": r"D:\File storage\AI\AI Agent\Nuoc uong",
        "col_code": 2,
        "col_amount": 8,
        "sheet_index": 2,
        "data_start_row": 5,
        "desc": "Doi soat ma KH + so luong giua PDF va Excel",
    },
}

TESSERACT_PATHS = [r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]

# ==========================================
# LOGIC
# ==========================================
def add_months(ds, n):
    try:
        d = datetime.strptime(ds.strip(), '%d/%m/%Y')
        t = d.year * 12 + d.month - 1 + n
        y, m = t // 12, t % 12 + 1
        return d.replace(year=y, month=m, day=min(d.day, calendar.monthrange(y, m)[1])).strftime('%d/%m/%Y')
    except: return None

def shift_dr(s, n):
    if not s: return None
    m = re.match(r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', str(s).strip())
    if not m: return None
    a, b = add_months(m.group(1), n), add_months(m.group(2), n)
    return f"{a} - {b}" if a and b else None

def _code(t):
    if not t: return ""
    t = str(t).strip()
    m = re.match(r'^([A-Z]{2}[0-9A-Z]+)', t) or re.search(r'\b([A-Z]{2}\d{8,})\b', t)
    return m.group(1) if m else ""

def _amt(d, c):
    s = ""
    if d and d != '0': s = d.replace(',', '').replace('.00', '').strip()
    elif c and c != '0': s = c.replace(',', '').replace('.00', '').strip()
    return float(s) if s else 0

def extract_pdf(fp):
    data = []
    with pdfplumber.open(fp) as pdf:
        for p in pdf.pages:
            for t in p.extract_tables():
                for r in t:
                    if not r or not r[0]: continue
                    if not str(r[0]).strip().isdigit(): continue
                    c = _code(str(r[5]) if r[5] else "")
                    a = _amt(str(r[6]) if r[6] else "", str(r[7]) if r[7] else "")
                    if c and c != 'MSB' and a > 0: data.append({'code': c, 'amount': a})
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
    """Extract Số PGH (9-digit) -> quantity from water PDF.
    Handles 2 formats:
    1. Bảng kê chi tiết: col1=date, col2=Số PGH, col6=qty
    2. Phiếu giao hàng: Số Phiếu in text, qty in Table1 R1 col2
    """
    result = {}

    with pdfplumber.open(fp) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # Detect format: Phiếu giao hàng has "PHIẾU GIAO HÀNG" in text
            if 'PHIẾU GIAO HÀNG' in text or 'DELIVERY TICKET' in text:
                # Format 2: Phiếu giao hàng
                # Find Số Phiếu (9-digit number after "SỐ PHIẾU" or "NO.")
                m = re.search(r'SỐ PHIẾU[/\s]*NO[\.\s]*\s*(\d{9})', text)
                if not m:
                    m = re.search(r'NO\.\s*(\d{9})', text)
                if not m:
                    # Try finding 9-digit number near "SỐ PHIẾU"
                    lines = text.split('\n')
                    for i, line in enumerate(lines):
                        if 'SỐ PHIẾU' in line or 'NO.' in line:
                            # Check next few lines for 9-digit number
                            for j in range(i, min(i + 4, len(lines))):
                                nums = re.findall(r'\d{9}', lines[j])
                                if nums:
                                    m = re.match(r'(\d{9})', nums[0])
                                    break
                            if m: break

                if m:
                    pgh = m.group(1) if m else (m.group(0) if m else "")
                    # Find quantity in Table 1, R1, col2
                    tables = page.extract_tables()
                    qty = 0
                    for table in tables:
                        if len(table) >= 2:
                            row1 = table[1]  # Data row
                            # Find "Số lượng" column from header
                            header = table[0]
                            qty_col = -1
                            for ci, h in enumerate(header):
                                if h and 'lượng' in str(h).lower():
                                    qty_col = ci
                                    break
                            if qty_col >= 0 and len(row1) > qty_col and row1[qty_col]:
                                try:
                                    qty = float(str(row1[qty_col]).replace(',', ''))
                                except:
                                    qty = 0
                            break

                    if pgh and qty > 0 and pgh not in result:
                        result[pgh] = qty
            else:
                # Format 1: Bảng kê chi tiết bán hàng
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
                                try:
                                    qty = float(col6.replace(',', ''))
                                except:
                                    qty = 0
                                if pgh not in result:
                                    result[pgh] = qty

    return [{'code': k, 'amount': v} for k, v in result.items()]

def extract_xl_by_pgh(fp):
    """Extract Số PGH -> quantity from Excel.
    Bảng kê chi tiết bán hàng: Số PGH at column E (index 4), Số lượng at column P (index 15)
    """
    result = {}
    wb = openpyxl.load_workbook(fp, keep_vba=False, data_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
            if not row: continue
            # Column E = index 4 (Số PGH), Column P = index 15 (Số lượng)
            pgh_val = row[4] if len(row) > 4 else None
            qty_val = row[15] if len(row) > 15 else None
            if pgh_val:
                pgh_str = str(pgh_val).strip()
                m = re.search(r'\d{9}', pgh_str)
                if m and qty_val:
                    pgh = m.group(0)
                    try:
                        qty = float(str(qty_val).replace(',', ''))
                    except:
                        qty = 0
                    if qty > 0 and pgh not in result:
                        result[pgh] = qty
    wb.close()
    return [{'code': k, 'amount': v} for k, v in result.items()]

def extract_xl(fp):
    data = []
    wb = openpyxl.load_workbook(fp, keep_vba=False, data_only=True)
    for ws in wb.worksheets:
        for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=5, values_only=True):
            if not r or not r[0]: continue
            c = str(r[0]).strip(); a = 0
            if len(r) > 2 and r[2]:
                try: a = float(str(r[2]).replace(',', '').replace('.00', '').strip())
                except: pass
            if c and a > 0: data.append({'code': c, 'amount': a})
    wb.close(); return data

def extract_img(fp):
    if not HAS_OCR: raise Exception("pytesseract chua cai")
    tp = next((p for p in TESSERACT_PATHS if os.path.exists(p)), None)
    if not tp: raise Exception("Tesseract chua cai")
    pytesseract.pytesseract.tesseract_cmd = tp
    text = pytesseract.image_to_string(Image.open(fp), lang='eng')
    data = []
    for line in text.split('\n'):
        c = _code(line)
        nums = re.findall(r'[\d,]+\.?\d*', line)
        a = float(nums[-1].replace(',', '')) if nums else 0
        if c and c != 'MSB' and a > 0: data.append({'code': c, 'amount': a})
    return data

def extract_api(url):
    if not HAS_REQUESTS: raise Exception("requests chua cai")
    r = requests.get(url, timeout=30); r.raise_for_status()
    res = r.json()
    items = res if isinstance(res, list) else res.get('data', res.get('items', []))
    return [{'code': str(i.get('code', i.get('ma', ''))).strip(), 'amount': float(i.get('amount', i.get('so_luong', 0)))}
            for i in items if i.get('code') and i.get('amount')]

def find_xl(d):
    for f in os.listdir(d):
        if not f.startswith('~$') and 'backup' not in f.lower() and f.endswith('.xlsx') and 'gốc' in f.lower():
            return os.path.join(d, f)
    for f in os.listdir(d):
        if not f.startswith('~$') and 'backup' not in f.lower() and f.endswith('.xlsx'):
            return os.path.join(d, f)
    return None

def build_lk(data):
    lk = {}
    for i in data:
        if i['code'] not in lk: lk[i['code']] = i
    return lk

def process_xl(ep, lk, r, ms=0):
    wb = openpyxl.load_workbook(ep, keep_vba=False)
    ws = wb.worksheets[r["sheet_index"]]
    res = []; u = sd = sn = 0; cc = {}
    merged_ranges = list(ws.merged_cells.ranges)
    if r["split_duplicates"]:
        for ri in range(r["data_start_row"], ws.max_row + 1):
            c = ws.cell(row=ri, column=r["col_code"]).value
            if c: cc[str(c).strip()] = cc.get(str(c).strip(), 0) + 1
    for ri in range(r["data_start_row"], ws.max_row + 1):
        ec = ws.cell(row=ri, column=r["col_code"]).value
        if not ec or not str(ec).strip(): continue
        ec = str(ec).strip()
        if "HO HCM" in ec.upper():
            res.append({'row': ri, 'code': ec, 'action': 'KEEP FORMULA', 'amount': '(HO HCM)'}); continue
        cell = ws.cell(row=ri, column=r["col_fill"])
        if any(cell.coordinate in mr for mr in merged_ranges): continue
        if r["has_di_doi"]:
            c13 = ws.cell(row=ri, column=r["col_ky_tt"]).value or ""
            c14 = ws.cell(row=ri, column=r["col_ky_tt_truoc"]).value or ""
            if any(k in str(c13).lower() for k in ['di d', 'dời', 'ngừng']) or \
               any(k in str(c14).lower() for k in ['di d', 'dời', 'ngừng']):
                cell.value = None; sd += 1
                res.append({'row': ri, 'code': ec, 'action': 'SKIP', 'amount': 'blank'}); continue
            if ms and r["has_date_shift"]:
                n13 = shift_dr(c13, ms)
                if n13: ws.cell(row=ri, column=r["col_ky_tt"]).value = n13
                n14 = shift_dr(c14, ms)
                if n14: ws.cell(row=ri, column=r["col_ky_tt_truoc"]).value = n14
        if ec in lk:
            a = lk[ec]['amount']
            if r["split_duplicates"] and cc.get(ec, 1) > 1: a = a / cc[ec]
            cell.value = a; u += 1
            res.append({'row': ri, 'code': ec, 'action': 'UPDATED', 'amount': f"{a:,.0f}"})
        else:
            cell.value = 0; sn += 1
            res.append({'row': ri, 'code': ec, 'action': 'NO MATCH', 'amount': '0'})
    # Update title
    if ms != 0:
        mn = str(r["ref_month"] + ms).zfill(2)
        ws.cell(row=3, column=1).value = r["title_format"].format(month=mn)
        ws.title = r["sheet_format"].format(month=mn)
    return wb, res, u, sd, sn

def doi_soat(ep, lk, ds_cfg):
    """Compare PDF data with Excel data. Return mismatches."""
    wb = openpyxl.load_workbook(ep, keep_vba=False, data_only=True)
    ws = wb.worksheets[ds_cfg["sheet_index"]]
    results = []
    matched = mismatched = only_excel = only_pdf = 0

    # Get Excel data
    excel_data = {}
    for ri in range(ds_cfg["data_start_row"], ws.max_row + 1):
        code = ws.cell(row=ri, column=ds_cfg["col_code"]).value
        if not code or not str(code).strip(): continue
        code = str(code).strip()
        if "HO HCM" in code.upper(): continue
        amt = ws.cell(row=ri, column=ds_cfg["col_amount"]).value
        try: amt = float(amt) if amt else 0
        except: amt = 0
        excel_data[code] = amt

    # Compare
    all_codes = set(list(excel_data.keys()) + list(lk.keys()))
    for code in sorted(all_codes):
        pdf_amt = lk.get(code, {}).get('amount', None)
        xl_amt = excel_data.get(code, None)

        if pdf_amt is not None and xl_amt is not None:
            if abs(pdf_amt - xl_amt) < 1:
                matched += 1
                results.append({'code': code, 'pdf': pdf_amt, 'excel': xl_amt, 'status': 'MATCH'})
            else:
                mismatched += 1
                results.append({'code': code, 'pdf': pdf_amt, 'excel': xl_amt, 'status': 'MISMATCH'})
        elif pdf_amt is not None and xl_amt is None:
            only_pdf += 1
            results.append({'code': code, 'pdf': pdf_amt, 'excel': None, 'status': 'ONLY_PDF'})
        elif xl_amt is not None and pdf_amt is None:
            only_excel += 1
            results.append({'code': code, 'pdf': None, 'excel': xl_amt, 'status': 'ONLY_EXCEL'})

    wb.close()
    return results, matched, mismatched, only_pdf, only_excel

def get_out(ip, bd, r):
    bn = os.path.basename(ip[0])
    m = re.search(r'(?:thang|thang|T)\s*(\d+)', bn, re.IGNORECASE)
    if m:
        ms = m.group(1)
        return os.path.join(bd, f"{r['excel_prefix']} thang {ms}.xlsx"), ms
    return None, None

def load_set():
    try:
        with open(SETTINGS_FILE, 'r') as f: return json.load(f)
    except: return {}

def save_set(s):
    try:
        with open(SETTINGS_FILE, 'w') as f: json.dump(s, f)
    except: pass

def load_hist():
    try:
        with open(HISTORY_FILE, 'r') as f: return json.load(f)
    except: return []

def add_hist(e):
    h = load_hist(); h.insert(0, e); h = h[:20]
    try:
        with open(HISTORY_FILE, 'w') as f: json.dump(h, f, ensure_ascii=False)
    except: pass

# ==========================================
# GUI
# ==========================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1100x750")
        self.input_paths = []
        self.excel_path = None
        self.wb = None
        self.results = []
        self.output_path = None
        self.settings = load_set()
        self.current_view = "I.dien"
        self._build_ui()

    def _build_ui(self):
        # HEADER
        header = tk.Frame(self.root, bg='#1a237e', height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text=f"  {APP_NAME}", bg='#1a237e', fg='white',
                font=('Segoe UI', 16, 'bold')).pack(side='left', pady=10)
        tk.Label(header, text="MSB AI Hackathon 2026  ", bg='#1a237e', fg='#64ffda',
                font=('Segoe UI', 9)).pack(side='right', pady=10)

        # BODY
        body = ttk.Frame(self.root)
        body.pack(fill='both', expand=True)

        # SIDEBAR
        sidebar = tk.Frame(body, bg='#263238', width=230)
        sidebar.pack(side='left', fill='y')
        sidebar.pack_propagate(False)

        self.nav_buttons = {}

        # I. Doi soat du lieu
        tk.Label(sidebar, text="  I. DOI SOAT DU LIEU", bg='#263238', fg='#18BC9C',
                font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(10, 3))
        for key, ds in DOI_SOAT.items():
            btn = tk.Button(sidebar, text=f"    {ds['menu']}. {ds['name']}",
                           bg='#263238', fg='#b0bec5', bd=0, anchor='w',
                           font=('Segoe UI', 9), padx=10, pady=5,
                           activebackground='#37474f', activeforeground='white',
                           command=lambda k=f"I.{key}": self._nav(k))
            btn.pack(fill='x')
            self.nav_buttons[f"I.{key}"] = btn

        # II. Phan bo chi phi
        tk.Label(sidebar, text="  II. PHAN BO CHI PHI", bg='#263238', fg='#18BC9C',
                font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(15, 3))
        for key, r in RULES.items():
            btn = tk.Button(sidebar, text=f"    {r['menu']}. {r['name']}",
                           bg='#263238', fg='#b0bec5', bd=0, anchor='w',
                           font=('Segoe UI', 9), padx=10, pady=5,
                           activebackground='#37474f', activeforeground='white',
                           command=lambda k=f"II.{key}": self._nav(k))
            btn.pack(fill='x')
            self.nav_buttons[f"II.{key}"] = btn

        # III. Quan ly chi phi
        tk.Label(sidebar, text="  III. QUAN LY CHI PHI", bg='#263238', fg='#18BC9C',
                font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(15, 3))
        tk.Button(sidebar, text="    (Sap co)", bg='#263238', fg='#546e7a',
                 bd=0, anchor='w', font=('Segoe UI', 9, 'italic'),
                 padx=10, pady=5, state='disabled').pack(fill='x')

        # History
        tk.Label(sidebar, text="  LICH SU", bg='#263238', fg='#90a4ae',
                font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(20, 5))
        self.hist_label = tk.Label(sidebar, text="Chua co", bg='#263238', fg='#78909c',
                                  font=('Segoe UI', 8), justify='left', wraplength=210)
        self.hist_label.pack(anchor='w', padx=5)
        self._refresh_hist()

        # CONTENT
        self.content = ttk.Frame(body, padding=15)
        self.content.pack(side='left', fill='both', expand=True)
        self._build_content()

    def _build_content(self):
        for w in self.content.winfo_children():
            w.destroy()
        c = self.content
        view = self.current_view

        if view.startswith("I."):
            self._build_doi_soat(c, view.split(".")[1])
        elif view.startswith("II."):
            self._build_phan_bo(c, view.split(".")[1])

    def _build_phan_bo(self, c, key):
        r = RULES[key]
        ttk.Label(c, text=f"{r['menu']}. {r['name']}", font=('Segoe UI', 14, 'bold')).pack(anchor='w')
        ttk.Label(c, text=r['desc'], foreground='gray', font=('Segoe UI', 9)).pack(anchor='w', pady=(0, 10))

        self._step(c, "1", "Chon file dau vao (PDF / Excel / Image)")
        f1 = ttk.Frame(c); f1.pack(fill='x', pady=(0, 3))
        ttk.Button(f1, text="Chon file", command=self._sel_files).pack(side='left', padx=(0, 5))
        ttk.Button(f1, text="Chon nhieu", command=self._sel_multi).pack(side='left', padx=(0, 5))
        ttk.Button(f1, text="Xoa", command=self._clear).pack(side='left', padx=(0, 15))
        tk.Label(f1, text="API:").pack(side='left', padx=(0, 5))
        self.api_entry = ttk.Entry(f1, width=35)
        self.api_entry.pack(side='left', fill='x', expand=True)
        self.input_lb = tk.Listbox(c, height=2, font=('Segoe UI', 9))
        self.input_lb.pack(fill='x', pady=(0, 5))

        self._step(c, "2", "File Excel goc (template)")
        f2 = ttk.Frame(c); f2.pack(fill='x', pady=(0, 3))
        ttk.Button(f2, text="Tu dong tim", command=self._auto_xl).pack(side='left', padx=(0, 5))
        ttk.Button(f2, text="Chon file", command=self._sel_xl).pack(side='left')
        self.xl_label = ttk.Label(f2, text="Chua chon", foreground='gray')
        self.xl_label.pack(side='left', padx=15)

        self._step(c, "3", "Xu ly & Luu")
        f3 = ttk.Frame(c); f3.pack(fill='x', pady=(0, 3))
        self.proc_btn = ttk.Button(f3, text="Xu ly", command=self._process, state='disabled')
        self.proc_btn.pack(side='left', padx=(0, 10))
        self.save_btn = ttk.Button(f3, text="Luu Excel", command=self._save, state='disabled')
        self.save_btn.pack(side='left', padx=(0, 10))
        self.auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(f3, text="Tu mo", variable=self.auto_var).pack(side='left', padx=10)
        self.st_label = ttk.Label(f3, text="", foreground='gray')
        self.st_label.pack(side='left', padx=15)

        self._step(c, "4", "Ket qua")
        self.sum_label = ttk.Label(c, text="Chua xu ly", foreground='gray')
        self.sum_label.pack(anchor='w', pady=(0, 3))
        cols = ('row', 'code', 'action', 'amount')
        self.tree = ttk.Treeview(c, columns=cols, show='headings', height=10)
        for col, txt, w, a in [('row','Dong',50,'center'),('code','Code',150,'w'),
                               ('action','Hanh dong',120,'center'),('amount','Gia tri',120,'e')]:
            self.tree.heading(col, text=txt); self.tree.column(col, width=w, anchor=a)
        self.tree.tag_configure('updated', foreground='green')
        self.tree.tag_configure('skip', foreground='orange')
        self.tree.tag_configure('nomatch', foreground='red')
        sb = ttk.Scrollbar(c, orient='v', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        self._check()

    def _build_doi_soat(self, c, key):
        ds = DOI_SOAT[key]
        ttk.Label(c, text=f"{ds['menu']}. {ds['name']}", font=('Segoe UI', 14, 'bold')).pack(anchor='w')
        ttk.Label(c, text="So sanh 2 file bat ky (PDF / Excel / Image)", foreground='gray', font=('Segoe UI', 9)).pack(anchor='w', pady=(0, 10))

        # File 1
        self._step(c, "1", "Chon file thu 1 (PDF / Excel / Image)")
        f1 = ttk.Frame(c); f1.pack(fill='x', pady=(0, 3))
        ttk.Button(f1, text="Chon file 1", command=lambda: self._sel_doi_soat_file(1)).pack(side='left', padx=(0, 10))
        self.ds_file1_label = ttk.Label(f1, text="Chua chon", foreground='gray')
        self.ds_file1_label.pack(side='left')

        # File 2
        self._step(c, "2", "Chon file thu 2 (PDF / Excel / Image)")
        f2 = ttk.Frame(c); f2.pack(fill='x', pady=(0, 3))
        ttk.Button(f2, text="Chon file 2", command=lambda: self._sel_doi_soat_file(2)).pack(side='left', padx=(0, 10))
        self.ds_file2_label = ttk.Label(f2, text="Chua chon", foreground='gray')
        self.ds_file2_label.pack(side='left')

        # Compare
        self._step(c, "3", "Doi soat & Luu")
        f3 = ttk.Frame(c); f3.pack(fill='x', pady=(0, 3))
        self.proc_btn = ttk.Button(f3, text="Doi soat", command=self._doi_soat, state='disabled')
        self.proc_btn.pack(side='left', padx=(0, 10))
        self.ds_save_btn = ttk.Button(f3, text="Luu ket qua", command=self._doi_soat_save, state='disabled')
        self.ds_save_btn.pack(side='left', padx=(0, 10))
        self.auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(f3, text="Tu mo", variable=self.auto_var).pack(side='left', padx=10)
        self.st_label = ttk.Label(f3, text="", foreground='gray')
        self.st_label.pack(side='left', padx=15)

        # Results
        self._step(c, "4", "Ket qua doi soat")
        self.sum_label = ttk.Label(c, text="Chua doi soat", foreground='gray')
        self.sum_label.pack(anchor='w', pady=(0, 3))
        cols = ('code', 'file1', 'file2', 'status')
        self.tree = ttk.Treeview(c, columns=cols, show='headings', height=12)
        for col, txt, w, a in [('code','Ma KH',150,'w'),('file1','File 1',120,'e'),
                               ('file2','File 2',120,'e'),('status','Trang thai',100,'center')]:
            self.tree.heading(col, text=txt); self.tree.column(col, width=w, anchor=a)
        self.tree.tag_configure('match', foreground='green')
        self.tree.tag_configure('mismatch', foreground='red')
        self.tree.tag_configure('only_file1', foreground='orange')
        self.tree.tag_configure('only_file2', foreground='blue')
        sb = ttk.Scrollbar(c, orient='v', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.ds_file1 = None
        self.ds_file2 = None
        self._check_doi_soat()

    def _step(self, parent, num, title):
        f = ttk.Frame(parent); f.pack(fill='x', pady=(8, 0))
        tk.Label(f, text=f" {num} ", bg='#1a237e', fg='white',
                font=('Segoe UI', 9, 'bold'), width=3).pack(side='left', padx=(0, 8))
        ttk.Label(f, text=title, font=('Segoe UI', 10, 'bold')).pack(side='left')

    def _nav(self, key):
        self.current_view = key
        self.input_paths = []; self.excel_path = None
        for k, btn in self.nav_buttons.items():
            btn.config(bg='#37474f' if k == key else '#263238',
                      fg='white' if k == key else '#b0bec5')
        self._build_content()

    def _cur_rule(self):
        key = self.current_view.split(".")[1]
        if self.current_view.startswith("I."):
            return RULES[key]
        return DOI_SOAT[key]

    def _sel_files(self):
        r = self._cur_rule()
        p = filedialog.askopenfilename(title="Chon file", initialdir=r["dir"],
            filetypes=[("All","*.pdf *.xlsx *.xls *.jpg *.jpeg *.png *.gif"),("PDF","*.pdf"),("Excel","*.xlsx *.xls"),("Image","*.jpg *.jpeg *.png *.gif")])
        if p: self.input_paths = [p]; self._ref_in(); self._check()

    def _sel_multi(self):
        r = self._cur_rule()
        ps = filedialog.askopenfilenames(title="Chon nhieu", initialdir=r["dir"],
            filetypes=[("All","*.pdf *.xlsx *.xls *.jpg *.jpeg *.png *.gif")])
        if ps: self.input_paths = list(ps); self._ref_in(); self._check()

    def _clear(self):
        self.input_paths = []; self._ref_in(); self._check()

    def _ref_in(self):
        self.input_lb.delete(0, tk.END)
        for p in self.input_paths:
            e = os.path.splitext(p)[1].upper()
            self.input_lb.insert(tk.END, f"[{e[1:]}] {os.path.basename(p)}")

    def _auto_xl(self):
        r = self._cur_rule(); e = find_xl(r["dir"])
        if e:
            self.excel_path = e
            self.xl_label.config(text=os.path.basename(e), foreground='green')
            self._check()
        else:
            messagebox.showwarning("Canh bao", "Khong tim thay file Excel goc!")

    def _sel_xl(self):
        r = self._cur_rule()
        p = filedialog.askopenfilename(title="Chon Excel", initialdir=r["dir"], filetypes=[("Excel","*.xlsx")])
        if p:
            self.excel_path = p
            self.xl_label.config(text=os.path.basename(p), foreground='green')
            self._check()

    def _check(self):
        ok = len(self.input_paths) > 0 and self.excel_path
        if hasattr(self, 'proc_btn'):
            self.proc_btn.config(state='normal' if ok else 'disabled')

    def _process(self):
        self.proc_btn.config(state='disabled'); self.save_btn.config(state='disabled')
        self.st_label.config(text="Dang xu ly...", foreground='blue'); self.root.update()
        try:
            r = RULES[self.current_view.split(".")[1]]; data = []
            for p in self.input_paths:
                e = os.path.splitext(p)[1].lower()
                if e == '.pdf':
                    if self.current_view.split(".")[1] == 'nuoc': data.extend(extract_pdf_nuoc(p))
                    else: data.extend(extract_pdf(p))
                elif e in ('.xlsx','.xls'): data.extend(extract_xl(p))
                elif e in ('.jpg','.jpeg','.png','.gif'): data.extend(extract_img(p))
            lk = build_lk(data); ms = 0
            if self.input_paths:
                bn = os.path.basename(self.input_paths[0])
                m = re.search(r'(?:thang|thang|T)\s*(\d+)', bn, re.IGNORECASE)
                if m: ms = int(m.group(1)) - r["ref_month"]
            self.wb, self.results, u, sd, sn = process_xl(self.excel_path, lk, r, ms)
            self.output_path, mstr = get_out(self.input_paths, os.path.dirname(self.excel_path), r)
            if not self.output_path: self.output_path = self.excel_path
            self.tree.delete(*self.tree.get_children())
            for x in self.results:
                t = 'updated' if x['action']=='UPDATED' else ('skip' if 'SKIP' in x['action'] or 'KEEP' in x['action'] else 'nomatch')
                self.tree.insert('', 'end', values=(x['row'], x['code'], x['action'], x['amount']), tags=(t,))
            self.sum_label.config(text=f"Input: {len(data)} | Updated: {u} | Skip: {sd} | No match: {sn} | Total: {u+sd+sn}", foreground='black')
            self.st_label.config(text="Xong!", foreground='green')
            self.save_btn.config(state='normal')
        except Exception as e:
            messagebox.showerror("Loi", str(e))
            self.st_label.config(text="Loi!", foreground='red')
        finally:
            self.proc_btn.config(state='normal')

    def _sel_doi_soat_file(self, num):
        key = self.current_view.split(".")[1]
        ds = DOI_SOAT[key]
        p = filedialog.askopenfilename(title=f"Chon file {num}", initialdir=ds["dir"],
            filetypes=[("All","*.pdf *.xlsx *.xls *.jpg *.jpeg *.png *.gif"),("PDF","*.pdf"),("Excel","*.xlsx *.xls"),("Image","*.jpg *.jpeg *.png *.gif")])
        if p:
            if num == 1:
                self.ds_file1 = p
                self.ds_file1_label.config(text=os.path.basename(p), foreground='green')
            else:
                self.ds_file2 = p
                self.ds_file2_label.config(text=os.path.basename(p), foreground='green')
            self._check_doi_soat()

    def _check_doi_soat(self):
        ok = self.ds_file1 and self.ds_file2
        if hasattr(self, 'proc_btn'):
            self.proc_btn.config(state='normal' if ok else 'disabled')

    def _extract_by_type(self, filepath, key, by_pgh=False):
        """Extract data from any file type. If by_pgh, extract by Số PGH."""
        ext = os.path.splitext(filepath)[1].lower()
        if ext == '.pdf':
            if key == 'nuoc':
                return extract_pdf_nuoc_by_pgh(filepath) if by_pgh else extract_pdf_nuoc(filepath)
            else:
                return extract_pdf(filepath)
        elif ext in ('.xlsx', '.xls'):
            if by_pgh:
                return extract_xl_by_pgh(filepath)
            return extract_xl(filepath)
        elif ext in ('.jpg', '.jpeg', '.png', '.gif'):
            return extract_img(filepath)
        return []

    def _doi_soat(self):
        self.proc_btn.config(state='disabled')
        self.ds_save_btn.config(state='disabled')
        self.st_label.config(text="Dang doi soat...", foreground='blue'); self.root.update()
        try:
            key = self.current_view.split(".")[1]

            # Extract Số PGH -> quantity from BOTH files
            data1 = self._extract_by_type(self.ds_file1, key, by_pgh=True)
            data2 = self._extract_by_type(self.ds_file2, key, by_pgh=True)
            lk1 = build_lk(data1)  # Số PGH -> qty (file 1)
            lk2 = build_lk(data2)  # Số PGH -> qty (file 2)

            # Compare quantities per Số PGH
            pgh_result = {}
            for pgh in set(list(lk1.keys()) + list(lk2.keys())):
                q1 = lk1.get(pgh, {}).get('amount', None)
                q2 = lk2.get(pgh, {}).get('amount', None)
                if q1 is not None and q2 is not None:
                    pgh_result[pgh] = 'Đúng' if abs(q1 - q2) < 1 else 'Sai'
                elif q1 is not None and q2 is None:
                    pgh_result[pgh] = 'Không có kết quả'  # only in file 1
                elif q2 is not None and q1 is None:
                    pgh_result[pgh] = 'Không có kết quả'  # only in file 2

            # Extract ALL rows from file 1 for Excel output
            ext1 = os.path.splitext(self.ds_file1)[1].lower()
            all_rows = []
            if ext1 == '.pdf':
                with pdfplumber.open(self.ds_file1) as pdf:
                    for page in pdf.pages:
                        for table in page.extract_tables():
                            for row in table:
                                if not row: continue
                                col0 = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                                col1 = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                                if 'Mã KH' in col0 or 'Cust' in col0: continue
                                if 'Theo mặt hàng' in col1 or 'By SKU' in col1: continue
                                if not col0 and not col1: continue
                                all_rows.append(row)
            elif ext1 in ('.xlsx', '.xls'):
                wb_tmp = openpyxl.load_workbook(self.ds_file1, keep_vba=False, data_only=True)
                ws_tmp = wb_tmp.active
                for row in ws_tmp.iter_rows(min_row=1, max_row=ws_tmp.max_row, values_only=True):
                    if not row: continue
                    all_rows.append(list(row))
                wb_tmp.close()

            # Build row results - compare by Số PGH
            results = []
            matched = mismatched = no_result = 0

            # Determine which column has Số PGH based on file type
            is_excel = ext1 in ('.xlsx', '.xls')
            pgh_col = 4 if is_excel else 2  # Excel: col E=4, PDF: col 2

            for row in all_rows:
                col0 = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                col_pgh = str(row[pgh_col]).strip() if len(row) > pgh_col and row[pgh_col] else ""

                m = re.match(r'^(C\d+)', col0)
                if m:
                    # Customer header row - no Số PGH, leave empty
                    results.append({'row': row, 'result': ''})
                    continue

                # Look for 9-digit Số PGH
                m_pgh = re.search(r'\d{9}', col_pgh)
                if m_pgh:
                    pgh = m_pgh.group(0)
                    r = pgh_result.get(pgh, 'Không có kết quả')
                    results.append({'row': row, 'result': r})
                    if r == 'Đúng': matched += 1
                    elif r == 'Sai': mismatched += 1
                    else: no_result += 1
                else:
                    # No Số PGH in this row - leave empty
                    results.append({'row': row, 'result': ''})

            self.ds_results = results
            self.pgh_result = pgh_result  # Store for save method

            self.tree.delete(*self.tree.get_children())
            for x in results:
                row = x['row']
                code = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                ticket = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                qty = str(row[6]).strip() if len(row) > 6 and row[6] else ""
                result = x['result']
                tag = 'match' if result == 'Đúng' else ('mismatch' if result == 'Sai' else '')
                self.tree.insert('', 'end', values=(f"{code} {ticket}", qty, "", result), tags=(tag,))

            self.sum_label.config(text=f"File1: {len(data1)} | File2: {len(data2)} | Dung: {matched} | Sai: {mismatched} | Khong co ket qua: {no_result} | Total: {len(results)}", foreground='black')
            self.st_label.config(text="Doi soat xong! Bam 'Luu ket qua' de luu.", foreground='green')
            self.ds_save_btn.config(state='normal')
        except Exception as e:
            messagebox.showerror("Loi", str(e))
            self.st_label.config(text="Loi!", foreground='red')
        finally:
            self.proc_btn.config(state='normal')
            if hasattr(self, 'ds_results') and self.ds_results:
                self.ds_save_btn.config(state='normal')

    def _doi_soat_save(self):
        if not hasattr(self, 'pgh_result') or not self.pgh_result:
            messagebox.showwarning("Canh bao", "Chua co ket qua doi soat!")
            return
        try:
            # Check if either file is Excel
            ext1 = os.path.splitext(self.ds_file1)[1].lower()
            ext2 = os.path.splitext(self.ds_file2)[1].lower()
            excel_file = None
            if ext1 in ('.xlsx', '.xls'):
                excel_file = self.ds_file1
            elif ext2 in ('.xlsx', '.xls'):
                excel_file = self.ds_file2

            if excel_file:
                # CASE: PDF vs Excel -> use Excel format, add result column
                wb = openpyxl.load_workbook(excel_file, keep_vba=False)
                ws = wb.active

                # Find last column and add result header
                max_col = ws.max_column
                result_col = max_col + 1
                ws.cell(row=1, column=result_col).value = "Kết quả đối soát"
                ws.cell(row=1, column=result_col).font = openpyxl.styles.Font(bold=True)

                # Write results directly - scan column E for 9-digit Số PGH
                written = 0
                for row_idx in range(2, ws.max_row + 1):
                    pgh_val = ws.cell(row=row_idx, column=5).value  # Column E
                    if pgh_val:
                        m = re.search(r'\d{9}', str(pgh_val))
                        if m:
                            pgh = m.group(0)
                            result = self.pgh_result.get(pgh, 'Không có kết quả')
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

                # Save
                output_dir = os.path.dirname(excel_file)
                excel_name = os.path.splitext(os.path.basename(excel_file))[0]
                output_name = f"Doi soat {excel_name}.xlsx"
                output_path = os.path.join(output_dir, output_name)
                wb.save(output_path)
                wb.close()

                self.st_label.config(text=f"Da luu: {output_name} ({written} ket qua)", foreground='green')
                messagebox.showinfo("Thanh cong", f"File da luu:\n{output_name}\n{written} ket qua")
                add_hist({"time": datetime.now().strftime("%d/%m %H:%M"), "rule": f"Doi soat {self.current_view}", "file": output_name})
                self._refresh_hist()

                if self.auto_var.get():
                    os.startfile(output_path)

            else:
                # CASE: PDF vs PDF -> create new Excel from PDF (existing code)
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Doi soat"

                headers = ['Mã KH\nCust. Account', 'Ngày\nDate', 'Số PGH\nTicket No.',
                           'Mã HH\nItem Code', 'Tên hàng hóa, dịch vụ\nDescription',
                           'ĐVT\nUnit', 'Số Lượng\nQuantity', 'Đơn giá\nUnit Price',
                           'Thành tiền\nAmount', 'Chiết khấu\nDiscount',
                           'Thành tiền sau chiết khấu\nNet Amount', 'Kết quả đối soát']
                for col, h in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col)
                    cell.value = h
                    cell.font = openpyxl.styles.Font(bold=True)
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

                current_code = None
                code_start_row = None
                merge_ranges = []

                for ri, item in enumerate(self.ds_results, 2):
                    row = item['row']
                    result = item['result']
                    col0 = str(row[0]).strip() if len(row) > 0 and row[0] else ""

                    m = re.match(r'^(C\d+)', col0)
                    if m:
                        if current_code and code_start_row and code_start_row < ri - 1:
                            merge_ranges.append((code_start_row, ri - 1))
                        current_code = m.group(1)
                        code_start_row = ri
                    elif not col0:
                        pass
                    else:
                        if current_code and code_start_row and code_start_row < ri - 1:
                            merge_ranges.append((code_start_row, ri - 1))
                        current_code = None
                        code_start_row = None

                    for ci in range(min(11, len(row))):
                        val = row[ci]
                        if val is not None and str(val).strip():
                            ws.cell(row=ri, column=ci + 1).value = str(val).strip()

                    # Column 12: result (only if has value)
                    if result:
                        cell = ws.cell(row=ri, column=12)
                        cell.value = result
                        if result == 'Đúng':
                            cell.font = openpyxl.styles.Font(color='008000', bold=True)
                        elif result == 'Sai':
                            cell.font = openpyxl.styles.Font(color='FF0000', bold=True)
                        else:
                            cell.font = openpyxl.styles.Font(color='808080')
                        cell.alignment = openpyxl.styles.Alignment(horizontal='center')

                if current_code and code_start_row and code_start_row < len(self.ds_results) + 1:
                    merge_ranges.append((code_start_row, len(self.ds_results) + 1))

                for start, end in merge_ranges:
                    if end > start:
                        ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
                        ws.merge_cells(start_row=start, start_column=2, end_row=end, end_column=2)

                for col in range(1, 13):
                    max_len = 10
                    for row in range(1, min(50, len(self.ds_results) + 2)):
                        val = ws.cell(row=row, column=col).value
                        if val:
                            max_len = max(max_len, min(len(str(val)), 30))
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = max_len + 2

                output_dir = os.path.dirname(self.ds_file1)
                f1_name = os.path.splitext(os.path.basename(self.ds_file1))[0]
                output_name = f"Doi soat {f1_name}.xlsx"
                output_path = os.path.join(output_dir, output_name)
                wb.save(output_path)
                wb.close()

            self.st_label.config(text=f"Da luu: {output_name}", foreground='green')
            messagebox.showinfo("Thanh cong", f"File da luu:\n{output_name}")
            add_hist({"time": datetime.now().strftime("%d/%m %H:%M"), "rule": f"Doi soat {self.current_view}", "file": output_name})
            self._refresh_hist()

            if self.auto_var.get():
                os.startfile(output_path)
        except Exception as e:
            messagebox.showerror("Loi", str(e))

    def _save(self):
        if not self.wb or not self.output_path: return
        try:
            bp = self.output_path.replace('.xlsx', '_backup.xlsx')
            if os.path.exists(self.output_path): shutil.copy2(self.output_path, bp)
            self.wb.save(self.output_path); self.wb.close()
            fn = os.path.basename(self.output_path)
            self.st_label.config(text=f"Da luu: {fn}", foreground='green')
            messagebox.showinfo("Thanh cong", f"File da luu:\n{fn}")
            add_hist({"time": datetime.now().strftime("%d/%m %H:%M"), "rule": self._cur_rule()["name"], "file": fn})
            self._refresh_hist()
            if self.auto_var.get(): os.startfile(self.output_path)
            self.save_btn.config(state='disabled')
        except PermissionError:
            messagebox.showerror("Loi", "File dang mo! Dong file roi thu lai.")
        except Exception as e:
            messagebox.showerror("Loi", str(e))

    def _refresh_hist(self):
        h = load_hist()
        if h:
            txt = "\n".join([f"{x['time']} {x['rule'][:15]} {x['file'][:25]}" for x in h[:5]])
            self.hist_label.config(text=txt)
        else:
            self.hist_label.config(text="Chua co lich su")

if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
