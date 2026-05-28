"""
comparador.py
Lógica de extracción PDF + Excel y generación del reporte consolidado.
"""
import logging
import re
from pathlib import Path

import fitz  # pymupdf
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz

log = logging.getLogger("comparador")

# ── Patrones ──────────────────────────────────────────────────────────────────
GUIA_PATTERNS = [
    r'\b(\d{10,20})\b',
    r'(?:guia|guía|guide|tracking|#|no\.?)\s*:?\s*([A-Z0-9\-]{6,25})',
    r'\b([A-Z]{2,4}\d{8,15})\b',
]
MONTO_PATTERNS = [
    r'Q\.?\s*([\d,]+\.?\d*)',
    r'\$\s*([\d,]+\.?\d*)',
    r'([\d,]+\.\d{2})',
]
ESTADO_MAP = {
    'entregad': 'ENTREGADO', 'deliver': 'ENTREGADO',
    'devuelt': 'DEVUELTO',   'return': 'DEVUELTO',
    'transit': 'EN TRANSITO','en ruta': 'EN TRANSITO',
    'pendient': 'PENDIENTE', 'pending': 'PENDIENTE',
    'cancelad': 'CANCELADO', 'cancel': 'CANCELADO',
    'retenid': 'RETENIDO',   'held': 'RETENIDO',
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def normalize(val):
    return str(val).strip().upper() if val is not None else ""

def normalize_amount(val):
    s = str(val).replace(",", "").replace("Q", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0

def detect_estado(text):
    t = str(text).lower()
    for key, label in ESTADO_MAP.items():
        if key in t:
            return label
    return normalize(text) or "DESCONOCIDO"

def extract_guia(text):
    for pat in GUIA_PATTERNS:
        m = re.search(pat, str(text), re.IGNORECASE)
        if m:
            return m.group(1).strip().upper()
    return None

def extract_amount(text):
    for pat in MONTO_PATTERNS:
        m = re.search(pat, str(text).replace(" ", ""))
        if m:
            return normalize_amount(m.group(1))
    return None

def detect_currency(texts):
    combined = " ".join(str(t) for t in texts)
    return "Q" if "Q" in combined else ("$" if "$" in combined else "Q")

def _find_col(headers, candidates):
    for c in candidates:
        for i, h in enumerate(headers):
            if c in normalize(h):
                return i
    return None

def _find_col_df(columns, candidates):
    for c in candidates:
        for col in columns:
            if c in normalize(col):
                return col
    return None

# ── Extracción PDF ────────────────────────────────────────────────────────────
def extract_pdf_data(pdf_path: str) -> pd.DataFrame:
    log.info("Extrayendo PDF: %s", pdf_path)
    records = []
    doc = fitz.open(pdf_path)
    for page in doc:
        # Intentar extraer como tabla usando bloques
        blocks = page.get_text("blocks")
        text = page.get_text()
        _parse_text_lines(text, records)
    doc.close()

    df = pd.DataFrame(records).drop_duplicates(subset=["guia"])
    log.info("  → %d registros extraídos", len(df))
    return df

def _parse_text_lines(text, records):
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    for line in lines:
        guia = extract_guia(line)
        if not guia:
            continue
        estado_raw = next((k for k in ESTADO_MAP if k in line.lower()), "")
        monto = extract_amount(line)
        records.append({
            "guia": normalize(guia),
            "estado_pdf": detect_estado(estado_raw) if estado_raw else "DESCONOCIDO",
            "costo_pdf": monto if monto is not None else 0.0,
        })

# ── Lectura Excel ─────────────────────────────────────────────────────────────
def read_excel_data(excel_path: str) -> pd.DataFrame:
    log.info("Leyendo Excel: %s", excel_path)
    xl = pd.ExcelFile(excel_path)
    sheet = xl.sheet_names[0]
    raw = pd.read_excel(excel_path, sheet_name=sheet, header=None, dtype=str)

    header_row = 0
    for i, row in raw.iterrows():
        vals = [normalize(v) for v in row if pd.notna(v)]
        if any(kw in v for v in vals for kw in ["GUIA","TRACKING","ESTADO","COSTO","VALOR","NUMERO"]):
            header_row = i
            break

    df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row, dtype=str)
    df.columns = [normalize(c) for c in df.columns]

    col_guia   = _find_col_df(df.columns, ["GUIA","GUÍA","TRACKING","NUMERO","N°","GUIDE","AWB","CODIGO"])
    col_estado = _find_col_df(df.columns, ["ESTADO","STATUS","ESTATUS","CONDICION"])
    col_costo  = _find_col_df(df.columns, ["COSTO","VALOR","MONTO","PRECIO","IMPORTE","TOTAL","CARGO","AMOUNT"])

    records = []
    for _, row in df.iterrows():
        guia = normalize(row.get(col_guia, "")) if col_guia else None
        if not guia:
            for val in row.values:
                g = extract_guia(str(val))
                if g:
                    guia = g
                    break
        if not guia or guia in ("NAN", "", "NONE"):
            continue

        estado   = row.get(col_estado, "") if col_estado else ""
        costo_raw = row.get(col_costo, "") if col_costo else ""
        if not costo_raw or str(costo_raw).upper() in ("NAN","NONE",""):
            for val in row.values:
                a = extract_amount(str(val))
                if a is not None and a > 0:
                    costo_raw = a
                    break

        records.append({
            "guia":         normalize(guia),
            "estado_excel": detect_estado(estado),
            "costo_excel":  normalize_amount(costo_raw),
        })

    result = pd.DataFrame(records).drop_duplicates(subset=["guia"])
    log.info("  → %d registros Excel", len(result))
    return result

# ── Matching ──────────────────────────────────────────────────────────────────
def match_datasets(df_pdf: pd.DataFrame, df_excel: pd.DataFrame):
    pdf_set   = set(df_pdf["guia"])
    excel_set = set(df_excel["guia"])
    both      = pdf_set & excel_set
    only_pdf  = pdf_set - excel_set
    only_excel = excel_set - pdf_set

    # Fuzzy match para guías no encontradas exactamente
    unmatched_pdf   = only_pdf.copy()
    unmatched_excel = only_excel.copy()
    for g_pdf in list(unmatched_pdf):
        for g_excel in list(unmatched_excel):
            if fuzz.ratio(g_pdf, g_excel) >= 85:
                log.info("Fuzzy match: %s <-> %s", g_pdf, g_excel)
                both.add(g_pdf)
                unmatched_pdf.discard(g_pdf)
                unmatched_excel.discard(g_excel)
                df_excel.loc[df_excel["guia"] == g_excel, "guia"] = g_pdf
                break

    merged = df_pdf[df_pdf["guia"].isin(both)].merge(
        df_excel[df_excel["guia"].isin(both)], on="guia", how="inner"
    )
    df_only_excel = df_excel[df_excel["guia"].isin(unmatched_excel)].copy()
    df_only_pdf   = df_pdf[df_pdf["guia"].isin(unmatched_pdf)].copy()
    return merged, df_only_excel, df_only_pdf

# ── Estilos openpyxl ──────────────────────────────────────────────────────────
H_FILL  = PatternFill("solid", fgColor="1F3864")
H_FONT  = Font(color="FFFFFF", bold=True)
T_FILL  = PatternFill("solid", fgColor="D9D9D9")
T_FONT  = Font(bold=True)
G_FILL  = PatternFill("solid", fgColor="C6EFCE")
R_FILL  = PatternFill("solid", fgColor="FFC7CE")
BORDER  = Border(*[Side(style="thin")] * 0,
                 left=Side(style="thin"), right=Side(style="thin"),
                 top=Side(style="thin"), bottom=Side(style="thin"))
CUR_FMT = '"Q"#,##0.00'

def _hdr(ws):
    for cell in ws[1]:
        cell.fill = H_FILL; cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = BORDER

def _tot(ws):
    for cell in ws[ws.max_row]:
        cell.fill = T_FILL; cell.font = T_FONT; cell.border = BORDER

def _borders(ws):
    for row in ws.iter_rows():
        for cell in row:
            cell.border = BORDER

def _currency(ws, cols, start=2):
    for row in ws.iter_rows(min_row=start, max_row=ws.max_row):
        for i in cols:
            if i <= len(row):
                row[i-1].number_format = CUR_FMT

def _autofit(ws):
    for col in ws.columns:
        L = get_column_letter(col[0].column)
        mx = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[L].width = min(mx + 4, 45)

# ── Generación de reporte consolidado ─────────────────────────────────────────
def generar_reporte(pdf_paths: list[str], excel_path: str, output_path: str) -> dict:
    """
    Procesa múltiples PDFs contra un Excel y genera el reporte consolidado.
    Retorna dict con resumen de conteos.
    """
    df_excel = read_excel_data(excel_path)
    currency = detect_currency(df_excel["costo_excel"].astype(str).tolist())

    all_entregadas      = []
    all_estado_diferente = []
    all_solo_pdf        = []
    resumen_rows        = []

    # Guardar "Solo en Excel" una sola vez al final (global)
    global_only_excel = None

    for pdf_path in pdf_paths:
        pdf_name = Path(pdf_path).name
        df_pdf = extract_pdf_data(pdf_path)
        if df_pdf.empty:
            log.warning("Sin registros en %s", pdf_name)
            continue

        # Trabajar con copia del excel para no contaminar entre PDFs
        df_excel_copy = df_excel.copy()
        merged, df_only_excel_local, df_only_pdf_local = match_datasets(df_pdf, df_excel_copy)

        # Acumular "Solo en Excel" (intersección de todos los PDFs)
        if global_only_excel is None:
            global_only_excel = df_only_excel_local
        else:
            # Solo mantener guías que no aparezcan en ningún PDF procesado
            matched_so_far = set(all_entregadas[0]["guia"]) | set() if all_entregadas else set()
            # Filtrar guías ya encontradas
            found_guias = (
                set(merged["guia"]) |
                set(pd.concat([r["guia"] for r in [{"guia": pd.Series(x["guia"] for x in all_entregadas + all_estado_diferente + all_solo_pdf)}]]).values)
                if all_entregadas or all_estado_diferente or all_solo_pdf else set(merged["guia"])
            )
            global_only_excel = global_only_excel[~global_only_excel["guia"].isin(found_guias)]

        entregadas = merged[(merged["estado_pdf"] == "ENTREGADO") & (merged["estado_excel"] == "ENTREGADO")].copy()
        estado_dif = merged[~((merged["estado_pdf"] == "ENTREGADO") & (merged["estado_excel"] == "ENTREGADO"))].copy()

        entregadas["factura"] = pdf_name
        entregadas["diferencia"] = entregadas["costo_pdf"] - entregadas["costo_excel"]
        estado_dif["factura"] = pdf_name
        df_only_pdf_local["factura"] = pdf_name

        all_entregadas.append(entregadas)
        all_estado_diferente.append(estado_dif)
        all_solo_pdf.append(df_only_pdf_local)

        total_guias = len(merged) + len(df_only_pdf_local)
        resumen_rows.append({
            "factura":         pdf_name,
            "entregadas":      len(entregadas),
            "estado_diferente":len(estado_dif),
            "solo_pdf":        len(df_only_pdf_local),
            "total_guias":     total_guias,
            "total_costo_pdf": (entregadas["costo_pdf"].sum() + estado_dif["costo_pdf"].sum()
                                + df_only_pdf_local["costo_pdf"].sum()),
            "total_costo_excel":(entregadas["costo_excel"].sum() + estado_dif["costo_excel"].sum()),
        })

    # Consolidar DataFrames
    df_ent  = pd.concat(all_entregadas, ignore_index=True)      if all_entregadas      else pd.DataFrame()
    df_dif  = pd.concat(all_estado_diferente, ignore_index=True) if all_estado_diferente else pd.DataFrame()
    df_spdf = pd.concat(all_solo_pdf, ignore_index=True)        if all_solo_pdf        else pd.DataFrame()
    df_sxls = global_only_excel if global_only_excel is not None else pd.DataFrame()

    wb = Workbook()
    wb.remove(wb.active)

    _write_entregadas(wb, df_ent)
    _write_estado_diferente(wb, df_dif)
    _write_solo_excel(wb, df_sxls)
    _write_solo_pdf(wb, df_spdf)
    _write_resumen(wb, resumen_rows, df_ent, df_dif, df_spdf, df_sxls)

    wb.save(output_path)
    log.info("Reporte guardado: %s", output_path)

    return {
        "entregadas":       len(df_ent),
        "estado_diferente": len(df_dif),
        "solo_excel":       len(df_sxls),
        "solo_pdf":         len(df_spdf),
    }

# ── Hojas ─────────────────────────────────────────────────────────────────────
def _write_entregadas(wb, df: pd.DataFrame):
    ws = wb.create_sheet("Entregadas")
    ws.append(["N° Guía", "Factura PDF", "Estado", "Costo PDF", "Costo Excel", "Diferencia"])
    _hdr(ws)
    if not df.empty:
        for _, r in df.iterrows():
            ws.append([r["guia"], r.get("factura",""), r["estado_pdf"],
                       r["costo_pdf"], r["costo_excel"], r["diferencia"]])
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=6, max_col=6):
            for cell in row:
                cell.fill = G_FILL if (cell.value == 0) else R_FILL
        ws.append(["TOTAL","","",
                   df["costo_pdf"].sum(), df["costo_excel"].sum(), df["diferencia"].sum()])
    else:
        ws.append(["Sin registros","","",0,0,0])
        ws.append(["TOTAL","","",0,0,0])
    _tot(ws); _currency(ws, [4,5,6]); _borders(ws); _autofit(ws)

def _write_estado_diferente(wb, df: pd.DataFrame):
    ws = wb.create_sheet("Estado Diferente")
    ws.append(["N° Guía", "Factura PDF", "Estado PDF", "Estado Excel", "Costo PDF", "Costo Excel"])
    _hdr(ws)
    if not df.empty:
        for _, r in df.iterrows():
            ws.append([r["guia"], r.get("factura",""),
                       r["estado_pdf"], r["estado_excel"],
                       r["costo_pdf"], r["costo_excel"]])
        ws.append(["TOTAL","","","", df["costo_pdf"].sum(), df["costo_excel"].sum()])
    else:
        ws.append(["Sin registros","","","",0,0])
        ws.append(["TOTAL","","","",0,0])
    _tot(ws); _currency(ws, [5,6]); _borders(ws); _autofit(ws)

def _write_solo_excel(wb, df: pd.DataFrame):
    ws = wb.create_sheet("Solo en Excel")
    ws.append(["N° Guía", "Estado", "Costo Excel"])
    _hdr(ws)
    if not df.empty:
        for _, r in df.iterrows():
            ws.append([r["guia"], r["estado_excel"], r["costo_excel"]])
        ws.append(["TOTAL","", df["costo_excel"].sum()])
    else:
        ws.append(["Sin registros","",0])
        ws.append(["TOTAL","",0])
    _tot(ws); _currency(ws, [3]); _borders(ws); _autofit(ws)

def _write_solo_pdf(wb, df: pd.DataFrame):
    ws = wb.create_sheet("Solo en PDF")
    ws.append(["N° Guía", "Factura PDF", "Estado", "Costo PDF"])
    _hdr(ws)
    if not df.empty:
        for _, r in df.iterrows():
            ws.append([r["guia"], r.get("factura",""), r["estado_pdf"], r["costo_pdf"]])
        ws.append(["TOTAL","","", df["costo_pdf"].sum()])
    else:
        ws.append(["Sin registros","","",0])
        ws.append(["TOTAL","","",0])
    _tot(ws); _currency(ws, [4]); _borders(ws); _autofit(ws)

def _write_resumen(wb, rows, df_ent, df_dif, df_spdf, df_sxls):
    ws = wb.create_sheet("Resumen General")
    ws.append(["Factura PDF","Entregadas","Estado Diferente","Solo PDF",
               "Total Guías","Total Costo PDF","Total Costo Excel"])
    _hdr(ws)
    for r in rows:
        ws.append([r["factura"], r["entregadas"], r["estado_diferente"],
                   r["solo_pdf"], r["total_guias"],
                   r["total_costo_pdf"], r["total_costo_excel"]])

    total_cnt = sum(r["total_guias"] for r in rows)
    total_cpdf  = (df_ent["costo_pdf"].sum()  if not df_ent.empty  else 0) + \
                  (df_dif["costo_pdf"].sum()   if not df_dif.empty  else 0) + \
                  (df_spdf["costo_pdf"].sum()  if not df_spdf.empty else 0)
    total_cexcel = (df_ent["costo_excel"].sum() if not df_ent.empty  else 0) + \
                   (df_dif["costo_excel"].sum()  if not df_dif.empty  else 0) + \
                   (df_sxls["costo_excel"].sum() if not df_sxls.empty else 0)

    ent_sum  = sum(r["entregadas"]       for r in rows)
    dif_sum  = sum(r["estado_diferente"] for r in rows)
    spdf_sum = sum(r["solo_pdf"]         for r in rows)

    ws.append(["TOTAL GENERAL", ent_sum, dif_sum, spdf_sum,
               total_cnt, total_cpdf, total_cexcel])
    _tot(ws); _currency(ws, [6,7]); _borders(ws); _autofit(ws)
