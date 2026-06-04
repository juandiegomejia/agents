"""
Transformador de Estados Financieros NIF - SUI Gas Natural 2025
================================================================
Toma los CSV crudos XBRL descargados por descargar_estados_financieros.py
y produce un Excel con estados financieros presentados y comparativos.

Uso:
    python transformar_estados_financieros.py

Salida:
    Estados_Financieros_Presentados_2025.xlsx
      - Hoja 1: Estado de Situacion Financiera  (Balance)
      - Hoja 2: Estado de Resultados
      - Hoja 3: Estado de Flujos de Efectivo
"""

from pathlib import Path
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
BASE = Path(__file__).parent
YEAR = 2025

COMPANIES = {
    "Vanti":      ("EF_2025_Vanti.csv",                        "VANTI S.A. ESP"),
    "GNC":        ("EF_2025_Gas_Natural_Cundiboyacense.csv",    "GAS NATURAL CUNDIBOYACENSE SA ESP"),
    "Gasoriente": ("EF_2025_Gas_Natural_Oriente.csv",           "GAS NATURAL DEL ORIENTE SA ESP"),
    "Gasnacer":   ("EF_2025_Gas_Natural_Cesar_Gasnacer.csv",    "GAS NATURAL DEL CESAR S.A. EMPRESA DE SERVICIOS PUBLICOS"),
}

# The "Acueducto [miembro]" member is the XBRL grand-total axis for the P&L
PL_MEMBER    = "Acueducto [miembro]"
TOTAL_MEMBER = "Total [miembro]"


# ---------------------------------------------------------------------------
# Line-item definitions (exact stripped concept names, in display order)
# Each tuple: (display_label, xbrl_concept_substring, is_header, is_total)
# ---------------------------------------------------------------------------

BS_ITEMS = [
    ("ACTIVOS",                                                        None,                                                                  True,  False),
    ("Activos corrientes",                                             None,                                                                  True,  False),
    ("  Efectivo y equivalentes al efectivo",                          "Efectivo y equivalentes al efectivo",                                 False, False),
    ("  Inventarios",                                                  "Inventarios corrientes",                                              False, False),
    ("  Cuentas por cobrar y otras (corriente)",                       "Total cuentas comerciales por cobrar y otras cuentas por cobrar corrientes", False, False),
    ("  Activos por impuestos corrientes",                             "Activos por impuestos corrientes",                                    False, False),
    ("  Otros activos corrientes",                                     "Otros activos no financieros corrientes",                             False, False),
    ("  Total activos corrientes",                                     "Total de activos corrientes",                                         False, True),
    ("Activos no corrientes",                                          None,                                                                  True,  False),
    ("  Propiedades, planta y equipo",                                 "Propiedades, planta y equipo",                                        False, False),
    ("  Activos intangibles",                                          "Activos intangibles distintos de la plusvalia",                       False, False),
    ("  Inversiones en subsidiarias",                                  "Inversiones en subsidiarias presentadas",                             False, False),
    ("  Inversiones en asociadas",                                     "Inversiones en asociadas presentadas",                                False, False),
    ("  Inversiones (metodo participacion)",                           "Inversiones contabilizadas utilizando el metodo",                     False, False),
    ("  Activos por impuestos diferidos",                              "Activos por impuestos diferidos",                                     False, False),
    ("  Otros activos no corrientes",                                  "Otros activos no financieros no corrientes",                          False, False),
    ("  Total activos no corrientes",                                  "Total de activos no corrientes",                                      False, True),
    ("TOTAL ACTIVOS",                                                  "Total de activos",                                                    True,  True),
    ("",                                                               None,                                                                  False, False),
    ("PASIVOS",                                                        None,                                                                  True,  False),
    ("  Pasivos corrientes",                                           None,                                                                  True,  False),
    ("    Obligaciones financieras corrientes",                        "Obligaciones financieras corrientes",                                 False, False),
    ("    Cuentas por pagar y otras (corriente)",                      "Total cuentas comerciales por pagar y otras cuentas por pagar corrientes", False, False),
    ("    Pasivos por impuestos corrientes",                           "Pasivos por impuestos corrientes, corriente",                         False, False),
    ("    Otros pasivos corrientes",                                   "Otros pasivos no financieros corrientes",                             False, False),
    ("    Total pasivos corrientes",                                   "Total pasivos corrientes",                                            False, True),
    ("  Pasivos no corrientes",                                        None,                                                                  True,  False),
    ("    Obligaciones financieras no corrientes",                     "Obligaciones financieras no corrientes",                              False, False),
    ("    Otros pasivos no corrientes",                                "Otros pasivos no financieros no corrientes",                          False, False),
    ("    Total pasivos no corrientes",                                "Total pasivos no corrientes",                                         False, True),
    ("TOTAL PASIVOS",                                                  "Total pasivos",                                                       True,  True),
    ("",                                                               None,                                                                  False, False),
    ("PATRIMONIO",                                                     None,                                                                  True,  False),
    ("  Capital emitido",                                              "Capital emitido",                                                     False, False),
    ("  Prima de emision",                                             "Prima de emision",                                                    False, False),
    ("  Reservas",                                                     "Otras reservas",                                                      False, False),
    ("  Reservas",                                                     "Reserva legal",                                                       False, False),
    ("  Ganancias acumuladas",                                         "Ganancias acumuladas",                                                False, False),
    ("  Efectos por adopcion NIF",                                     "Efectos por adopcion NIF",                                            False, False),
    ("TOTAL PATRIMONIO",                                               "Total patrimonio",                                                    True,  True),
    ("TOTAL PASIVOS + PATRIMONIO",                                     "Total patrimonio y pasivos",                                          True,  True),
]

PL_ITEMS = [
    ("Ingresos de actividades ordinarias",  "Ingresos de actividades ordinarias",                       False, False),
    ("(-) Costo de ventas",                 "Costo de ventas",                                          False, False),
    ("UTILIDAD BRUTA",                      "Ganancia bruta",                                           True,  True),
    ("(-) Gastos de administracion",        "Gastos de administracion",                                 False, False),
    ("(-) Otros gastos",                    "Otros gastos",                                             False, False),
    ("Otros ingresos operacionales",        "Otros ingresos",                                           False, False),
    ("UTILIDAD OPERACIONAL",                "Ganancia (perdida) por actividades de operacion",          True,  True),
    ("(+) Ingresos financieros",            "Ingresos financieros",                                     False, False),
    ("(-) Costos financieros",              "Costos financieros",                                       False, False),
    ("Otros ingresos (gastos) subsidiarias","Otros ingresos (gastos) procedentes de subsidiarias",      False, False),
    ("EBT (Antes de impuestos)",            "Ganancia (perdida), antes de impuestos",                   True,  True),
    ("(-) Impuesto de renta",               "Gasto (ingreso) por impuesto a las ganancias corriente",   False, False),
    ("UTILIDAD NETA",                       "Ganancia (perdida)",                                       True,  True),
]

CF_ITEMS = [
    ("Flujos de operacion",                 "Flujos de efectivo netos procedentes de (utilizados en) actividades de operacion", True,  True),
    ("Flujos de inversion",                 "Flujos de efectivo netos procedentes de (utilizados en) actividades de inversion", True,  True),
    ("Flujos de financiacion",              "Flujos de efectivo netos procedentes de (utilizados en) actividades de financiacion", True,  True),
    ("",                                    None,                                                                                False, False),
    ("CAPEX (compras de PPE)",              "Compras de propiedades, planta y equipo",                  False, False),
    ("Dividendos pagados",                  "Dividendos pagados",                                       False, False),
    ("Importes prestamos",                  "Importes procedentes de prestamos",                        False, False),
    ("Reembolsos prestamos",                "Reembolsos de prestamos",                                  False, False),
    ("",                                    None,                                                                                False, False),
    ("Incremento (disminucion) neto efectivo",
     "Incremento (disminucion) neto del efectivo y equivalentes de efectivo despues",                   True,  True),
    ("Efectivo inicio del periodo",         "Efectivo y equivalentes al efectivo al principio del periodo", False, False),
    ("Efectivo fin del periodo",            "Efectivo y equivalentes al efectivo al final del periodo", False, False),
    ("",                                    None,                                                                                False, False),
    ("Depreciacion y amortizacion",         "Ajustes por gastos de depreciacion y amortizacion",        False, False),
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_valor(v: str) -> float | None:
    """'$ 123,456,789.00'  →  123456789.0   (pesos)"""
    if pd.isna(v) or v is None:
        return None
    clean = re.sub(r"[$ ,]", "", str(v).strip())
    try:
        return float(clean)
    except ValueError:
        return None


def load_company(csv_path: Path) -> pd.DataFrame:
    df = pd.read_csv(csv_path, dtype=str)
    df["_valor_num"] = df["VALOR"].apply(parse_valor)
    df["_concepto_stripped"] = df["CONCEPTO"].str.strip()
    return df


def get_value(
    df: pd.DataFrame,
    eje: str,
    miembro: str,
    concepto_substr: str,
    period: int,
) -> float | None:
    mask = (
        (df["EJE"] == eje)
        & (df["MIEMBRO"] == miembro)
        & (df["PERIODO COMPARADO"] == str(period))
        & df["_valor_num"].notna()
        & df["_concepto_stripped"].str.contains(concepto_substr, regex=False, na=False)
    )
    rows = df[mask]
    if rows.empty:
        return None
    # If multiple rows, take the one whose stripped concept most closely matches
    # (prefer shortest = most specific match)
    rows = rows.copy()
    rows["_match_len"] = rows["_concepto_stripped"].str.len()
    return rows.sort_values("_match_len").iloc[0]["_valor_num"]


def build_table(items, eje, miembro, companies_data, years=(2025, 2024)):
    """
    Returns list of dicts: {label, is_header, is_total, (company, year): value, ...}
    """
    rows = []
    for label, substr, is_header, is_total in items:
        row = {"label": label, "is_header": is_header, "is_total": is_total}
        if substr is None:
            # Section header or spacer
            for short, (_, df) in companies_data.items():
                for yr in years:
                    row[(short, yr)] = None
        else:
            for short, (_, df) in companies_data.items():
                for yr in years:
                    val = get_value(df, eje, miembro, substr, yr)
                    row[(short, yr)] = val / 1e9 if val is not None else None  # → miles de millones COP
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

HEADER_FILL    = PatternFill("solid", fgColor="1F4E79")
SUBHEADER_FILL = PatternFill("solid", fgColor="2E75B6")
TOTAL_FILL     = PatternFill("solid", fgColor="D6E4F0")
EVEN_FILL      = PatternFill("solid", fgColor="F5F5F5")
WHITE_FILL     = PatternFill("solid", fgColor="FFFFFF")

HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
SUBHDR_FONT= Font(bold=True, color="FFFFFF", size=10)
TOTAL_FONT = Font(bold=True, size=10)
NORMAL_FONT= Font(size=10)
LABEL_FONT = Font(size=10)

thin   = Side(style="thin",   color="BFBFBF")
medium = Side(style="medium",  color="1F4E79")
THIN_BORDER   = Border(bottom=thin)
TOTAL_BORDER  = Border(top=medium, bottom=medium)

NUM_FMT = '#,##0.0'  # one decimal, thousands separator


def write_sheet(wb, sheet_name, table_rows, company_names, years=(2025, 2024)):
    ws = wb.create_sheet(title=sheet_name)

    shorts = list(company_names.keys())
    # --- Column layout ---
    # Col A: label
    # Col B+: (company, year) pairs
    col_headers = [(s, yr) for s in shorts for yr in years]
    total_cols = 1 + len(col_headers)

    # --- Header rows ---
    # Row 1: company groups
    ws.cell(1, 1, "CONCEPTO").font = HDR_FONT
    ws.cell(1, 1).fill = HEADER_FILL
    ws.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 20

    for i, s in enumerate(shorts):
        start_col = 2 + i * len(years)
        end_col   = start_col + len(years) - 1
        ws.cell(1, start_col, company_names[s]).font = HDR_FONT
        ws.cell(1, start_col).fill = HEADER_FILL
        ws.cell(1, start_col).alignment = Alignment(horizontal="center")
        if end_col > start_col:
            ws.merge_cells(
                start_row=1, start_column=start_col,
                end_row=1,   end_column=end_col
            )

    # Row 2: year sub-headers
    ws.cell(2, 1, "Miles de millones COP").font = SUBHDR_FONT
    ws.cell(2, 1).fill = SUBHEADER_FILL
    ws.cell(2, 1).alignment = Alignment(horizontal="left")
    ws.row_dimensions[2].height = 18
    for col_idx, (s, yr) in enumerate(col_headers, start=2):
        c = ws.cell(2, col_idx, str(yr))
        c.font    = SUBHDR_FONT
        c.fill    = SUBHEADER_FILL
        c.alignment = Alignment(horizontal="right")

    # --- Data rows ---
    for row_idx, row in enumerate(table_rows, start=3):
        label      = row["label"]
        is_header  = row["is_header"]
        is_total   = row["is_total"]

        ws.row_dimensions[row_idx].height = 16

        label_cell = ws.cell(row_idx, 1, label)
        label_cell.alignment = Alignment(horizontal="left", indent=0, vertical="center", wrap_text=False)

        if label == "":
            # Spacer row
            for col_idx in range(1, total_cols + 1):
                ws.cell(row_idx, col_idx).fill = WHITE_FILL
            continue

        if is_header:
            label_cell.font = Font(bold=True, size=10)
            label_cell.fill = TOTAL_FILL
            for col_idx, (s, yr) in enumerate(col_headers, start=2):
                c = ws.cell(row_idx, col_idx)
                c.fill = TOTAL_FILL
        else:
            label_cell.font = LABEL_FONT
            fill = EVEN_FILL if (row_idx % 2 == 0) else WHITE_FILL
            label_cell.fill = fill

        for col_idx, (s, yr) in enumerate(col_headers, start=2):
            val = row.get((s, yr))
            c   = ws.cell(row_idx, col_idx)

            if val is not None:
                c.value          = val
                c.number_format  = NUM_FMT
                c.alignment      = Alignment(horizontal="right")
            else:
                c.value = "-"
                c.alignment = Alignment(horizontal="right")

            if is_header:
                c.font = TOTAL_FONT
                c.fill = TOTAL_FILL
            else:
                fill = EVEN_FILL if (row_idx % 2 == 0) else WHITE_FILL
                c.fill = fill
                c.font = NORMAL_FONT

            if is_total:
                c.border = TOTAL_BORDER
                label_cell.border = TOTAL_BORDER

    # --- Column widths ---
    ws.column_dimensions["A"].width = 42
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14

    # Freeze top 2 rows + label column
    ws.freeze_panes = "B3"

    # --- Note row at top ---
    ws.insert_rows(1)
    ws.row_dimensions[1].height = 14
    note = f"Fuente: SUI – Superservicios. Datos NIF/XBRL. Año: {YEAR}. Valores en miles de millones de pesos colombianos."
    ws.cell(1, 1, note).font = Font(italic=True, size=9, color="595959")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Load all companies
    companies_data = {}
    for short, (fname, label) in COMPANIES.items():
        path = BASE / fname
        if not path.exists():
            print(f"  ADVERTENCIA: {fname} no encontrado. Ejecuta primero descargar_estados_financieros.py")
            continue
        companies_data[short] = (label, load_company(path))
        print(f"Cargado: {fname} ({len(companies_data[short][1])} filas)")

    if not companies_data:
        print("No hay datos. Abortando.")
        return

    company_names = {s: COMPANIES[s][1] for s in companies_data}

    # Build tables
    print("\nGenerando tablas...")

    bs_rows = build_table(
        BS_ITEMS, "Estado de situación financiera", TOTAL_MEMBER, companies_data
    )
    pl_rows = build_table(
        PL_ITEMS, "Estado de resultados - Resultado del periodo", PL_MEMBER, companies_data
    )
    cf_rows = build_table(
        CF_ITEMS, "Estado de flujos de efectivo", TOTAL_MEMBER, companies_data
    )

    # Write Excel
    wb = Workbook()
    wb.remove(wb.active)

    write_sheet(wb, "Situación Financiera", bs_rows, company_names)
    write_sheet(wb, "Estado de Resultados",  pl_rows, company_names)
    write_sheet(wb, "Flujos de Efectivo",    cf_rows, company_names)

    out = BASE / f"Estados_Financieros_Presentados_{YEAR}.xlsx"
    wb.save(out)
    print(f"\nGuardado: {out.name}")


if __name__ == "__main__":
    main()
