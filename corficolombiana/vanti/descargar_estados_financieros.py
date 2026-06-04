"""
Estados Financieros 2025 - SUI Gas Natural
===========================================
Descarga datos NIF/XBRL de las 4 distribuidoras de gas desde el reporte
Power BI del SUI (sui.superservicios.gov.co).

Instalacion:
    pip install playwright pandas openpyxl
    playwright install chromium

Uso:
    python descargar_estados_financieros.py
    python descargar_estados_financieros.py --year 2024   # otro año

Salida (misma carpeta del script):
    EF_2025_Vanti.csv
    EF_2025_Gas_Natural_Cundiboyacense.csv
    EF_2025_Gas_Natural_Oriente.csv
    EF_2025_Gas_Natural_Cesar_Gasnacer.csv
    Estados_Financieros_Gas_2025.xlsx   (una hoja por empresa)
"""

import argparse
import asyncio
from io import StringIO
from pathlib import Path

import pandas as pd
from playwright.async_api import async_playwright

REPORT_URL = (
    "https://wa-reportsui.azurewebsites.net/home/report/"
    "878dc8c1-9584-42ec-8545-69b184a65216"
)

# Nombres exactos tal como aparecen en el SUI (confirmados 2026-06-04)
COMPANIES = {
    "Vanti": "VANTI S.A. ESP",
    "Gas_Natural_Cundiboyacense": "GAS NATURAL CUNDIBOYACENSE SA ESP",
    "Gas_Natural_Oriente": "GAS NATURAL DEL ORIENTE SA ESP",
    # Gasnacer opera bajo razón social Gas Natural del Cesar en el SUI
    "Gas_Natural_Cesar_Gasnacer": (
        "GAS NATURAL DEL CESAR S.A. EMPRESA DE SERVICIOS PUBLICOS"
    ),
}

_EXPORT_JS = """
async ([company, yr]) => {
    const report = window.powerbi.embeds[0];

    await report.setFilters([
        {
            "$schema": "http://powerbi.com/product/schema#basic",
            "target": {"table": "Gas", "column": "NOMBRE_EMPRESA"},
            "filterType": 1,
            "operator": "In",
            "values": [company]
        },
        {
            "$schema": "http://powerbi.com/product/schema#basic",
            "target": {"table": "Gas", "column": "AÑO CARGUE"},
            "filterType": 1,
            "operator": "In",
            "values": [yr]
        }
    ]);

    // Wait for the visual to re-render with filtered data
    await new Promise(r => setTimeout(r, 5000));

    const pages = await report.getPages();
    const visuals = await pages[0].getVisuals();
    const table = visuals.find(v => v.type === 'tableEx');
    if (!table) throw new Error('Table visual not found');

    const result = await table.exportData(0);
    return result.data;
}
"""


async def fetch(page, company_name: str, year: int) -> pd.DataFrame:
    csv_data = await page.evaluate(_EXPORT_JS, [company_name, year])
    return pd.read_csv(StringIO(csv_data), dtype=str)


async def run(year: int, output_dir: Path) -> None:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        print(f"Abriendo reporte SUI ({REPORT_URL})...")
        await page.goto(REPORT_URL, timeout=90_000)

        print("Esperando que Power BI cargue...")
        await page.wait_for_function(
            "window.powerbi"
            " && window.powerbi.embeds"
            " && window.powerbi.embeds.length > 0",
            timeout=120_000,
        )
        # Give the report extra time to finish rendering all visuals
        await page.wait_for_timeout(8_000)
        print("Reporte listo.\n")

        dfs: dict[str, tuple[str, pd.DataFrame]] = {}

        for file_key, company_name in COMPANIES.items():
            print(f"Descargando {year} → {company_name} ...")
            try:
                df = await fetch(page, company_name, year)
                dfs[file_key] = (company_name, df)
                csv_path = output_dir / f"EF_{year}_{file_key}.csv"
                df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                print(f"  Guardado: {csv_path.name}  ({len(df):,} registros)")
            except Exception as exc:
                print(f"  ERROR: {exc}")

        if not dfs:
            print("No se descargo ninguna empresa.")
            await browser.close()
            return

        excel_path = output_dir / f"Estados_Financieros_Gas_{year}.xlsx"
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for file_key, (company_name, df) in dfs.items():
                df.to_excel(writer, sheet_name=file_key[:31], index=False)
        print(f"\nExcel combinado: {excel_path.name}")

        await browser.close()
        print("Listo.")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Descarga estados financieros NIF 2025 del SUI"
    )
    parser.add_argument(
        "--year", type=int, default=2025, help="Año a descargar (default: 2025)"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).parent,
        help="Carpeta de salida (default: misma carpeta del script)",
    )
    args = parser.parse_args()
    asyncio.run(run(args.year, args.output))


if __name__ == "__main__":
    main()
