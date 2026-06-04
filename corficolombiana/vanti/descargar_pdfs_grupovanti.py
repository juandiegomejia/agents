"""
Descargador de PDFs — Informes de Gestión Grupo Vanti 2025
==========================================================
Descarga los informes anuales auditados de las 4 empresas de Grupo Vanti
desde grupovanti.com y los guarda en la carpeta de destino.

Uso:
    python descargar_pdfs_grupovanti.py
    python descargar_pdfs_grupovanti.py --output /ruta/personalizada

Fuente: grupovanti.com (Contentful CDN — assets.ctfassets.net / downloads.ctfassets.net)
Nota: GasNacer no tiene PDF publicado en el sitio (ver Supersociedades SIIS, NIT 804000551).
"""

import argparse
import urllib.request
from pathlib import Path

# URLs confirmadas el 2026-06-04 navegando grupovanti.com
PDFS = [
    {
        "empresa": "Vanti S.A. ESP",
        "archivo": "Informe_Gestion_Vanti_2025.pdf",
        "url": "https://downloads.ctfassets.net/3brzg7q3bvg1/2ayUeuDq2YpWQLqvSHeJwJ/98ce78a1ae2526869f65df5f53009f3e/vfINFORME_VANTI_S.A._ESP_2025.pdf",
        "pagina": "https://www.grupovanti.com/conocenos/vanti-sa-esp/informacion-financiera-y-de-la-sociedad",
    },
    {
        "empresa": "Gas Natural Cundiboyacense S.A. ESP",
        "archivo": "Informe_Gestion_GNC_2025.pdf",
        "url": "https://downloads.ctfassets.net/3brzg7q3bvg1/6LXdeaIflVbjsPgIszltQD/351bff78cfed347d3096ca8744e78a50/IG_Cundiboyacense_2025__1_.pdf",
        "pagina": "https://www.grupovanti.com/conocenos/gas-natural-cundiboyacense-sa-esp/informacion-financiera-y-de-la-sociedad",
    },
    {
        "empresa": "Gas Natural del Oriente S.A. ESP (Gasoriente)",
        "archivo": "Informe_Gestion_Gasoriente_2025.pdf",
        "url": "https://downloads.ctfassets.net/3brzg7q3bvg1/1Gy8YuvixoQsSF0ayW7FEV/e00c50ef37310e88d42dce2e1afc5629/INFORME_GAS_NATURAL_DEL_ORIENTE_V2.pdf",
        "pagina": "https://www.grupovanti.com/conocenos/gasoriente-sa-esp/informacion-financiera-y-de-la-sociedad",
    },
    # GasNacer (Gas Natural del Cesar): sin PDF publicado a 2026-06-04
    # Estados financieros individuales disponibles en:
    #   Supersociedades SIIS → https://siis.ia.supersociedades.gov.co/ (NIT 804000551)
    # Estados como subsidiaria incluidos en el consolidado de Gasoriente (pdf anterior).
]

HEADERS = {"User-Agent": "Mozilla/5.0"}


def descargar(url: str, dest: Path) -> int:
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=120) as r, open(dest, "wb") as f:
        data = r.read()
        f.write(data)
    return len(data)


def main() -> None:
    parser = argparse.ArgumentParser(description="Descarga PDFs Grupo Vanti 2025")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).parent,
        help="Carpeta de destino (default: misma carpeta del script)",
    )
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)

    for item in PDFS:
        dest = args.output / item["archivo"]
        if dest.exists():
            print(f"Ya existe: {item['archivo']} — omitiendo")
            continue
        print(f"Descargando: {item['empresa']} ...")
        try:
            size = descargar(item["url"], dest)
            print(f"  Guardado: {item['archivo']}  ({size/1024/1024:.1f} MB)")
        except Exception as e:
            print(f"  ERROR: {e}")
            print(f"  Descarga manual: {item['pagina']}")

    print("\nGasNacer: sin PDF publicado.")
    print("  → Supersociedades SIIS: https://siis.ia.supersociedades.gov.co/ (NIT 804000551)")


if __name__ == "__main__":
    main()
