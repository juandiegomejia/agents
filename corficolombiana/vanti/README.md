# Estados Financieros Gas Natural — SUI

Descarga y transforma estados financieros NIF/XBRL de las 4 distribuidoras de gas natural de Grupo Vanti desde el SUI (Superservicios).

## Empresas

| Empresa | NIT | Nombre en SUI |
|---|---|---|
| Vanti S.A. ESP | 800007813 | VANTI S.A. ESP |
| Gas Natural Cundiboyacense S.A. ESP | — | GAS NATURAL CUNDIBOYACENSE SA ESP |
| Gas Natural del Oriente S.A. ESP | — | GAS NATURAL DEL ORIENTE SA ESP |
| Gas Natural del Cesar (Gasnacer) | 804000551 | GAS NATURAL DEL CESAR S.A. EMPRESA DE SERVICIOS PUBLICOS |

Fuente de datos: [SUI — Información financiera bajo NIF](https://sui.superservicios.gov.co/Reportes-del-sector/Gas-natural/Reportes-financieros/Informacion-financiera-bajo-NIF-de-los-prestadores-del-servicio-de-Gas)

## Instalación

```bash
pip install playwright pandas openpyxl pypdf
playwright install chromium
```

## Uso

```bash
# 1. Descargar datos crudos XBRL del SUI
python descargar_estados_financieros.py

# 2. Transformar a estados financieros presentados
python transformar_estados_financieros.py

# Año diferente
python descargar_estados_financieros.py --year 2024
```

## Salida

`descargar_estados_financieros.py` genera:
- `EF_{año}_{empresa}.csv` — datos crudos XBRL (una fila por concepto)

`transformar_estados_financieros.py` genera:
- `Estados_Financieros_Presentados_{año}.xlsx` con 3 hojas:
  - Situación Financiera (Balance)
  - Estado de Resultados
  - Flujos de Efectivo

Valores en miles de millones de pesos colombianos. Comparativo 2 años.

## Notas técnicas

- El reporte SUI es un Power BI embebido. El script usa el SDK de Power BI (`window.powerbi.embeds[0]`) para aplicar filtros y exportar datos directamente, sin depender de credenciales.
- El miembro `Acueducto [miembro]` en el XBRL funciona como total consolidado del Estado de Resultados (peculiaridad de la taxonomía colombiana).
- Datos actualizados diariamente en el SUI.
