# Input Contract

El insumo real del primer data product es un workbook Excel ubicado en `input/raw/`.

## Archivo actual

- Nombre actual: `BASE - COMPETENCIA - ENERO AL 22 DE MARZO.xlsx`
- Hoja utilizada por el pipeline: `BASE BRUTA`
- Formato soportado actualmente: `.xlsx`

## Columnas fuente esperadas

La hoja `BASE BRUTA` debe incluir, al menos, estas columnas:

- `Año`
- `Mes`
- `Fecha`
- `Tipo de medio`
- `Marca`
- `Inversión Neta`

El pipeline también normaliza el resto de las columnas observadas hoy, entre ellas `Anunciante`, `Producto`, `Medio`, `Programa`, `Hora`, `Inversión` y `Multimedia`.

## Normalización aplicada

- `Fecha` se interpreta como serial de Excel y se convierte a `YYYY-MM-DD`.
- `Mes` se transforma a llave mensual `YYYY-MM`.
- `Fecha` tambien se transforma a corte semanal con cierre al domingo (`week_ending`).
- `Marca` y `Tipo de medio` se llevan a mayúsculas estandarizadas.
- `Inversión` e `Inversión Neta` se convierten a valores numéricos.

## Tablas generadas

El pipeline genera una tabla detallada normalizada, un data product semanal y un producto de cambios respecto del corte anterior.

### Detalle normalizado

- `input/processed/latest_base_bruta.csv`
- `output/master/master_investment_detail.csv`
- `output/master/master_investment_detail.json`

### Data product

Carpeta: `output/data_products/inversion_mensual_por_casino_ilegal/`
Carpeta actual: `output/data_products/inversion_semanal_por_casino_ilegal/`

Archivos:

- `total.csv`
- `tv_abierta.csv`
- `tv_cable.csv`
- `radio.csv`
- `via_publica.csv`
- `digital.csv`
- `prensa.csv`

Cada CSV tiene una fila por marca y una columna por semana de cierre observada, más una columna `total`.

### Producto de cambios entre cortes

Carpeta: `output/data_products/cambios_vs_corte_anterior_semanal/`

Archivos:

- `changes_report.json`
- `total.csv`
- `tv_abierta.csv`
- `tv_cable.csv`
- `radio.csv`
- `via_publica.csv`
- `digital.csv`
- `prensa.csv`

Cada CSV muestra la diferencia entre el workbook actual y el corte inmediatamente anterior disponible en `input/raw/`.

## Regla editorial actual

Para este primer producto, se excluyen `MONTICELLO` y `XPERTO` del universo publicado como “casino/apuesta ilegal”. La lista vive en [schema.py](../src/schema.py).

## Validación

El pipeline falla si faltan campos críticos o si aparece un `Tipo de medio` no soportado.

El reporte queda en:

- `output/master/validation_report.json`
- `output/master/qa_report.json`

## QA de consistencia

Despues de generar los CSV, el pipeline compara los agregados publicados contra las hojas de control del workbook:

- `RESUMEN` para el total mensual por marca.
- hojas por marca para el desglose mensual por medio.

Si algun valor difiere por mas de `0.01`, el pipeline falla y deja el detalle en `output/master/qa_report.json`.
