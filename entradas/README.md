# Datos de entrada

Esta carpeta **no se versiona**: los archivos los publica la CNBV y cambian
cada mes. El generador los espera con esta estructura exacta:

```
entradas/
├── tablero/
│   └── sh_datos_csv_27_i/
│       ├── 027_datos.csv                 serie mensual (en pesos)
│       ├── catalogo_conceptos_027.xlsx   1,502 conceptos
│       └── catalogo_instituciones_027.xlsx
└── BE_SOFIPOS_202606.xlsx                hoja Universo (junio 2026)
```

## Cómo obtenerlos

1. **Tablero de información financiera** de entidades supervisadas:
   <https://app.powerbi.com/view?r=eyJrIjoiMTlhN2U5Y2QtYTcxNi00MzQ3LTg2MGYtNWI5ZjBhNDk2MmYyIiwidCI6IjVlMmM0OTc3LTEwN2QtNDBhMy04YWY3LTcwMDc0ODFhNjBkNCIsImMiOjR9>

   Descarga la serie **`sh_datos_csv_27_i`** (SOFIPOs) y descomprímela en
   `entradas/tablero/sh_datos_csv_27_i/`. Debe quedar ahí el archivo
   `027_datos.csv` junto con los dos catálogos.

2. **Boletines electrónicos (BE)** de SOFIPOs:
   <https://portafolioinfo.cnbv.gob.mx/Paginas/Contenidos.aspx?ID=27&Contenido=Boletines&Titulo=Sociedades%20Financieras%20Populares>

   Descarga el BE del mes que corresponda y guárdalo en `entradas/` con el
   nombre `BE_SOFIPOS_AAAAMM.xlsx`. Del boletín solo se usa la hoja
   `Universo`; el resto de las cifras sale del tablero.

3. **Nivel de capitalización y alertas tempranas** (opcional, solo consulta):
   <https://portafolioinfo.cnbv.gob.mx/Paginas/Contenidos.aspx?ID=27&Contenido=Nivel%20de%20Capitalizaci%C3%B3n%20y%20Alertas%20Tempranas&Titulo=Sociedades%20Financieras%20Populares>

## Notas

- El `027_datos.csv` pesa ~54 MB y viene en **pesos**; el generador lo
  convierte a **miles de pesos**.
- Solo se usan saldos **sin consolidar** (`tipo_saldo = 130`).
- La fila con `institucion = 27` es el total del sector y se excluye.
- Si actualizas el tablero o el BE, revisa que el nombre del BE en
  `generar_sofipos_2026.py` (`BE_UNIVERSO`) coincida con el archivo descargado.
