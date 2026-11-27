# Estructura de financiamiento, costo de fondeo y rentabilidad de la banca peruana, 2017-2025

**Finanzas I (055D) · Unidad II · Tema N.º 60 · Periodo 2026-II**
**Universidad Nacional del Centro del Perú, Facultad de Economía**

| Dato | Detalle |
|---|---|
| Autora | Mirian Angela Romero Ramón |
| Código de matrícula | `2024200523C` |
| Tema | N.º 60: Estructura de financiamiento y costo de capital en el sistema bancario peruano |
| Docente | Dr. Ciro Iván Machacuay Meza |
| Periodo de estudio | Enero de 2017 a diciembre de 2025 (108 meses) |
| Fecha de corte | `FECHA_CORTE = 2025-12-31` (igual en los cuatro scripts) |
| Fecha de extracción | 26 de noviembre de 2026 (detalle en `log_ejecucion.txt`) |
| Repositorio | https://github.com/RomeroRamonMirianAngela/BASE-DE-DATOS-Y-CODIGOS_ROMERO-RAMON-MIRIAN-ANGELA_U2 |

## 1. De qué trata

El proyecto arma un panel banco-mes de la banca múltiple peruana con datos públicos de la SBS y del BCRP, y lo usa para responder una pregunta: ¿cambia el costo de fondeo de un banco cuando cambian los montos de sus fuentes de financiamiento (depósitos, adeudados, deuda subordinada y patrimonio efectivo), una vez que se toma en cuenta la tasa pasiva del mercado? Como segundo paso, mide cómo se relaciona ese costo con la rentabilidad sobre el patrimonio (ROE).

El diseño es observacional, así que los coeficientes se leen como asociaciones y no como efectos causales.

## 2. Fuentes y vías de extracción

El proyecto usa las dos vías que pide la consigna. Ninguna necesita clave de acceso.

| Vía | Institución | Qué se descarga | Endpoint |
|---|---|---|---|
| 1. API REST | BCRP (BCRPData) | `PN07816NM`: tasa pasiva promedio en moneda nacional (TIPMN) | `https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PN07816NM/csv/2017-1/2025-12/esp` |
| 1. API REST | BCRP (BCRPData) | `PN07836NM`: tasa pasiva promedio en moneda extranjera (TIPMEX) | `https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PN07836NM/csv/2017-1/2025-12/esp` |
| 2. Scraping y descarga programática | SBS | Reportes mensuales de banca múltiple B-2201, B-2310, B-2342, B-2370, B-2390, B-2401 y B-2402 (840 archivos, 2016-2025) | Página: `https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c={reporte}` · Archivo: `https://intranet2.sbs.gob.pe/estadistica/financiera/{AAAA}/{Mes}/{reporte}-{mm}{AAAA}.XLS` |

Los meses de 2016 se descargan solo para poder calcular el costo de fondeo móvil de 2017. No entran como observaciones.

## 3. Estructura de la carpeta

```text
BASE DE DATOS Y CÓDIGOS_ROMERO RAMON MIRIAN ANGELA_U2/
├── codigo/
│   ├── 01_extraccion_api.py      # vía 1: API de BCRPData
│   ├── 02_scraping_web.py        # vía 2: scraping y descarga de la SBS
│   ├── 03_limpieza_datos.py      # une, limpia y calcula las variables del modelo
│   └── 04_analisis.py            # 10 tablas y 9 figuras del artículo
├── datos_crudos/                 # archivos originales, sin editar
├── datos_procesados/
│   ├── datos_procesados_2024200523C.csv   # base oficial lista para el análisis
│   └── (base maestra, incidencias y auditorías que deja el 03)
├── salidas/                      # solo las 10 tablas y 9 figuras del artículo
├── diccionario_variables.md
├── requirements.txt
├── log_ejecucion.txt
└── README.md
```

## 4. Cómo replicar

1. Instalar Python 3.14 y las librerías:
   ```bash
   pip install -r requirements.txt
   ```
2. Correr los scripts en este orden, desde la carpeta del proyecto:
   ```bash
   python codigo/01_extraccion_api.py --modo completo
   python codigo/02_scraping_web.py --modo completo
   python codigo/03_limpieza_datos.py
   python codigo/04_analisis.py
   ```
   Sin `--modo completo`, los scripts 01 y 02 corren en modo de demostración y bajan solo un tramo corto del periodo. La descarga completa de la SBS toma unos 20 minutos por la pausa entre consultas.
3. Comprobar la huella de la base oficial (debe coincidir con la de la sección 6):
   ```bash
   certutil -hashfile datos_procesados\datos_procesados_2024200523C.csv SHA256   # Windows
   sha256sum datos_procesados/datos_procesados_2024200523C.csv                   # Linux o macOS
   ```

Cada script escribe en `log_ejecucion.txt` la fecha, la hora, la URL consultada y el código HTTP de cada descarga.

## 5. Qué hace cada script

| Script | Entrada | Proceso | Salida |
|---|---|---|---|
| `01_extraccion_api.py` | Códigos de serie y fechas | Consulta la API de BCRPData y verifica que lleguen los meses de cada serie | 2 CSV originales en `datos_crudos/` |
| `02_scraping_web.py` | Reportes y meses | Revisa `robots.txt`, lee las páginas de la SBS y descarga los Excel mensuales con una pausa entre consultas | 840 Excel originales en `datos_crudos/` |
| `03_limpieza_datos.py` | `datos_crudos/` | Homologa los nombres de los bancos, calcula el costo de fondeo, revisa la cobertura, registra incidencias y agrega las 12 variables del modelo | Base oficial de 25 columnas, base maestra, incidencias y auditorías en `datos_procesados/` |
| `04_analisis.py` | Solo la base oficial | Arma los paneles mensual y anual, aplica las pruebas de selección y diagnóstico, estima los modelos y la robustez | 10 tablas (CSV) y 9 figuras (PNG) en `salidas/` |

El script 04 no modifica la base: todas las variables que usa ya vienen calculadas desde el 03.

## 6. Base oficial

| Dato | Valor |
|---|---|
| Archivo | `datos_procesados/datos_procesados_2024200523C.csv` |
| Filas | 1 718 (banco-mes) |
| Columnas | 25: 13 originales y 12 calculadas para los modelos |
| Bancos | 19 |
| Meses | 108 (enero de 2017 a diciembre de 2025) |
| SHA-256 | `7b3c188a51051d43782c146589e61611fdc773f4be9b9302849eea824800513b` |

La definición de cada columna está en `diccionario_variables.md`.

## 7. Modelos

- **Modelo principal:** panel anual con el dato de diciembre de cada banco, estimado en primeras diferencias, con errores agrupados por banco. Usa el logaritmo de los cuatro montos y los promedios de 12 meses de la TIPMN y la TIPMEX. Tiene 124 observaciones de 16 bancos (2018-2025).
- **Comparaciones:** panel mensual con efectos fijos por banco y errores de Driscoll-Kraay (12 rezagos), y panel anual con efectos fijos.
- **Modelo de rentabilidad:** ROE de diciembre con efectos fijos por banco, en función del costo de fondeo, la participación de la deuda subordinada, la capitalización, el tamaño y la TIPMN de 12 meses.
- **Pruebas:** F de efectos fijos, Breusch-Pagan LM, Mundlak, Wooldridge (efectos fijos frente a primeras diferencias), VIF, Breusch-Pagan, AR(1) de los residuos, CD de Pesaran y Jarque-Bera.

Se usa el dato anual porque el costo de fondeo es un promedio móvil de 12 meses: en datos mensuales cada valor repite once meses del anterior y los residuos quedan casi perfectamente autocorrelacionados.

## 8. Salidas (todas aparecen en el artículo)

| Archivo | Contenido | En el artículo |
|---|---|---|
| `tabla1_descriptivos.csv` | Estadísticos descriptivos del panel banco-mes | Tabla 4 |
| `tabla2_estructura_bancos.csv` | Estructura del fondeo, costo y ROE por banco | Tabla 5 |
| `tabla3_construccion_muestra.csv` | Construcción de los paneles | Tabla 3 |
| `tabla4_seleccion_modelo.csv` | Pruebas para elegir el estimador | Tabla 6 |
| `tabla5_modelos_costo_fondeo.csv` | Tres estimaciones del costo de fondeo | Tabla 7 |
| `tabla6_pruebas_diagnostico.csv` | Pruebas econométricas de los tres modelos | Tabla 9 |
| `tabla7_robustez.csv` | Variantes del modelo principal | Tabla 10 |
| `tabla8_efectos_economicos.csv` | Efecto de 1 D. E. y betas estandarizados | Tabla 11 |
| `tabla9_modelo_roe.csv` | Costo de fondeo y ROE | Tabla 8 |
| `tabla10_contraste_hipotesis.csv` | Decisión sobre H1 a H6 | Tabla 12 |
| `figura1_composicion_fondeo.png` | Composición del fondeo: sistema y bancos | Figura 1 |
| `figura2_costo_y_tasas.png` | Costo de fondeo y tasas pasivas del BCRP | Figura 2 |
| `figura3_costo_por_banco.png` | Distribución del costo por banco | Figura 3 |
| `figura4_mapa_calor_banco_anio.png` | Costo promedio por banco y año | Figura 4 |
| `figura5_correlaciones.png` | Matriz de correlaciones | Figura 5 |
| `figura6_regresion_parcial.png` | Regresión parcial del modelo principal | Figura 6 |
| `figura7_comparacion_modelos.png` | Coeficientes e intervalos de los tres modelos | Figura 7 |
| `figura8_costo_y_roe.png` | Costo de fondeo y ROE | Figura 8 |
| `figura9_diagnostico_residuos.png` | Diagnóstico de los residuos | Figura 9 |

## 9. Entorno

Python 3.14 con las versiones fijadas en `requirements.txt`: requests 2.34.2, pandas 3.1.1, numpy 2.5.5, beautifulsoup4 4.15.0, openpyxl 3.1.5, xlrd 2.0.2, lxml 6.1.3, matplotlib 3.11.2, statsmodels 0.15.0, scipy 1.18.1 y linearmodels 7.1.

## 10. Advertencias

- El costo de fondeo es una medida contable (gastos de 12 meses entre saldos promedio de 12 meses), no un precio de mercado ni el WACC.
- El patrimonio efectivo es el capital regulatorio de la SBS y no el patrimonio contable.
- No se imputan datos: las celdas que la fuente no reporta quedan vacías.
- No se usan claves de API ni se recogen datos personales; las fuentes publican información agregada por banco.
