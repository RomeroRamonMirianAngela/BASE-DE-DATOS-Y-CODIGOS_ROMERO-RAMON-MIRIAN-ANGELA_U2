# Diccionario de variables

**Base:** `datos_procesados/datos_procesados_2024200523C.csv`
**Autora:** Mirian Angela Romero Ramón · **Código:** `2024200523C`
**Curso:** Finanzas I (055D), Unidad II, tema N.º 60 · Universidad Nacional del Centro del Perú

## 1. Datos generales

| Dato | Valor |
|---|---|
| Unidad de observación | Banco-mes (un registro por banco y por último día de cada mes) |
| Periodo | Enero de 2017 a diciembre de 2025 |
| Filas | 1 718 |
| Columnas | 25: columnas 1 a 13 con los valores originales de la SBS y el BCRP, y columnas 14 a 25 calculadas por `03_limpieza_datos.py` para los modelos |
| Bancos | 19 (ALFIN, BANBIF, BANCOM, BANK_OF_CHINA, BBVA, BCI, BCP, CENCOSUD, CITIBANK, COMPARTAMOS, FALABELLA, GNB, ICBC, INTERBANK, MIBANCO, PICHINCHA, RIPLEY, SANTANDER, SCOTIABANK) |
| Llave única | `banco_id` + `fecha` |
| Valores faltantes | Celda vacía. No se imputa ningún dato |
| SHA-256 | `7b3c188a51051d43782c146589e61611fdc773f4be9b9302849eea824800513b` |

Los porcentajes están en puntos: `3.12` significa 3.12 %, no 0.0312. Los montos están en miles de soles.

## 2. Columnas originales (1 a 13)

| N.º | Variable | Definición | Unidad | Frecuencia | Fuente exacta | URL o endpoint | Obs. no vacías | Rango |
|---:|---|---|---|---|---|---|---:|---|
| 1 | `banco_id` | Código del banco. Une los nombres históricos de una misma entidad (por ejemplo, Azteca y Alfin; Financiero y Pichincha; Continental y BBVA) | Texto | Banco-mes | SBS, B-2201 (encabezados de columna) | [B-2201](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2201) | 1 718 | 19 bancos |
| 2 | `banco_nombre` | Nombre del banco tal como aparece en el reporte de ese mes | Texto | Banco-mes | SBS, B-2201 | [B-2201](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2201) | 1 718 | |
| 3 | `fecha` | Último día del mes de referencia. Sirve para unir la SBS con el BCRP | AAAA-MM-DD | Mensual | Periodo del archivo SBS y de la serie BCRP | | 1 718 | 2017-01-31 a 2025-12-31 |
| 4 | `anio` | Año de `fecha` | Año | Mensual | Calculada en 03 | | 1 718 | 2017 a 2025 |
| 5 | `mes` | Mes de `fecha` (1 = enero, 12 = diciembre) | 1 a 12 | Mensual | Calculada en 03 | | 1 718 | 1 a 12 |
| 6 | `costo_fondeo` | Variable dependiente. Gastos por intereses de los últimos 12 meses (obligaciones con el público, depósitos del sistema financiero y organismos internacionales, y adeudados) divididos entre el saldo promedio de esas mismas cuentas en los 12 meses, por 100 | % anual | Mensual (ventana móvil de 12 meses) | SBS, B-2201 (saldos y gastos) y B-2390 (control de gastos financieros) | [B-2201](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2201), [B-2390](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2390) | 1 718 | 0.065 a 8.266 |
| 7 | `depositos` | Obligaciones con el público menos otras obligaciones con el público, más depósitos del sistema financiero y organismos internacionales | Miles de S/ | Saldo de fin de mes | SBS, B-2201 (balance); B-2342 como control | [B-2201](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2201) | 1 718 | 148 891 a 143 893 031 |
| 8 | `adeudados` | Adeudos y obligaciones financieras con otras entidades | Miles de S/ | Saldo de fin de mes | SBS, B-2201; B-2310 como control | [B-2310](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2310) | 1 718 | 0 a 12 220 967 |
| 9 | `deuda_subordinada` | Valores, títulos y obligaciones en circulación de carácter subordinado | Miles de S/ | Saldo de fin de mes | SBS, B-2201 | [B-2201](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2201) | 1 718 | 0 a 9 938 984 |
| 10 | `patrimonio_efectivo` | Patrimonio efectivo total reconocido por la SBS (capital regulatorio, no el patrimonio contable) | Miles de S/ | Mensual | SBS, B-2370; B-2402 para conciliar el ratio de capital | [B-2370](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2370) | 1 718 | 35 757 a 32 651 442 |
| 11 | `roe` | Rentabilidad sobre el patrimonio publicada por la SBS (utilidad neta anualizada entre patrimonio promedio) | % | Mensual | SBS, B-2401 | [B-2401](https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2401) | 1 684 | −395.559 a 37.404 |
| 12 | `tipmn_bcrp` | Tasa de interés pasiva promedio en moneda nacional (TIPMN). Es la misma para todos los bancos en un mes | % anual | Mensual | BCRPData, serie PN07816NM | `https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PN07816NM/csv/2017-1/2025-12/esp` | 1 718 | 0.763 a 3.963 |
| 13 | `tipmex_bcrp` | Tasa de interés pasiva promedio en moneda extranjera (TIPMEX). Es la misma para todos los bancos en un mes | % anual | Mensual | BCRPData, serie PN07836NM | `https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PN07836NM/csv/2017-1/2025-12/esp` | 1 718 | 0.200 a 2.003 |

## 3. Columnas calculadas para los modelos (14 a 25)

Todas se calculan en `03_limpieza_datos.py` con los valores originales de la misma fila o con la serie completa del BCRP. El script 04 las usa tal cual.

| N.º | Variable | Fórmula | Unidad | Frecuencia | Calculada a partir de | Uso en el artículo | Obs. no vacías | Rango |
|---:|---|---|---|---|---|---|---:|---|
| 14 | `fondeo_total` | `depositos + adeudados + deuda_subordinada` | Miles de S/ | Mensual | Columnas 7, 8 y 9 (SBS) | Denominador de las participaciones | 1 718 | 198 474 a 159 893 760 |
| 15 | `part_depositos` | `100 × depositos / fondeo_total` | % | Mensual | Columnas 7 y 14 | Tabla 4, Tabla 5, Figura 1 | 1 718 | 49.585 a 100 |
| 16 | `part_adeudados` | `100 × adeudados / fondeo_total` | % | Mensual | Columnas 8 y 14 | Tabla 5, Figura 1, robustez (Tabla 10) | 1 718 | 0 a 30.632 |
| 17 | `part_subordinada` | `100 × deuda_subordinada / fondeo_total` | % | Mensual | Columnas 9 y 14 | Modelo de ROE (Tabla 8), robustez | 1 718 | 0 a 20.528 |
| 18 | `capitalizacion` | `100 × patrimonio_efectivo / fondeo_total` | % | Mensual | Columnas 10 y 14 | Modelo de ROE, robustez | 1 718 | 5.954 a 121.131 |
| 19 | `tamano` | `ln(fondeo_total)` | Logaritmo natural | Mensual | Columna 14 | Modelo de ROE, robustez | 1 718 | 12.198 a 18.890 |
| 20 | `ln_depositos` | `ln(1 + depositos)` | Logaritmo natural | Mensual | Columna 7 | Modelo principal (H1) | 1 718 | 11.911 a 18.785 |
| 21 | `ln_adeudados` | `ln(1 + adeudados)` | Logaritmo natural | Mensual | Columna 8 | Modelo principal (H2) | 1 718 | 0 a 16.319 |
| 22 | `ln_deuda_subordinada` | `ln(1 + deuda_subordinada)` | Logaritmo natural | Mensual | Columna 9 | Modelo principal (H3) | 1 718 | 0 a 16.112 |
| 23 | `ln_patrimonio_efectivo` | `ln(1 + patrimonio_efectivo)` | Logaritmo natural | Mensual | Columna 10 | Modelo principal (H4) | 1 718 | 10.485 a 17.301 |
| 24 | `tipmn_ma12` | Promedio de la TIPMN de los últimos 12 meses (incluido el mes actual) | % anual | Mensual | Serie BCRP PN07816NM completa | Modelo principal (H5) y modelo de ROE | 1 553 | 0.867 a 3.685 |
| 25 | `tipmex_ma12` | Promedio de la TIPMEX de los últimos 12 meses (incluido el mes actual) | % anual | Mensual | Serie BCRP PN07836NM completa | Modelo principal (control) | 1 553 | 0.223 a 1.957 |

## 4. Notas

- **Costo de fondeo.** Los gastos del estado de resultados de la SBS son acumulados en el año, así que el script obtiene el gasto de cada mes por diferencia y reinicia el acumulado en enero. El indicador solo se calcula con 12 meses seguidos y con denominador positivo. Por eso se descargan también los reportes de 2016.
- **Promedios de 12 meses de las tasas.** Se calculan solo con 12 meses seguidos de datos. Como la serie del BCRP empieza en enero de 2017, `tipmn_ma12` y `tipmex_ma12` están vacías de enero a noviembre de 2017 (165 filas) y el modelo mensual empieza en diciembre de 2017. Se usan para que la tasa del mercado mida el mismo periodo que el costo de fondeo.
- **Logaritmos con 1 + x.** Algunos bancos no tienen adeudados o deuda subordinada en ciertos meses. Sumar 1 permite tomar el logaritmo cuando el saldo es cero sin borrar esas filas.
- **ROE vacío.** Faltan 34 valores de ROE, todos de Bank of China, porque el indicador no aparece para ese banco en el reporte B-2401 de esos meses.
- **Participaciones.** `part_depositos + part_adeudados + part_subordinada = 100` en cada fila; el script lo comprueba y se detiene si no se cumple.
- **Muestra de estimación.** El script 04 deja fuera a los bancos con menos de 24 meses (solo Compartamos, con 1 mes), así que trabaja con 1 717 filas de 18 bancos. En el panel anual exige al menos 4 años por banco.

## 5. Fuentes

| Institución | Reporte o serie | Contenido | Página |
|---|---|---|---|
| SBS | B-2201 | Balance general y estado de resultados por banco | https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2201 |
| SBS | B-2310 | Adeudos y obligaciones financieras | https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2310 |
| SBS | B-2342 | Estructura del pasivo | https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2342 |
| SBS | B-2370 | Patrimonio efectivo | https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2370 |
| SBS | B-2390 | Estructura de los gastos financieros | https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2390 |
| SBS | B-2401 | Indicadores financieros (ROE) | https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2401 |
| SBS | B-2402 | Requerimiento de patrimonio efectivo y ratio de capital | https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=B-2402 |
| BCRP | PN07816NM | TIPMN | https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PN07816NM/csv/2017-1/2025-12/esp |
| BCRP | PN07836NM | TIPMEX | https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PN07836NM/csv/2017-1/2025-12/esp |

Plantilla de los archivos históricos de la SBS: `https://intranet2.sbs.gob.pe/estadistica/financiera/{AAAA}/{Mes}/{reporte}-{mm}{AAAA}.XLS`. Las URL exactas de cada descarga quedan en `log_ejecucion.txt`.
