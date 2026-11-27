# Nombres y apellidos completos: ROMERO RAMON MIRIAN ANGELA
# Codigo de matricula: 2024200523C
# Tema 60: Estructura de financiamiento y costo de capital en el sistema bancario peruano
# Fecha de extraccion: 2026-11-26

"""FINANZAS I - UNIDAD II | 03: LIMPIEZA E INTEGRACION SBS + BCRP.

Ejecutar desde la raiz del proyecto:
  python codigo/03_limpieza_datos.py --solo-anio 2025
  python codigo/03_limpieza_datos.py
  python codigo/03_limpieza_datos.py --max-bancos 0  # todas las entidades
  python codigo/03_limpieza_datos.py --solo-visualizar # CSV oficial + Excel desde base maestra

VERSION CORREGIDA (v5, 2026-11-26). Base de v3: BCP/Interbank con la entidad "con sucursales"
  (igual que los cuadros auxiliares SBS) y flujos negativos conservados.
Cambios de la v2, todos auditables:
  1. Por defecto se usan TODAS las entidades bancarias (--max-bancos 0).
  2. Alias historicos de bancos renombrados (Azteca->Alfin, Financiero->Pichincha,
     BanBif, Cencosud) para no partir ni perder series de la misma entidad.
  3. El guion "-" de la SBS en cuentas de B-2201 se lee como saldo/gasto CERO
     (convencion de presentacion SBS), solo si el banco reporta TOTAL ACTIVO
     ese mes. Cada caso queda en auditoria (GUION_SBS_CERO). Desactivable con
     --guion-como-faltante. Las celdas vacias siguen siendo NA.
  4. Periodo de estudio: enero de 2017 a diciembre de 2025 (108 meses), igual
     que en los scripts 01, 02 y 04. La ventana previa 2016 (si existen los
     crudos SBS de 2016) se lee solo para calcular el costo de fondeo movil
     de 12 meses desde enero de 2017. Las filas 2016 NUNCA salen en el panel.
  5. BCP e Interbank: si el B-2201 trae ambas coberturas, se prioriza
      "con sucursales" como en unificacion v3. Su consistencia con los
      reportes auxiliares se comprueba mediante las conciliaciones y la
      cobertura elegida se documenta en la base maestra.
  6. Atipicos: se siguen MARCANDO (MAD/IQR) sin borrar. La winsorizacion pasa a
     ser GLOBAL 1%-99% (pooled, estandar en la literatura bancaria); la version
     por banco sigue disponible con --winsor-modo banco. Los valores fuera de
     dominio (p. ej. ratio > 100%) quedan NA en la version *_w y se reportan.
  7. Variables ln_* (log natural) de los montos para el modelo: reducen la
     asimetria por tamano sin recortar datos.
  8. Embudo de observaciones (embudo_observaciones_<codigo>.csv) que muestra
     en que filtro se pierden filas, y verificacion de atipicos residuales.
  Nada de esto imputa ni inventa datos: solo cambia que filas/bancos entran.

Produce tres bases diferentes, sin editar los originales:
  1) datos_procesados_2024200523C.csv: BASE LISTA PARA EL ANALISIS.
     Columnas 1-13: cinco identificadores y ocho variables SBS/BCRP con sus
     valores ORIGINALES (cotejables con la fuente). Columnas 14-25: doce
     variables calculadas para el modelo (participaciones, capitalizacion,
     tamano, logaritmos y promedios de 12 meses de TIPMN/TIPMEX).
     Selecciona CF + cuatro montos SBS + TIPMN completos; ROE/TIPMEX pueden
     ser NA.
  2) panel_analisis_winsorizado_2024200523C.csv: la misma muestra,
     con valores originales, logaritmos, ratios, columnas *_w y banderas.
  3) base_maestra_2024200523C.csv: todas las filas y las auditorias.
Los atipicos se conservan y se documentan; comparar modelos entre
originales y winsorizados permite medir sensibilidad.

NO hace solicitudes web ni modifica datos_crudos/. Sus resultados y auditorias
se guardan solo en datos_procesados/, salvo el log tecnico en la raiz.
Una observacion por banco y por ultimo dia de mes; panel no balanceado real.
Las celdas no reportadas quedan vacias (NA), nunca se inventan.

Advertencias metodologicas:
- CF es una aproximacion contable de 12 meses moviles, no el WACC.
- El indicador de fondeo combina gastos sobre obligaciones con el publico,
  depositos interbancarios y adeudados; su denominador usa esos mismos saldos.
- 'depositos' EXCLUYE otras obligaciones del publico, mientras que el
  denominador CF conserva obligaciones con el publico completas para mantener
  la correspondencia con los gastos publicados por la SBS.
- El patrimonio efectivo SBS B-2370 usa tres encabezados verificados:
  NIVEL 1/2/3 + TOTAL en E (2018,2022); NIVEL 1/2 + TOTAL en D
  (enero-mayo 2023); Patrimonio Efectivo Total en E (junio 2023,2025).
  El formato de cada mes se guarda en formato_pe_b2370. La comparabilidad
  normativa debe discutirse en la investigacion; no es patrimonio contable.
- INTERBANK se distingue por cobertura: la columna individual es INTERBANK,
  y "Interbank Peru con Sucursales en el Exterior" es INTERBANK_EXT.
  Nunca se mezclan sus saldos ni sus gastos, ni se suman entre si.
- Los valores atipicos originales se conservan en la base oficial.
  Las banderas y los valores tratados estan en el panel de sensibilidad.
- La winsorizacion se aplica SOLO a tasas/ratios financieros (costo de fondeo,
  ROE y ratios de estructura), con percentiles 1 y 99 de toda la muestra
  (modo global, por defecto) o de cada banco (--winsor-modo banco); no modifica
  montos monetarios ni tasas BCRP. Los originales permanecen junto a los *_w.
- La winsorizacion es una version alternativa para analisis, no una correccion
  de errores ni una sustitucion de la base oficial.
- Las cuentas y la cobertura de los Excel anteriores a diciembre de 2025
  deben ser auditadas: los layouts historicos pueden cambiar.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import logging
import math
import os
import re
import shutil
import sys
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import load_workbook

CODIGO_MATRICULA = "2024200523C"
# El mismo periodo se declara en los scripts 01, 02 y 04.
FECHA_INICIO = "2017-01-01"
# Los crudos 2016 (si existen) se leen SOLO para la ventana movil de 12 meses
# del costo de fondeo de 2017; nunca se exportan como observaciones.
FECHA_INICIO_LECTURA = "2016-01-01"
FECHA_CORTE = "2025-12-31"
RAIZ = Path(__file__).resolve().parent.parent
CRUDOS = RAIZ / "datos_crudos"
PROCESADOS = RAIZ / "datos_procesados"
ARCHIVO_LOG = RAIZ / "log_ejecucion.txt"
SBS_REPORTES = ("B-2201", "B-2310", "B-2342", "B-2370", "B-2390", "B-2401", "B-2402")
BCRP_SERIES = {"PN07816NM": "tipmn_bcrp", "PN07836NM": "tipmex_bcrp"}
# Estas cuentas son identificaciones comprobadas sobre el Excel SBS 2025-12.
# Si cambia una etiqueta o encabezado no se sustituyen automaticamente.
TOLERANCIA_CONCILIACION = 0.015  # 1.5%; marca diferencia, nunca ajusta valores
MIN_CASOS_ATIPICO = 12
UMBRAL_MAD = 4.5
# Winsorizacion global p1-p99, alternativa de sensibilidad; por banco opcional.
WINSOR_LIM_INF = 0.01
WINSOR_LIM_SUP = 0.99
MIN_CASOS_WINSOR = 24
WINSOR_MODO = "global"  # "global" (pooled) o "banco"
# Cuentas de B-2201 donde el guion SBS "-" significa saldo o gasto nulo.
# No incluye totales: si TOTAL ACTIVO/PASIVO viene con guion, queda NA.
CUENTAS_GUION_CERO = (
    "otras_oblig_publico", "depositos_sistema", "adeudados", "deuda_subordinada",
    "gasto_publico_acum", "gasto_depsistema_acum", "gasto_adeudados_acum",
)
GUIONES_SBS = {"-", "--", "\u2013", "\u2014"}
OBJETIVO_OBSERVACIONES = 1000
VARIABLES_WINSOR = (
    "costo_fondeo", "roe", "dep_ratio", "ade_ratio",
    "ds_ratio", "pe_ratio"
)

RE_SBS = re.compile(r"^sbs_original_(B-\d+)_(\d{6})_2024200523C\.(xls|xlsx)$", re.I)
MISSING = {"", "-", "--", "n.d.", "n.d", "nd", "na", "n/a", "null", "none", "#div/0!", "#n/a", "#value!", "#ref!"}
MONTHS = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,"jul":7,"ago":8,"sep":9,"set":9,"oct":10,"nov":11,"dic":12}

# Vista Excel: identificadores y variables del estudio; la base maestra
# conserva todas las columnas historicas con faltantes e incidencias.
COLUMNAS_BASE_ANALISIS = [
    "banco_nombre", "anio", "mes",
    "costo_fondeo", "depositos", "adeudados",
    "deuda_subordinada", "patrimonio_efectivo", "roe",
]
VARIABLES_PRINCIPALES = [
    "costo_fondeo", "depositos", "adeudados",
    "deuda_subordinada", "patrimonio_efectivo",
]

VARIABLES_SEIS = VARIABLES_PRINCIPALES + ["roe"]
VARIABLES_BCRP = ["tipmn_bcrp", "tipmex_bcrp"]
# El archivo oficial conserva EXCLUSIVAMENTE los valores originales de estudio.
# Las transformaciones, las banderas y la auditoria permanecen en archivos
# paralelos. La separacion ayuda al cotejo SBS/BCRP exigido por la consigna.
COLUMNAS_CSV_OFICIAL = [
    "banco_id", "banco_nombre", "fecha", "anio", "mes",
    "costo_fondeo", "depositos", "adeudados", "deuda_subordinada",
    "patrimonio_efectivo", "roe", "tipmn_bcrp", "tipmex_bcrp",
]
# Criterio de la muestra principal: CF, 4 X originales y TIPMN.
# ROE y TIPMEX son complementarios: la ausencia no elimina banco-mes.
COLUMNAS_MUESTRA_PRINCIPAL = [
    "costo_fondeo", "depositos", "adeudados",
    "deuda_subordinada", "patrimonio_efectivo", "tipmn_bcrp",
]
# Transformaciones auxiliares SOLO en panel paralelo.
COLUMNAS_LOG_MODELO = [
    "ln_depositos", "ln_adeudados", "ln_deuda_subordinada",
    "ln_patrimonio_efectivo",
]
# Variables calculadas que se AGREGAN al final del CSV oficial (version
# 2026-11-26). No reemplazan a ninguna columna original: se calculan con
# los 13 valores originales y con la serie BCRP completa del crudo.
COLUMNAS_DERIVADAS_MODELO = [
    "fondeo_total",        # depositos + adeudados + deuda_subordinada (S/ mil)
    "part_depositos",      # 100 * depositos / fondeo_total (%)
    "part_adeudados",      # 100 * adeudados / fondeo_total (%)
    "part_subordinada",    # 100 * deuda_subordinada / fondeo_total (%)
    "capitalizacion",      # 100 * patrimonio_efectivo / fondeo_total (%)
    "tamano",              # ln(fondeo_total)
    "ln_depositos", "ln_adeudados", "ln_deuda_subordinada",
    "ln_patrimonio_efectivo",  # ln(1 + monto), porque puede haber ceros
    "tipmn_ma12",          # promedio de 12 meses de la TIPMN (BCRP)
    "tipmex_ma12",         # promedio de 12 meses de la TIPMEX (BCRP)
]
VENTANA_TASAS = 12  # mismo horizonte de 12 meses que el costo de fondeo


def norm(s: Any) -> str:
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9]+", " ", s.casefold())
    return " ".join(s.split())


def banca_id(nombre: Any) -> str | None:
    """Alias historicos explicitos; NO une bancos distintos silenciosamente."""
    s = norm(nombre)
    if not s or "total banca" in s or "total sistema" in s:
        return None
    if "credito" in s and ("peru" in s or "sucursales" in s):
        return "BCP_EXT" if "sucursales" in s or "exterior" in s else "BCP_DOM"
    # B-2201 historico presenta simultaneamente INTERBANK individual y la
    # variante "Interbank Peru con Sucursales en el Exterior". Son COBERTURAS
    # contables distintas, no encabezados duplicados (ver 2018-01 hoja 1/2).
    # Solo INTERBANK individual se integra con las tablas auxiliares SBS que
    # identifican la empresa como "Interbank"; la variante EXT queda separada.
    if "interbank" in s:
        return "INTERBANK_EXT" if "sucursales" in s or "exterior" in s else "INTERBANK"
    reglas = [
        # Identidad historica verificada en el B-2370 de dic-2018:
        # "B. Continental" corresponde a la actual entidad BBVA Peru.
        ("b continental", "BBVA"),
        ("banco continental", "BBVA"),
        ("bbva", "BBVA"),
        # Banco Financiero del Peru cambio su denominacion a Banco Pichincha
        # (2019): misma entidad supervisada, se une la serie.
        ("b financiero", "PICHINCHA"), ("banco financiero", "PICHINCHA"),
        # Banco Azteca del Peru paso a denominarse Alfin Banco (2021).
        ("azteca", "ALFIN"),
        ("banbif", "BANBIF"),
        ("cencosud", "CENCOSUD"),
        ("bancom", "BANCOM"), ("banco de comercio", "BANCOM"),
        ("b de comercio", "BANCOM"),
        ("pichincha", "PICHINCHA"),
        ("interamericano de finanzas", "BANBIF"),
        ("scotiabank", "SCOTIABANK"),
        ("citibank", "CITIBANK"),
        ("interbank", "INTERBANK"),
        ("mibanco", "MIBANCO"),
        ("falabella", "FALABELLA"),
        ("santander consumer", "SANTANDER_CONSUMER"),
        ("santander", "SANTANDER"),
        ("ripley", "RIPLEY"),
        ("alfin", "ALFIN"),
        ("icbc", "ICBC"),
        ("bank of china", "BANK_OF_CHINA"),
        ("bci", "BCI"),
        ("compartamos", "COMPARTAMOS"),
        ("gnb", "GNB"),
        ("hsbc", "HSBC"),
    ]
    for marca, codigo in reglas:
        if marca in s:
            return codigo
    return None


# ---------------------------------------------------------------------------
# CORRECCION v3 (2026-11-25): BCP e Interbank
# B-2201 trae dos columnas para BCP e Interbank: la individual y la "con
# sucursales en el exterior". Los reportes auxiliares SBS (B-2310, B-2342,
# B-2370, B-2390, B-2401) tienen UNA sola fila por banco, y sus cifras
# coinciden EXACTAMENTE con la columna "con sucursales" de B-2201
# (verificado en auditoria_cuentas: adeudos B-2310 = adeudados B-2201 _EXT y
# pasivo B-2342 = pasivo B-2201 _EXT, todos los meses). Ademas, la columna
# individual del BCP daba deuda subordinada negativa en 2018-2020.
# Regla: se usa la entidad "con sucursales" (la entidad supervisada que
# reportan todos los cuadros SBS) y se le da un solo ID: BCP / INTERBANK.
# ---------------------------------------------------------------------------
ALIAS_COBERTURA = {"BCP_EXT": "BCP", "BCP_DOM": "BCP", "INTERBANK_EXT": "INTERBANK"}


def unificar_coberturas(d: dict[str, dict]) -> dict[str, dict]:
    """Homologa BCP e Interbank, priorizando 'con sucursales' si existe.

    Solo se elige una cobertura del B-2201 por banco-mes. Se registra la
    procedencia, que debe conciliarse con los cuadros auxiliares SBS; no
    se suman saldos ni se imputan datos de una cobertura a otra.
    """
    for base, individual, sucursales in (("BCP", "BCP_DOM", "BCP_EXT"),
                                         ("INTERBANK", "INTERBANK", "INTERBANK_EXT")):
        if sucursales in d:
            d.pop(individual, None)
            registro = d.pop(sucursales)
            cobertura = "CON_SUCURSALES"
        elif individual in d:
            registro = d.pop(individual)
            cobertura = "INDIVIDUAL_UNICA"
        else:
            continue
        registro["banco_id"] = base
        registro["cobertura_b2201"] = cobertura
        d[base] = registro
    return d


def num(v: Any) -> float:
    """Convierte celdas sin rellenar no disponibles con cero."""
    if isinstance(v, (int, float, np.number)) and not isinstance(v, bool):
        z = float(v)
        return z if math.isfinite(z) else np.nan
    if v is None:
        return np.nan
    s = str(v).strip().lower()
    if s in MISSING:
        return np.nan
    s = s.replace(" ", "")
    if re.fullmatch(r"\(.*\)", s):
        s = "-" + s[1:-1]
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        z = float(s)
        return z if math.isfinite(z) else np.nan
    except ValueError:
        return np.nan


def es_guion_sbs(v: Any) -> bool:
    return isinstance(v, str) and v.strip() in GUIONES_SBS


def num_sbs(v: Any, variable: str, guion_cero: bool) -> tuple[float, bool]:
    """Como num(), pero el guion SBS en cuentas no totales se lee como 0.

    Devuelve (valor, fue_guion). La SBS publica "-" cuando la entidad no tiene
    saldo/gasto en esa cuenta (p. ej. bancos sin deuda subordinada). Leerlo
    como NA eliminaba del panel a bancos que SI reportaron su balance.
    """
    if guion_cero and variable in CUENTAS_GUION_CERO and es_guion_sbs(v):
        return 0.0, True
    return num(v), False


def log_setup() -> logging.Logger:
    logger = logging.getLogger("limpieza_finanzas_u2")
    logger.handlers.clear()
    logger.setLevel(logging.INFO)
    logger.propagate = False
    h = logging.FileHandler(ARCHIVO_LOG, encoding="utf-8")
    h.setFormatter(logging.Formatter("%(asctime)s | LIMPIEZA | %(levelname)s | %(message)s"))
    logger.addHandler(h)
    return logger


def buscar_archivos(solo_anio: int | None) -> tuple[dict[pd.Timestamp,dict[str,Path]], list[dict]]:
    """Cataloga crudos SBS. Incluye el anio previo solo como ventana del CF."""
    catalogo = defaultdict(dict)
    problemas = []
    for f in sorted(CRUDOS.iterdir()):
        if not f.is_file():
            continue
        m = RE_SBS.fullmatch(f.name)
        if not m:
            continue
        reporte, anomes = m.group(1).upper(), m.group(2)
        if reporte not in SBS_REPORTES:
            continue
        y, mo = int(anomes[:4]), int(anomes[4:])
        if not 1 <= mo <= 12:
            problemas.append({"fecha":"", "banco_id":"", "tipo":"NOMBRE_ARCHIVO", "detalle":f.name})
            continue
        if solo_anio and y not in (solo_anio, solo_anio - 1):
            continue
        fecha = pd.Timestamp(y, mo, 1) + pd.offsets.MonthEnd(0)
        if not (pd.Timestamp(FECHA_INICIO_LECTURA) <= fecha <= pd.Timestamp(FECHA_CORTE)):
            continue
        if reporte in catalogo[fecha]:
            raise ValueError(f"Dos archivos para {reporte} {anomes}: {f.name}")
        catalogo[fecha][reporte] = f
    return catalogo, problemas


def leer_libro(f: Path) -> dict[str, list[list[Any]]]:
    """Lee XLS real con xlrd y XLSX con openpyxl; NUNCA modifica el original."""
    sig = f.open("rb").read(8)
    if sig.startswith(b"PK\x03\x04"):
        wb = load_workbook(f, read_only=True, data_only=True)
        try:
            hojas = {}
            for ws in wb.worksheets:
                ws.reset_dimensions()  # evita hojas con dimensiones declaradas incorrectas
                hojas[ws.title] = [list(r) for r in ws.iter_rows(values_only=True)]
            return hojas
        finally:
            wb.close()
    if sig == bytes.fromhex("D0CF11E0A1B11AE1"):
        try:
            import xlrd
        except ImportError as exc:
            raise RuntimeError("Los XLS de la SBS requieren: python -m pip install xlrd==2.0.2") from exc
        wb = xlrd.open_workbook(str(f), on_demand=True)
        try:
            hojas = {}
            for sh in wb.sheets():
                hojas[sh.name] = [sh.row_values(r) for r in range(sh.nrows)]
            return hojas
        finally:
            wb.release_resources()
    raise ValueError(f"El archivo no tiene firma XLS/XLSX valida: {f.name}")


def cel(m: list[list[Any]], fila: int, columna: int) -> Any:
    if fila < 0 or columna < 0 or fila >= len(m) or columna >= len(m[fila]):
        return None
    return m[fila][columna]


def encontrar_fila(m: list[list[Any]], nombre: str, exacto: bool = True, desde: int = 0) -> int | None:
    objetivo = norm(nombre)
    for i in range(desde, len(m)):
        texto = norm(cel(m, i, 0))
        if (texto == objetivo) if exacto else texto.startswith(objetivo):
            return i
    return None


def primer_renglones(m: list[list[Any]], conceptos: list[str], desde: int = 0) -> int | None:
    for label in conceptos:
        r = encontrar_fila(m, label, desde=desde)
        if r is not None:
            return r
    return None


def encabezados_balance(m: list[list[Any]], audit: list[dict] | None = None,
                       archivo: str = "", hoja: str = "") -> dict[str,tuple[int,str]]:
    """Recupera bancos aun con rotulos repetidos; NO elige entre cifras ambiguas.

    En algunos archivos historicos un banco figura dos veces en cabecera. Si
    ambas referencias conducen a la misma columna TOTAL, son una repeticion de
    rotulo y se consolidan. Si apuntan a columnas distintas con montos
    diferentes, se omite SOLO ese banco/hoja y se registra la incidencia para
    revision manual. Los otros bancos/meses se conservan intactos.
    """
    candidatos = []
    for r in range(min(15, len(m))):
        ids = [banca_id(v) for v in m[r] if isinstance(v, str)]
        ids = [x for x in ids if x]
        if len(set(ids)) >= 6:
            # Priorizar diversidad de bancos y presencia de columna TOTAL.
            n_total = sum(norm(x) == "total" for x in (m[r+1] if r+1 < len(m) else []))
            candidatos.append((len(set(ids)), n_total, r))
    if not candidatos:
        raise ValueError("No se identifico el encabezado de bancos en B-2201")
    _, _, fila = max(candidatos)
    encontrados: dict[str, list[tuple[int, str]]] = defaultdict(list)
    for c, valor in enumerate(m[fila]):
        bid = banca_id(valor) if isinstance(valor, str) else None
        if not bid:
            continue
        total = None
        for j in range(c + 1, min(c + 5, len(m[fila+1]) if fila+1 < len(m) else c+5)):
            if norm(cel(m, fila + 1, j)) == "total":
                total = j
                break
        if total is not None:
            encontrados[bid].append((total, str(valor).strip().replace("\n", " ")))

    def registrar(banco: str, problema: str, detalle: str) -> None:
        if audit is not None:
            audit.append({"archivo": archivo, "reporte": "B-2201", "banco_id": banco,
                          "variable": "cabecera", "problema": problema,
                          "ubicacion": f"hoja {hoja} / fila {fila+1}", "valor": detalle})

    salida: dict[str, tuple[int, str]] = {}
    for bid, candidatos_banco in encontrados.items():
        # Mismo TOTAL significa la misma cuenta aunque el titulo se repita.
        distintos: dict[int, str] = {}
        for col, nombre in candidatos_banco:
            distintos.setdefault(col, nombre)
        if len(distintos) == 1:
            col, nombre = next(iter(distintos.items()))
            salida[bid] = (col, nombre)
            if len(candidatos_banco) > 1:
                registrar(bid, "CABECERA_REPETIDA_MISMO_TOTAL", f"columna={col_letra(col+1)}")
            continue

        # En una duplicacion real puede haber dos columnas con datos identicos.
        # Solo las equiparamos si se contrastaron varias cuentas numericas.
        columnas = list(distintos)
        muestras: dict[int, dict[int, float]] = {}
        for col in columnas:
            muestras[col] = {r: num(cel(m, r, col)) for r in range(fila + 2, len(m))
                            if math.isfinite(num(cel(m, r, col)))}
        representativa = max(columnas, key=lambda col: len(muestras[col]))
        misma_cifra = True
        for col in columnas:
            if col == representativa:
                continue
            inter = set(muestras[representativa]).intersection(muestras[col])
            if len(inter) < 10 or any(not math.isclose(muestras[representativa][r], muestras[col][r],
                                                   rel_tol=1e-9, abs_tol=0.01) for r in inter):
                misma_cifra = False
                break
        if misma_cifra:
            salida[bid] = (representativa, distintos[representativa])
            registrar(bid, "CABECERA_REPETIDA_MISMOS_IMPORTES",
                      f"columnas={','.join(col_letra(c+1) for c in columnas)}; verificado contra >=10 celdas")
        else:
            # Nunca descartar TODO el mes por un banco sin correspondencia clara.
            # Tampoco elegir silenciosamente una de las dos columnas.
            registrar(bid, "CABECERA_AMBIGUA_BANCO_OMITIDO",
                      f"columnas={','.join(col_letra(c+1) for c in columnas)}; cifras distintas")

    if len(salida) < 4:
        raise ValueError(f"Muy pocas columnas TOTAL sin ambiguedad ({len(salida)}) en B-2201")
    return salida


def extraer_balance(path: Path, audit: list[dict], guion_cero: bool = True) -> dict[str,dict]:
    hojas = leer_libro(path)
    if "1" not in hojas or "2" not in hojas:
        raise ValueError("B-2201 no contiene hojas '1' (balance) y '2' (resultados)")
    bal,pyg=hojas["1"],hojas["2"]
    bancos=encabezados_balance(bal,audit,path.name,"1")
    bancos_gastos=encabezados_balance(pyg,audit,path.name,"2")
    pasivo_ini = encontrar_fila(bal, "OBLIGACIONES CON EL PUBLICO")
    if pasivo_ini is None: raise ValueError("No aparece la cuenta Obligaciones con el publico")
    campos = {
        "activos_totales": (bal, encontrar_fila(bal, "TOTAL ACTIVO"), "1"),
        "oblig_publico": (bal, encontrar_fila(bal, "OBLIGACIONES CON EL PUBLICO",desde=pasivo_ini), "1"),
        "otras_oblig_publico": (bal, encontrar_fila(bal,"Otras Obligaciones", desde=pasivo_ini), "1"),
        "depositos_sistema": (bal, primer_renglones(bal,["DEPOSITOS DEL SISTEMA FINANCIERO Y ORGANISMOS INTERNACIONALES", "DEPOSITOS DEL SISTEMA FINANCIERO Y ORGANISMOS FINANCIEROS INTERNACIONALES"],desde=pasivo_ini),"1"),
        "adeudados": (bal, encontrar_fila(bal,"ADEUDOS Y OBLIGACIONES FINANCIERAS",desde=pasivo_ini), "1"),
        "deuda_subordinada": (bal, primer_renglones(bal,["OBLIGACIONES EN CIRCULACION SUBORDINADAS 1","OBLIGACIONES EN CIRCULACION SUBORDINADAS"],desde=pasivo_ini),"1"),
        "pasivos_totales": (bal, encontrar_fila(bal,"TOTAL PASIVO",desde=pasivo_ini), "1"),
        "gasto_publico_acum": (pyg, encontrar_fila(pyg,"Obligaciones con el Publico"), "2"),
        "gasto_depsistema_acum": (pyg, primer_renglones(pyg,["Depositos del Sistema Financiero y Organismos Financieros Internacionales","Depositos del Sistema Financiero y Organismos Internacionales"]),"2"),
        "gasto_adeudados_acum": (pyg, encontrar_fila(pyg,"Adeudos y Obligaciones Financieras"),"2"),
        "gasto_total_acum": (pyg, encontrar_fila(pyg,"GASTOS FINANCIEROS"),"2"),
    }
    faltan=[k for k,(_,r,_) in campos.items() if r is None]
    if faltan:
        audit.append({"archivo":path.name,"reporte":"B-2201","banco_id":"","variable":",".join(faltan),"problema":"ETIQUETA_NO_IDENTIFICADA","ubicacion":"","valor":""})
    salida={}
    fila_activo=campos["activos_totales"][1]
    for bid,(col,nombre) in bancos.items():
        # Evitar mezclar BCP sin sucursales y con sucursales.
        fila={"banco_id":bid,"banco_nombre":nombre}
        # El guion solo se lee como cero si la entidad reporta su balance ese mes.
        activo=num(cel(bal,fila_activo,col)) if fila_activo is not None else np.nan
        banco_reporta=math.isfinite(activo) and activo>0
        for variable,(matriz,ri,hoja) in campos.items():
            ci = col
            if hoja=="2":
                if bid not in bancos_gastos:
                    ci=-1
                else: ci = bancos_gastos[bid][0]
            if ri is not None and ci>=0:
                v,fue_guion=num_sbs(cel(matriz,ri,ci),variable,guion_cero and banco_reporta)
            else:
                v,fue_guion=np.nan,False
            fila[variable]=v
            audit.append({"archivo":path.name,"reporte":"B-2201","banco_id":bid,"variable":variable,
                          "problema":"GUION_SBS_CERO" if fue_guion else ("" if math.isfinite(v) else "CELDA_SIN_DATO"),
                          "ubicacion": f"{hoja}!{col_letra(ci+1)}{ri+1}" if ri is not None and ci>=0 else "",
                          "valor":v if math.isfinite(v) else ""})
        op,otras,df=fila["oblig_publico"],fila["otras_oblig_publico"],fila["depositos_sistema"]
        fila["depositos"] = op-otras+df if all(math.isfinite(x) for x in (op,otras,df)) else np.nan
        salida[bid]=fila
    return salida


def col_letra(i: int) -> str:
    s=""
    while i>0:
        i,a=divmod(i-1,26)
        s=chr(65+a)+s
    return s


def encabezado_patrimonio_b2370(m: list[list[Any]]) -> tuple[int, str]:
    """Identifica el total SBS B-2370 por la estructura REAL del encabezado.

    EVIDENCIA SBS (archivos originales verificados):
    - 2018-12 y 2022-12: B/C/D=Nivel 1/2/3 y E=TOTAL.
    - 2023-01 y 2023-05: B/C=Nivel 1/2 y D=TOTAL.
      El total NO esta en E: los libros solo contienen cuatro columnas.
    - 2023-06 y 2025-12: E=Patrimonio Efectivo Total.

    Se busca la disposicion por sus etiquetas, NO por fecha supuesta.
    La variante de dos niveles se contrasta adicionalmente con B + C = D.
    Si no se reconoce la estructura, se detiene ese reporte; no se
    sustituyen valores por cero, patrimonio contable ni interpolaciones.
    """
    limite = min(20, len(m))

    def titulo_padre(r: int) -> bool:
        return any(norm(cel(m,k,1)) == "patrimonio efectivo"
                   for k in range(max(0,r-3),r))

    def fila_entidad(r: int) -> bool:
        return any(norm(cel(m,k,0)) == "entidad"
                   for k in range(r,min(len(m),r+4)))

    # Disposicion nueva: E contiene literalmente el total.
    for r in range(limite):
        if (norm(cel(m,r,4)) == "patrimonio efectivo total" and
                any(norm(cel(m,k,0)) == "entidad"
                    for k in range(max(0,r-2),min(len(m),r+4)))):
            return 4, "TOTAL_EXPLICITO_E"

    # Disposicion antigua: NIVEL 1, NIVEL 2, NIVEL 3, TOTAL (columna E).
    for r in range(1,limite):
        if (norm(cel(m,r,1)) == "nivel 1" and
            norm(cel(m,r,2)) == "nivel 2" and
            norm(cel(m,r,3)) == "nivel 3" and
            norm(cel(m,r,4)) == "total" and
            titulo_padre(r) and fila_entidad(r)):
            return 4, "TOTAL_TRES_NIVELES_E"

    # Transicion de enero-mayo 2023: NIVEL 1, NIVEL 2, TOTAL (columna D).
    for r in range(1,limite):
        if not (norm(cel(m,r,1)) == "nivel 1" and
                norm(cel(m,r,2)) == "nivel 2" and
                norm(cel(m,r,3)) == "total" and
                titulo_padre(r) and fila_entidad(r)):
            continue
        comprobados = 0
        for i in range(r+1,len(m)):
            if not isinstance(cel(m,i,0),str) or not banca_id(cel(m,i,0)):
                continue
            nivel_1, nivel_2, total = (num(cel(m,i,k)) for k in (1,2,3))
            if not all(math.isfinite(z) for z in (nivel_1,nivel_2,total)):
                continue
            if abs((nivel_1+nivel_2)-total) > max(0.02,abs(total)*0.000001):
                raise ValueError(
                    "B-2370: estructura de dos niveles detectada, "
                    f"pero NIVEL1+NIVEL2 no concilia con TOTAL en fila {i+1}")
            comprobados += 1
            if comprobados >= 3:
                return 3, "TOTAL_DOS_NIVELES_D"
        raise ValueError(
            "B-2370: estructura de dos niveles detectada, "
            "pero sin tres bancos con montos conciliables; revisar original")

    raise ValueError(
        "B-2370: formato de patrimonio efectivo no reconocido. "
        "Se verifican TOTAL_E de tres niveles, TOTAL_D de dos niveles "
        "y Patrimonio Efectivo Total_E; no se imputan valores")


def detectar_columnas_b2402(m: list[list[Any]]) -> dict[str, Any]:
    """Ubica APR y ratio por sus TITULOS; no por la fecha ni una columna fija.

    Formatos contrastados con crudos SBS auténticos del 2019-07, 2021-04 y
    2023-06. Para esquemas históricos con solo requerimiento total,
    la fórmula publicada (p. ej. /10%) permite derivar APR; si falta
    el requerimiento de crédito no se interpreta el parcial como total.
    """
    cab = range(min(12, len(m)))
    ancho = max((len(fila) for fila in m[:12]), default=0)
    ratio_candidatos = []
    for c in range(1, ancho):
        for r in cab:
            etiqueta = norm(cel(m, r, c))
            if etiqueta == "ratio de capital" or etiqueta.startswith("ratio de capital global"):
                ratio_candidatos.append(c)
                break
    ratio_candidatos = sorted(set(ratio_candidatos))
    if len(ratio_candidatos) != 1:
        raise ValueError(f"B-2402: columna de ratio ambigua: {ratio_candidatos}")
    ratio_col = ratio_candidatos[0]

    # Disposición nueva: el rótulo APR aparece explícito sobre la columna I.
    apr_exp = []
    for c in range(1, ratio_col):
        if any(norm(cel(m, r, c)).startswith(
                "activos y contingentes ponderados por riesgo totales") for r in cab):
            apr_exp.append(c)
    if len(apr_exp) == 1:
        return {"apr_col": apr_exp[0], "ratio_col": ratio_col,
                "formula": None, "formato": "APR_EXPLICITO"}
    if len(apr_exp) > 1:
        raise ValueError(f"B-2402: APR explícito ambiguo: {apr_exp}")

    # Disposición intermedia (ej. 2021-04): F encabeza el bloque APR
    # (crédito, mercado, operacional); I está rotulada TOTAL.
    bloques_apr = [(r, c) for r in cab for c in range(1, ratio_col)
                   if norm(cel(m, r, c)).startswith(
                       "activos y contingentes ponderados por riesgo")]
    for r0, c0 in bloques_apr:
        total_cols = [c for c in range(c0 + 1, ratio_col)
                      if any(norm(cel(m, rr, c)) == "total"
                             for rr in range(r0 + 1, min(r0 + 4, len(m))))]
        if len(total_cols) == 1:
            return {"apr_col": total_cols[0], "ratio_col": ratio_col,
                    "formula": None, "formato": "APR_TOTAL_BLOQUE"}
        if len(total_cols) > 1:
            raise ValueError(f"B-2402: múltiples totales del bloque APR: {total_cols}")

    # Disposición antigua (ej. 2019-07): E = TOTAL requerimiento,
    # F = PATRIMONIO EFECTIVO TOTAL y G = RATIO DE CAPITAL GLOBAL.
    requerimientos = [(r, c) for r in cab for c in range(1, ratio_col)
                      if norm(cel(m, r, c)).startswith(
                          "requerimiento de patrimonio efectivo")]
    if not requerimientos:
        raise ValueError("B-2402: no se identifica bloque APR ni requerimiento")
    totales = [c for c in range(1, ratio_col)
               if any(norm(cel(m, r, c)) == "total" for r in cab)]
    pe_cols = [c for c in range(1, ratio_col)
               if any(norm(cel(m, r, c)) == "patrimonio efectivo total" for r in cab)]
    candidatos = [c for c in totales if c + 1 in pe_cols]
    if len(candidatos) != 1:
        raise ValueError(f"B-2402 antiguo: total del requerimiento ambiguo: {candidatos}")

    # Leer la tasa de la FÓRMULA del propio Excel. No asumir siempre 10%.
    tasas = set()
    for r in cab:
        x = cel(m, r, ratio_col)
        if isinstance(x, str):
            found = re.search(r"/\s*(\d{1,2}(?:[.,]\d+)?)\s*%", x)
            if found:
                tasas.add(float(found.group(1).replace(",", ".")))
    if len(tasas) != 1 or next(iter(tasas)) <= 0:
        raise ValueError(f"B-2402 antiguo: no se identifica tasa de APR en formula: {tasas}")
    return {"apr_col": candidatos[0], "ratio_col": ratio_col,
            "formula": next(iter(tasas)), "formato": "REQUERIMIENTO_A_APR"}


def tabla_por_empresa(path: Path, code: str, audit: list[dict]) -> dict[str,dict]:
    """SBS B-2310,2342,2370,2390,2401,2402; lectura dinamica de empresas."""
    hojas=leer_libro(path)
    out={}

    if code=="B-2401":
        # El historico 2020-06 trae dos hojas ("3" y "Hoja1"). Seleccionar
        # por CONTENIDO, no por el numero de hojas ni por la posicion.
        validas=[]
        for nombre_hoja,m in hojas.items():
            cab=next((r for r in range(min(15,len(m)))
                      if len({banca_id(v) for v in m[r]
                              if isinstance(v,str) and banca_id(v)})>=6),None)
            idx=primer_renglones(m,[
                "Utilidad Neta Anualizada / Patrimonio Promedio",
                "Rentabilidad sobre el Patrimonio (ROE)",
            ])
            if cab is None or idx is None:
                continue
            por_banco={}
            for col,v in enumerate(m[cab]):
                bid=banca_id(v) if isinstance(v,str) else None
                if bid and bid not in por_banco:
                    por_banco[bid]={"roe":num(cel(m,idx,col))}
            n_numericos=sum(math.isfinite(z["roe"]) for z in por_banco.values())
            if n_numericos>=6:
                validas.append((n_numericos,len(por_banco),nombre_hoja,idx,por_banco))
        if not validas:
            raise ValueError(f"B-2401: ninguna hoja contiene 6 bancos y una fila ROE valida; hojas={list(hojas)}")
        validas.sort(key=lambda t:(t[0],t[1],t[2]),reverse=True)
        n_validos,n_bancos,nombre_hoja,idx,out=validas[0]
        # Dos hojas completas con indicadores contradictorios NO se mezclan
        # ni se elige silenciosamente una como si fuera la otra.
        for _,_,otra_hoja,_,otros in validas[1:]:
            divergentes=[bid for bid in set(out)&set(otros)
                         if math.isfinite(out[bid]["roe"])
                         and math.isfinite(otros[bid]["roe"])
                         and not math.isclose(out[bid]["roe"],otros[bid]["roe"],rel_tol=1e-8,abs_tol=1e-6)]
            if divergentes:
                raise ValueError(f"B-2401: ROE discordante entre hojas {nombre_hoja} y {otra_hoja}; bancos={divergentes}")
        if len(hojas)>1:
            audit.append({
                "archivo":path.name,"reporte":code,"banco_id":"",
                "variable":"hoja_roe","problema":"HOJA_ROE_VERIFICADA",
                "ubicacion":f"{nombre_hoja}!A{idx+1}",
                "valor":f"hojas={list(hojas)}; bancos={n_bancos}; roe_numericos={n_validos}",
            })
        return out

    if len(hojas)!=1:
        raise ValueError(f"{code}: se esperaba una hoja, hay {list(hojas)}")
    nombre_hoja,m=next(iter(hojas.items()))

    # Reportes con una fila por banco, salvo 2401.
    hdr={
        "B-2310": ("total adeudos", 5, "adeudados_b2310"),
        "B-2342": ("total pasivo", 8, "pasivos_b2342"),
        "B-2370": ("patrimonio efectivo total", 4, "patrimonio_efectivo"),
        "B-2390": ("total gastos financieros", 13, "gasto_total_b2390"),
    }
    formato_pe = None
    formato_2402 = None
    if code == "B-2402":
        variable = "apr"
        formato_2402 = detectar_columnas_b2402(m)
        col = formato_2402["apr_col"]
        audit.append({"archivo":path.name,"reporte":code,"banco_id":"",
                      "variable":"columnas_b2402","problema":"FORMATO_B2402_IDENTIFICADO",
                      "ubicacion":f"{nombre_hoja}!{col_letra(col+1)} / {col_letra(formato_2402['ratio_col']+1)}",
                      "valor":str(formato_2402)})
    else:
        titulo,col,variable=hdr[code]
        if code == "B-2370":
            col,formato_pe = encabezado_patrimonio_b2370(m)
            audit.append({"archivo":path.name,"reporte":code,"banco_id":"",
                          "variable":"formato_patrimonio_efectivo","problema":"FORMATO_VERIFICADO",
                          "ubicacion":f"{nombre_hoja}!{col_letra(col+1)}", "valor":formato_pe})
        else:
            texto_cab=norm(" ".join(str(c) for row in m[:12] for c in row if isinstance(c,str)))
            if norm(titulo) not in texto_cab:
                raise ValueError(f"{code}: encabezado '{titulo}' no identificado; posible cambio de formato")
    for i,row in enumerate(m):
        if i < 7:continue
        ident = banca_id(row[0]) if row and isinstance(row[0],str) else None
        if not ident:continue
        v=num(cel(m,i,col))
        aviso_apr = None
        if code == "B-2402" and formato_2402["formula"] is not None:
            # En el formato antiguo el APR no está publicado como saldo.
            # Si falta el riesgo de crédito B, no reconstruir APR a partir
            # de un TOTAL que solo refleja mercado + operacional.
            if not math.isfinite(num(cel(m,i,1))):
                aviso_apr = ("El requerimiento por riesgo de credito (columna B) "
                             "no es numerico: APR no reconstruible sin fuente adicional")
                v = np.nan
            elif math.isfinite(v):
                v = 100.0 * v / formato_2402["formula"]
        if ident in out:
            audit.append({"archivo":path.name,"reporte":code,"banco_id":ident,"variable":variable,"problema":"EMPRESA_DUPLICADA","ubicacion":f"{nombre_hoja}!A{i+1}","valor":""})
            continue
        out[ident]={variable:v}
        if code=="B-2370":
            out[ident]["formato_pe_b2370"] = formato_pe
        if code=="B-2390":
            # B,C,E = cuotas (%) de intereses publico, depositos SF, adeudados.
            for nom,ix in (("pct_gasto_publico",1),("pct_gasto_depsistema",2),("pct_gasto_adeudados",4)):
                out[ident][nom]=num(cel(m,i,ix))
            out[ident]["gasto_cf_b2390"] = v * sum(out[ident][k] for k in ("pct_gasto_publico","pct_gasto_depsistema","pct_gasto_adeudados"))/100 if math.isfinite(v) and all(math.isfinite(out[ident][k]) for k in ("pct_gasto_publico","pct_gasto_depsistema","pct_gasto_adeudados")) else np.nan
        elif code=="B-2342":
            out[ident]["pct_dep_publico_b2342"]=num(cel(m,i,1))
            out[ident]["pct_dep_sistema_b2342"]=num(cel(m,i,2))
        elif code=="B-2402":
            out[ident]["ratio_capital_global"] = num(cel(m,i,formato_2402["ratio_col"]))
            out[ident]["formato_b2402"] = formato_2402["formato"]
            if aviso_apr:
                out[ident]["aviso_apr_b2402"] = aviso_apr
                audit.append({"archivo":path.name,"reporte":code,"banco_id":ident,
                              "variable":"apr","problema":"APR_NO_RECONSTRUIBLE",
                              "ubicacion":f"{nombre_hoja}!B{i+1}", "valor":aviso_apr})
        audit.append({"archivo":path.name,"reporte":code,"banco_id":ident,"variable":variable,
                      "problema":"" if math.isfinite(v) else "CELDA_SIN_DATO",
                      "ubicacion":f"{nombre_hoja}!{col_letra(col+1)}{i+1}","valor":v if math.isfinite(v) else ""})
    return out


def num_mes_csv(periodo: str) -> pd.Timestamp | None:
    p=norm(periodo)
    mat=re.fullmatch(r"(20\d{2}) (\d{1,2})",p)
    if mat:
        year,mes=map(int,mat.groups())
    else:
        mat=re.fullmatch(r"(ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic) ?(\d{2,4})",p)
        if not mat: return None
        mes=MONTHS[mat.group(1)];year=int(mat.group(2))
        if year<100:year+=2000 if year<70 else 1900
    if 1<=mes<=12:
        return pd.Timestamp(year,mes,1)+pd.offsets.MonthEnd(0)
    return None


def bcrp_serie(path: Path) -> dict[pd.Timestamp,float]:
    content=path.read_bytes()
    try: data=content.decode("utf-8-sig")
    except UnicodeDecodeError: data=content.decode("latin-1")
    raw=html.unescape(re.sub(r"<br\s*/?>","\n",data,flags=re.I))
    rows={}
    mes_pat=r"(?:20\d{2}[-/]\d{1,2}|(?:ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)\.?\s*[- /.]?\s*\d{2,4})"
    pat=re.compile(rf"^\s*[\"']?({mes_pat})[\"']?\s*[,;]\s*[\"']?([^,;\"'\s]+)",re.I)
    for linea in raw.splitlines():
        z=pat.match(linea)
        if not z:continue
        fecha=num_mes_csv(z.group(1))
        if fecha is None:continue
        if fecha in rows:raise ValueError(f"Periodo BCRP duplicado en {path.name}: {fecha}")
        rows[fecha]=num(z.group(2))
    if not rows:
        raise ValueError(f"No se pudo interpretar CSV BCRP original: {path.name}")
    return rows


def archivo_bcrp(codigo: str) -> Path:
    """Crudo BCRP del periodo declarado (mismo nombre que crea el script 01)."""
    inicio = FECHA_INICIO[:7].replace("-", "")
    fin = FECHA_CORTE[:7].replace("-", "")
    return CRUDOS/f"bcrp_original_{codigo}_{inicio}_{fin}_{CODIGO_MATRICULA}.csv"


def leer_bcrp() -> pd.DataFrame:
    series={}
    for codigo,variable in BCRP_SERIES.items():
        path=archivo_bcrp(codigo)
        if not path.is_file():
            raise FileNotFoundError(f"Falta BCRP original {path.name}: ejecuta primero codigo/01_extraccion_api.py")
        vals=bcrp_serie(path)
        series[variable]=pd.Series(vals,dtype="float64")
    df=pd.DataFrame(series).rename_axis("fecha").reset_index()
    if df["fecha"].duplicated().any():raise ValueError("Fechas duplicadas BCRP")
    return df


def discrepancia(a: Any,b: Any, tol:float=TOLERANCIA_CONCILIACION) -> bool:
    if not all(math.isfinite(x) for x in (num(a),num(b))): return False
    a,b=float(a),float(b)
    return abs(a-b) > max(1.0, tol*max(abs(a),abs(b)))


def construir_flujos(df:pd.DataFrame,incidencias:list[dict]) -> pd.DataFrame:
    df=df.sort_values(["banco_id","fecha"]).copy()
    gastos=("gasto_publico_acum","gasto_depsistema_acum","gasto_adeudados_acum")
    for k in gastos:df[k.replace("_acum","_mes")]=np.nan
    for bid,idx in df.groupby("banco_id",sort=False).groups.items():
        grupo=df.loc[idx].sort_values("fecha")
        for _,r in grupo.iterrows():
            fecha=r["fecha"]; es_enero=fecha.month==1
            prev_fecha=fecha-pd.offsets.MonthEnd(1)
            prev=grupo.loc[grupo["fecha"].eq(prev_fecha)]
            for k in gastos:
                actual=r[k]
                if not math.isfinite(num(actual)):continue
                if es_enero: flu=float(actual)
                elif len(prev)==1 and math.isfinite(num(prev.iloc[0][k])):
                    flu=float(actual)-float(prev.iloc[0][k])
                else:
                    incidencias.append({"fecha":fecha,"banco_id":bid,"tipo":"GASTO_MES_NO_CALCULABLE", "detalle":f"Falta acumulado del mes anterior para {k}"})
                    continue
                if flu < -max(1.0,abs(float(actual))*0.0001):
                    # CORRECCION v3: un flujo mensual negativo es una reversion contable
                    # publicada por la SBS (el acumulado bajo). Se registra como incidencia
                    # pero NO se borra: asi la suma de 12 meses coincide exactamente con
                    # los acumulados publicados y no se pierden 12 meses de costo de fondeo.
                    incidencias.append({"fecha":fecha,"banco_id":bid,"tipo":"FLUJO_NEGATIVO", "detalle":f"{k}: {flu:.3f}; reversion conservada"})
                df.loc[r.name,k.replace("_acum","_mes")]=flu
    cols_g=[k.replace("_acum","_mes") for k in gastos]
    df["gasto_fondeo_mes"]=df[cols_g].sum(axis=1,min_count=len(cols_g))
    df["base_fondeo"] = df[["oblig_publico","depositos_sistema","adeudados"]].sum(axis=1,min_count=3)
    # 12 meses MOVILES exactos y consecutivos, nunca sobre saltos de fechas.
    df["costo_fondeo"]=np.nan
    df["gastos_fondeo_12m"]=np.nan
    df["saldo_fondeo_prom_12m"]=np.nan
    for bid,indices in df.groupby("banco_id").groups.items():
        g=df.loc[indices].sort_values("fecha")
        for j in range(11,len(g)):
            ventana=g.iloc[j-11:j+1]
            periodos=ventana["fecha"].dt.to_period("M").astype(int).to_numpy()
            if len(periodos)!=12 or not np.all(np.diff(periodos)==1): continue
            if ventana["gasto_fondeo_mes"].notna().sum()!=12 or ventana["base_fondeo"].notna().sum()!=12:continue
            saldo_prom=ventana["base_fondeo"].mean()
            gasto_12=ventana["gasto_fondeo_mes"].sum()
            if saldo_prom <=0 or gasto_12<0:continue
            idx=g.index[j]
            df.loc[idx,"gastos_fondeo_12m"]=gasto_12
            df.loc[idx,"saldo_fondeo_prom_12m"]=saldo_prom
            df.loc[idx,"costo_fondeo"]=100*gasto_12/saldo_prom
    return df


def marcar_atipicos(df:pd.DataFrame) -> pd.DataFrame:
    """Hampel/MAD bancario. No winsoriza ni elimina originales."""
    res=df.copy()
    res["flag_atipico"]=False
    res["detalle_atipico"]=""
    # Los ratios son comparables en el tiempo por banco, no los montos entre bancos.
    indicadores=["costo_fondeo","dep_ratio","ade_ratio","ds_ratio","pe_ratio","roe"]
    for variable in indicadores:
        for bid,sub in res.groupby("banco_id"):
            x=sub[variable].dropna().astype(float)
            if len(x)<MIN_CASOS_ATIPICO:continue
            med=x.median();mad=np.median(np.abs(x-med))
            if mad>1e-9:
                z=0.67448975*(x-med)/mad
                bad=z.abs()>UMBRAL_MAD
            else:
                q1,q3=x.quantile([0.25,0.75]);iqr=q3-q1
                bad=((x<q1-3*iqr)|(x>q3+3*iqr)) if iqr>1e-9 else (x-med).abs().gt(1e-8)
            for idx in x.index[bad]:
                res.at[idx,"flag_atipico"]=True
                res.at[idx,"detalle_atipico"] += ("; " if res.at[idx,"detalle_atipico"] else "")+f"MAD/IQR:{variable}"
    # Variacion de saldos: saltos abruptos >= triple; se marcan, nunca cambian.
    for col in ("depositos","adeudados","deuda_subordinada","patrimonio_efectivo"):
        for bid,sub in res.groupby("banco_id"):
            g=sub.sort_values("fecha")
            anterior=g[col].shift(1)
            consecutivo=g["fecha"].dt.to_period("M").astype(int).diff().eq(1)
            fuerte=consecutivo & g[col].gt(0) & anterior.gt(0) & (((g[col]/anterior)>3.0) | ((anterior/g[col])>3.0))
            for idx in g.index[fuerte.fillna(False)]:
                res.at[idx,"flag_atipico"]=True
                res.at[idx,"detalle_atipico"] += ("; " if res.at[idx,"detalle_atipico"] else "")+f"SALTO_3X:{col}"
    return res




def winsorizar_variables(
    df: pd.DataFrame,
    limite_inferior: float = WINSOR_LIM_INF,
    limite_superior: float = WINSOR_LIM_SUP,
    min_casos: int = MIN_CASOS_WINSOR,
    modo: str = WINSOR_MODO,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Crea versiones winsorizadas (*_w) sin reemplazar valores originales.

    modo="global" (por defecto): percentiles 1-99 de TODA la muestra banco-mes,
    practica habitual en estudios de panel bancario. Con ~85 meses por banco,
    el modo "banco" recorta como maximo una observacion por cola y deja
    intactos los extremos entre entidades (p. ej. ROE muy negativo de un banco
    pequeno), que son justamente los que distorsionan una regresion de panel.
    modo="banco": comportamiento anterior, limites por entidad.

    Valores fuera de dominio (ratio de pasivo <0 o >100, costo de fondeo
    negativo) NO se recortan: quedan NA en *_w y se marcan en detalle, porque
    son probables errores de lectura y no extremos legitimos.

    Se winsorizan costo_fondeo, ROE y ratios financieros. Los montos absolutos
    (depositos, adeudados, deuda subordinada y patrimonio efectivo) NO se
    winsorizan porque sus diferencias reflejan, en gran parte, el tamaño real
    de cada entidad. TIPMN/TIPMEX tampoco se winsorizan: son series macro comunes.

    Los limites se calculan globalmente por defecto; --winsor-modo banco
    utiliza limites propios de cada entidad. Los montos NO se winsorizan.
    Los registros fuera de dominio quedan originales en la base maestra y NA
    en *_w, con alerta para verificacion individual.
    """
    if not (0 <= limite_inferior < limite_superior <= 1):
        raise ValueError("Los limites de winsorizacion deben cumplir 0 <= inferior < superior <= 1")
    if modo not in ("global", "banco"):
        raise ValueError("modo de winsorizacion debe ser 'global' o 'banco'")

    res = df.copy()
    res["flag_winsorizado"] = False
    res["detalle_winsorizado"] = ""
    resumen: list[dict[str, Any]] = []

    # Los ratios de fuentes del pasivo deben estar entre 0 y 100 para ser
    # candidatos. Patrimonio efectivo/activos se restringe solo a no negativo.
    dominios = {
        "costo_fondeo": lambda x: x >= 0,
        "roe": lambda x: np.isfinite(x),
        "dep_ratio": lambda x: x.between(0, 100),
        "ade_ratio": lambda x: x.between(0, 100),
        "ds_ratio": lambda x: x.between(0, 100),
        "pe_ratio": lambda x: x >= 0,
    }

    for variable in VARIABLES_WINSOR:
        if variable not in res.columns:
            continue
        nueva = f"{variable}_w"
        res[nueva] = pd.to_numeric(res[variable], errors="coerce")

        grupos = (res.groupby("banco_id", sort=False, dropna=False) if modo == "banco"
                  else [("TODOS", res)])
        for banco_id, sub in grupos:
            x = pd.to_numeric(sub[variable], errors="coerce")
            validos = x.notna() & np.isfinite(x)
            dominio = dominios.get(variable)
            if dominio is not None:
                try:
                    validos &= dominio(x).fillna(False)
                except AttributeError:
                    validos &= pd.Series(dominio(x), index=x.index).fillna(False)

            valores = x.loc[validos]
            n = int(valores.size)
            fuera = x.notna() & ~validos
            excluidos_dominio = int(fuera.sum())
            # Fuera de dominio: no es un extremo, es un dato invalido para el modelo.
            for idx in x.index[fuera]:
                res.at[idx, nueva] = np.nan
                res.at[idx, "flag_winsorizado"] = True
                detalle = res.at[idx, "detalle_winsorizado"]
                res.at[idx, "detalle_winsorizado"] = (detalle + "; " if detalle else "") + f"FUERA_DOMINIO:{variable}"

            if n < min_casos or valores.nunique() < 2:
                resumen.append({
                    "banco_id": banco_id,
                    "variable": variable,
                    "n_elegible": n,
                    "limite_inferior": np.nan,
                    "limite_superior": np.nan,
                    "n_inferior": 0,
                    "n_superior": 0,
                    "n_excluido_dominio": excluidos_dominio,
                    "estado": "NO_APLICA_MUESTRA_INSUFICIENTE_O_SIN_VARIACION",
                })
                continue

            q_inf = float(valores.quantile(limite_inferior))
            q_sup = float(valores.quantile(limite_superior))
            if not (math.isfinite(q_inf) and math.isfinite(q_sup) and q_inf <= q_sup):
                resumen.append({
                    "banco_id": banco_id, "variable": variable, "n_elegible": n,
                    "limite_inferior": np.nan, "limite_superior": np.nan,
                    "n_inferior": 0, "n_superior": 0,
                    "n_excluido_dominio": excluidos_dominio,
                    "estado": "NO_APLICA_LIMITES_INVALIDOS",
                })
                continue

            recortados = valores.clip(lower=q_inf, upper=q_sup)
            cambiados = (recortados - valores).abs() > 1e-12
            res.loc[valores.index, nueva] = recortados

            indices_cambiados = valores.index[cambiados]
            for idx in indices_cambiados:
                res.at[idx, "flag_winsorizado"] = True
                detalle = res.at[idx, "detalle_winsorizado"]
                res.at[idx, "detalle_winsorizado"] = (detalle + "; " if detalle else "") + variable

            n_inf = int((valores < q_inf).sum())
            n_sup = int((valores > q_sup).sum())
            resumen.append({
                "banco_id": banco_id,
                "variable": variable,
                "n_elegible": n,
                "limite_inferior": q_inf,
                "limite_superior": q_sup,
                "n_inferior": n_inf,
                "n_superior": n_sup,
                "n_excluido_dominio": excluidos_dominio,
                "estado": "WINSORIZADO" if (n_inf + n_sup) else "SIN_CAMBIOS",
                "modo": modo,
            })

    return res, pd.DataFrame(resumen)


def agregar_logaritmos(df: pd.DataFrame) -> pd.DataFrame:
    """ln(1+monto en miles de S/) para montos >= 0; negativos o NA -> NA.

    Los montos difieren por el TAMANO real de cada banco: no son atipicos que
    deban recortarse. El logaritmo comprime esa escala para la regresion.
    Se usa log1p porque deuda_subordinada y adeudados pueden ser cero.
    """
    res = df.copy()
    for col in ("depositos", "adeudados", "deuda_subordinada", "patrimonio_efectivo"):
        if col not in res:
            continue
        x = pd.to_numeric(res[col], errors="coerce")
        res[f"ln_{col}"] = np.where(x >= 0, np.log1p(x.clip(lower=0)), np.nan)
    return res


def atipicos_residuales(panel: pd.DataFrame, umbral_iqr: float = 3.0) -> pd.DataFrame:
    """Cuenta extremos que AUN quedan en las variables del modelo (*_w y ln_*).

    Regla de Tukey 'extremos' (Q1-3*IQR, Q3+3*IQR) sobre la muestra pooled.
    Es solo un diagnostico para reportar en la tesis; no modifica datos.
    """
    filas = []
    cols = [f"{v}_w" for v in VARIABLES_WINSOR] + [
        "ln_depositos", "ln_adeudados", "ln_deuda_subordinada", "ln_patrimonio_efectivo"]
    for col in cols:
        if col not in panel:
            continue
        x = pd.to_numeric(panel[col], errors="coerce").dropna()
        if len(x) < 4:
            continue
        q1, q3 = x.quantile([0.25, 0.75])
        iqr = q3 - q1
        n_ext = int(((x < q1 - umbral_iqr * iqr) | (x > q3 + umbral_iqr * iqr)).sum()) if iqr > 0 else 0
        filas.append({"variable": col, "n": len(x), "min": x.min(), "p1": x.quantile(.01),
                      "mediana": x.median(), "p99": x.quantile(.99), "max": x.max(),
                      "extremos_tukey_3iqr": n_ext})
    return pd.DataFrame(filas)


def embudo_observaciones(df: pd.DataFrame) -> pd.DataFrame:
    """Explica perdidas de la muestra PRINCIPAL, no de variables opcionales.

    ROE y TIPMEX se contabilizan por separado porque no condicionan
    la muestra principal del modelo de costo de fondeo.
    """
    orden = COLUMNAS_MUESTRA_PRINCIPAL + COLUMNAS_LOG_MODELO
    pasos = [("Filas banco-mes del panel", pd.Series(True, index=df.index))]
    acumulado = pd.Series(True, index=df.index)
    for col in orden:
        acumulado = acumulado & (df[col].notna() if col in df else False)
        pasos.append((f"+ {col} disponible", acumulado.copy()))
    salida = []
    anterior = len(df)
    for descripcion, mascara in pasos:
        n = int(mascara.sum())
        salida.append({"paso": descripcion, "observaciones": n,
                       "perdidas_en_paso": anterior - n})
        anterior = n
    for col in VARIABLES_SEIS + VARIABLES_BCRP:
        if col in df:
            salida.append({"paso": f"[aislado] faltantes en {col}",
                           "observaciones": int(df[col].isna().sum()),
                           "perdidas_en_paso": np.nan})
    return pd.DataFrame(salida)


def _mascara_muestra_principal(panel: pd.DataFrame) -> pd.Series:
    """Misma regla para los dos CSV, sin filtrar ROE, TIPMEX ni atipicos.

    Cada fila conservada tiene CF no faltante, los 4 montos y TIPMN.
    Las transformaciones logaritmicas de los montos deben ser posibles,
    porque son las cuatro regresoras del modelo (x>=0, sin datos inventados).
    """
    requeridas = COLUMNAS_MUESTRA_PRINCIPAL + COLUMNAS_LOG_MODELO
    faltan = sorted(set(requeridas) - set(panel.columns))
    if faltan:
        raise ValueError(f"No se puede seleccionar muestra principal: faltan {faltan}")
    return panel[requeridas].notna().all(axis=1)


def crear_panel_analisis_winsorizado(
    panel: pd.DataFrame,
    archivo: Path,
    resumen: pd.DataFrame,
    archivo_resumen: Path,
    logger: logging.Logger | None = None,
) -> Path:
    """Panel PARALELO con las mismas filas del oficial y columnas derivadas.

    No sobrescribe los originales y no filtra por ROE/TIPMEX ni costo_fondeo_w.
    El script 04 restringe a casos con Y_w para robustez sobre muestra comun.
    """
    originales = COLUMNAS_CSV_OFICIAL + [
        "dep_ratio", "ade_ratio", "ds_ratio", "pe_ratio",
    ]
    winsor = [f"{v}_w" for v in VARIABLES_WINSOR if f"{v}_w" in panel.columns]
    winsor += [c for c in COLUMNAS_LOG_MODELO if c in panel.columns]
    controles = ["flag_atipico", "detalle_atipico", "flag_winsorizado",
                 "detalle_winsorizado", "completo_modelo", "cobertura_b2201"]
    columnas = list(dict.fromkeys(c for c in originales + winsor + controles
                                   if c in panel.columns))
    p = panel.copy()
    p["fecha"] = pd.to_datetime(p["fecha"], errors="raise")
    muestra = p.loc[_mascara_muestra_principal(p), columnas].copy()
    muestra = muestra.sort_values(["banco_id", "fecha"], kind="stable")
    if muestra.duplicated(["banco_id", "fecha"]).any():
        raise ValueError("Duplicados banco-fecha en panel de analisis paralelo")
    muestra["fecha"] = muestra["fecha"].dt.strftime("%Y-%m-%d")
    archivo.parent.mkdir(parents=True, exist_ok=True)
    temporal = archivo.with_suffix(archivo.suffix + ".tmp")
    try:
        muestra.to_csv(temporal, index=False, encoding="utf-8-sig", float_format="%.9f")
        os.replace(temporal, archivo)
    finally:
        if temporal.exists():
            temporal.unlink()
    archivo_resumen.parent.mkdir(parents=True, exist_ok=True)
    resumen.to_csv(archivo_resumen, index=False, encoding="utf-8-sig", float_format="%.9f")
    if logger:
        logger.info("PANEL_PARALELO | archivo=%s | filas=%d | columnas=%d | "
                    "registros_modificados=%d", archivo.name, len(muestra),
                    len(muestra.columns), int(muestra["flag_winsorizado"].sum()))
    return archivo


def _valor_excel(valor):
    """Mantiene tipos y faltantes: NaN -> celda vacia, sin imputar."""
    if isinstance(valor, (np.bool_, bool)):
        return bool(valor)
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime()
    if isinstance(valor, np.integer):
        return int(valor)
    if isinstance(valor, np.floating):
        return float(valor)
    return valor


def promedio_movil_bcrp(ventana: int = VENTANA_TASAS) -> pd.DataFrame:
    """Promedio de 12 meses de TIPMN y TIPMEX desde el crudo BCRP completo.

    El costo de fondeo es un promedio movil de 12 meses; para que la tasa de
    mercado mida el mismo periodo se usa su promedio de los ultimos 12 meses.
    Solo se calcula con 12 meses consecutivos observados (no se imputa): por
    eso enero-noviembre de 2017 quedan vacios en tipmn_ma12/tipmex_ma12.
    """
    tasas = leer_bcrp().sort_values("fecha")
    tasas["fecha"] = pd.to_datetime(tasas["fecha"]) + pd.offsets.MonthEnd(0)
    tasas = tasas.set_index("fecha").asfreq("ME")
    salida = pd.DataFrame(index=tasas.index)
    for col, nuevo in (("tipmn_bcrp", "tipmn_ma12"), ("tipmex_bcrp", "tipmex_ma12")):
        salida[nuevo] = tasas[col].rolling(ventana, min_periods=ventana).mean()
    return salida.rename_axis("fecha").reset_index()


def agregar_variables_modelo(oficial: pd.DataFrame) -> pd.DataFrame:
    """Agrega las 12 variables calculadas al final de la base oficial.

    Usa solo los valores originales de la misma fila y la serie BCRP; no
    cambia las 13 columnas originales ni elimina filas.
    """
    d = oficial.copy()
    montos = ["depositos", "adeudados", "deuda_subordinada"]
    d["fondeo_total"] = d[montos].sum(axis=1, min_count=3)
    base = d["fondeo_total"].where(d["fondeo_total"] > 0)
    d["part_depositos"] = 100 * d["depositos"] / base
    d["part_adeudados"] = 100 * d["adeudados"] / base
    d["part_subordinada"] = 100 * d["deuda_subordinada"] / base
    d["capitalizacion"] = 100 * d["patrimonio_efectivo"] / base
    d["tamano"] = np.log(base)
    for col in ("depositos", "adeudados", "deuda_subordinada", "patrimonio_efectivo"):
        x = pd.to_numeric(d[col], errors="coerce")
        d[f"ln_{col}"] = np.where(x >= 0, np.log1p(x.clip(lower=0)), np.nan)
    fechas = pd.to_datetime(d["fecha"]) + pd.offsets.MonthEnd(0)
    tasas = promedio_movil_bcrp().set_index("fecha")
    for col in ("tipmn_ma12", "tipmex_ma12"):
        d[col] = fechas.map(tasas[col]).to_numpy()
    suma = d[["part_depositos", "part_adeudados", "part_subordinada"]].sum(axis=1)
    if not np.allclose(suma.dropna(), 100, atol=1e-6):
        raise AssertionError("Las participaciones del fondeo no suman 100%")
    return d[COLUMNAS_CSV_OFICIAL + COLUMNAS_DERIVADAS_MODELO]


def crear_csv_oficial(panel: pd.DataFrame, archivo: Path,
                      logger: logging.Logger | None = None) -> Path:
    """Exporta 13 columnas originales (sin winsorizacion) + 12 calculadas.

    Selecciona casos de la regresion principal sin requerir ROE ni TIPMEX.
    Para cotejar con la fuente: CF es una proxy calculada (NO una celda SBS),
    mientras que montos, ROE y tasas BCRP derivan de los originales declarados.
    """
    faltan = sorted(set(COLUMNAS_CSV_OFICIAL + ["completo_modelo"]) - set(panel.columns))
    if faltan:
        raise ValueError(f"Faltan columnas en base maestra: {faltan}")
    p = panel.copy()
    p["fecha"] = pd.to_datetime(p["fecha"], errors="raise")
    if p.duplicated(["banco_id", "fecha"]).any():
        raise ValueError("Duplicados banco-fecha en base maestra")
    completos5 = p[VARIABLES_PRINCIPALES].notna().all(axis=1)
    guardado = p["completo_modelo"].astype(str).str.strip().str.lower().map(
        {"true": True, "false": False, "1": True, "0": False})
    if guardado.isna().any() or not (completos5.to_numpy() == guardado.to_numpy()).all():
        raise ValueError("La marca completo_modelo no coincide con las cinco variables")
    validos = _mascara_muestra_principal(p)
    oficial = p.loc[validos, COLUMNAS_CSV_OFICIAL].sort_values(
        ["banco_id", "fecha"], kind="stable").copy()
    if oficial[COLUMNAS_MUESTRA_PRINCIPAL].isna().any().any():
        raise AssertionError("La base oficial no puede omitir Y, X1-X4 o TIPMN")
    if oficial.duplicated(["banco_id", "fecha"]).any():
        raise AssertionError("Duplicados banco-fecha en base oficial")
    oficial = agregar_variables_modelo(oficial)
    oficial["fecha"] = oficial["fecha"].dt.strftime("%Y-%m-%d")
    archivo.parent.mkdir(parents=True, exist_ok=True)
    temporal = archivo.with_suffix(archivo.suffix + ".tmp")
    try:
        oficial.to_csv(temporal, index=False, encoding="utf-8-sig", float_format="%.9f")
        os.replace(temporal, archivo)
    finally:
        if temporal.exists():
            temporal.unlink()
    if logger:
        logger.info("CSV_OFICIAL_25 | %s | filas=%d | columnas=%d | "
                    "filas_maestras=%d | ROE_faltante=%d | TIPMEX_faltante=%d",
                    archivo.name, len(oficial), len(oficial.columns), len(p),
                    int(oficial["roe"].isna().sum()), int(oficial["tipmex_bcrp"].isna().sum()))
    return archivo


def crear_vista_excel(panel:pd.DataFrame, incidencias:pd.DataFrame, archivo:Path, logger:logging.Logger|None=None) -> Path:
    """Copia VISUAL del CSV procesado. No altera ninguna observacion ni los crudos.

    Organiza el panel existente en hojas legibles, con filtros y unidades.
    Las estadisticas del resumen corresponden a la base leida, nunca se
    emplean para sustituir observaciones de la base oficial.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.workbook.properties import CalcProperties

    p=panel.copy()
    flags=("flag_atipico","flag_winsorizado","completo_modelo","fuente_bcrp_disponible","cobertura_4_saldos")
    for key in flags:
        if key in p:
            p[key]=p[key].apply(lambda x: (str(x).strip().lower() in ("true","1","si","sí")) if pd.notna(x) else False).astype(bool)
    if "fecha" not in p or "banco_id" not in p:
        raise ValueError("No se encuentran identificadores banco_id y fecha en datos_procesados")
    p["fecha"]=pd.to_datetime(p["fecha"],errors="raise")
    if p.duplicated(["banco_id","fecha"]).any():
        raise ValueError("El panel contiene duplicados banco-mes; no generar la vista")
    columnas_clave=("costo_fondeo","depositos","adeudados","deuda_subordinada","patrimonio_efectivo")
    faltantes=[c for c in columnas_clave if c not in p]
    if faltantes:raise ValueError(f"Faltan variables principales: {faltantes}")
    completo=p[list(columnas_clave)].notna().all(axis=1)
    if "completo_modelo" in p and not (completo.to_numpy()==p["completo_modelo"].to_numpy()).all():
        raise ValueError("completo_modelo no coincide con las cinco variables; revisar integridad del CSV")
    p=p.sort_values(["banco_id","fecha"],kind="stable").reset_index(drop=True)
    orden_vista=COLUMNAS_BASE_ANALISIS
    completos6=p[VARIABLES_SEIS].notna().all(axis=1)
    completos_oficial=_mascara_muestra_principal(p)
    prim=p.loc[completos_oficial,orden_vista].copy()
    sub=p.loc[completos_oficial,COLUMNAS_CSV_OFICIAL].copy()
    outliers=p.loc[p["flag_atipico"],[c for c in ("banco_id","banco_nombre","fecha","detalle_atipico")+columnas_clave+("dep_ratio","ade_ratio","ds_ratio","pe_ratio","roe") if c in p]] if "flag_atipico" in p else p.iloc[0:0]
    if incidencias is None:
        incidencias=pd.DataFrame(columns=["fecha","banco_id","tipo","detalle"])
    else:
        incidencias=incidencias.copy()
        if "fecha" in incidencias:
            incidencias["fecha"]=pd.to_datetime(incidencias["fecha"],errors="coerce")

    navy="17324D"; navy2="284B63"; pale="E9F1F6"; green="087F70"; orange="FFF3DF"; red="FCE8E5"; white="FFFFFF"
    wb=Workbook()
    resumen=wb.active; resumen.title="Resumen"
    wb.calculation=CalcProperties(calcMode="auto",fullCalcOnLoad=True,forceFullCalc=True)
    resumen.sheet_view.showGridLines=False
    resumen.merge_cells("A1:H2")
    c=resumen["A1"]; c.value="PANEL BANCARIO PERUANO  |  FINANZAS I";c.font=Font(name="Aptos Display",size=19,bold=True,color=white);c.fill=PatternFill("solid",fgColor=navy);c.alignment=Alignment(vertical="center",indent=1)
    for row in resumen["A1:H2"]:
        for cell in row:cell.fill=PatternFill("solid",fgColor=navy)
    resumen.merge_cells("A3:H3");resumen["A3"]="Vista de consulta generada automaticamente por 03_limpieza_datos.py";resumen["A3"].font=Font(color=navy2,italic=True,size=11)
    resumen.merge_cells("A5:H5");resumen["A5"]="CSV oficial: 13 columnas ORIGINALES (CF+4 X+TIPMN completos). ROE/TIPMEX opcionales; maestra conserva NA.";resumen["A5"].font=Font(color="794D00",size=10)
    resumen["A7"]="INDICADOR";resumen["B7"]="RESULTADO";resumen["C7"]="INTERPRETACION"
    for cell in resumen[7][:3]:cell.fill=PatternFill("solid",fgColor=navy2);cell.font=Font(bold=True,color=white)
    metrics=[
        ("Filas banco-mes",len(p),"Una observacion por banco y mes"),
        ("Bancos distintos",int(p["banco_id"].nunique()),"Panel bancario no necesariamente balanceado"),
        ("Meses distintos",int(p["fecha"].nunique()),"Periodo de estudio 2017-2025"),
        ("Casos completos (5 variables)",int(completo.sum()),"Referencia con las cinco variables financieras centrales"),
        ("Atipicos marcados",int(p["flag_atipico"].sum()) if "flag_atipico" in p else 0,"Valores originales conservados; panel winsorizado separado"),
        ("Incidencias registradas",len(incidencias),"Ver hoja Incidencias"),
        ("% de filas completas","=IFERROR(B11/B8,0)","Calculado con formula; formato porcentaje"),
        ("Filas con costo de fondeo",int(p["costo_fondeo"].notna().sum()),"No sustituir faltantes por cero"),
        ("Inicio de ventana",p["fecha"].min(),"Primera fecha observada en el panel"),
        ("Corte de ventana",p["fecha"].max(),"Ultima fecha observada en el panel"),
        ("Completos incluyendo ROE (6)",int(completos6.sum()),
         "ROE puede faltar: el CSV oficial requiere CF+cuatro X+TIPMN"),
    ]
    for r,(k,v,desc) in enumerate(metrics,start=8):
        resumen.cell(r,1,k);resumen.cell(r,2,_valor_excel(v));resumen.cell(r,3,desc)
        if r%2==0:
            for x in resumen[r][:3]:x.fill=PatternFill("solid",fgColor="F2F6F9")
        resumen.cell(r,1).font=Font(bold=True,color=navy)
        if r in (16,17):resumen.cell(r,2).number_format="yyyy-mm-dd"
    resumen["B14"].number_format="0.0%"
    resumen.merge_cells("A19:H19");resumen["A19"]="FUENTES Y UNIDADES";resumen["A19"].font=Font(bold=True,color=white);resumen["A19"].fill=PatternFill("solid",fgColor=navy2)
    for rr,desc in enumerate((
        "SBS: Balance B-2201, adeudados B-2310, pasivos B-2342, patrimonio B-2370, gastos B-2390, indicadores B-2401 y B-2402.",
        "BCRP: TIPMN (PN07816NM) y TIPMEX (PN07836NM); son tasas agregadas, no tasas individuales de cada banco.",
        "Montos monetarios: miles de soles. Tasas/ratios del CSV se almacenan en unidades porcentuales (por ejemplo, 3.5 = 3.5%).",
        "El costo de fondeo es una aproximacion contable movil de 12 meses, no el WACC bancario.",
        "Las 5 variables se consideran completas solo si no estan vacias; esto no sustituye la auditoria contable.",
        "Las tablas y figuras del articulo deben generarse posteriormente desde 04_analisis.py y guardarse en salidas/."
    ),start=20):
        resumen.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=8)
        resumen.cell(rr,1,desc).alignment=Alignment(wrap_text=True,vertical="center")
        resumen.row_dimensions[rr].height=25
    resumen.column_dimensions["A"].width=33
    resumen.column_dimensions["B"].width=23
    resumen.column_dimensions["C"].width=67
    for k in "DEFGH":resumen.column_dimensions[k].width=13
    resumen.freeze_panes="A8"

    percent_cols={"costo_fondeo","tipmn_bcrp","tipmex_bcrp","roe","dep_ratio","ade_ratio","ds_ratio","pe_ratio","ratio_capital_global","pct_gasto_publico","pct_gasto_depsistema","pct_gasto_adeudados","pct_dep_publico_b2342","pct_dep_sistema_b2342"}
    money_cols={"depositos","adeudados","deuda_subordinada","patrimonio_efectivo","activos_totales","oblig_publico","otras_oblig_publico","depositos_sistema","pasivos_totales","gasto_publico_acum","gasto_depsistema_acum","gasto_adeudados_acum","gasto_total_acum","adeudados_b2310","pasivos_b2342","gasto_total_b2390","gasto_cf_b2390","apr","gasto_publico_mes","gasto_depsistema_mes","gasto_adeudados_mes","gasto_fondeo_mes","base_fondeo","gastos_fondeo_12m","saldo_fondeo_prom_12m"}
    def hoja_dataframe(nombre, data, anchos=None):
        sh=wb.create_sheet(nombre)
        sh.sheet_view.showGridLines=False
        sh.freeze_panes="F2" if nombre in ("Vista principal","Muestra modelo","Panel completo") else "C2"
        cols=list(data.columns)
        sh.append(cols)
        for c in sh[1]:
            c.fill=PatternFill("solid",fgColor=navy);c.font=Font(name="Aptos",size=10,bold=True,color=white)
            c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        sh.row_dimensions[1].height=38
        for row in data.itertuples(index=False,name=None):
            sh.append([_valor_excel(v) for v in row])
        for idx,col in enumerate(cols,start=1):
            letter=get_column_letter(idx)
            if col=="fecha":width=15
            elif col=="banco_nombre":width=37
            elif col in ("detalle","detalle_atipico"):width=62
            elif col in ("fuente_sbs","fuente_bcrp","formato_pe_b2370","indicador_calidad","problema"):width=37
            elif col in ("banco_id","tipo","variable"):width=23
            elif col in ("anio","mes"):width=11
            elif col in money_cols or col in percent_cols:width=23
            else:width=20
            if anchos and col in anchos:width=anchos[col]
            sh.column_dimensions[letter].width=width
            if col=="fecha":formato="yyyy-mm-dd"
            elif col in percent_cols:formato='0.000"%";[Red](0.000"%");"-"'
            elif col in money_cols:formato='#,##0.00;[Red](#,##0.00);"-"'
            elif col in ("anio","mes"):formato="0"
            else:formato=None
            if formato:
                for cells in sh.iter_cols(min_col=idx,max_col=idx,min_row=2,max_row=sh.max_row):
                    for cell in cells:cell.number_format=formato
        # Filtros Excel y bandas, sin cambiar el contenido.
        sh.auto_filter.ref=sh.dimensions
        if nombre in ("Vista principal","Muestra modelo","Panel completo") and sh.max_row>1:
            fmap={c:i for i,c in enumerate(cols,start=1)}
            end=get_column_letter(len(cols))
            if "flag_atipico" in fmap:
                fl=get_column_letter(fmap["flag_atipico"])
                sh.conditional_formatting.add(f"A2:{end}{sh.max_row}",FormulaRule(formula=[f"${fl}2=TRUE"],fill=PatternFill("solid",fgColor=orange)))
            if "completo_modelo" in fmap:
                fl=get_column_letter(fmap["completo_modelo"])
                sh.conditional_formatting.add(f"A2:{end}{sh.max_row}",FormulaRule(formula=[f"${fl}2=FALSE"],fill=PatternFill("solid",fgColor=red)))
        return sh

    hoja_vista=hoja_dataframe("Vista principal",prim)
    hoja_modelo=hoja_dataframe("Muestra modelo",sub)
    # Resaltar faltantes sin convertirlos a cero. Esta es una marca visual.
    for sh in (hoja_vista,hoja_modelo):
        sh.freeze_panes="D2"
        columnas_hoja=[c.value for c in sh[1]]
        for col_name in VARIABLES_SEIS:
            if col_name not in columnas_hoja:continue
            idx=columnas_hoja.index(col_name)+1
            for cells in sh.iter_cols(min_col=idx,max_col=idx,min_row=2,max_row=sh.max_row):
                for cell in cells:
                    if cell.value is None:
                        cell.fill=PatternFill("solid",fgColor="FFF3DF")
    # Mostrar la tabla minima al abrir el Excel. La trazabilidad queda en otras hojas.
    hoja_dataframe("Panel completo",p)
    grupos=p.groupby("banco_id",dropna=False,as_index=False).agg(
        banco_nombre=("banco_nombre", lambda s: s.dropna().astype(str).iloc[-1] if len(s.dropna()) else "SIN NOMBRE"),
        meses=("fecha","nunique"),casos_completos=("completo_modelo","sum"),
        cf_disponible=("costo_fondeo",lambda x:x.notna().sum()),
        patrimonio_disponible=("patrimonio_efectivo",lambda x:x.notna().sum()),
        atipicos=("flag_atipico","sum"))
    grupos.columns=["Banco ID","Banco","Meses","Casos completos","CF disponible","PE disponible","Atipicos"]
    gr=hoja_dataframe("Cobertura por banco",grupos)
    gr.cell(1,8,"% completos")
    gr.column_dimensions["H"].width=18
    for rr in range(2,gr.max_row+1):
        gr.cell(rr,8,f"=IFERROR(D{rr}/C{rr},0)")
        gr.cell(rr,8).number_format="0.0%"
    gr.auto_filter.ref=f"A1:H{gr.max_row}"
    hoja_dataframe("Atipicos",outliers)
    hoja_dataframe("Incidencias",incidencias)
    dic=[
        ("banco_id","Identificador armonizado de entidad","Texto","Mensual","SBS","B-2201"),
        ("fecha","Ultimo dia del mes informado","Fecha","Mensual","SBS y BCRP","Llave temporal de integracion"),
        ("costo_fondeo","Gastos de fondeo de 12 meses / saldo promedio de 12 meses x 100","Porcentaje anual","Mensual","SBS","B-2201, calculo de 12 meses"),
        ("depositos","Obligaciones con el publico menos otras obligaciones, mas depositos del sistema financiero","Miles de S/","Mensual","SBS","B-2201"),
        ("adeudados","Adeudos y obligaciones financieras, saldo","Miles de S/","Mensual","SBS","B-2201, contrastado con B-2310"),
        ("deuda_subordinada","Obligaciones en circulacion subordinadas","Miles de S/","Mensual","SBS","B-2201"),
        ("patrimonio_efectivo","Patrimonio efectivo regulatorio de la SBS, no patrimonio contable","Miles de S/","Mensual","SBS","B-2370"),
        ("tipmn_bcrp","Tasa de interes pasiva promedio de mercado en moneda nacional","Porcentaje","Mensual","BCRP","PN07816NM"),
        ("tipmex_bcrp","Tasa de interes pasiva promedio de mercado en moneda extranjera","Porcentaje","Mensual","BCRP","PN07836NM"),
        ("roe","Rentabilidad sobre patrimonio","Porcentaje","Mensual","SBS","B-2401"),
        ("flag_atipico","Marca estadistica MAD/IQR o salto de saldo >3x","Booleano","Mensual","Calculada","03_limpieza_datos.py"),
        ("*_w","Version winsorizada p1-p99 (global por defecto) de CF, ROE y ratios; originales intactos","Porcentaje","Mensual","Calculada","03_limpieza_datos.py"),
        ("ln_*","ln(1 + monto en miles de S/) de depositos, adeudados, deuda subordinada y patrimonio efectivo","Log","Mensual","Calculada","03_limpieza_datos.py"),
        ("completo_modelo","Indica si Y y X1-X4 estan presentes","Booleano","Mensual","Calculada","03_limpieza_datos.py")
    ]
    dd=pd.DataFrame(dic,columns=["Variable","Definicion operativa","Unidad","Frecuencia","Fuente","Referencia"])
    hoja_dataframe("Guia de variables",dd,anchos={"Definicion operativa":79,"Referencia":42,"Variable":27})
    wb.active=1  # Abrir vista oficial compacta, con las seis variables completas.
    archivo.parent.mkdir(parents=True,exist_ok=True)
    wb.save(archivo)
    if logger:logger.info("EXCEL_VISUAL | archivo=%s | filas_maestro=%s | filas_vista_completa=%s",archivo.name,len(p),len(prim))
    return archivo


def generar_vista_desde_csv(solo_anio:int|None,logger:logging.Logger,winsor_modo:str=WINSOR_MODO) -> Path:
    """Crea CSV oficial y Excel desde el maestro; sin releer libros SBS.

    Migracion: si aun no existe base_maestra_<codigo>.csv, COPIA (sin editar)
    el CSV historico amplio antes de reemplazar datos_procesados.
    No se permite crear una maestra ficticia desde un CSV ya reducido.
    """
    sufijo=f"_prueba_{solo_anio}" if solo_anio is not None else ""
    maestra=PROCESADOS/f"base_maestra_{CODIGO_MATRICULA}{sufijo}.csv"
    oficial=PROCESADOS/f"datos_procesados_{CODIGO_MATRICULA}{sufijo}.csv"
    if not maestra.is_file():
        if not oficial.is_file():
            raise FileNotFoundError(f"Falta la base maestra y el CSV historico: {maestra} / {oficial}")
        encabezados=pd.read_csv(oficial,encoding='utf-8-sig',nrows=0).columns
        if 'completo_modelo' not in encabezados or 'flag_atipico' not in encabezados:
            raise ValueError('El CSV disponible ya no tiene las columnas de auditoria; ejecutar limpieza completa desde datos_crudos/')
        shutil.copy2(oficial,maestra)  # conserva bytes y hash del antiguo CSV de 54 columnas
        logger.info('MIGRACION | copiada maestra integra desde %s',oficial.name)
        print(f"[RESPALDO] CSV historico copiado intacto a: {maestra.relative_to(RAIZ)}")
    hash_maestra=hashlib.sha256(maestra.read_bytes()).hexdigest()
    panel=pd.read_csv(maestra,encoding='utf-8-sig',low_memory=False)
    # Un maestro antiguo 2018-2025 no sustituye la reconstruccion 2017-2025.
    fechas_maestra = pd.to_datetime(panel["fecha"], errors="raise")
    if fechas_maestra.min() > pd.Timestamp(FECHA_INICIO) + pd.offsets.MonthEnd(0):
        raise ValueError("Base maestra empieza despues de enero 2017. Regenera desde datos_crudos/ ejecutando 03 SIN --solo-visualizar")
    if (fechas_maestra < pd.Timestamp(FECHA_INICIO)).any():
        raise ValueError("La base maestra contiene anios previos al estudio; 2016 solo debe entrar al calculo auxiliar")
    incid_path=PROCESADOS/f"incidencias_limpieza_{CODIGO_MATRICULA}{sufijo}.csv"
    inc=pd.read_csv(incid_path,encoding='utf-8-sig') if incid_path.exists() else pd.DataFrame(columns=['fecha','banco_id','tipo','detalle'])
    salida=PROCESADOS/f"VISTA_PANEL_BANCARIO_{CODIGO_MATRICULA}{sufijo}.xlsx"
    panel['fecha']=pd.to_datetime(panel['fecha'],errors='raise')
    panel=marcar_atipicos(panel)
    panel,resumen_w=winsorizar_variables(panel,modo=winsor_modo)
    panel=agregar_logaritmos(panel)
    analisis=PROCESADOS/f"panel_analisis_winsorizado_{CODIGO_MATRICULA}{sufijo}.csv"
    resumen_path=PROCESADOS/f"resumen_winsorizacion_{CODIGO_MATRICULA}{sufijo}.csv"
    crear_csv_oficial(panel,oficial,logger)
    crear_panel_analisis_winsorizado(panel,analisis,resumen_w,resumen_path,logger)
    crear_vista_excel(panel,inc,salida,logger)
    if hash_maestra!=hashlib.sha256(maestra.read_bytes()).hexdigest():
        raise RuntimeError('La base maestra cambio durante la generacion: abortar')
    print(f"[BASE OFICIAL] {oficial.relative_to(RAIZ)}")
    print(f"[BASE MAESTRA] {maestra.relative_to(RAIZ)} | SHA-256 {hash_maestra}")
    print(f"[VISTA EXCEL] {salida.relative_to(RAIZ)}")
    print(f"[PANEL WINSORIZADO] {analisis.relative_to(RAIZ)}")
    print(f"[RESUMEN WINSORIZACION] {resumen_path.relative_to(RAIZ)}")
    return salida


def normalizar_auxiliar_por_cobertura(
    aux: dict[str, dict],
    balance: dict[str, dict],
    reporte: str,
    fecha: pd.Timestamp,
    incidencias: list[dict],
) -> dict[str, dict]:
    """Unifica auxiliares sin mezclar coberturas ni perder bancos ajenos.

    Si un auxiliar contiene coberturas individual y con sucursales, se usa
    EXCLUSIVAMENTE la cobertura ya seleccionada en el balance B-2201.
    Para adeudados y pasivos, se verifica adicionalmente la conciliacion.
    Una ambiguedad sin criterio verificable omite solo ese banco/reporte.
    """
    grupos: dict[str, list[tuple[str, dict]]] = defaultdict(list)
    for original, valores in aux.items():
        destino = ALIAS_COBERTURA.get(original, original)
        grupos[destino].append((original, valores))

    def datos_equivalentes(registros: list[tuple[str, dict]]) -> bool:
        """Comparacion estricta de campos compartidos; NA solo equivale a NA."""
        a = registros[0][1]
        for _, b in registros[1:]:
            if set(a) != set(b):
                return False
            for clave in a:
                va, vb = a[clave], b[clave]
                na, nb = num(va), num(vb)
                if math.isfinite(na) or math.isfinite(nb):
                    if not (math.isfinite(na) and math.isfinite(nb)
                            and math.isclose(na, nb, rel_tol=1e-9, abs_tol=0.01)):
                        return False
                elif pd.isna(va) and pd.isna(vb):
                    continue
                elif str(va) != str(vb):
                    return False
        return True

    res: dict[str, dict] = {}
    for banco, candidatos in grupos.items():
        if len(candidatos) == 1:
            res[banco] = candidatos[0][1]
            continue

        origenes = [nombre for nombre, _ in candidatos]
        preferida = None
        cobertura = balance.get(banco, {}).get("cobertura_b2201", "")
        if banco == "BCP":
            if cobertura == "CON_SUCURSALES":
                preferida = "BCP_EXT"
            elif cobertura == "INDIVIDUAL_UNICA":
                preferida = "BCP_DOM"
        elif banco == "INTERBANK":
            if cobertura == "CON_SUCURSALES":
                preferida = "INTERBANK_EXT"
            elif cobertura == "INDIVIDUAL_UNICA":
                preferida = "INTERBANK"

        elegidos = [(nombre, registro) for nombre, registro in candidatos
                    if nombre == preferida] if preferida else []
        if len(elegidos) == 1:
            origen, registro = elegidos[0]
            # Verificacion contra B-2201 cuando existe cuenta comparable.
            comparables = {
                "B-2310": ("adeudados_b2310", "adeudados"),
                "B-2342": ("pasivos_b2342", "pasivos_totales"),
            }
            if reporte in comparables:
                aux_col, bal_col = comparables[reporte]
                a = num(registro.get(aux_col))
                b = num(balance.get(banco, {}).get(bal_col))
                if not (math.isfinite(a) and math.isfinite(b)
                        and not discrepancia(a, b)):
                    incidencias.append({
                        "fecha": fecha, "banco_id": banco,
                        "tipo": "COBERTURA_AUX_NO_CONCILIA",
                        "detalle": (f"{reporte}: {origen} seleccionado en B-2201, "
                                    f"pero {aux_col}={a} difiere de {bal_col}={b}; "
                                    "se omite solo este banco/reporte")})
                    continue
            res[banco] = registro
            incidencias.append({
                "fecha": fecha, "banco_id": banco,
                "tipo": "COBERTURA_AUX_HOMOLOGADA",
                "detalle": (f"{reporte}: {origen} elegido por cobertura_b2201="
                            f"{cobertura}; alternativas={origenes}; "
                            "sin sumar ni combinar coberturas")})
            continue

        if datos_equivalentes(candidatos):
            res[banco] = candidatos[0][1]
            incidencias.append({
                "fecha": fecha, "banco_id": banco,
                "tipo": "COBERTURAS_AUX_EQUIVALENTES",
                "detalle": f"{reporte}: {origenes}; registros cotejados y equivalentes"})
            continue

        incidencias.append({
            "fecha": fecha, "banco_id": banco,
            "tipo": "COBERTURA_AUX_AMBIGUA",
            "detalle": (f"{reporte}: {origenes}; no coinciden con cobertura "
                        "documentada ni tienen importes equivalentes; "
                        "se omite solo este banco/reporte")})

    return res


def procesa(solo_anio:int|None,max_bancos:int, logger:logging.Logger,
            guion_cero:bool=True, winsor_modo:str=WINSOR_MODO) -> dict[str,Path]:
    PROCESADOS.mkdir(parents=True,exist_ok=True)
    catalogo,incidencias=buscar_archivos(solo_anio)
    if not catalogo:raise FileNotFoundError("No se encontraron SBS originales con nombre esperado en datos_crudos")
    audit=[];datos=[];n_validos=defaultdict(int)
    for fecha,reportes in sorted(catalogo.items()):
        if "B-2201" not in reportes:
            incidencias.append({"fecha":fecha,"banco_id":"","tipo":"FALTA_BALANCE", "detalle":"B-2201 no disponible: sin registros nuevos para este mes"})
            continue
        try:
            d=extraer_balance(reportes["B-2201"],audit,guion_cero)
            d=unificar_coberturas(d)  # CORRECCION v3: un solo BCP y un solo INTERBANK
            n_validos["B-2201"]+=1
        except (Exception) as e:
            logger.error("B-2201 | %s | %s",fecha,e)
            incidencias.append({"fecha":fecha,"banco_id":"","tipo":"ERROR_BALANCE", "detalle":f"{reportes['B-2201'].name}: {e}"})
            continue
        for code in SBS_REPORTES:
            if code=="B-2201":continue
            path=reportes.get(code)
            if not path:
                incidencias.append({"fecha":fecha,"banco_id":"","tipo":"FALTA_REPORTE", "detalle":code})
                continue
            try:
                aux=tabla_por_empresa(path,code,audit)
                n_validos[code]+=1
                # Un conflicto de cobertura en UN banco no debe descartar
                # toda la informacion de los otros bancos del mismo reporte.
                # Priorizar la misma cobertura seleccionada en B-2201; registrar
                # la decision y no fusionar nunca montos de coberturas distintas.
                aux_normalizado = normalizar_auxiliar_por_cobertura(
                    aux, d, code, fecha, incidencias)
                for bid, reg in aux_normalizado.items():
                    if bid in d:
                        d[bid].update(reg)
                    else:
                        incidencias.append({"fecha": fecha, "banco_id": bid, "tipo": "BANCO_SIN_BALANCE", "detalle": f"{code}: entidad no presente en B-2201"})
            except Exception as e:
                incidencias.append({"fecha":fecha,"banco_id":"","tipo":"ERROR_FORMATO", "detalle":f"{path.name}: {e}"})
                logger.error("%s | %s | %s",code,fecha,e)
        for bid,r in d.items():
            r["fecha"]=fecha
            # Comprobar concordancia entre fuentes SIN cambiar cifras.
            comp=[
                ("ADEUDADOS_B2310",r.get("adeudados"),r.get("adeudados_b2310")),
                ("PASIVOS_B2342",r.get("pasivos_totales"),r.get("pasivos_b2342")),
                ("GF_B2390", sum((r.get("gasto_publico_acum",np.nan),r.get("gasto_depsistema_acum",np.nan),r.get("gasto_adeudados_acum",np.nan))) if all(math.isfinite(num(r.get(k))) for k in ("gasto_publico_acum","gasto_depsistema_acum","gasto_adeudados_acum")) else np.nan, r.get("gasto_cf_b2390")),
            ]
            for etiqueta,a,b in comp:
                if discrepancia(a,b):
                    incidencias.append({"fecha":fecha,"banco_id":bid,"tipo":"DESCUADRE_"+etiqueta,"detalle":f"balance={a}; reporte={b}"})
            # La estructura B-2342 representa participaciones, no saldos.
            if all(math.isfinite(num(r.get(k))) for k in ("depositos","pasivos_totales","pct_dep_publico_b2342","pct_dep_sistema_b2342")) and r["pasivos_totales"]>0:
                pct=100*r["depositos"]/r["pasivos_totales"]
                reportado=r["pct_dep_publico_b2342"]+r["pct_dep_sistema_b2342"]
                if abs(pct-reportado)>0.35:
                    incidencias.append({"fecha":fecha,"banco_id":bid,"tipo":"DESCUADRE_DEP_B2342","detalle":f"ratio_balance={pct:.4f}; suma_porcentajes={reportado:.4f}"})
            if r.get("aviso_apr_b2402"):
                incidencias.append({"fecha":fecha,"banco_id":bid,
                                    "tipo":"APR_NO_RECONSTRUIBLE_B2402",
                                    "detalle":r["aviso_apr_b2402"]})
            pe=r.get("patrimonio_efectivo",np.nan); apr=r.get("apr",np.nan);ratio=r.get("ratio_capital_global",np.nan)
            if all(math.isfinite(num(x)) for x in (pe,apr,ratio)) and apr>0 and abs(pe/apr*100-ratio)>0.35:
                incidencias.append({"fecha":fecha,"banco_id":bid,"tipo":"DESCUADRE_CAPITAL","detalle":f"PE/APR={pe/apr*100:.3f}; ratio={ratio:.3f}"})
            datos.append(r)
        print(f"[LIMPIEZA] {fecha:%Y-%m}: {len(d)} bancos desde B-2201")

    if not datos:raise ValueError("No hubo filas extraibles del balance SBS; revisar auditoria y hojas")
    df=pd.DataFrame(datos)
    fechas=pd.to_datetime(df["fecha"])
    df["fecha"]=fechas
    if df.duplicated(["banco_id","fecha"]).any():
        raise ValueError("Duplicados banco_id + fecha en SBS; detener para corregir")
    for v in ("oblig_publico","otras_oblig_publico","depositos_sistema","adeudados","deuda_subordinada","patrimonio_efectivo","activos_totales","pasivos_totales","gasto_publico_acum","gasto_depsistema_acum","gasto_adeudados_acum","roe"):
        if v not in df: df[v]=np.nan
        df[v]=pd.to_numeric(df[v],errors="coerce")

    # unificar_coberturas() ya elegio una sola cobertura B-2201 por
    # banco-mes ANTES de unir los reportes auxiliares. No invertir la regla aqui.
    inesperados = {"BCP_DOM", "BCP_EXT", "INTERBANK_EXT"} & set(df["banco_id"])
    if inesperados:
        raise ValueError(f"Coberturas no homologadas tras unificar_coberturas: {inesperados}")
    if "cobertura_b2201" in df:
        logger.info("COBERTURAS_B2201 | %s", df.loc[df["banco_id"].isin(["BCP", "INTERBANK"])].groupby(["banco_id", "cobertura_b2201"]).size().to_dict())

    if max_bancos>0:
        df["cobertura_4_saldos"]=df[["depositos","adeudados","deuda_subordinada","patrimonio_efectivo"]].notna().all(axis=1)
        ranking=df.groupby("banco_id").agg(
            meses=("fecha","nunique"),
            completos=("cobertura_4_saldos","sum")
        ).sort_values(["meses","completos"],ascending=False)
        # Seleccion reproducible por cobertura historica, empates por ID.
        ranking=ranking.assign(id_orden=ranking.index).sort_values(["meses","completos","id_orden"],ascending=[False,False,True])
        elegidos=set(ranking.head(max_bancos).index)
        df=df.loc[df["banco_id"].isin(elegidos)].copy()
        logger.info("SELECCION | max_bancos=%s | seleccionados=%s",max_bancos,sorted(elegidos))

    df=df.sort_values(["banco_id","fecha"]).reset_index(drop=True)
    # Comprobaciones de signos; no forzar a cero ni borrar las cifras originales.
    for _,r in df.iterrows():
        for v in ("depositos","adeudados","deuda_subordinada","patrimonio_efectivo","activos_totales","pasivos_totales"):
            if pd.notna(r[v]) and r[v]<0:
                incidencias.append({"fecha":r["fecha"],"banco_id":r["banco_id"],"tipo":"SALDO_NEGATIVO", "detalle":f"{v}={r[v]}"})

    df=construir_flujos(df,incidencias)
    # Las filas de la ventana previa (2016 o anio anterior en modo prueba) solo
    # sirvieron para el costo de fondeo movil; no son observaciones del estudio.
    inicio_panel = pd.Timestamp(solo_anio,1,1) if solo_anio else pd.Timestamp(FECHA_INICIO)
    n_ventana = int((df["fecha"] < inicio_panel).sum())
    if n_ventana:
        logger.info("VENTANA_PREVIA | filas usadas solo para CF y excluidas=%s", n_ventana)
        df = df.loc[df["fecha"] >= inicio_panel].reset_index(drop=True)
        incidencias[:] = [i for i in incidencias
                          if not (isinstance(i.get("fecha"), pd.Timestamp) and i["fecha"] < inicio_panel)]
    for variable,numerador,denominador in (
        ("dep_ratio","depositos","pasivos_totales"),
        ("ade_ratio","adeudados","pasivos_totales"),
        ("ds_ratio","deuda_subordinada","pasivos_totales"),
        ("pe_ratio","patrimonio_efectivo","activos_totales"),
    ):
        n=pd.to_numeric(df[numerador],errors="coerce")
        d=pd.to_numeric(df[denominador],errors="coerce")
        df[variable]=np.where(d>0,100*n/d,np.nan)
        # ratios negativos no se arreglan; bandera de auditoria.
        malo=df[variable].lt(0)|(df[variable].gt(100) & variable.startswith(("dep","ade","ds")))
        for _,r in df.loc[malo.fillna(False)].iterrows():
            incidencias.append({"fecha":r["fecha"],"banco_id":r["banco_id"],"tipo":"RATIO_SOSPECHOSO", "detalle":f"{variable}={r[variable]}"})

    tasas=leer_bcrp()
    df=df.merge(tasas,on="fecha",how="left",validate="many_to_one",indicator=True)
    df["fuente_bcrp_disponible"]=df.pop("_merge").eq("both") & df["tipmn_bcrp"].notna() & df["tipmex_bcrp"].notna()
    df["anio"]=df["fecha"].dt.year
    df["mes"]=df["fecha"].dt.month
    df=marcar_atipicos(df)
    # Version winsorizada paralela: los valores originales nunca se reemplazan.
    df,resumen_winsor=winsorizar_variables(df,modo=winsor_modo)
    df=agregar_logaritmos(df)
    # Definir calidad base en funcion de las cinco variables, no ocultar faltantes.
    principales=["costo_fondeo","depositos","adeudados","deuda_subordinada","patrimonio_efectivo"]
    df["completo_modelo"]=df[principales].notna().all(axis=1)
    df["indicador_calidad"]=np.select(
        [~df["completo_modelo"],df["flag_atipico"],~df["fuente_bcrp_disponible"]],
        ["INCOMPLETO","ATIPICO_REVISAR","SIN_BCRP"],default="OK")
    # Rompimiento regulatorio potencial: comparar 2017-2022 vs. 2023-2025.
    df["regimen_patrimonio"] = np.where(df["fecha"].dt.year>=2023,"Basilea_III_desde_2023","regimen_pre_2023")
    df["fuente_sbs"]="SBS: B-2201; B-2310; B-2342; B-2370; B-2390; B-2401; B-2402"
    df["fuente_bcrp"]="BCRP: PN07816NM; PN07836NM"
    # Seleccion consistente del modelo: no duplicar columna banco de resultados.
    primeras=["banco_id","banco_nombre","fecha","anio","mes","costo_fondeo","depositos","adeudados","deuda_subordinada","patrimonio_efectivo","dep_ratio","ade_ratio","ds_ratio","pe_ratio","tipmn_bcrp","tipmex_bcrp","roe","flag_atipico","detalle_atipico","completo_modelo","indicador_calidad","fuente_sbs","fuente_bcrp"]
    for col in primeras:
        if col not in df: df[col]=np.nan
    df=df[primeras+[c for c in df.columns if c not in primeras]]
    df["fecha"]=df["fecha"].dt.strftime("%Y-%m-%d")

    # La prueba parcial nunca reemplaza el CSV definitivo.
    sufijo = f"_prueba_{solo_anio}" if solo_anio is not None else ""
    paths={
        "diagnostico_cabeceras":PROCESADOS/f"diagnostico_cabeceras_{CODIGO_MATRICULA}{sufijo}.csv",
        "panel": PROCESADOS/f"datos_procesados_{CODIGO_MATRICULA}{sufijo}.csv",
        "maestra": PROCESADOS/f"base_maestra_{CODIGO_MATRICULA}{sufijo}.csv",
        "auditoria":PROCESADOS/f"auditoria_cuentas_{CODIGO_MATRICULA}{sufijo}.csv",
        "incidencias":PROCESADOS/f"incidencias_limpieza_{CODIGO_MATRICULA}{sufijo}.csv",
        "atipicos":PROCESADOS/f"revision_atipicos_{CODIGO_MATRICULA}{sufijo}.csv",
        "cobertura":PROCESADOS/f"cobertura_bancaria_{CODIGO_MATRICULA}{sufijo}.csv",
        "sensibilidad":PROCESADOS/f"modelo_sensibilidad_sin_atipicos_{CODIGO_MATRICULA}{sufijo}.csv",
        "analisis_winsorizado":PROCESADOS/f"panel_analisis_winsorizado_{CODIGO_MATRICULA}{sufijo}.csv",
        "resumen_winsorizacion":PROCESADOS/f"resumen_winsorizacion_{CODIGO_MATRICULA}{sufijo}.csv",
        "inventario":PROCESADOS/f"inventario_crudos_{CODIGO_MATRICULA}{sufijo}.csv",
        "embudo":PROCESADOS/f"embudo_observaciones_{CODIGO_MATRICULA}{sufijo}.csv",
        "atipicos_residuales":PROCESADOS/f"atipicos_residuales_{CODIGO_MATRICULA}{sufijo}.csv",
        "vista_excel":PROCESADOS/f"VISTA_PANEL_BANCARIO_{CODIGO_MATRICULA}{sufijo}.xlsx",
    }
    df.to_csv(paths["maestra"],index=False,encoding="utf-8-sig",float_format="%.9f")
    crear_csv_oficial(df,paths["panel"],logger)
    crear_panel_analisis_winsorizado(
        df, paths["analisis_winsorizado"], resumen_winsor, paths["resumen_winsorizacion"], logger
    )
    df_audit=pd.DataFrame(audit,columns=["archivo","reporte","banco_id","variable","problema","ubicacion","valor"])
    df_audit.to_csv(paths["auditoria"],index=False,encoding="utf-8-sig")
    df_audit[df_audit["problema"].astype(str).str.startswith("CABECERA_")].to_csv(
        paths["diagnostico_cabeceras"],index=False,encoding="utf-8-sig")
    inc=pd.DataFrame(incidencias,columns=["fecha","banco_id","tipo","detalle"])
    inc.to_csv(paths["incidencias"],index=False,encoding="utf-8-sig")
    df.loc[df["flag_atipico"], ["banco_id","fecha","detalle_atipico"]+principales+["dep_ratio","ade_ratio","ds_ratio","pe_ratio","roe"]].to_csv(paths["atipicos"],index=False,encoding="utf-8-sig")
    # Submuestra EXCLUSIVAMENTE para pruebas de sensibilidad; no reemplaza panel oficial.
    df.loc[df["completo_modelo"] & ~df["flag_atipico"]].to_csv(paths["sensibilidad"],index=False,encoding="utf-8-sig",float_format="%.9f")
    cov=df.groupby("banco_id",as_index=False).agg(banco_nombre=("banco_nombre","first"),meses=("fecha","nunique"),meses_modelo_completo=("completo_modelo","sum"),cf_disponible=("costo_fondeo",lambda s:s.notna().sum()),atipicos=("flag_atipico","sum"))
    cov.to_csv(paths["cobertura"],index=False,encoding="utf-8-sig")
    inventario=[]
    for fecha,reportes in sorted(catalogo.items()):
        for code,f in reportes.items():
            inventario.append({"fecha":fecha.strftime("%Y-%m-%d"),"reporte":code,"archivo":f.name,
                "tamano_bytes":f.stat().st_size,"sha256":hashlib.sha256(f.read_bytes()).hexdigest()})
    pd.DataFrame(inventario).to_csv(paths["inventario"],index=False,encoding="utf-8-sig")
    crear_vista_excel(df,inc,paths["vista_excel"],logger)
    embudo=embudo_observaciones(df)
    embudo.to_csv(paths["embudo"],index=False,encoding="utf-8-sig")
    muestra_oficial=df.loc[_mascara_muestra_principal(df)]
    residuales=atipicos_residuales(muestra_oficial)
    residuales.to_csv(paths["atipicos_residuales"],index=False,encoding="utf-8-sig",float_format="%.6f")
    n_guion=int(df_audit["problema"].eq("GUION_SBS_CERO").sum())
    sha=hashlib.sha256(paths["panel"].read_bytes()).hexdigest()
    sha_maestra=hashlib.sha256(paths["maestra"].read_bytes()).hexdigest()
    logger.info("FIN | meses_sbs=%s | bancos=%s | filas=%s | completos=%s | atipicos=%s | winsorizados=%s | sha256=%s | reporte=%s",len(catalogo),df.banco_id.nunique(),len(df),int(df.completo_modelo.sum()),int(df.flag_atipico.sum()),int(df.flag_winsorizado.sum()),sha,dict(n_validos))
    print("\n=== RESULTADO DE LA LIMPIEZA ===")
    print(f"Bancos: {df.banco_id.nunique()} | Filas panel: {len(df)} | Filas con 5 variables: {df.completo_modelo.sum()}")
    print(f"Atipicos marcados: {df.flag_atipico.sum()} | Incidencias: {len(inc)}")
    print("\nCobertura anual del panel (filas / casos completos):")
    print(df.groupby("anio").agg(filas=("fecha","size"), completos=("completo_modelo","sum")).to_string())
    if not inc.empty:
        print("\nIncidencias por tipo:")
        print(inc["tipo"].value_counts().to_string())
    ambiguas=df_audit[df_audit["problema"].eq("CABECERA_AMBIGUA_BANCO_OMITIDO")]
    if len(ambiguas):
        print(f"[ATENCION] {len(ambiguas)} cabeceras ambiguas: revisar {paths['diagnostico_cabeceras'].name}")
    print(f"Base OFICIAL lista para el analisis (13 originales + 12 calculadas): {paths['panel'].relative_to(RAIZ)}")
    print(f"Base maestra (datos y control): {paths['maestra'].relative_to(RAIZ)}")
    print(f"Panel PARALELO con winsorizacion y logaritmos: {paths['analisis_winsorizado'].relative_to(RAIZ)}")
    print(f"Auditoria de winsorizacion: {paths['resumen_winsorizacion'].relative_to(RAIZ)}")
    print(f"Vista Excel: {paths['vista_excel'].relative_to(RAIZ)}")
    print(f"Registros con al menos una variable winsorizada: {int(df['flag_winsorizado'].sum())}")
    n_6=int(df[["costo_fondeo","depositos","adeudados","deuda_subordinada","patrimonio_efectivo","roe"]].notna().all(axis=1).sum())
    print(f"Filas con 6 variables (incluyendo ROE): {n_6}")
    n_oficial=pd.read_csv(paths['panel'],encoding='utf-8-sig',usecols=['banco_id']).shape[0]
    print(f"Observaciones del modelo principal en CSV oficial (CF + 4 X + TIPMN): {n_oficial} | columnas: {13 + len(COLUMNAS_DERIVADAS_MODELO)}")
    if solo_anio is None and n_6 < 1000:
        print("[AVISO ROE] Menos de 1000 filas con ROE: el modelo principal conserva su propia muestra; revisar reportes SBS si se utiliza ROE.")
    print("\nEmbudo de observaciones (donde se pierden filas):")
    print(embudo.to_string(index=False))
    print(f"\nCeldas SBS con guion leidas como cero: {n_guion} (ver auditoria, problema=GUION_SBS_CERO)")
    print(f"Winsorizacion: modo={winsor_modo}, percentiles {WINSOR_LIM_INF:.0%}-{WINSOR_LIM_SUP:.0%}")
    if not residuales.empty:
        print("Extremos residuales (Tukey 3*IQR) en variables del modelo *_w / ln_*:")
        print(residuales[["variable","n","extremos_tukey_3iqr"]].to_string(index=False))
    if solo_anio is None:
        estado = "CUMPLE" if n_oficial >= OBJETIVO_OBSERVACIONES else "NO CUMPLE"
        print(f"[OBJETIVO >={OBJETIVO_OBSERVACIONES} OBS] {estado}: {n_oficial} observaciones utilizables en modelo principal")
    print(f"SHA-256 oficial: {sha}")
    print(f"SHA-256 maestra: {sha_maestra}")
    if n_oficial < OBJETIVO_OBSERVACIONES and solo_anio is None:
        print("[ADVERTENCIA] Menos de 1000 observaciones completas para el modelo; revisar cobertura y definiciones.")
    return paths


def main():
    p=argparse.ArgumentParser(description="Limpieza SBS+BCRP / panel bancario 2017-2025")
    p.add_argument("--solo-anio",type=int,default=None,help="Prueba sin descargar: procesa un anio SBS")
    p.add_argument("--max-bancos",type=int,default=0,help="0=todas las entidades (por defecto); N=las N con mayor cobertura")
    p.add_argument("--guion-como-faltante",action="store_true",
                   help="Trata el guion SBS '-' como NA (comportamiento anterior) en lugar de cero")
    p.add_argument("--winsor-modo",choices=("global","banco"),default=WINSOR_MODO,
                   help="Percentiles 1-99 de toda la muestra (global) o de cada banco")
    p.add_argument("--solo-visualizar",action="store_true",help="Genera CSV oficial (13 originales + 12 calculadas) y Excel desde base maestra previamente generada")
    args=p.parse_args()
    anio_ini,anio_fin=int(FECHA_INICIO[:4]),int(FECHA_CORTE[:4])
    if args.solo_anio is not None and not anio_ini<=args.solo_anio<=anio_fin:
        p.error(f"--solo-anio debe estar entre {anio_ini} y {anio_fin}")
    if args.max_bancos<0:p.error("--max-bancos debe ser >=0")
    logger=log_setup()
    logger.info("INICIO | fecha_real=%s | %s a %s | ventana_previa_desde=%s | solo_anio=%s | max_bancos=%s | guion_cero=%s | winsor=%s",
                datetime.now().astimezone().isoformat(),FECHA_INICIO,FECHA_CORTE,FECHA_INICIO_LECTURA,
                args.solo_anio,args.max_bancos,
                not args.guion_como_faltante,args.winsor_modo)
    if args.solo_visualizar:
        generar_vista_desde_csv(args.solo_anio,logger,args.winsor_modo)
        return
    if not CRUDOS.exists():raise SystemExit("No existe datos_crudos/ junto a codigo/")
    try:
        procesa(args.solo_anio,args.max_bancos,logger,
                guion_cero=not args.guion_como_faltante, winsor_modo=args.winsor_modo)
    except Exception:
        logger.exception("FALLO_EN_LIMPIEZA")
        raise


if __name__=="__main__":
    main()
