# Nombres y apellidos completos: ROMERO RAMON MIRIAN ANGELA
# Codigo de matricula: 2024200523C
# Tema 60: Estructura de financiamiento y costo de capital en el sistema bancario peruano
# Fecha de extraccion: 16-11-2026

"""
FINANZAS I - UNIDAD II
VIA 2: SCRAPING WEB Y DESCARGA PROGRAMATICA DE REPORTES SBS
====================================================================

OBJETIVO
--------
Localizar en las paginas oficiales de la Superintendencia de Banca,
Seguros y AFP (SBS) los enlaces a reportes mensuales de banca multiple,
descargarlos y conservarlos en su formato Excel original (XLS/XLSX).

ALCANCE
-------
- Fuente: portal estadistico de la SBS.
- Periodo de estudio: enero de 2018 a diciembre de 2025.
- Reportes considerados: siete tipos de reportes SBS.
- Carpeta de destino: datos_crudos/.
- Registro de ejecucion: log_ejecucion.txt, en la raiz del proyecto.

IMPORTANTE
----------
Este script SOLO obtiene, comprueba y almacena los archivos originales.
No limpia datos, no transforma variables y no calcula costo de fondeo.
La interpretacion contable y la integracion con el BCRP se realizan
posteriormente en codigo/03_limpieza_datos.py.


EJECUCION DESDE LA RAIZ DEL PROYECTO
------------------------------------
Modo demostracion (un reporte de diciembre de 2025):
    python codigo/02_scraping_web.py --modo demo

Modo completo (siete reportes, enero de 2018 a diciembre de 2025):
    python codigo/02_scraping_web.py --modo completo

Solo el anio 2025:
    python codigo/02_scraping_web.py --modo completo --solo-anio 2025

Un reporte particular:
    python codigo/02_scraping_web.py --modo completo --solo-reporte B-2401


"""

from __future__ import annotations

# =====================================================================
# 0. BIBLIOTECAS
# =====================================================================
# Biblioteca estandar: opciones de terminal, archivos, red, fechas,
# codificacion, validacion ZIP y huellas digitales SHA-256.
import argparse
import hashlib
import io
import logging
import sys
import time
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse
from urllib.robotparser import RobotFileParser

# Bibliotecas externas:
# - requests: realiza las solicitudes HTTP.
# - BeautifulSoup: interpreta el HTML y permite buscar enlaces.
import requests
from bs4 import BeautifulSoup


# =====================================================================
# 1. DATOS DEL PROYECTO, RUTAS Y PARAMETROS FIJOS
# =====================================================================
# Estos valores corresponden a la investigacion y se conservan tal
# como estaban definidos en el codigo original.
CODIGO_MATRICULA = "2024200523C"
FECHA_INICIO = "2018-01-01"
FECHA_CORTE = "2025-12-31"

# __file__ es la ruta del propio script. Su carpeta superior es codigo/;
# la siguiente es la raiz del proyecto, desde donde se guardan los datos.
RAIZ = Path(__file__).resolve().parent.parent
DATOS_CRUDOS = RAIZ / "datos_crudos"
LOG_COMPLETO = RAIZ / "log_ejecucion.txt"

# Pagina oficial de resultados estadisticos de la SBS.
PORTAL_SBS = (
    "https://www.sbs.gob.pe/app/stats_net/stats/"
    "EstadisticaSistemaFinancieroResultados.aspx"
)

# Cada codigo identifica un reporte publicado en el portal de banca
# multiple. Las etiquetas describen su utilidad general; NO sustituyen
# la identificacion precisa de cuentas dentro de cada archivo Excel.
REPORTES = {
    "B-2390": "estructura_gastos_financieros",       # Posible numerador del CF.
    "B-2342": "estructura_pasivo",                  # Depositos y obligaciones.
    "B-2310": "adeudos_obligaciones_financieras",   # Adeudados.
    "B-2201": "balance_general_y_resultados",      # Balance y resultados.
    "B-2370": "estructura_patrimonio_efectivo",    # Patrimonio efectivo.
    "B-2402": "requerimiento_patrimonio_ratio",    # Indicadores de capital.
    "B-2401": "indicadores_financieros",           # ROE, entre otros.
}

# Los enlaces de la SBS muestran los meses escritos en espanol.
# Se reconocen las dos grafias habituales de septiembre.
MESES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "setiembre": 9,
    "septiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

# Solo se aceptan URLs HTTPS cuyo dominio pertenezca a la SBS.
SERVIDORES_AUTORIZADOS = {"www.sbs.gob.pe", "intranet2.sbs.gob.pe"}

# El identificador informa que la consulta tiene finalidad academica.
USER_AGENT = (
    f"UNCP-Finanzas-I-2026-II/{CODIGO_MATRICULA} "
    "(investigacion-academica)"
)

# Controles de acceso y descarga. PAUSA_MINIMA_SEG evita solicitudes
# demasiado seguidas; MAX_BYTES se comprueba al recibir la respuesta.
TIMEOUT_SEG = 30
PAUSA_MINIMA_SEG = 1.3
MAX_BYTES = 20_000_000


# =====================================================================
# 2. FECHA Y REGISTRO DE EJECUCION
# =====================================================================
def fecha_lima() -> str:
    """Devuelve la fecha/hora con zona local configurada en el equipo.

    Se registra la fecha real de ejecucion para distinguirla del periodo
    historico de los reportes consultados. No se modifica el reloj local.
    """
    return datetime.now().astimezone().isoformat(timespec="seconds")


def crear_logger(modo: str) -> logging.Logger:
    """Prepara el registro de acciones en log_ejecucion.txt.

    El archivo permite demostrar cuando se ejecutaron las consultas,
    cuales enlaces respondieron y que archivos originales se guardaron.
    Se utiliza un mismo registro tecnico para ambas vias de extraccion.

    El parametro ``modo`` se conserva por compatibilidad con el script
    original; el modo concreto se documenta al iniciar ``ejecutar``.
    """
    destino = LOG_COMPLETO  # Unico log tecnico en la raiz del proyecto.
    destino.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("extraccion_sbs")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    # Evita duplicar registros si se invoca el codigo mas de una vez
    # dentro del mismo proceso de Python.
    for h in logger.handlers[:]:
        logger.removeHandler(h)
        h.close()

    h = logging.FileHandler(destino, encoding="utf-8")
    h.setFormatter(
        logging.Formatter("%(asctime)s | SBS | %(levelname)s | %(message)s")
    )
    logger.addHandler(h)
    return logger


# =====================================================================
# 3. CLIENTE HTTP: SESION, PAUSAS Y POLITICAS DE ACCESO
# =====================================================================
class ClienteSBS:
    """Realiza solicitudes a SBS con controles de acceso y registro.

    Una unica sesion reutiliza conexiones entre solicitudes. Antes de
    consultar las paginas se revisa robots.txt para el servidor usado.
    """

    def __init__(self, logger: logging.Logger):
        """Inicializa encabezados HTTP y estado del cliente de descarga."""
        self.logger = logger
        self.sesion = requests.Session()
        self.sesion.headers.update({
            "User-Agent": USER_AGENT,
            "Accept": (
                "text/html,application/vnd.ms-excel,"
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet,*/*;q=0.8"
            ),
        })
        self.ultimo_inicio = 0.0
        self.robots: dict[str, RobotFileParser | bool] = {}

    def esperar(self) -> None:
        """Respeta una pausa minima entre inicios de solicitudes HTTP."""
        restante = PAUSA_MINIMA_SEG - (time.monotonic() - self.ultimo_inicio)
        if restante > 0:
            time.sleep(restante)
        self.ultimo_inicio = time.monotonic()

    def comprobar_robots(self, url: str) -> None:
        """Comprueba dominio HTTPS y las restricciones de robots.txt.

        Si robots.txt no existe (HTTP 404), se permite continuar. Si
        no puede consultarse o prohibe el acceso, se detiene el proceso.
        La politica de cada servidor se consulta una sola vez por sesion.
        """
        parsed = urlparse(url)
        servidor = parsed.hostname

        # Primera barrera: no consultar servidores ajenos a la SBS.
        if parsed.scheme != "https" or servidor not in SERVIDORES_AUTORIZADOS:
            raise PermissionError(f"URL fuera de servidores SBS autorizados: {url}")

        if servidor not in self.robots:
            robots_url = f"https://{servidor}/robots.txt"
            self.esperar()
            try:
                r = self.sesion.get(robots_url, timeout=TIMEOUT_SEG)
            except requests.RequestException as exc:
                raise RuntimeError(f"No se pudo revisar robots.txt: {exc}") from exc

            self.logger.info("ROBOTS | HTTP=%s | url=%s", r.status_code, robots_url)
            if r.status_code == 404:
                self.robots[servidor] = True
            elif r.status_code == 200:
                parser = RobotFileParser()
                parser.parse(r.text.splitlines())
                self.robots[servidor] = parser
            else:
                raise PermissionError(
                    f"No se pudo verificar robots.txt en {servidor}: "
                    f"HTTP {r.status_code}"
                )

        regla = self.robots[servidor]
        if not (regla if isinstance(regla, bool) else regla.can_fetch(USER_AGENT, url)):
            raise PermissionError(f"robots.txt prohibe este acceso: {url}")

    def obtener(self, url: str) -> bytes:
        """Solicita una pagina HTML o un archivo Excel por HTTP GET.

        - Se comprueba el acceso antes de descargar.
        - Se permiten hasta dos intentos frente a errores de red.
        - No se reintentan respuestas de acceso denegado o limitado.
        - Se rechazan respuestas mayores a MAX_BYTES.
        - Los bytes descargados se devuelven sin transformarlos.
        """
        self.comprobar_robots(url)

        for intento in range(1, 3):
            self.esperar()
            try:
                r = self.sesion.get(
                    url, timeout=TIMEOUT_SEG, allow_redirects=True
                )
                servidor_final = urlparse(r.url).hostname
                if (
                    urlparse(r.url).scheme != "https"
                    or servidor_final not in SERVIDORES_AUTORIZADOS
                ):
                    raise PermissionError("Redireccion fuera de la SBS: " + r.url)

                self.logger.info(
                    "GET | intento=%s | HTTP=%s | bytes=%s | url=%s",
                    intento, r.status_code, len(r.content), url,
                )
                print(f"    HTTP {r.status_code} | {len(r.content)} bytes")

                if r.status_code in (401, 403, 429):
                    raise PermissionError(
                        f"Acceso rechazado o limitado (HTTP {r.status_code}); "
                        "se detiene"
                    )

                r.raise_for_status()
                if len(r.content) > MAX_BYTES:
                    raise ValueError("Respuesta excede el limite de seguridad de tamaño")
                return r.content

            except PermissionError:
                raise
            except (
                requests.ConnectionError, requests.Timeout, requests.HTTPError
            ) as exc:
                self.logger.warning("Error de red (intento %d): %s", intento, exc)
                if intento == 2:
                    raise
                time.sleep(2)

        raise RuntimeError("Error inesperado de descarga")


# =====================================================================
# 4. SCRAPING: LOCALIZAR LOS ENLACES DE EXCEL EN EL HTML
# =====================================================================
def sin_tildes(texto: str) -> str:
    """Normaliza textos para identificar nombres de meses con o sin tilde."""
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    ).lower()


def detectar_anio_mes(url: str) -> tuple[int, int] | None:
    """Identifica el anio y el mes presentes en una ruta de descarga SBS.

    Ejemplo de criterio: /2025/diciembre/reporte.xls -> (2025, 12).
    Si la ruta no contiene un anio seguido de un mes identificable,
    devuelve None. Este paso NO abre ni modifica el Excel.
    """
    partes = unquote(urlparse(url).path).strip("/").split("/")

    for i in range(len(partes) - 1):
        if partes[i].isdigit() and len(partes[i]) == 4:
            anio = int(partes[i])
            mes = MESES.get(sin_tildes(partes[i + 1]))
            if mes and 2017 <= anio <= 2026:
                return anio, mes
    return None


def extraer_enlaces_excel(
    html: bytes, pagina: str
) -> dict[tuple[int, int], str]:
    """Extrae del HTML enlaces Excel y los asocia a cada anio-mes.

    Se buscan exclusivamente enlaces existentes en la pagina oficial.
    Se excluyen otros dominios y archivos que no terminen en XLS/XLSX.
    No se inventan rutas de descarga. Si dos enlaces diferentes
    anuncian el mismo periodo, se solicita una revision manual.
    """
    soup = BeautifulSoup(html, "html.parser")
    encontrados: dict[tuple[int, int], str] = {}

    for enlace in soup.select("a[href]"):
        url = urljoin(pagina, enlace.get("href", "").strip())
        parseada = urlparse(url)

        if parseada.scheme != "https" or parseada.hostname not in SERVIDORES_AUTORIZADOS:
            continue
        if not unquote(parseada.path).lower().endswith((".xls", ".xlsx")):
            continue

        fecha = detectar_anio_mes(url)
        if fecha is None:
            continue
        if fecha in encontrados and encontrados[fecha] != url:
            raise ValueError(
                f"Existen dos enlaces diferentes para {fecha} en {pagina}; "
                "revisar manualmente"
            )
        encontrados[fecha] = url

    return encontrados


# =====================================================================
# 5. VALIDAR EL EXCEL SIN LIMPIAR SUS CIFRAS
# =====================================================================
def identificar_excel(contenido: bytes) -> tuple[str, list[str]]:
    """Confirma el formato real del archivo y obtiene sus nombres de hojas.

    Los XLS antiguos y los XLSX modernos tienen firmas de archivo
    diferentes. El libro se abre en memoria SOLO para verificar que
    sea legible y que contenga al menos una hoja. No se guarda ningun
    Excel transformado y no se interpretan aqui cuentas bancarias.
    """
    if contenido.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        import xlrd
        libro = xlrd.open_workbook(file_contents=contenido, on_demand=True)
        hojas = libro.sheet_names()
        libro.release_resources()
        formato = "xls"

    elif contenido.startswith(b"PK\x03\x04") and zipfile.is_zipfile(io.BytesIO(contenido)):
        from openpyxl import load_workbook
        libro = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
        hojas = libro.sheetnames
        libro.close()
        formato = "xlsx"

    else:
        raise ValueError(
            "La respuesta no es XLS/XLSX valido (podria ser HTML de error)"
        )

    if not hojas:
        raise ValueError("Libro Excel sin hojas")
    return formato, hojas


def guardar_sin_sobrescribir(contenido: bytes, ruta: Path) -> str:
    """Conserva los bytes originales y devuelve la huella SHA-256.

    Si el archivo ya existe y es identico, se conserva. Si el mismo
    nombre tiene contenido diferente, se detiene ese guardado para
    evitar reemplazar silenciosamente una fuente historica.
    """
    huella = hashlib.sha256(contenido).hexdigest()
    ruta.parent.mkdir(parents=True, exist_ok=True)

    if ruta.exists():
        actual = hashlib.sha256(ruta.read_bytes()).hexdigest()
        if actual != huella:
            raise FileExistsError(
                f"Revision detectada: {ruta.name} ya existe y difiere; "
                "conservar original y registrar incidencia"
            )
    else:
        ruta.write_bytes(contenido)
    return huella


# =====================================================================
# 6. SELECCION DEL PERIODO DE DESCARGA
# =====================================================================
def intervalo(
    modo: str, incluir_2017: bool, solo_anio: int | None
) -> list[tuple[int, int]]:
    """Devuelve los pares (anio, mes) que se intentaran descargar.

    Demo: exclusivamente diciembre de 2025.
    Completo: todos los meses del periodo 2018-2025 (o desde 2017 si
    se ha solicitado expresamente). --solo-anio limita el conjunto.
    """
    if modo == "demo":
        return [(2025, 12)]

    desde = 2017 if incluir_2017 else int(FECHA_INICIO[:4])
    hasta = int(FECHA_CORTE[:4])
    return [
        (a, m)
        for a in range(desde, hasta + 1)
        for m in range(1, 13)
        if solo_anio is None or a == solo_anio
    ]


# =====================================================================
# 7. EJECUCION: DESCUBRIR, DESCARGAR, COMPROBAR Y REGISTRAR
# =====================================================================
def ejecutar(args: argparse.Namespace) -> int:
    """Coordina la descarga de los reportes y presenta un resumen final.

    Para cada reporte:
      1) consulta su pagina HTML oficial;
      2) identifica enlaces mensuales de Excel publicados;
      3) busca los meses que corresponden al periodo elegido;
      4) descarga y verifica cada libro disponible;
      5) guarda el archivo original y registra su SHA-256;
      6) informa enlaces faltantes y errores sin inventar datos.

    Devuelve 0 si no hubo errores y se guardo al menos una descarga;
    en los demas casos, devuelve 1. Esta etapa NO genera un panel.
    """
    modo = args.modo
    logger = crear_logger(modo)
    logger.info(
        "INICIO | fecha=%s | modo=%s | inicio=%s | corte=%s",
        fecha_lima(), modo, FECHA_INICIO, FECHA_CORTE,
    )
    cliente = ClienteSBS(logger)
    periodos = intervalo(modo, args.incluir_2017, args.solo_anio)
    if not periodos:
        raise ValueError("No existen meses seleccionados para esos parametros")

    reportes = (
        [(args.solo_reporte, REPORTES[args.solo_reporte])]
        if args.solo_reporte
        else list(REPORTES.items())
    )
    if modo == "demo" and not args.solo_reporte:
        reportes = reportes[:1]  # Solo un reporte para una demostracion corta.

    destino = DATOS_CRUDOS
    destino.mkdir(parents=True, exist_ok=True)

    # Los contadores resumen archivos descargados, periodos sin enlace
    # y errores en la obtencion o comprobacion de reportes.
    ok = 0
    faltantes = 0
    errores = 0

    try:
        for codigo, nombre in reportes:
            pagina = f"{PORTAL_SBS}?c={codigo}"
            print(f"\n[SBS] Reporte {codigo}: {nombre} | HTML oficial")
            html = cliente.obtener(pagina)
            enlaces = extraer_enlaces_excel(html, pagina)
            print(f"    Enlaces mensuales descubiertos: {len(enlaces)}")

            if not enlaces:
                raise ValueError(
                    f"No se encontraron enlaces Excel publicados para {codigo}"
                )

            for anio, mes in periodos:
                periodo = f"{anio}-{mes:02d}"
                url = enlaces.get((anio, mes))
                if not url:
                    faltantes += 1
                    logger.warning(
                        "SIN ENLACE | reporte=%s | periodo=%s | pagina=%s",
                        codigo, periodo, pagina,
                    )
                    continue

                print(f"[SBS] {codigo} | {periodo}")
                try:
                    contenido = cliente.obtener(url)
                    formato, hojas = identificar_excel(contenido)

                    # El modo demo agrega un sufijo para no confundirse
                    # con los originales del periodo completo.
                    sufijo = "_demo" if modo == "demo" else ""
                    archivo = (
                        f"sbs_original_{codigo}_{anio}{mes:02d}_"
                        f"{CODIGO_MATRICULA}{sufijo}.{formato}"
                    )
                    ruta = destino / archivo
                    huella = guardar_sin_sobrescribir(contenido, ruta)
                    logger.info(
                        "ORIGINAL | reporte=%s | mes=%s | bytes=%d | "
                        "sha256=%s | url=%s",
                        codigo, periodo, len(contenido), huella, url,
                    )
                    ok += 1
                    print(f"    Guardado: {ruta.name} | hojas={hojas}")

                except PermissionError:
                    raise
                except (
                    requests.RequestException, ValueError, RuntimeError,
                    OSError, zipfile.BadZipFile,
                ) as exc:
                    errores += 1
                    logger.error("ERROR | %s %s | %s", codigo, periodo, exc)
                    print(f"    ERROR: {exc}")

    except PermissionError as exc:
        errores += 1
        logger.error("DETENIDO POR RESTRICCION: %s", exc)
        print(f"[SBS] Se detuvo el rastreo: {exc}")
    except (requests.RequestException, ValueError, RuntimeError) as exc:
        errores += 1
        logger.error("ERROR GENERAL: %s", exc)
        print(f"[SBS] Error: {exc}")
    finally:
        logger.info(
            "FIN | modo=%s | correctos=%d | sin_enlace=%d | errores=%d",
            modo, ok, faltantes, errores,
        )
        print(
            f"\n[SBS] Descargas verificadas: {ok}; "
            f"faltantes sin enlace: {faltantes}; errores: {errores}"
        )
        print("[SBS] Registro tecnico: log_ejecucion.txt (raiz del proyecto)")
        print("[SBS] Los XLS/XLSX originales no fueron limpiados ni transformados.")

    return 0 if errores == 0 and ok > 0 else 1


# =====================================================================
# 8. PARAMETROS DE TERMINAL Y PUNTO DE ENTRADA
# =====================================================================
def main() -> None:
    """Lee las opciones que introduce el usuario y ejecuta el proceso SBS."""
    parser = argparse.ArgumentParser(
        description="SBS: scraping HTML y descarga automatica de XLS originales"
    )
    parser.add_argument(
        "--modo", choices=["demo", "completo"], default="demo"
    )
    parser.add_argument(
        "--incluir-2017", action="store_true",
        help="Descargar 2017 para la ventana movil de 12 meses",
    )
    parser.add_argument(
        "--solo-anio", type=int, choices=range(2017, 2026), metavar="AAAA"
    )
    parser.add_argument(
        "--solo-reporte", choices=REPORTES.keys(), metavar="CODIGO"
    )
    args = parser.parse_args()

    # El modo demo esta definido exclusivamente para diciembre de 2025.
    if args.modo == "demo" and (
        args.incluir_2017 or args.solo_anio is not None
    ):
        parser.error("En modo demo se utiliza exclusivamente diciembre 2025")
    if args.solo_anio == 2017 and not args.incluir_2017:
        parser.error("Para 2017 incluir --incluir-2017")

    sys.exit(ejecutar(args))


# La proteccion __main__ permite importar las funciones para pruebas
# sin iniciar descargas por accidente.
if __name__ == "__main__":
    main()
