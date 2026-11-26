# Nombres y apellidos completos: ROMERO RAMON MIRIAN ANGELA
# Codigo de matricula: 2024200523C
# Tema 60: Estructura de financiamiento y costo de capital en el sistema bancario peruano
# Fecha de extraccion: 16-11-2026

"""
FINANZAS I - UNIDAD II | VIA 1: EXTRACCION MEDIANTE API REST DEL BCRP
====================================================================
Nombre y Apellidos completos: ROMERO RAMON MIRIAN ANGELA
Codigo matricula : 2024200523C
Tema             : 60 - Estructura de financiamiento y costo de capital
                   en el sistema bancario peruano
Fecha de extraccion: 16/11/2026

OBJETIVO DEL SCRIPT
-------------------
Descargar, mediante la API oficial del Banco Central de Reserva del Peru
(BCRP), dos series mensuales de tasas pasivas para el periodo 2018-2025.

SERIES CONSULTADAS
------------------
PN07816NM : TIPMN - Tasa pasiva promedio en moneda nacional.
PN07836NM : TIPMEX - Tasa pasiva promedio en moneda extranjera.

IMPORTANTE: ambas tasas describen el sistema bancario a nivel agregado.
No representan el costo de fondeo individual de cada banco.

EJECUCION (desde la carpeta principal del proyecto)
---------------------------------------------------
    python codigo/01_extraccion_api.py
    python codigo/01_extraccion_api.py --modo demo

Modo completo : consulta las dos series desde 2018-01 hasta 2025-12.
Modo demo     : consulta solo TIPMN desde 2018-01 hasta 2018-03.

SALIDAS
-------
    datos_crudos/bcrp_original_PN07816NM_2024200523C.csv
    datos_crudos/bcrp_original_PN07836NM_2024200523C.csv
    log_ejecucion.txt

El modo demo anade '_demo' antes de '.csv' y no sustituye los originales.
"""

from __future__ import annotations

# ============================================================================
# 0. BIBLIOTECAS
# ============================================================================
# Bibliotecas de Python: parametros, hashes, registros, expresiones regulares,
# errores, pausas entre solicitudes, fecha de ejecucion y rutas de archivos.
import argparse
import hashlib
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path

# Biblioteca externa para realizar solicitudes HTTP a la API del BCRP.
import requests


# ============================================================================
# 1. IDENTIFICACION DEL PROYECTO Y PARAMETROS DE EXTRACCION
# ============================================================================
# Codigo utilizado en los nombres de los CSV para identificar al estudiante.
CODIGO_MATRICULA = "2024200523C"

# Limites temporales del estudio. El script convierte estas fechas a los
# periodos mensuales aceptados por el servicio del BCRP.
FECHA_INICIO = "2018-01-01"
FECHA_CORTE = "2025-12-31"
FECHA_CORTE_DEMO = "2018-03-31"  # Tres meses para la demostracion en vivo.

# Series oficiales de tasas pasivas. Se mantienen sus codigos de origen.
SERIES_BCRP = {
    "PN07816NM": "TIPMN - tasa pasiva promedio en moneda nacional",
    "PN07836NM": "TIPMEX - tasa pasiva promedio en moneda extranjera",
}

# Direccion base de la API y rutas relativas al proyecto.
# __file__ es este programa: su carpeta superior es 'codigo/' y la siguiente
# corresponde a la carpeta principal del proyecto.
URL_BASE = "https://estadisticas.bcrp.gob.pe/estadisticas/series/api"
RAIZ_PROYECTO = Path(__file__).resolve().parent.parent
CARPETA_CRUDOS = RAIZ_PROYECTO / "datos_crudos"
LOG_PRINCIPAL = RAIZ_PROYECTO / "log_ejecucion.txt"

# Parametros tecnicos: tiempo maximo de respuesta, reintentos y pausa de
# cortesia entre las dos series. No modifican los datos descargados.
TIMEOUT_SEG = 40
MAX_INTENTOS = 3
PAUSA_ENTRE_SERIES_SEG = 1.2
USER_AGENT = f"UNCP-Finanzas-I-2026-II/{CODIGO_MATRICULA} (uso-academico)"


# ============================================================================
# 2. CONFIGURACION DEL REGISTRO DE EJECUCION (LOG)
# ============================================================================
def configurar_log(modo: str) -> logging.Logger:
    """Prepara el registro de fecha, consultas HTTP y archivos generados.

    Parametro:
        modo: 'completo' o 'demo'. Se recibe para conservar la interfaz
              del codigo base; los mensajes detallan el modo en ejecutar().

    Retorna:
        Logger que escribe en 'log_ejecucion.txt' sin borrar sus entradas
        anteriores. Ese archivo documenta la trazabilidad de la extraccion.
    """
    ruta = LOG_PRINCIPAL  # Log tecnico compartido por ambas vias de extraccion.
    ruta.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("extraccion_bcrp")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    # Evita duplicar mensajes cuando se configura el logger mas de una vez.
    for manejador in logger.handlers[:]:
        logger.removeHandler(manejador)
        manejador.close()

    archivo = logging.FileHandler(ruta, encoding="utf-8")
    archivo.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    )
    logger.addHandler(archivo)
    return logger


# ============================================================================
# 3. CONSTRUCCION DE LA CONSULTA A LA API DEL BCRP
# ============================================================================
def convertir_periodo(fecha_iso: str) -> str:
    """Convierte 'AAAA-MM-DD' a 'AAAA-M' para el parametro de la API.

    Ejemplo: '2018-01-01' se convierte en '2018-1'.
    Esta conversion NO modifica ningun archivo descargado.
    """
    return f"{fecha_iso[:4]}-{int(fecha_iso[5:7])}"


def construir_endpoint(codigo: str, fecha_fin: str) -> str:
    """Construye la URL CSV usando codigo de serie y periodo solicitado.

    La ruta resultante identifica: serie / formato / inicio / fin / idioma.
    """
    inicio = convertir_periodo(FECHA_INICIO)
    fin = convertir_periodo(fecha_fin)
    return f"{URL_BASE}/{codigo}/csv/{inicio}/{fin}/esp"


# ============================================================================
# 4. SOLICITUD HTTP Y CONTROL DE ERRORES DE CONEXION
# ============================================================================
def solicitar_csv(url: str, logger: logging.Logger) -> bytes:
    """Solicita el archivo CSV por GET y devuelve sus bytes sin editarlos.

    - Registra el codigo de respuesta HTTP, el intento y los bytes recibidos.
    - No reintenta los errores permanentes 400, 401, 403 y 404.
    - Ante errores de conexion u otros HTTP, realiza hasta tres intentos.
    - El formato del contenido se verifica posteriormente, en memoria.
    """
    encabezados = {
        "User-Agent": USER_AGENT,
        "Accept": "text/csv,text/plain;q=0.9,*/*;q=0.8",
    }
    ultimo_error: Exception | None = None

    for intento in range(1, MAX_INTENTOS + 1):
        try:
            # requests.get() consulta la API; la respuesta aun no se guarda.
            respuesta = requests.get(url, headers=encabezados, timeout=TIMEOUT_SEG)
            logger.info(
                "GET | intento=%d | HTTP=%d | bytes=%d | url=%s",
                intento,
                respuesta.status_code,
                len(respuesta.content),
                url,
            )
            print(f"[BCRP] Intento {intento}: HTTP {respuesta.status_code}")

            # Un endpoint inexistente o un acceso prohibido no se trata como
            # un problema temporal; se informa el error de inmediato.
            if respuesta.status_code in (400, 401, 403, 404):
                raise RuntimeError(
                    f"HTTP {respuesta.status_code}: acceso o endpoint rechazado ({url})"
                )

            # Detecta otras respuestas HTTP fallidas (por ejemplo, HTTP 500).
            respuesta.raise_for_status()
            return respuesta.content  # Bytes originales, sin normalizar.

        except RuntimeError:
            raise
        except requests.RequestException as error:
            ultimo_error = error
            logger.warning("Error de conexion/HTTP, intento %d: %s", intento, error)
            if intento < MAX_INTENTOS:
                espera = 2 ** intento
                print(f"[BCRP] Reintentando en {espera} segundos...")
                time.sleep(espera)

    raise ConnectionError(
        f"No fue posible descargar el CSV tras {MAX_INTENTOS} intentos: {ultimo_error}"
    )


# ============================================================================
# 5. REVISION BASICA DE LA RESPUESTA, SIN LIMPIAR LOS ARCHIVOS
# ============================================================================
def contar_periodos_visibles(contenido: bytes) -> int:
    """Cuenta periodos mensuales en memoria antes de guardar el CSV.

    La respuesta puede separar filas mediante saltos de linea o etiquetas
    '<br>'. Se reconocen etiquetas como 'Ene18', 'Ene.18', 'Ene-2018' y
    periodos de tipo '2018-01'.

    El conteo NO implica limpiar, recalcular ni reescribir el CSV original.
    Si la respuesta parece HTML, JSON o no tiene periodos, se rechaza.
    """
    # Solo se decodifica una COPIA en memoria para revisar su contenido.
    try:
        texto = contenido.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = contenido.decode("latin-1")

    if not texto.strip():
        raise ValueError("La API devolvio un archivo vacio")

    # Previene guardar respuestas HTML o JSON como si fueran datos bancarios.
    comienzo = texto.lstrip().lower()[:1200]
    if any(marca in comienzo for marca in ("<!doctype html", "<html", "<head", "<body")):
        raise ValueError("Se recibio una pagina HTML en lugar de un CSV")
    if comienzo.startswith("{") or comienzo.startswith("["):
        raise ValueError("Se recibio JSON y no CSV; comprobar el endpoint")

    # Se identifican filas unicamente para contar periodos visibles.
    filas = re.split(r"<br\s*/?>|\r?\n", texto, flags=re.IGNORECASE)
    meses = "ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic"
    patron = re.compile(
        rf"^\s*[\"']?(?:(?:{meses})\.?\s*[-/. ]?\s*\d{{2,4}}|\d{{4}}[-/]\d{{1,2}})[\"']?\s*[,;]",
        flags=re.IGNORECASE,
    )
    cantidad = sum(bool(patron.search(fila)) for fila in filas)

    if cantidad == 0:
        raise ValueError(
            "No se reconocieron filas mensuales en la respuesta. "
            "Revisar el formato oficial antes de guardar el archivo."
        )
    return cantidad


# ============================================================================
# 6. CONSERVACION DE LA RESPUESTA ORIGINAL Y HUELLA SHA-256
# ============================================================================
def guardar_original(contenido: bytes, destino: Path) -> str:
    """Guarda el contenido original sin cambios y retorna su hash SHA-256.

    Si existe el mismo archivo y el hash coincide, se conserva tal como esta.
    Si existe con distinto contenido, se evita sobrescribirlo y se avisa.
    """
    destino.parent.mkdir(parents=True, exist_ok=True)
    huella = hashlib.sha256(contenido).hexdigest()

    if destino.exists():
        huella_previa = hashlib.sha256(destino.read_bytes()).hexdigest()
        if huella_previa != huella:
            raise FileExistsError(
                f"El original {destino.name} ya existe y la fuente lo reviso; "
                "no se sobrescribe. Registra la incidencia en el log."
            )
    else:
        destino.write_bytes(contenido)

    return huella


# ============================================================================
# 7. DESCARGA DOCUMENTADA DE CADA SERIE
# ============================================================================
def descargar_serie(
    codigo: str,
    descripcion: str,
    fecha_fin: str,
    carpeta: Path,
    logger: logging.Logger,
    modo: str,
) -> None:
    """Descarga una serie, comprueba su formato y guarda el archivo original.

    El modo 'demo' utiliza el mismo proceso, pero produce un archivo aparte
    con el sufijo '_demo' para no alterar la evidencia del periodo completo.
    """
    url = construir_endpoint(codigo, fecha_fin)
    sufijo = "_demo" if modo == "demo" else ""
    nombre = f"bcrp_original_{codigo}_{CODIGO_MATRICULA}{sufijo}.csv"
    destino = carpeta / nombre

    print(f"\n[BCRP] Serie: {codigo} | {descripcion}")
    print(f"[BCRP] Endpoint: {url}")

    # Paso 1: consultar la API y mantener la respuesta en memoria.
    contenido = solicitar_csv(url, logger)

    # Paso 2: comprobar que la respuesta tiene periodos mensuales visibles.
    cantidad = contar_periodos_visibles(contenido)

    # Paso 3: guardar los bytes originales y calcular el hash identificador.
    # Si la comprobacion anterior falla, NO se guarda un falso archivo crudo.
    hash_sha256 = guardar_original(contenido, destino)
    logger.info(
        "GUARDADO | serie=%s | periodos_detectados=%d | archivo=%s | sha256=%s",
        codigo,
        cantidad,
        destino.relative_to(RAIZ_PROYECTO),
        hash_sha256,
    )
    print(f"[BCRP] Periodos detectados: {cantidad}")
    print(f"[BCRP] Archivo original: {destino.relative_to(RAIZ_PROYECTO)}")
    print(f"[BCRP] SHA-256: {hash_sha256}")

    # Paso 4: comparar el numero visible de meses con la cobertura solicitada.
    # Este control solo genera una advertencia: no modifica los CSV.
    esperados = 3 if modo == "demo" else 96
    if cantidad != esperados:
        aviso = (
            f"La respuesta tiene {cantidad} periodos detectados y se esperaban "
            f"{esperados}. Revisar la cobertura en el script 03."
        )
        logger.warning(aviso)
        print(f"[ADVERTENCIA] {aviso}")


# ============================================================================
# 8. ORGANIZACION GENERAL DE LA EJECUCION
# ============================================================================
def ejecutar(modo: str) -> None:
    """Coordina fecha, carpeta, series y registro de inicio y finalizacion."""
    logger = configurar_log(modo)
    fecha_real = datetime.now().astimezone().isoformat(timespec="seconds")
    fecha_fin = FECHA_CORTE_DEMO if modo == "demo" else FECHA_CORTE
    carpeta = CARPETA_CRUDOS  # Ambos modos usan esta carpeta; nombre diferente.

    # Se documentan el periodo y el modo realmente utilizados.
    logger.info(
        "INICIO | fecha_real=%s | modo=%s | FECHA_INICIO=%s | FECHA_CORTE=%s",
        fecha_real,
        modo,
        FECHA_INICIO,
        fecha_fin,
    )
    print(f"Extraccion BCRP | modo={modo} | fecha real: {fecha_real}")
    print(f"Periodo solicitado: {FECHA_INICIO} a {fecha_fin}")

    CARPETA_CRUDOS.mkdir(parents=True, exist_ok=True)
    series = list(SERIES_BCRP.items())
    if modo == "demo":
        series = series[:1]  # Solo TIPMN, en archivo de demostracion separado.

    # Descarga secuencial con pausa de cortesia entre las dos series.
    for indice, (codigo, descripcion) in enumerate(series):
        if indice > 0:
            time.sleep(PAUSA_ENTRE_SERIES_SEG)
        descargar_serie(codigo, descripcion, fecha_fin, carpeta, logger, modo)

    logger.info("FIN CORRECTO | modo=%s | series_descargadas=%d", modo, len(series))
    print("\n[BCRP] Descarga finalizada. No se limpiaron los archivos originales.")


# ============================================================================
# 9. PARAMETROS DE LINEA DE COMANDOS Y PUNTO DE ENTRADA
# ============================================================================
def main() -> None:
    """Permite escoger entre el modo completo y una demostracion corta."""
    parser = argparse.ArgumentParser(
        description="Extraccion API BCRP, CSV crudos sin editar"
    )
    parser.add_argument(
        "--modo",
        choices=("completo", "demo"),
        default="completo",
        help="completo: 2018-2025 y dos series; demo: tres meses y una serie",
    )
    args = parser.parse_args()

    try:
        ejecutar(args.modo)
    except Exception as error:
        logging.getLogger("extraccion_bcrp").exception("FALLO DE EXTRACCION")
        print(f"\n[ERROR] {error}", file=sys.stderr)
        sys.exit(1)


# Este bloque evita que la descarga se ejecute al importar el archivo.
# Solo se llama a main() si ejecutamos 'python codigo/01_extraccion_api.py'.
if __name__ == "__main__":
    main()
