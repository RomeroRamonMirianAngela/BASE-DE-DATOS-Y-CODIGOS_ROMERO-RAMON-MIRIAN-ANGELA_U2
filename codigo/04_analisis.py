# Nombres y apellidos completos: ROMERO RAMON MIRIAN ANGELA
# Codigo de matricula: 2024200523C
# Tema 60: Estructura de financiamiento y costo de capital en el sistema bancario peruano
# Fecha de extraccion: 2026-11-26 (fecha historica; ejecuciones en log_ejecucion.txt)
"""Finanzas I - Unidad II | 04: analisis econometrico SBS + BCRP.

Lee SOLO datos_procesados/datos_procesados_2024200523C.csv, generado por
03_limpieza_datos.py (13 columnas originales + 12 variables calculadas).
No modifica datos. Todas las salidas van al articulo:

  Tablas (.csv y .tex)
    tabla1_descriptivos             Estadisticos descriptivos
    tabla2_estructura_bancos        Estructura del fondeo, costo y ROE por banco
    tabla3_construccion_muestra     Paneles usados en cada modelo
    tabla4_seleccion_modelo         F de efectos fijos, LM de Breusch-Pagan,
                                    Mundlak y Wooldridge (EF frente a PD)
    tabla5_modelos_costo_fondeo     Tres estimaciones del costo de fondeo
    tabla6_pruebas_diagnostico      Pruebas econometricas de los tres modelos
    tabla7_robustez                 Cinco variantes del modelo principal
    tabla8_efectos_economicos       Magnitud economica de cada coeficiente
    tabla9_modelo_roe               Costo de fondeo y rentabilidad (H6)
    tabla10_contraste_hipotesis     Decision sobre H1-H6
  Figuras (.png 300 dpi y .pdf vectorial)
    figura1_composicion_fondeo      (a) sistema en el tiempo, (b) por banco
    figura2_costo_y_tasas           Costo de fondeo, rango entre bancos y tasas BCRP
    figura3_costo_por_banco         Distribucion del costo por banco
    figura4_mapa_calor_banco_anio   Costo de fondeo por banco y ano
    figura5_correlaciones           Matriz de correlaciones con significancia
    figura6_regresion_parcial       Graficos de variable anadida del modelo principal
    figura7_comparacion_modelos     Coeficientes e IC 95 % de los tres modelos
    figura8_costo_y_roe             Relacion costo de fondeo - ROE (H6)
    figura9_diagnostico_residuos    Cuatro diagnosticos del modelo principal

Modelos del costo de fondeo (asociativos, NO causales):
  (1) Mensual, efectos fijos por banco, errores Driscoll-Kraay de 12 meses.
  (2) Anual (diciembre), efectos fijos por banco, errores agrupados por banco.
  (3) PRINCIPAL. Anual en primeras diferencias, errores agrupados por banco.
      Diciembre se usa porque el costo de fondeo es un promedio movil de
      12 meses: cada dato anual resume 12 meses distintos, sin traslape.
  Regresoras: ln(1+depositos), ln(1+adeudados), ln(1+deuda subordinada),
  ln(1+patrimonio efectivo), TIPMN y TIPMEX en promedio de 12 meses.
Modelo complementario (H6): ROE anual con efectos fijos por banco.
Las figuras no llevan titulo interno: el numero, el titulo y la nota van en
el articulo, segun APA 7.
"""

from __future__ import annotations

import math
import warnings
from datetime import datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
from matplotlib.patches import Patch
import numpy as np
import pandas as pd
from scipy import stats
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.nonparametric.smoothers_lowess import lowess
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.sandwich_covariance import S_hac_simple

warnings.filterwarnings("ignore", category=FutureWarning)

CODIGO_MATRICULA = "2024200523C"
FECHA_INICIO = "2017-01-01"
FECHA_CORTE = "2025-12-31"
RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "datos_procesados" / f"datos_procesados_{CODIGO_MATRICULA}.csv"
SALIDAS = RAIZ / "salidas"
LOG = RAIZ / "log_ejecucion.txt"

MIN_MESES = 24          # meses minimos por banco en el panel mensual
MIN_ANIOS = 4           # anos minimos por banco en el panel anual
MES_ANUAL = 12          # diciembre: cierre del promedio movil de 12 meses
DK_LAGS = 12            # ancho Driscoll-Kraay = horizonte del costo de fondeo
ALFA = 0.05
BANCOS_GRANDES = ["BCP", "BBVA", "SCOTIABANK", "INTERBANK"]

Y = "costo_fondeo"
X_FONDEO = ["ln_depositos", "ln_adeudados", "ln_deuda_subordinada", "ln_patrimonio_efectivo"]
X_TASAS = ["tipmn_ma12", "tipmex_ma12"]
X = X_FONDEO + X_TASAS
X_ESTRUCTURA = ["part_adeudados", "part_subordinada", "capitalizacion", "tamano"] + X_TASAS
X_ROE = [Y, "part_subordinada", "capitalizacion", "tamano", "tipmn_ma12"]

HIPOTESIS = [  # (hipotesis, variable, signo esperado, enunciado de H1)
    ("H1", "ln_depositos", -1, "Más depósitos se asocian con menor costo de fondeo"),
    ("H2", "ln_adeudados", 1, "Más adeudados se asocian con mayor costo de fondeo"),
    ("H3", "ln_deuda_subordinada", 1, "Más deuda subordinada se asocia con mayor costo de fondeo"),
    ("H4", "ln_patrimonio_efectivo", -1, "Más patrimonio efectivo se asocia con menor costo de fondeo"),
    ("H5", "tipmn_ma12", 1, "Una TIPMN más alta se asocia con mayor costo de fondeo"),
    ("H6", Y, -1, "Un mayor costo de fondeo se asocia con menor ROE"),
]

ETIQUETAS = {
    "costo_fondeo": "Costo de fondeo (%)",
    "roe": "ROE (%)",
    "depositos": "Depósitos (millones de S/)",
    "adeudados": "Adeudados (millones de S/)",
    "deuda_subordinada": "Deuda subordinada (millones de S/)",
    "patrimonio_efectivo": "Patrimonio efectivo (millones de S/)",
    "part_depositos": "Participación de depósitos (%)",
    "part_adeudados": "Participación de adeudados (%)",
    "part_subordinada": "Participación de deuda subordinada (%)",
    "capitalizacion": "Patrimonio efectivo / fondeo (%)",
    "tamano": "Tamaño: ln(fondeo total)",
    "ln_depositos": "ln(1 + depósitos)",
    "ln_adeudados": "ln(1 + adeudados)",
    "ln_deuda_subordinada": "ln(1 + deuda subordinada)",
    "ln_patrimonio_efectivo": "ln(1 + patrimonio efectivo)",
    "tipmn_bcrp": "TIPMN mensual (%)",
    "tipmex_bcrp": "TIPMEX mensual (%)",
    "tipmn_ma12": "TIPMN, promedio 12 meses (%)",
    "tipmex_ma12": "TIPMEX, promedio 12 meses (%)",
}
CORTAS = {
    "costo_fondeo": "Costo de fondeo", "roe": "ROE",
    "ln_depositos": "ln Depósitos", "ln_adeudados": "ln Adeudados",
    "ln_deuda_subordinada": "ln Deuda subord.", "ln_patrimonio_efectivo": "ln Patrimonio efect.",
    "tipmn_ma12": "TIPMN 12m", "tipmex_ma12": "TIPMEX 12m",
    "part_adeudados": "% Adeudados", "part_subordinada": "% Deuda subord.",
    "capitalizacion": "Capitalización", "tamano": "Tamaño",
}
MODELOS = {"m1": "(1) Mensual EF, DK12", "m2": "(2) Anual EF", "m3": "(3) Anual PD"}

# Paleta editorial, legible impresa en color y en gris
MARINO, TURQUESA, ORO, ACERO, ROJO = "#183247", "#168C90", "#C99538", "#6B8192", "#A7484B"
CELESTE, ARENA, GRILLA, TEXTO = "#D6E6EC", "#F3E7CF", "#E7EDF1", "#26323C"
COLORES_MODELO = {"m1": ACERO, "m2": TURQUESA, "m3": MARINO}
FUENTE = "Fuente: elaboración propia con datos de la SBS (balances y reportes por banco) y del BCRP (BCRPData)."


# ------------------------------------------------------------------ utilidades
def avisar(texto: str) -> None:
    linea = f"{datetime.now():%Y-%m-%d %H:%M:%S} | ANALISIS | {texto}"
    print(linea)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(linea + "\n")


def etiqueta(v: str) -> str:
    return ETIQUETAS.get(v, v)


def estrellas(p: float) -> str:
    if not np.isfinite(p):
        return ""
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.10 else ""


def guardar_tabla(df: pd.DataFrame, nombre: str, titulo: str, nota: str) -> None:
    df.to_csv(SALIDAS / f"{nombre}.csv", encoding="utf-8-sig")
    tex = df.to_latex(index=True, escape=True, na_rep="--", caption=titulo,
                      label=f"tab:{nombre}", float_format=lambda v: f"{v:.3f}")
    tex = tex.replace("\\end{tabular}", "\\end{tabular}\n\\par\\footnotesize\\textit{Nota.} "
                      + nota.replace("%", "\\%"))
    (SALIDAS / f"{nombre}.tex").write_text(tex, encoding="utf-8")
    avisar(f"TABLA | {nombre} | filas={len(df)}")


def estilo() -> None:
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 9.5, "axes.titlesize": 10.5,
        "axes.titleweight": "bold", "axes.titlecolor": MARINO, "axes.labelcolor": TEXTO,
        "axes.edgecolor": "#B9C6CF", "axes.linewidth": .8, "xtick.color": TEXTO,
        "ytick.color": TEXTO, "grid.color": GRILLA, "grid.linewidth": .8,
        "legend.frameon": False, "legend.fontsize": 8.5, "pdf.fonttype": 42,
        "savefig.facecolor": "white", "axes.facecolor": "white",
    })


def ejes(ax, eje="y") -> None:
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    if eje:
        ax.grid(axis=eje, zorder=0)
    ax.set_axisbelow(True)


def panel(ax, letra: str, titulo: str) -> None:
    """Rotulo de panel tipo revista: (a) Titulo."""
    ax.set_title(f"({letra}) {titulo}", loc="left", pad=8)


def caja_texto(ax, texto: str, x=.02, y=.97, ha="left", va="top") -> None:
    ax.text(x, y, texto, transform=ax.transAxes, ha=ha, va=va, fontsize=8, color=TEXTO,
            bbox=dict(boxstyle="round,pad=.45", fc="white", ec="#C9D4DB", lw=.8, alpha=.95))


def eje_anios(ax) -> None:
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.margins(x=.01)


def mes_es(f) -> str:
    meses = "ene feb mar abr may jun jul ago sep oct nov dic".split()
    return f"{meses[f.month - 1]}-{f.year}"


def guardar_figura(fig, nombre: str) -> None:
    fig.savefig(SALIDAS / f"{nombre}.png", dpi=300, bbox_inches="tight")
    fig.savefig(SALIDAS / f"{nombre}.pdf", bbox_inches="tight")
    plt.close(fig)
    avisar(f"FIGURA | {nombre}")


# ------------------------------------------------------------------ datos
def cargar_base() -> pd.DataFrame:
    if not DATOS.is_file():
        raise FileNotFoundError(f"Falta {DATOS.name}: ejecuta primero codigo/03_limpieza_datos.py")
    df = pd.read_csv(DATOS, encoding="utf-8-sig", parse_dates=["fecha"])
    faltan = sorted(set([Y, "roe", "part_depositos"] + X + X_ESTRUCTURA + X_ROE) - set(df.columns))
    if faltan:
        raise ValueError(f"datos_procesados no tiene {faltan}: vuelve a correr el 03 actualizado")
    if df.duplicated(["banco_id", "fecha"]).any():
        raise ValueError("Duplicados banco-fecha en datos_procesados")
    if not df["fecha"].between(pd.Timestamp(FECHA_INICIO), pd.Timestamp(FECHA_CORTE)).all():
        raise ValueError("Hay filas fuera del periodo 2017-2025")
    meses = df.groupby("banco_id")["fecha"].nunique()
    df = df[df["banco_id"].isin(meses[meses >= MIN_MESES].index)].copy()
    avisar(f"DATOS | filas={len(df)} | bancos={df['banco_id'].nunique()} | meses={df['fecha'].nunique()}")
    return df.sort_values(["banco_id", "fecha"]).reset_index(drop=True)


def panel_anual(df: pd.DataFrame, columnas: list[str]) -> pd.DataFrame:
    a = df[df["mes"] == MES_ANUAL].dropna(subset=columnas).copy()
    n = a.groupby("banco_id")["anio"].nunique()
    return a[a["banco_id"].isin(n[n >= MIN_ANIOS].index)].sort_values(["banco_id", "anio"]).reset_index(drop=True)


def diferenciar(a: pd.DataFrame, columnas: list[str]) -> pd.DataFrame:
    d = a.copy()
    for c in columnas:
        d["d_" + c] = d.groupby("banco_id")[c].diff()
    consecutivo = d.groupby("banco_id")["anio"].diff().eq(1)
    return d[consecutivo].dropna(subset=["d_" + c for c in columnas]).reset_index(drop=True)


def r2_within(modelo, datos, y) -> float:
    yw = datos[y] - datos.groupby("banco_id")[y].transform("mean")
    return float(1 - (modelo.resid ** 2).sum() / (yw ** 2).sum())


# ------------------------------------------------------------------ errores estandar
def errores_dk(modelo, datos: pd.DataFrame, lags: int = DK_LAGS):
    """Driscoll-Kraay (Bartlett) sobre scores sumados por mes calendario.

    Un mes sin observaciones aporta un vector cero: no se inventan filas."""
    inicio = pd.Timestamp(FECHA_INICIO)
    t = (datos["fecha"].dt.year * 12 + datos["fecha"].dt.month - (inicio.year * 12 + inicio.month)).to_numpy()
    total = int(t.max()) + 1
    t_obs = len(np.unique(t))
    X_ = np.asarray(modelo.model.exog, float)
    e = np.asarray(modelo.resid, float)
    n, k = X_.shape
    scores = np.zeros((total, k))
    np.add.at(scores, t, X_ * e[:, None])
    pan = np.asarray(modelo.normalized_cov_params, float)
    v = pan @ S_hac_simple(scores, nlags=lags) @ pan * (t_obs / (t_obs - 1)) * ((n - 1) / (n - k))
    ee = np.sqrt(np.diag(v))
    p = 2 * stats.t.sf(np.abs(modelo.params.to_numpy() / ee), t_obs - 1)
    return pd.Series(ee, modelo.params.index), pd.Series(p, modelo.params.index), t_obs - 1


def errores_cluster(modelo, grupos):
    r = modelo.get_robustcov_results(cov_type="cluster", groups=grupos, use_correction=True)
    g = len(np.unique(grupos)) - 1
    return pd.Series(r.bse, modelo.params.index), pd.Series(r.pvalues, modelo.params.index), g


# ------------------------------------------------------------------ pruebas
def ar1_residuos(datos, e, unidad, periodo):
    r = pd.DataFrame({"u": datos[unidad].to_numpy(), "t": datos[periodo].to_numpy(),
                      "e": np.asarray(e)}).sort_values(["u", "t"])
    r["e1"] = r.groupby("u")["e"].shift()
    r.loc[r.groupby("u")["t"].diff().ne(1), "e1"] = np.nan
    z = r.dropna()
    ar = sm.OLS(z["e"].to_numpy(), z[["e1"]].to_numpy()).fit(
        cov_type="cluster", cov_kwds={"groups": z["u"].to_numpy()})
    gl = z["u"].nunique() - 1
    return float(ar.params[0]), float(ar.bse[0]), gl, r, z


def cd_pesaran(r: pd.DataFrame) -> tuple[float, float]:
    tabla = r.pivot(index="t", columns="u", values="e")
    suma, pares = 0.0, 0
    cols = list(tabla.columns)
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            q = tabla[[cols[i], cols[j]]].dropna()
            if len(q) >= 4 and q.iloc[:, 0].std() > 0 and q.iloc[:, 1].std() > 0:
                suma += math.sqrt(len(q)) * q.iloc[:, 0].corr(q.iloc[:, 1])
                pares += 1
    cd = suma / math.sqrt(pares)
    return cd, 2 * stats.norm.sf(abs(cd))


def pruebas(datos, modelo, xs, unidad, periodo, efectos_fijos):
    """VIF, Breusch-Pagan, AR(1) de residuos, CD de Pesaran y Jarque-Bera."""
    e = np.asarray(modelo.resid)
    Z = datos[xs].astype(float)
    if efectos_fijos:
        Z = Z - datos.groupby(unidad)[xs].transform("mean")
    vif = max(variance_inflation_factor(Z.to_numpy(), i) for i in range(len(xs)))
    aux = sm.OLS(e ** 2, sm.add_constant(datos[xs].astype(float))).fit()
    bp_lm = len(e) * aux.rsquared
    rho, ee, gl, r, _ = ar1_residuos(datos, e, unidad, periodo)
    cd, cd_p = cd_pesaran(r)
    jb = stats.jarque_bera(e)
    return {"VIF máximo": vif, "Breusch-Pagan (LM)": bp_lm,
            "Breusch-Pagan p": stats.chi2.sf(bp_lm, len(xs)),
            "AR(1) residuos (rho)": rho, "AR(1) p": 2 * stats.t.sf(abs(rho / ee), gl),
            "CD de Pesaran": cd, "CD p": cd_p,
            "Jarque-Bera": jb.statistic, "Jarque-Bera p": jb.pvalue}


def pasa_todo(q: dict) -> bool:
    return (q["VIF máximo"] < 5 and q["Breusch-Pagan p"] > ALFA and q["AR(1) p"] > ALFA
            and q["CD p"] > ALFA and q["Jarque-Bera p"] > ALFA)


# ------------------------------------------------------------------ modelos
def ajustar_pd(a: pd.DataFrame, xs: list[str]) -> dict:
    d = diferenciar(a, [Y] + xs)
    dx = ["d_" + c for c in xs]
    fd = smf.ols(f"d_{Y} ~ " + " + ".join(dx), data=d).fit()
    ee, p, gl = errores_cluster(fd, d["banco_id"])
    return dict(modelo=fd, datos=d, xs=dx, base=xs, ee=ee, p=p, gl=gl, r2=fd.rsquared)


def estimar_costo(df: pd.DataFrame) -> dict:
    res = {}
    # (1) mensual con efectos fijos por banco y Driscoll-Kraay de 12 meses
    m = df.dropna(subset=[Y] + X).copy()
    m["periodo"] = m["fecha"].dt.year * 12 + m["fecha"].dt.month
    fe = smf.ols(f"{Y} ~ " + " + ".join(X) + " + C(banco_id)", data=m).fit()
    ee, p, gl = errores_dk(fe, m)
    res["m1"] = dict(modelo=fe, datos=m, xs=X, ee=ee, p=p, gl=gl, r2=r2_within(fe, m, Y),
                     prueba=pruebas(m, fe, X, "banco_id", "periodo", True), error="Driscoll-Kraay (12)")
    # (2) anual (diciembre) con efectos fijos por banco
    a = panel_anual(df, [Y] + X)
    fe_a = smf.ols(f"{Y} ~ " + " + ".join(X) + " + C(banco_id)", data=a).fit()
    ee, p, gl = errores_cluster(fe_a, a["banco_id"])
    res["m2"] = dict(modelo=fe_a, datos=a, xs=X, ee=ee, p=p, gl=gl, r2=r2_within(fe_a, a, Y),
                     prueba=pruebas(a, fe_a, X, "banco_id", "anio", True), error="Agrupados por banco")
    # (3) anual en primeras diferencias: PRINCIPAL
    r3 = ajustar_pd(a, X)
    r3["prueba"] = pruebas(r3["datos"], r3["modelo"], r3["xs"], "banco_id", "anio", False)
    r3["error"] = "Agrupados por banco"
    res["m3"] = r3
    res["anual"] = a
    for k in ("m1", "m2", "m3"):
        r = res[k]
        avisar(f"MODELO | {MODELOS[k]} | n={int(r['modelo'].nobs)} | bancos={r['datos']['banco_id'].nunique()}")
    return res


def seleccion_modelo(res: dict) -> pd.DataFrame:
    """Pruebas para elegir entre MCO agrupado, efectos aleatorios, EF y PD."""
    a = res["anual"]
    fe = res["m2"]["modelo"]
    pool = smf.ols(f"{Y} ~ " + " + ".join(X), data=a).fit()
    n, nb, k = len(a), a["banco_id"].nunique(), len(X)
    # F de efectos fijos: H0 = todos los interceptos por banco son iguales
    f = ((pool.ssr - fe.ssr) / (nb - 1)) / (fe.ssr / (n - k - nb))
    f_p = stats.f.sf(f, nb - 1, n - k - nb)
    # LM de Breusch-Pagan (Baltagi-Li, panel no balanceado): H0 = var(efecto banco) = 0
    e = pool.resid
    suma_i = e.groupby(a["banco_id"]).sum()
    ti = a.groupby("banco_id").size()
    lm = (n ** 2 / (2 * ((ti ** 2).sum() - n))) * ((suma_i ** 2).sum() / (e ** 2).sum() - 1) ** 2
    lm_p = stats.chi2.sf(lm, 1)
    # Mundlak: H0 = efectos aleatorios suficientes (alternativa robusta a Hausman)
    mm = a.copy()
    medias = []
    for c in X_FONDEO:
        mm["media_" + c] = mm.groupby("banco_id")[c].transform("mean")
        medias.append("media_" + c)
    mu = smf.ols(f"{Y} ~ " + " + ".join(X + medias), data=mm).fit(
        cov_type="cluster", cov_kwds={"groups": mm["banco_id"]})
    w = mu.wald_test(", ".join(f"{c} = 0" for c in medias), use_f=True, scalar=True)
    # Wooldridge (2010, sec. 10.6): en PD, rho = -0.5 si EF es eficiente
    r3 = res["m3"]
    rho, ee, gl, _, _ = ar1_residuos(r3["datos"], r3["modelo"].resid, "banco_id", "anio")
    t_w = (rho + 0.5) / ee
    w_p = 2 * stats.t.sf(abs(t_w), gl)
    filas = [
        ("F de efectos fijos", "Los interceptos de todos los bancos son iguales (MCO agrupado basta)",
         f, f_p, "Se rechaza H0: hay diferencias propias de cada banco" if f_p < ALFA else "No se rechaza H0"),
        ("LM de Breusch-Pagan", "La varianza del efecto banco es cero (MCO agrupado basta)",
         lm, lm_p, "Se rechaza H0: conviene un modelo de panel" if lm_p < ALFA else "No se rechaza H0"),
        ("Mundlak (alternativa a Hausman)", "Efectos aleatorios son suficientes",
         float(w.statistic), float(w.pvalue),
         "Se rechaza H0: se prefieren efectos fijos" if w.pvalue < ALFA else "No se rechaza H0"),
        ("Wooldridge: EF frente a PD", "Errores en niveles sin autocorrelación (rho de PD = -0.5)",
         t_w, w_p, f"Se rechaza H0 (rho = {rho:.2f}): se prefieren primeras diferencias"
         if w_p < ALFA else "No se rechaza H0: EF es eficiente"),
    ]
    t = pd.DataFrame(filas, columns=["Prueba", "Hipótesis nula", "Estadístico", "Valor p", "Conclusión"])
    return t.set_index("Prueba")


def robustez(df: pd.DataFrame, res: dict) -> dict:
    a = res["anual"]
    salida = {"Base": res["m3"]}
    salida["Sin 2020-2021"] = ajustar_pd(a, X)
    d = salida["Sin 2020-2021"]["datos"]
    sub = d[~d["anio"].isin([2020, 2021])]
    fd = smf.ols(f"d_{Y} ~ " + " + ".join("d_" + c for c in X), data=sub).fit()
    ee, p, gl = errores_cluster(fd, sub["banco_id"])
    salida["Sin 2020-2021"] = dict(modelo=fd, datos=sub, xs=["d_" + c for c in X], base=X,
                                   ee=ee, p=p, gl=gl, r2=fd.rsquared)
    salida["Sin 4 bancos grandes"] = ajustar_pd(a[~a["banco_id"].isin(BANCOS_GRANDES)], X)
    salida["Solo TIPMN"] = ajustar_pd(a, X_FONDEO + ["tipmn_ma12"])
    a_est = panel_anual(df, [Y] + X_ESTRUCTURA)
    salida["Estructura (%)"] = ajustar_pd(a_est, X_ESTRUCTURA)
    for k, r in salida.items():
        avisar(f"ROBUSTEZ | {k} | n={int(r['modelo'].nobs)}")
    return salida


def estimar_roe(df: pd.DataFrame) -> dict:
    a = panel_anual(df, ["roe"] + X_ROE)
    fe = smf.ols("roe ~ " + " + ".join(X_ROE) + " + C(banco_id)", data=a).fit()
    ee, p, gl = errores_cluster(fe, a["banco_id"])
    avisar(f"MODELO | ROE anual EF | n={int(fe.nobs)} | bancos={a['banco_id'].nunique()}")
    return dict(modelo=fe, datos=a, ee=ee, p=p, gl=gl, r2=r2_within(fe, a, "roe"),
                prueba=pruebas(a, fe, X_ROE, "banco_id", "anio", True))


# ------------------------------------------------------------------ tablas
def tabla_descriptivos(df: pd.DataFrame) -> None:
    v = [Y, "roe", "depositos", "adeudados", "deuda_subordinada", "patrimonio_efectivo",
         "part_depositos", "part_adeudados", "part_subordinada", "capitalizacion",
         "tipmn_bcrp", "tipmex_bcrp", "tipmn_ma12", "tipmex_ma12"]
    d = df[v].copy()
    for c in ("depositos", "adeudados", "deuda_subordinada", "patrimonio_efectivo"):
        d[c] = d[c] / 1000  # miles -> millones de soles
    t = d.describe().T[["count", "mean", "std", "min", "50%", "max"]]
    t.columns = ["N", "Media", "D.E.", "Mínimo", "Mediana", "Máximo"]
    t["N"] = t["N"].astype(int)
    t["Asimetría"] = d.skew()
    t.index = [etiqueta(c) for c in t.index]
    guardar_tabla(t.round(3), "tabla1_descriptivos", "Estadísticos descriptivos del panel banco-mes",
                  f"Panel mensual 2017-2025 de {df['banco_id'].nunique()} bancos. Los montos están en millones "
                  "de soles. Los promedios de 12 meses de TIPMN y TIPMEX empiezan en diciembre de 2017. " + FUENTE)


def tabla_bancos(df: pd.DataFrame) -> None:
    t = df.groupby("banco_id").agg(
        meses=("fecha", "nunique"), costo=(Y, "mean"), dep=("part_depositos", "mean"),
        ade=("part_adeudados", "mean"), sub=("part_subordinada", "mean"),
        cap=("capitalizacion", "mean"), roe=("roe", "mean"))
    t = t.sort_values("costo")
    t.columns = ["Meses", "Costo de fondeo (%)", "Depósitos (%)", "Adeudados (%)",
                 "Deuda subordinada (%)", "PE / fondeo (%)", "ROE (%)"]
    t.index.name = "Banco"
    guardar_tabla(t.round(2), "tabla2_estructura_bancos",
                  "Estructura del fondeo, costo y rentabilidad promedio por banco",
                  "Promedios mensuales 2017-2025, ordenados de menor a mayor costo de fondeo. Las "
                  "participaciones se calculan sobre el fondeo total (depósitos + adeudados + deuda "
                  "subordinada). " + FUENTE)


def tabla_muestra(df: pd.DataFrame, res: dict, roe: dict) -> None:
    def fila(nombre, datos, periodo, regla):
        p = datos[periodo]
        if periodo == "fecha":
            rango = f"{p.min():%Y-%m} a {p.max():%Y-%m}"
        else:
            rango = f"{int(p.min())} a {int(p.max())}"
        return {"Panel": nombre, "Observaciones": len(datos), "Bancos": datos["banco_id"].nunique(),
                "Periodos": p.nunique(), "Rango": rango, "Regla de selección": regla}
    filas = [
        fila("Base banco-mes", df, "fecha", f"Bancos con al menos {MIN_MESES} meses completos"),
        fila(MODELOS["m1"], res["m1"]["datos"], "fecha", "Además, TIPMN y TIPMEX de 12 meses disponibles"),
        fila(MODELOS["m2"], res["m2"]["datos"], "anio", f"Diciembre de cada año; bancos con al menos {MIN_ANIOS} años"),
        fila(MODELOS["m3"], res["m3"]["datos"], "anio", "Cambio entre dos diciembres consecutivos"),
        fila("Modelo de ROE", roe["datos"], "anio", "Diciembre de cada año con ROE disponible"),
    ]
    t = pd.DataFrame(filas).set_index("Panel")
    guardar_tabla(t, "tabla3_construccion_muestra", "Construcción de los paneles de estimación",
                  "En el modelo (3), el primer año de cada banco se pierde al calcular la diferencia. "
                  + FUENTE)


def tabla_seleccion(sel: pd.DataFrame) -> None:
    t = sel.copy()
    t["Estadístico"] = t["Estadístico"].round(3)
    t["Valor p"] = t["Valor p"].round(4)
    guardar_tabla(t, "tabla4_seleccion_modelo", "Pruebas para elegir el estimador de panel",
                  "Las tres primeras pruebas usan el panel anual en niveles (modelo 2). La prueba de "
                  "Wooldridge usa los residuos del modelo en primeras diferencias (modelo 3). " + FUENTE)


def celda(b, ee, p):
    return f"{b:.4f}{estrellas(p)}", f"({ee:.4f})"


def tabla_modelos(res: dict) -> None:
    filas = {}
    for var in X:
        fb, fe = {}, {}
        for k in ("m1", "m2", "m3"):
            r = res[k]
            nombre = var if k != "m3" else "d_" + var
            fb[MODELOS[k]], fe[MODELOS[k]] = celda(r["modelo"].params[nombre], r["ee"][nombre], r["p"][nombre])
        filas[etiqueta(var)] = fb
        filas[etiqueta(var) + " EE"] = fe
    filas["Efectos fijos por banco"] = {MODELOS["m1"]: "Sí", MODELOS["m2"]: "Sí", MODELOS["m3"]: "Eliminados (PD)"}
    filas["Errores estándar"] = {MODELOS[k]: res[k]["error"] for k in ("m1", "m2", "m3")}
    filas["Observaciones"] = {MODELOS[k]: str(int(res[k]["modelo"].nobs)) for k in ("m1", "m2", "m3")}
    filas["Bancos"] = {MODELOS[k]: str(res[k]["datos"]["banco_id"].nunique()) for k in ("m1", "m2", "m3")}
    filas["R² (within en EF)"] = {MODELOS[k]: f"{res[k]['r2']:.3f}" for k in ("m1", "m2", "m3")}
    t = pd.DataFrame(filas).T
    t.index = [i if not i.endswith(" EE") else "" for i in t.index]
    t.index.name = "Variable"
    guardar_tabla(t, "tabla5_modelos_costo_fondeo",
                  "Estructura de financiamiento y costo de fondeo: tres estimaciones",
                  "Variable dependiente: costo de fondeo (%), en niveles en (1) y (2) y en primera diferencia "
                  "anual en (3). EF = efectos fijos por banco; PD = primeras diferencias; DK12 = Driscoll-Kraay "
                  "con 12 rezagos. Errores estándar entre paréntesis. *** p<0.01, ** p<0.05, * p<0.10. "
                  "El modelo (3) es el principal. " + FUENTE)


def tabla_pruebas(res: dict) -> None:
    t = pd.DataFrame({MODELOS[k]: res[k]["prueba"] for k in ("m1", "m2", "m3")}).round(3).astype(object)
    t.loc["¿Pasa todas las pruebas?"] = ["Sí" if pasa_todo(res[k]["prueba"]) else "No" for k in ("m1", "m2", "m3")]
    t.index.name = "Prueba"
    guardar_tabla(t, "tabla6_pruebas_diagnostico", "Pruebas econométricas de los modelos del costo de fondeo",
                  "VIF: multicolinealidad (problema si es mayor que 5). Breusch-Pagan: H0 homocedasticidad. "
                  "AR(1): H0 residuos sin autocorrelación (rho = 0). CD de Pesaran: H0 sin dependencia entre "
                  "bancos. Jarque-Bera: H0 normalidad. Se acepta cada prueba con p > 0.05. " + FUENTE)


def tabla_robustez(rob: dict) -> None:
    variables = X + [v for v in X_ESTRUCTURA if v not in X]
    filas = {}
    for var in variables:
        fb, fe = {}, {}
        for nombre, r in rob.items():
            dv = "d_" + var
            if dv in r["modelo"].params.index:
                fb[nombre], fe[nombre] = celda(r["modelo"].params[dv], r["ee"][dv], r["p"][dv])
            else:
                fb[nombre], fe[nombre] = "", ""
        filas["Δ " + etiqueta(var)] = fb
        filas["Δ " + etiqueta(var) + " EE"] = fe
    filas["Observaciones"] = {k: str(int(r["modelo"].nobs)) for k, r in rob.items()}
    filas["Bancos"] = {k: str(r["datos"]["banco_id"].nunique()) for k, r in rob.items()}
    filas["R²"] = {k: f"{r['r2']:.3f}" for k, r in rob.items()}
    t = pd.DataFrame(filas).T
    t = t.loc[~(t == "").all(axis=1)]
    t.index = [i if not i.endswith(" EE") else "" for i in t.index]
    t.index.name = "Variable"
    guardar_tabla(t, "tabla7_robustez", "Robustez del modelo principal en primeras diferencias",
                  "Todas las columnas son paneles anuales en primeras diferencias con errores agrupados por "
                  "banco. Sin 2020-2021 excluye los cambios de los años de la pandemia. Sin 4 bancos grandes "
                  f"excluye a {', '.join(BANCOS_GRANDES)}. Estructura (%) reemplaza los montos por "
                  "participaciones, capitalización y tamaño. *** p<0.01, ** p<0.05, * p<0.10. " + FUENTE)


def tabla_efectos(res: dict) -> None:
    r = res["m3"]
    d = r["datos"]
    filas = []
    de_y = d["d_" + Y].std()
    for var in X:
        dv = "d_" + var
        b, de = r["modelo"].params[dv], d[dv].std()
        filas.append({"Variable": "Δ " + etiqueta(var), "Coeficiente": b, "Valor p": r["p"][dv],
                      "D.E. del cambio anual": de, "Efecto de 1 D.E. (pp de costo)": b * de,
                      "Beta estandarizado": b * de / de_y})
    t = pd.DataFrame(filas).set_index("Variable").round(4)
    guardar_tabla(t, "tabla8_efectos_economicos", "Magnitud económica de los coeficientes del modelo principal",
                  "El efecto de 1 D.E. es el cambio en el costo de fondeo (en puntos porcentuales) cuando la "
                  "variable cambia en una desviación estándar de su cambio anual típico. El beta estandarizado "
                  f"divide ese efecto entre la D.E. del cambio anual del costo de fondeo ({de_y:.3f} pp). "
                  + FUENTE)


def tabla_roe(r: dict) -> None:
    t = pd.DataFrame({"Coeficiente": r["modelo"].params[X_ROE], "Error estándar": r["ee"][X_ROE],
                      "Valor p": r["p"][X_ROE]})
    t["Sig."] = [estrellas(p) for p in t["Valor p"]]
    t.index = [etiqueta(c) for c in t.index]
    t = t.round(4).astype(object)
    q = r["prueba"]
    t.loc["Observaciones / bancos"] = [int(r["modelo"].nobs), r["datos"]["banco_id"].nunique(), "", ""]
    t.loc["R² within"] = [round(r["r2"], 3), "", "", ""]
    t.loc["AR(1) p / VIF máximo"] = [round(q["AR(1) p"], 3), round(q["VIF máximo"], 2), "", ""]
    t.loc["Breusch-Pagan p / Jarque-Bera p"] = [round(q["Breusch-Pagan p"], 3), round(q["Jarque-Bera p"], 3), "", ""]
    t.index.name = "Variable"
    guardar_tabla(t, "tabla9_modelo_roe", "Costo de fondeo y rentabilidad (ROE): panel anual con efectos fijos",
                  "Variable dependiente: ROE (%) de diciembre de cada año. Errores agrupados por banco, que son "
                  "válidos aunque haya heterocedasticidad o residuos no normales. *** p<0.01, ** p<0.05, "
                  "* p<0.10. " + FUENTE)


def tabla_hipotesis(res: dict, roe: dict, rob: dict) -> pd.DataFrame:
    filas = []
    m3 = res["m3"]
    for h, var, signo, enunciado in HIPOTESIS:
        if h == "H6":
            b, p, modelo, robusto = roe["modelo"].params[var], roe["p"][var], "ROE anual EF", "--"
        else:
            b, p, modelo = m3["modelo"].params["d_" + var], m3["p"]["d_" + var], MODELOS["m3"]
            def acepta(bb, pp):  # 1 acepta H1, -1 signo contrario, 0 no significativo
                return 0 if pp >= ALFA else (1 if np.sign(bb) == signo else -1)
            base = acepta(b, p)
            otras = [acepta(r["modelo"].params["d_" + var], r["p"]["d_" + var])
                     for k, r in rob.items() if k != "Base" and "d_" + var in r["p"].index]
            robusto = f"{sum(o == base for o in otras)} de {len(otras)}"
        if p < ALFA and np.sign(b) == signo:
            decision = "Se rechaza H0: se acepta H1"
        elif p < ALFA:
            decision = "No se acepta H1: significativo con signo contrario"
        else:
            decision = "No se rechaza H0"
        filas.append({"Hipótesis": h, "Enunciado (H1)": enunciado, "Signo esperado": "+" if signo > 0 else "-",
                      "Coeficiente": round(b, 4), "Valor p": round(p, 4), "Modelo": modelo,
                      "Se mantiene en robustez": robusto, "Decisión": decision})
    t = pd.DataFrame(filas).set_index("Hipótesis")
    guardar_tabla(t, "tabla10_contraste_hipotesis", "Contraste de las hipótesis de investigación",
                  "Se rechaza H0 cuando p < 0.05 y el signo coincide con el esperado. H1-H5 usan el modelo "
                  "principal (3) y H6 el modelo de ROE (tabla 9). La columna de robustez cuenta en cuántas "
                  "variantes de la tabla 7 que incluyen la variable se mantiene la misma conclusión. " + FUENTE)
    return t


# ------------------------------------------------------------------ figuras
def figura_composicion(df: pd.DataFrame) -> None:
    g = df.groupby("fecha")[["depositos", "adeudados", "deuda_subordinada"]].sum()
    s = 100 * g.div(g.sum(axis=1), axis=0)
    fig = plt.figure(figsize=(11, 5.2))
    gs = fig.add_gridspec(1, 2, width_ratios=[1.45, 1], wspace=.32)
    ax = fig.add_subplot(gs[0])
    colores = [MARINO, TURQUESA, ORO]
    nombres = ["Depósitos", "Adeudados", "Deuda subordinada"]
    ax.stackplot(s.index, s.T.values, labels=nombres, colors=colores, alpha=.95, lw=0)
    ax.set_ylim(70, 100)
    ax.set_ylabel("% del fondeo total del sistema")
    eje_anios(ax)
    ini, fin = s.index[0], s.index[-1]
    caja_texto(ax, "Eje vertical desde 70 %.  Ene-2017 → dic-2025:\n"
                   f"Depósitos: {s['depositos'].iloc[0]:.1f} % → {s['depositos'].iloc[-1]:.1f} %\n"
                   f"Adeudados: {s['adeudados'].iloc[0]:.1f} % → {s['adeudados'].iloc[-1]:.1f} %\n"
                   f"Deuda subordinada: {s['deuda_subordinada'].iloc[0]:.1f} % → "
                   f"{s['deuda_subordinada'].iloc[-1]:.1f} %", y=.04, va="bottom")
    panel(ax, "a", "Sistema bancario, mes a mes")
    ejes(ax)

    ax2 = fig.add_subplot(gs[1])
    b = df.groupby("banco_id")[["part_depositos", "part_adeudados", "part_subordinada"]].mean()
    b = b.sort_values("part_depositos")
    izq = np.zeros(len(b))
    for col, colr in zip(b.columns, colores):
        ax2.barh(b.index, b[col], left=izq, color=colr, height=.72, edgecolor="white", lw=.4)
        izq += b[col].to_numpy()
    for i, v in enumerate(b["part_depositos"]):
        ax2.text(v - 1.5, i, f"{v:.0f}", ha="right", va="center", color="white", fontsize=7.2)
    ax2.set_xlim(0, 100)
    ax2.set_xlabel("% promedio del fondeo de cada banco (2017-2025)")
    ax2.tick_params(axis="y", labelsize=8)
    panel(ax2, "b", "Por banco")
    ejes(ax2, "x")
    fig.legend(handles=[Patch(color=c, label=n) for c, n in zip(colores, nombres)],
               loc="lower center", bbox_to_anchor=(.5, -.06), ncol=3)
    guardar_figura(fig, "figura1_composicion_fondeo")


def figura_costo_tasas(df: pd.DataFrame) -> None:
    g = df.groupby("fecha").agg(media=(Y, "mean"), p25=(Y, lambda x: x.quantile(.25)),
                                p75=(Y, lambda x: x.quantile(.75)),
                                tipmn=("tipmn_ma12", "first"), tipmex=("tipmex_ma12", "first"),
                                tipmn_m=("tipmn_bcrp", "first"))
    fig, ax = plt.subplots(figsize=(11, 5.2))
    tramos = [("2020-03-01", "2021-07-31", "Pandemia: tasa de referencia\nBCRP en mínimo histórico", "#EEF3F6"),
              ("2021-08-01", "2023-01-31", "Subidas de la tasa\nde referencia BCRP", ARENA),
              ("2023-09-01", "2025-12-31", "Recortes de la tasa\nde referencia BCRP", "#E3F1EF")]
    for (ini, fin, texto, color), y in zip(tramos, (7.25, 6.6, 7.25)):
        ax.axvspan(pd.Timestamp(ini), pd.Timestamp(fin), color=color, zorder=0, lw=0)
        ax.text(pd.Timestamp(ini) + pd.Timedelta(days=12), y, texto, fontsize=7.6, color=ACERO, va="top")
    ax.fill_between(g.index, g["p25"], g["p75"], color=CELESTE, alpha=.9, lw=0,
                    label="Rango entre bancos (percentil 25 a 75)")
    ax.plot(g.index, g["media"], color=MARINO, lw=2.6, label="Costo de fondeo promedio de los bancos")
    ax.plot(g.index, g["tipmn_m"], color=TURQUESA, lw=.9, alpha=.55, label="TIPMN mensual")
    ax.plot(g.index, g["tipmn"], color=TURQUESA, lw=2.1, label="TIPMN, promedio 12 meses")
    ax.plot(g.index, g["tipmex"], color=ORO, lw=2.1, ls="--", label="TIPMEX, promedio 12 meses")
    pico = g["media"].idxmax()
    ax.annotate(f"Máximo: {g.loc[pico, 'media']:.2f} %\n({mes_es(pico)})", (pico, g.loc[pico, "media"]),
                xytext=(35, 18), textcoords="offset points", fontsize=8.2, color=MARINO,
                arrowprops=dict(arrowstyle="-|>", color=MARINO, lw=.9))
    valle = g.loc["2020":"2023", "media"].idxmin()
    ax.annotate(f"Mínimo: {g.loc[valle, 'media']:.2f} %\n({mes_es(valle)})", (valle, g.loc[valle, "media"]),
                xytext=(-95, -30), textcoords="offset points", fontsize=8.2, color=MARINO,
                arrowprops=dict(arrowstyle="-|>", color=MARINO, lw=.9))
    c = g[["media", "tipmn"]].dropna().corr().iloc[0, 1]
    caja_texto(ax, f"Correlación costo de fondeo - TIPMN 12m: {c:.2f}", x=.985, y=.03, ha="right", va="bottom")
    ax.set_ylim(0, 7.4)
    ax.set_ylabel("Porcentaje anual (%)")
    eje_anios(ax)
    ax.legend(loc="lower center", bbox_to_anchor=(.5, 1.0), ncol=3)
    ejes(ax)
    guardar_figura(fig, "figura2_costo_y_tasas")


def figura_bancos(df: pd.DataFrame) -> None:
    orden = df.groupby("banco_id")[Y].median().sort_values().index
    fig, ax = plt.subplots(figsize=(10, 6.4))
    rng = np.random.default_rng(int(CODIGO_MATRICULA[:8]))
    for i, b in enumerate(orden, start=1):
        x = df.loc[df["banco_id"] == b, Y].dropna()
        color = MARINO if b in BANCOS_GRANDES else TURQUESA
        ax.scatter(x, i + rng.uniform(-.22, .22, len(x)), s=5, color=color, alpha=.28, lw=0, zorder=2)
        ax.boxplot(x, positions=[i], orientation="horizontal", widths=.55, patch_artist=True, showfliers=False,
                   boxprops=dict(facecolor="none", edgecolor=color, lw=1.3),
                   whiskerprops=dict(color=color), capprops=dict(color=color),
                   medianprops=dict(color=ORO, lw=2.4), zorder=3)
        ax.text(df[Y].max() * 1.02, i, f"{x.median():.2f}", va="center", fontsize=7.8, color=TEXTO)
    med = df[Y].median()
    ax.axvline(med, color=ROJO, ls="--", lw=1)
    ax.text(med, len(orden) + .9, f"Mediana del sistema: {med:.2f} %", color=ROJO, fontsize=8, ha="center")
    ax.text(df[Y].max() * 1.02, len(orden) + .9, "Mediana", fontsize=7.8, color=TEXTO, fontweight="bold")
    ax.set_yticks(range(1, len(orden) + 1), orden, fontsize=8.3)
    ax.set_ylim(.3, len(orden) + 1.3)
    ax.set_xlabel("Costo de fondeo mensual (%)")
    ax.legend(handles=[Patch(fc="none", ec=MARINO, label="Bancos grandes (BCP, BBVA, Scotiabank, Interbank)"),
                       Patch(fc="none", ec=TURQUESA, label="Bancos medianos y pequeños"),
                       plt.Line2D([], [], color=ORO, lw=2.4, label="Mediana del banco")],
              loc="upper center", bbox_to_anchor=(.5, -.09), ncol=3)
    ejes(ax, "x")
    guardar_figura(fig, "figura3_costo_por_banco")


def figura_mapa_calor(df: pd.DataFrame) -> None:
    t = df.pivot_table(index="banco_id", columns="anio", values=Y, aggfunc="mean")
    t = t.loc[t.mean(axis=1).sort_values().index]
    cmap = LinearSegmentedColormap.from_list("costo", ["#F4F8FA", "#9CCBCB", TURQUESA, MARINO])
    fig, ax = plt.subplots(figsize=(10, 6.6))
    im = ax.imshow(t.values, cmap=cmap, aspect="auto")
    for i in range(t.shape[0]):
        for j in range(t.shape[1]):
            v = t.values[i, j]
            if np.isfinite(v):
                ax.text(j, i, f"{v:.1f}", ha="center", va="center", fontsize=7.6,
                        color="white" if v > np.nanpercentile(t.values, 60) else TEXTO)
            else:
                ax.text(j, i, "–", ha="center", va="center", fontsize=8, color=ACERO)
    ax.set_xticks(range(t.shape[1]), t.columns)
    ax.set_yticks(range(t.shape[0]), t.index, fontsize=8.3)
    ax.set_xticks(np.arange(-.5, t.shape[1]), minor=True)
    ax.set_yticks(np.arange(-.5, t.shape[0]), minor=True)
    ax.grid(which="minor", color="white", lw=1.4)
    ax.tick_params(which="minor", length=0)
    for lado in ax.spines.values():
        lado.set_visible(False)
    prom = df.groupby("anio")[Y].mean()
    ax2 = ax.secondary_xaxis("top")
    ax2.set_xticks(range(t.shape[1]), [f"{v:.2f}" for v in prom.reindex(t.columns)], fontsize=7.6, color=ACERO)
    ax2.set_xlabel("Promedio del sistema (%)", fontsize=8, color=ACERO)
    fig.colorbar(im, ax=ax, shrink=.85, pad=.02, label="Costo de fondeo promedio del año (%)")
    guardar_figura(fig, "figura4_mapa_calor_banco_anio")


def figura_correlaciones(df: pd.DataFrame) -> None:
    v = [Y, "roe"] + X
    d = df[v].dropna()
    c = d.corr()
    n = len(v)
    fig, ax = plt.subplots(figsize=(8.6, 7.2))
    mascara = np.triu(np.ones((n, n), bool), k=1)
    im = ax.imshow(np.ma.array(c.values, mask=mascara), cmap="RdBu_r", norm=TwoSlopeNorm(0, -1, 1))
    for i in range(n):
        for j in range(i + 1):
            r = c.iloc[i, j]
            if i == j:
                txt = "1"
            else:
                p = stats.pearsonr(d[v[i]], d[v[j]])[1]
                txt = f"{r:.2f}{estrellas(p)}"
            ax.text(j, i, txt, ha="center", va="center", fontsize=8,
                    color="white" if abs(r) > .6 else TEXTO)
    nombres = [CORTAS[x] for x in v]
    ax.set_xticks(range(n), nombres, rotation=40, ha="right", fontsize=8.5)
    ax.set_yticks(range(n), nombres, fontsize=8.5)
    for lado in ax.spines.values():
        lado.set_visible(False)
    caja_texto(ax, f"N = {len(d)} banco-mes\n*** p<0.01  ** p<0.05  * p<0.10", x=.98, y=.98, ha="right")
    fig.colorbar(im, ax=ax, shrink=.78, pad=.02, label="Correlación de Pearson")
    guardar_figura(fig, "figura5_correlaciones")


def figura_regresion_parcial(res: dict) -> None:
    """Graficos de variable anadida: la pendiente es el coeficiente del modelo (3)."""
    r = res["m3"]
    d = r["datos"]
    dx = r["xs"]
    fig, axs = plt.subplots(2, 3, figsize=(11.5, 8))
    fig.subplots_adjust(hspace=.5, wspace=.32)
    anios = d["anio"].to_numpy()
    norm = plt.Normalize(anios.min(), anios.max())
    cmap = LinearSegmentedColormap.from_list("anios", [ORO, TURQUESA, MARINO])
    for ax, v, letra in zip(axs.flat, dx, "abcdef"):
        otras = [z for z in dx if z != v]
        Z = sm.add_constant(d[otras])
        ey = sm.OLS(d["d_" + Y], Z).fit().resid
        ex = sm.OLS(d[v], Z).fit().resid
        sc = ax.scatter(ex, ey, c=anios, cmap=cmap, norm=norm, s=22, alpha=.85, edgecolor="white", lw=.4)
        b, ee, p = r["modelo"].params[v], r["ee"][v], r["p"][v]
        xs = np.linspace(ex.min(), ex.max(), 50)
        t = stats.t.ppf(.975, r["gl"])
        ax.fill_between(xs, (b - t * ee) * xs, (b + t * ee) * xs, color=CELESTE, alpha=.6, lw=0)
        ax.plot(xs, b * xs, color=ROJO if p < ALFA else ACERO, lw=2)
        ax.axhline(0, color="#B9C6CF", lw=.7)
        ax.axvline(0, color="#B9C6CF", lw=.7)
        caja_texto(ax, f"β = {b:.3f}{estrellas(p)}\np = {p:.3f}")
        panel(ax, letra, "Δ " + CORTAS[v[2:]])
        ax.set_xlabel("Cambio de la variable (residual)", fontsize=8)
        ax.set_ylabel("Δ costo de fondeo (residual)", fontsize=8)
        ejes(ax, None)
    cb = fig.colorbar(sc, ax=axs, shrink=.6, pad=.02)
    cb.set_label("Año del cambio")
    fig.text(.01, -.01, "Banda celeste: intervalo de confianza al 95 % (errores agrupados por banco). "
             "Línea roja: coeficiente significativo al 5 %.", fontsize=8, color=ACERO)
    guardar_figura(fig, "figura6_regresion_parcial")


def figura_comparacion(res: dict) -> None:
    fig, ax = plt.subplots(figsize=(10, 6))
    y0 = np.arange(len(X))[::-1] * 1.0
    desplaz = {"m1": .24, "m2": 0, "m3": -.24}
    for k in ("m1", "m2", "m3"):
        r = res[k]
        nombres = [("d_" + v) if k == "m3" else v for v in X]
        b = r["modelo"].params[nombres].to_numpy()
        ee = r["ee"][nombres].to_numpy()
        p = r["p"][nombres].to_numpy()
        t = stats.t.ppf(.975, r["gl"])
        y = y0 + desplaz[k]
        ax.hlines(y, b - t * ee, b + t * ee, color=COLORES_MODELO[k], lw=2.2, alpha=.9)
        ax.scatter(b, y, color=[COLORES_MODELO[k] if q < ALFA else "white" for q in p],
                   edgecolor=COLORES_MODELO[k], s=46, lw=1.6, zorder=3, label=MODELOS[k])
    ax.axvline(0, color=ROJO, lw=1, ls="--")
    for yy in y0[:-1] - .5:
        ax.axhline(yy, color=GRILLA, lw=.8)
    ax.set_yticks(y0, [etiqueta(v) for v in X])
    ax.set_xlabel("Coeficiente e intervalo de confianza al 95 %")
    manejadores = [plt.Line2D([], [], color=COLORES_MODELO[k], marker="o", lw=2.2, markersize=7,
                              label=MODELOS[k]) for k in ("m1", "m2", "m3")]
    ax.legend(handles=manejadores, loc="lower center", bbox_to_anchor=(.5, 1.0), ncol=3)
    caja_texto(ax, "Punto relleno: p < 0.05\nPunto vacío: no significativo", x=.01, y=.03, va="bottom")
    ejes(ax, "x")
    guardar_figura(fig, "figura7_comparacion_modelos")


def figura_roe(df: pd.DataFrame, roe: dict) -> None:
    fig = plt.figure(figsize=(11.5, 5))
    gs = fig.add_gridspec(1, 2, width_ratios=[1, 1.25], wspace=.3)
    ax = fig.add_subplot(gs[0])
    a = roe["datos"]
    controles = " + ".join([c for c in X_ROE if c != Y] + ["C(banco_id)"])
    xd = smf.ols(f"{Y} ~ {controles}", data=a).fit().resid
    yd = smf.ols(f"roe ~ {controles}", data=a).fit().resid
    grande = a["banco_id"].isin(BANCOS_GRANDES)
    ax.scatter(xd[~grande], yd[~grande], s=24, color=TURQUESA, alpha=.75, edgecolor="white", lw=.4,
               label="Medianos y pequeños")
    ax.scatter(xd[grande], yd[grande], s=26, color=MARINO, alpha=.85, edgecolor="white", lw=.4,
               label="Grandes")
    b, p = roe["modelo"].params[Y], roe["p"][Y]
    xs = np.linspace(xd.min(), xd.max(), 50)
    tcrit = stats.t.ppf(.975, roe["gl"])
    ee = roe["ee"][Y]
    ax.fill_between(xs, (b - tcrit * ee) * xs, (b + tcrit * ee) * xs, color=CELESTE, alpha=.6, lw=0)
    ax.plot(xs, b * xs, color=ROJO, lw=2)
    lim = np.percentile(yd, [1, 99])
    margen = (lim[1] - lim[0]) * .15
    fuera = int(((yd < lim[0] - margen) | (yd > lim[1] + margen)).sum())
    ax.set_ylim(lim[0] - margen, lim[1] + margen)
    texto = f"β = {b:.2f}{estrellas(p)}, p = {p:.3f}\n(pendiente = coeficiente de la tabla 9)"
    if fuera:
        texto += f"\n{fuera} punto(s) extremo(s) fuera del eje"
    caja_texto(ax, texto)
    ax.axhline(0, color="#B9C6CF", lw=.7)
    ax.axvline(0, color="#B9C6CF", lw=.7)
    ax.set_xlabel("Costo de fondeo, residual (pp)")
    ax.set_ylabel("ROE, residual (pp)")
    ax.legend(loc="lower left")
    panel(ax, "a", "Regresión parcial del ROE anual")
    ejes(ax, None)

    ax2 = fig.add_subplot(gs[1])
    g = df.groupby("fecha").agg(costo=(Y, "mean"), roe=("roe", "median"))
    ax2.plot(g.index, g["costo"], color=MARINO, lw=2.3, label="Costo de fondeo promedio (eje izq.)")
    ax2.set_ylabel("Costo de fondeo (%)", color=MARINO)
    ax3 = ax2.twinx()
    ax3.plot(g.index, g["roe"], color=ORO, lw=2.3, label="ROE mediano (eje der.)")
    ax3.set_ylabel("ROE mediano de los bancos (%)", color=ORO)
    ax3.spines["top"].set_visible(False)
    eje_anios(ax2)
    ax2.tick_params(axis="x", labelsize=8)
    lineas = ax2.get_lines() + ax3.get_lines()
    ax2.legend(lineas, [l.get_label() for l in lineas], loc="upper left")
    panel(ax2, "b", "Evolución del sistema, mes a mes")
    ejes(ax2)
    guardar_figura(fig, "figura8_costo_y_roe")


def figura_residuos(res: dict) -> None:
    r = res["m3"]
    e = np.asarray(r["modelo"].resid)
    ajustado = np.asarray(r["modelo"].fittedvalues)
    q = r["prueba"]
    fig, axs = plt.subplots(2, 2, figsize=(10.5, 8))
    a1, a2, a3, a4 = axs.flat
    a1.hist(e, bins=18, density=True, color=CELESTE, edgecolor=MARINO, lw=.6)
    xs = np.linspace(e.min() * 1.1, e.max() * 1.1, 200)
    a1.plot(xs, stats.norm.pdf(xs, e.mean(), e.std(ddof=1)), color=ORO, lw=2, label="Normal teórica")
    a1.plot(xs, stats.gaussian_kde(e)(xs), color=MARINO, lw=1.6, ls="--", label="Densidad estimada")
    a1.legend(loc="upper right")
    caja_texto(a1, f"Jarque-Bera = {q['Jarque-Bera']:.2f}\np = {q['Jarque-Bera p']:.3f}")
    a1.set_xlabel("Residuo (pp)")
    a1.set_ylabel("Densidad")
    panel(a1, "a", "Histograma de residuos")

    (osm, osr), (pend, inter, _) = stats.probplot(e, dist="norm")
    n = len(e)
    pr = (np.arange(1, n + 1) - .5) / n
    se = pend * np.sqrt(pr * (1 - pr) / n) / stats.norm.pdf(stats.norm.ppf(pr))
    linea = pend * np.asarray(osm) + inter
    a2.fill_between(osm, linea - 1.96 * se, linea + 1.96 * se, color=CELESTE, alpha=.7, lw=0, label="Banda 95 %")
    a2.scatter(osm, osr, s=14, color=MARINO, alpha=.85, zorder=3)
    a2.plot(osm, linea, color=ORO, lw=2)
    a2.legend(loc="lower right")
    a2.set_xlabel("Cuantiles teóricos (normal)")
    a2.set_ylabel("Cuantiles de los residuos")
    panel(a2, "b", "Gráfico Q-Q")

    a3.scatter(ajustado, e, s=16, color=TURQUESA, alpha=.8, edgecolor="white", lw=.3)
    suave = lowess(e, ajustado, frac=.6)
    a3.plot(suave[:, 0], suave[:, 1], color=ROJO, lw=2, label="Tendencia (LOWESS)")
    a3.axhline(0, color=ACERO, lw=.9, ls="--")
    a3.legend(loc="upper right")
    caja_texto(a3, f"Breusch-Pagan = {q['Breusch-Pagan (LM)']:.2f}\np = {q['Breusch-Pagan p']:.3f}")
    a3.set_xlabel("Valor ajustado del Δ costo de fondeo (pp)")
    a3.set_ylabel("Residuo (pp)")
    panel(a3, "c", "Residuos frente a ajustados")

    rho, _, _, _, z = ar1_residuos(r["datos"], e, "banco_id", "anio")
    a4.scatter(z["e1"], z["e"], s=16, color=MARINO, alpha=.8, edgecolor="white", lw=.3)
    xs = np.linspace(z["e1"].min(), z["e1"].max(), 50)
    a4.plot(xs, rho * xs, color=ROJO, lw=2)
    a4.axhline(0, color=ACERO, lw=.7)
    a4.axvline(0, color=ACERO, lw=.7)
    caja_texto(a4, f"AR(1): rho = {q['AR(1) residuos (rho)']:.3f}, p = {q['AR(1) p']:.3f}\n"
                   f"CD de Pesaran = {q['CD de Pesaran']:.2f}, p = {q['CD p']:.3f}")
    a4.set_xlabel("Residuo del año anterior (pp)")
    a4.set_ylabel("Residuo del año (pp)")
    panel(a4, "d", "Autocorrelación de residuos")
    for a in axs.flat:
        ejes(a, None)
    fig.tight_layout()
    guardar_figura(fig, "figura9_diagnostico_residuos")


# ------------------------------------------------------------------ principal
def main() -> None:
    SALIDAS.mkdir(parents=True, exist_ok=True)
    estilo()
    avisar("INICIO | 04_analisis.py")
    df = cargar_base()
    res = estimar_costo(df)
    sel = seleccion_modelo(res)
    rob = robustez(df, res)
    roe = estimar_roe(df)

    tabla_descriptivos(df)
    tabla_bancos(df)
    tabla_muestra(df, res, roe)
    tabla_seleccion(sel)
    tabla_modelos(res)
    tabla_pruebas(res)
    tabla_robustez(rob)
    tabla_efectos(res)
    tabla_roe(roe)
    hip = tabla_hipotesis(res, roe, rob)

    figura_composicion(df)
    figura_costo_tasas(df)
    figura_bancos(df)
    figura_mapa_calor(df)
    figura_correlaciones(df)
    figura_regresion_parcial(res)
    figura_comparacion(res)
    figura_roe(df, roe)
    figura_residuos(res)

    print("\n=== RESUMEN ===")
    print(sel[["Valor p", "Conclusión"]].to_string())
    print(hip[["Coeficiente", "Valor p", "Se mantiene en robustez", "Decisión"]].to_string())
    avisar("FIN | 10 tablas y 9 figuras en salidas/")


if __name__ == "__main__":
    main()
