# -*- coding: utf-8 -*-
"""
generador_dash_v10.py
v09: TODAS las secciones de Asistencia/Ratios/Alertas reaccionan al filtro de "Fecha".
v10: Rediseño estético (sin cambios de lógica ni de cálculo):
  - Tipografía Inter, números tabulares, jerarquía más fina
  - Paleta sobria: verde profundo #1B4332 + acento dorado #C9A227; fundos en tonos
    desaturados coherentes (HP #B3453C, BH #31678F, CAO #1B4332, Consolidado #5B6660)
  - Hero sin animaciones de rayos; sello corporativo con línea dorada
  - Tarjetas con borde 1px + sombra suave; banners de sección con gradiente sutil
  - Tabs tipo segmented control; selects y tablas refinados
  - Gráficos Chart.js con la misma paleta por fundo y barras redondeadas
v11: Fix discrepancia Horas Promedio en gráfico vs tarjeta KPI. hrs_prom_chart()
     usaba AVERAGEX por trabajador (sum horas/días, promediado) en vez de
     SUMX(HorasEfectivas)/COUNTROWS filtrado EsUltimoRegistro=1 (la medida real).
     Ahora reutiliza horas_promedio() filtrado por fundo, igual que el DAX
     corregido: CALCULATE([Horas_Promedio], NombreFundo="Hp"/"Bh"/"Cao").
Salida: Dash_Indicadores.html

Fuentes (mismas que el pbix):
  1. Carpeta con los Excel mensuales de tareo  -> tabla DataMes
  2. DataActivosCesados.xlsx (hoja DataPersonal) -> tabla DataPersonal

Uso:  python generador_dash_v09.py
Requiere: pandas, openpyxl   (pip install pandas openpyxl)
"""

import json
import unicodedata
from datetime import date
from pathlib import Path

import pandas as pd

# ═══════════════ CONFIGURACIÓN (editar rutas aquí) ═══════════════
CARPETA_DATA_MES = r"C:\Users\ricardo.lizama\OneDrive - Hass Peru S.A\10. RRHH\04. Tareo_Data\2026\Data Mes"
ARCHIVO_PERSONAL = r"C:\Users\ricardo.lizama\OneDrive - Hass Peru S.A\10. RRHH\04. Tareo_Data\2026\DataActivosCesados.xlsx"
HOJA_PERSONAL    = "DataPersonal"
ARCHIVO_SALIDA   = r"Indicadores_Asistencia_Web.html"
UMBRAL_HE        = 9.6          # horas a partir de las cuales cuenta hora extra
FUNDOS           = ["Hp", "Bh", "Cao"]
ETIQ_FUNDO       = {"Hp": "HP", "Bh": "BH", "Cao": "Cao"}
ETIQ_CHART       = {"Hp": "Hass", "Bh": "Berry", "Cao": "Cao"}
# ══════════════════════════════════════════════════════════════════


# ─────────────────────────── utilidades ───────────────────────────
def fmt_miles(v):
    """FORMAT '#,##0' estilo Perú: 3.017"""
    if v is None or pd.isna(v):
        return "-"
    return f"{int(round(v)):,}".replace(",", ".")

def fmt_pct(v, dec=0):
    if v is None or pd.isna(v):
        return "-"
    return f"{v*100:.{dec}f}".replace(".", ",") + "%"

def fmt_dec(v, dec=1):
    if v is None or pd.isna(v):
        return "-"
    return f"{v:.{dec}f}".replace(".", ",")

def guion_si_vacio(txt, valor):
    """Replica IF(ISBLANK(x) || x=0, '-', FORMAT(x,...)) del DAX"""
    if valor is None or pd.isna(valor) or valor == 0:
        return "-"
    return txt

def js_num(v, dec=2):
    """Número para JavaScript (punto decimal)."""
    if v is None or pd.isna(v):
        return "0"
    return f"{v:.{dec}f}"

def sin_tildes(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()

def norm_txt(s):
    """Normaliza texto de ORIGEN (arefinal, Dia_Semana) para comparar sin importar
    mayúsculas/tildes/espacios -- así se replica la igualdad de texto de DAX (case-insensitive)."""
    return sin_tildes(str(s)).strip().upper()


# ─────────────────────── carga de datos ───────────────────────────
CANON_FUNDO = {"HP": "Hp", "BH": "Bh", "CAO": "Cao"}
CANON_CULTIVO = {"ARANDANO": "Arandano", "PALTA": "Palta", "GENERICO": "Generico"}
CANON_AREA = {
    "COSECHA": "Cosecha", "EMPAQUE": "Empaque", "LABORES": "Labores",
    "RIEGO": "Riego", "RIEGO Y FERTILIZACION": "Riego Y Fertilizacion",
    "SANIDAD": "Sanidad", "OTROS": "Otros",
}
DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

def _canoniza(serie, mapa, avisar_como=None):
    """Mapea cada valor a su forma canónica comparando en MAYÚSCULAS/sin tildes.
    Valores que no matchean ninguna clave del mapa se conservan tal cual (y se avisan)."""
    normalizados = serie.apply(norm_txt)
    resultado = normalizados.map(mapa)
    sin_match = resultado.isna()
    if sin_match.any():
        resultado = resultado.where(~sin_match, serie)
        if avisar_como:
            no_reconocidos = sorted(serie[sin_match].astype(str).str.strip().unique())
            print(f"  [!] {avisar_como}: valores no reconocidos (se dejan tal cual): {no_reconocidos}")
    return resultado

def cargar_data_mes(carpeta):
    """Replica el query M: Folder.Files -> solo .xlsx -> Table.Combine"""
    carpeta = Path(carpeta)
    archivos = sorted(p for p in carpeta.glob("*.xlsx") if not p.name.startswith("~$"))
    if not archivos:
        raise FileNotFoundError(f"No hay .xlsx en {carpeta}")
    frames = [pd.read_excel(p, sheet_name=0) for p in archivos]
    df = pd.concat(frames, ignore_index=True)

    df["FechaTareo"] = pd.to_datetime(df["FechaTareo"]).dt.normalize()
    df["HorasEfectivas"] = pd.to_numeric(df["HorasEfectivas"], errors="coerce").fillna(0)
    df["Semana_Correcta"] = pd.to_numeric(df["Semana_Correcta"], errors="coerce").astype("Int64")
    df["DNI_Trabajador"] = df["DNI_Trabajador"].astype(str).str.strip()

    # Columnas calculadas del modelo -- canonizadas (mayúsculas/tildes no deben romper el match,
    # igual que la comparación de texto de DAX, que es case-insensitive)
    cult = df["Cultivo"].astype(str).str.strip()
    df["NombreFundo"]   = _canoniza(cult.str.split(" ").str[0], CANON_FUNDO, "NombreFundo (prefijo de Cultivo)")
    df["NombreCultivo"] = _canoniza(cult.str.split(" ", n=1).str[1].fillna("").str.strip(), CANON_CULTIVO, "NombreCultivo (sufijo de Cultivo)")
    df["arefinal"]      = _canoniza(df["arefinal"].astype(str).str.strip(), CANON_AREA, "arefinal")
    # Dia_Semana se recalcula desde FechaTareo (no depende de que el Excel lo tenga bien escrito)
    df["Dia_Semana"] = df["FechaTareo"].dt.weekday.map(dict(enumerate(DIAS_ES)))
    return df


def cargar_personal(archivo, hoja):
    dp = pd.read_excel(archivo, sheet_name=hoja)
    dp["NumeroDocumento"] = dp["NumeroDocumento"].astype(str).str.strip()
    dp = dp.drop_duplicates("NumeroDocumento", keep="first")
    dp["FechaNacimiento"] = pd.to_datetime(dp["FechaNacimiento"], errors="coerce")
    dp["FechaIngresoEmpresa"] = pd.to_datetime(dp["FechaIngresoEmpresa"], errors="coerce")
    return dp


def preparar(df, dp):
    """RELATED() del modelo + columna EsUltimoRegistro"""
    df = df.merge(
        dp[["NumeroDocumento", "NACIONALIDAD", "Sexo", "FechaIngresoEmpresa", "FechaNacimiento"]],
        left_on="DNI_Trabajador", right_on="NumeroDocumento", how="left",
    ).rename(columns={
        "NACIONALIDAD": "Nacionalidad",
        "FechaIngresoEmpresa": "FechaIngreso",
        "Sexo": "SexoP",
    })
    df["SexoP"] = df["SexoP"].fillna("").astype(str).str.strip().str.upper()

    # EsUltimoRegistro: 1 si HorasEfectivas == max del grupo (DNI, FechaTareo)
    maxh = df.groupby(["DNI_Trabajador", "FechaTareo"])["HorasEfectivas"].transform("max")
    df["EsUltimo"] = (df["HorasEfectivas"] == maxh).astype(int)
    return df


# ─────────────────── medidas (mismo DAX en pandas) ───────────────────
def asistencia_maxima(d):
    if d.empty:
        return None
    return d.groupby("FechaTareo")["DNI_Trabajador"].nunique().max()

def rotacion(df_all, sem, fundo=None):
    base = df_all if fundo is None else df_all[df_all["NombreFundo"] == fundo]
    act = set(base.loc[base["Semana_Correcta"] == sem, "DNI_Trabajador"])
    ant = set(base.loc[base["Semana_Correcta"] == sem - 1, "DNI_Trabajador"])
    if not act and not ant:
        return None
    cesados = len(ant - act)
    prom = (len(act) + len(ant)) / 2
    return cesados / prom if prom else 0

def pct_ausentismo(d):
    lunes_rows = d[d["Dia_Semana"].astype(str).str.strip() == "Lunes"]
    if lunes_rows.empty:
        return None
    lunes = lunes_rows["FechaTareo"].min()
    base = d.loc[d["FechaTareo"] == lunes, "DNI_Trabajador"].nunique()
    if not base:
        return None
    resto = d[(d["FechaTareo"] > lunes)
              & (~d["Dia_Semana"].astype(str).str.strip().isin(["Sábado", "Sabado", "Domingo"]))]
    if resto.empty:
        return 0
    pres = resto.groupby("FechaTareo")["DNI_Trabajador"].nunique()
    aus = (base - pres).clip(lower=0).max()
    return aus / base

def horas_promedio(d):
    d1 = d[d["EsUltimo"] == 1]
    return d1["HorasEfectivas"].sum() / len(d1) if len(d1) else None

def genero_m(d):
    tot = d["DNI_Trabajador"].nunique()
    if not tot:
        return None
    m = d.loc[d["SexoP"] == "M", "DNI_Trabajador"].nunique()
    return m / tot

def antiguedad_0_1(d, hoy):
    """DAX: DISTINCTCOUNT(DNI) donde MAX(FechaIngreso) con EsUltimoRegistro=1 >= TODAY()-365"""
    fi = d.loc[d["EsUltimo"] == 1].groupby("DNI_Trabajador")["FechaIngreso"].max().dropna()
    if fi.empty:
        return None
    corte = pd.Timestamp(hoy) - pd.Timedelta(days=365)
    return int((fi >= corte).sum())

def pct_cierre(d):
    tot = d["DNI_Trabajador"].nunique()
    if not tot:
        return None
    dias = d.groupby("DNI_Trabajador")["FechaTareo"].nunique()
    return (dias == 7).sum() / tot

def horas_extras(d):
    d1 = d[(d["EsUltimo"] == 1) & (d["HorasEfectivas"] > UMBRAL_HE)]
    return (d1["HorasEfectivas"] - UMBRAL_HE).sum() if len(d1) else None

def jor_sup(d):
    d1 = d[d["EsUltimo"] == 1]
    sup = d1["Supervisor"].nunique()
    return d1["DNI_Trabajador"].nunique() / sup if sup else None

def mayor_60h(d):
    if d.empty:
        return None
    tot = d.groupby("DNI_Trabajador")["HorasEfectivas"].sum()
    return int((tot > 60).sum())

def bucket_horas(d, lo, hi=None):
    d1 = d[d["EsUltimo"] == 1]
    if d1.empty:
        return 0
    tot = d1.groupby("DNI_Trabajador")["HorasEfectivas"].sum()
    return int(((tot >= lo) & (tot < hi)).sum()) if hi else int((tot >= lo).sum())

def bucket_edad(d, hoy, lo, hi=None):
    ed = hoy.year - d.groupby("DNI_Trabajador")["FechaNacimiento"].max().dt.year
    ed = ed.dropna()
    if hi:
        return int(((ed >= lo) & (ed <= hi)).sum())
    return int((ed > lo).sum())

def asist_area_cultivo(d, fundo=None):
    """Tabla área x cultivo de HTML_KPI01. Devuelve dict[area][ara|pal|con]."""
    b = d[d["EsUltimo"] == 1]
    if fundo:
        b = b[b["NombreFundo"] == fundo]
    ncult = b["NombreCultivo"].astype(str)
    es_ara = ncult.isin(["Arandano", "Generico"])
    es_pal = ncult == "Palta"
    area = b["arefinal"].astype(str)
    mapa = {
        "Cosecha": area == "Cosecha",
        "Empaque": area == "Empaque",
        "Labores": area == "Labores",
        "Riego y Fert.": area.isin(["Riego", "Riego Y Fertilizacion"]),
        "Sanidad": area == "Sanidad",
        "Otros": area == "Otros",
    }
    out = {}
    for nombre, m in mapa.items():
        out[nombre] = {
            "ara": b.loc[m & es_ara, "DNI_Trabajador"].nunique(),
            "pal": b.loc[m & es_pal, "DNI_Trabajador"].nunique(),
            "con": b.loc[m, "DNI_Trabajador"].nunique(),
        }
    out["Total"] = {
        "ara": b.loc[es_ara, "DNI_Trabajador"].nunique(),
        "pal": b.loc[es_pal, "DNI_Trabajador"].nunique(),
        "con": b["DNI_Trabajador"].nunique(),
    }
    return out

def hrs_prom_chart(d, fundo):
    """Debe coincidir EXACTO con horas_promedio(): SUMX(HorasEfectivas)/COUNTROWS,
    filtrado EsUltimoRegistro=1. Antes usaba AVERAGEX por trabajador (sum/dias, promediado),
    lo que causaba discrepancia frente al KPI de la tarjeta."""
    return horas_promedio(d[d["NombreFundo"] == fundo])

def genero_chart(d, fundo):
    b = d[(d["NombreFundo"] == fundo) & (d["SexoP"] != "")]
    tot = b["DNI_Trabajador"].nunique()
    if not tot:
        return None, None
    m = b.loc[b["SexoP"] == "M", "DNI_Trabajador"].nunique()
    f = b.loc[b["SexoP"] == "F", "DNI_Trabajador"].nunique()
    return m / tot, f / tot


# ─────────── medidas nuevas v02: Asistencia diaria / Ratios / Alertas ───────────
DIAS_ORDEN = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
AREAS_RATIO = ["Cosecha", "Labores", "Empaque", "Riego y Fert.", "Sanidad", "Otros"]
RANGOS_ANTIG = ["[0 - 1 año]", "[1 - 3 años]", "[3 - 5 años]", "[> 5 años]"]

def asistencia_diaria_fundo(d, fundo):
    """Tabla Lun-Dom por área/cultivo, para el detalle de Grupo Asistencia (HTML_KPI02)."""
    b = d[(d["NombreFundo"] == fundo) & (d["EsUltimo"] == 1)]
    if b.empty:
        return None
    lunes = b.loc[b["Dia_Semana"].astype(str).str.strip() == "Lunes", "FechaTareo"].min()
    if pd.isna(lunes):
        lunes = b["FechaTareo"].min()
    fechas = [lunes + pd.Timedelta(days=i) for i in range(7)]

    area = b["arefinal"].astype(str).str.strip()
    ncult = b["NombreCultivo"].astype(str)
    es_ara = ncult.isin(["Arandano", "Generico", "Arándano", "Genérico"])
    es_pal = ncult == "Palta"
    filas_def = {
        "Cosecha - Arándano": (area == "Cosecha") & es_ara,
        "Cosecha - Palta":    (area == "Cosecha") & es_pal,
        "Empaque":            area == "Empaque",
        "Labores":            area == "Labores",
        "Riego y Fertilizacion": area.isin(["Riego", "Riego Y Fertilizacion"]),
        "Sanidad":            area == "Sanidad",
        "Otros":              area == "Otros",
    }
    filas = {}
    for nombre, mask in filas_def.items():
        serie = [b.loc[mask & (b["FechaTareo"] == f), "DNI_Trabajador"].nunique() for f in fechas]
        filas[nombre] = serie
    total_fundo = [b.loc[b["FechaTareo"] == f, "DNI_Trabajador"].nunique() for f in fechas]
    return {"fechas": [f.strftime("%d/%m") for f in fechas], "filas": filas, "total": total_fundo}

def supervisor_por_area(d, fundo):
    """Ratio trabajador/supervisor, horas y horas extra por área (HTML_KPI09_Personal)."""
    b = d[(d["NombreFundo"] == fundo) & (d["EsUltimo"] == 1)]
    area = b["arefinal"].astype(str).str.strip()
    mapa = {
        "Cosecha": area == "Cosecha", "Labores": area == "Labores", "Empaque": area == "Empaque",
        "Riego y Fert.": area.isin(["Riego", "Riego Y Fertilizacion"]),
        "Sanidad": area == "Sanidad", "Otros": area == "Otros",
    }
    out = {}
    for nombre, m in mapa.items():
        sub = b[m]
        sup = sub["Supervisor"].nunique()
        trab = sub["DNI_Trabajador"].nunique()
        horas = sub["HorasEfectivas"].sum()
        extra = (sub.loc[sub["HorasEfectivas"] > UMBRAL_HE, "HorasEfectivas"] - UMBRAL_HE).sum()
        out[nombre] = {"sup": sup, "trab": trab, "ratio": (trab / sup) if sup else 0,
                        "horas": horas, "extra": extra}
    return out

def antiguedad_rangos(d, hoy, fundo=None):
    b = d[d["EsUltimo"] == 1]
    if fundo:
        b = b[b["NombreFundo"] == fundo]
    fi = b.groupby("DNI_Trabajador")["FechaIngreso"].max().dropna()
    tot = b["DNI_Trabajador"].nunique()
    anios = (hoy - fi.dt.date).apply(lambda x: x.days / 365.25) if len(fi) else pd.Series(dtype=float)
    out = {}
    bordes = [(0, 1), (1, 3), (3, 5), (5, None)]
    for etiq, (lo, hi) in zip(RANGOS_ANTIG, bordes):
        n = int(((anios >= lo) & (anios < hi)).sum()) if hi else int((anios >= lo).sum())
        out[etiq] = {"n": n, "pct": (n / tot) if tot else 0}
    return out

def composicion_genero(d, fundo=None):
    b = d[d["EsUltimo"] == 1]
    if fundo:
        b = b[b["NombreFundo"] == fundo]
    tot = b["DNI_Trabajador"].nunique()
    m = b.loc[b["SexoP"] == "M", "DNI_Trabajador"].nunique()
    f = b.loc[b["SexoP"] == "F", "DNI_Trabajador"].nunique()
    return {"m": m, "f": f, "pm": (m / tot) if tot else 0, "pf": (f / tot) if tot else 0, "tot": tot}

def alertas_jubilacion(d, hoy, fundo=None):
    b = d[d["EsUltimo"] == 1]
    if fundo:
        b = b[b["NombreFundo"] == fundo]
    tot = b["DNI_Trabajador"].nunique()
    r60_65 = bucket_edad(b, hoy, 60, 65)
    r65_69 = bucket_edad(b, hoy, 65, 69)
    r69 = bucket_edad(b, hoy, 69)
    alerta = r60_65 + r65_69 + r69
    return {"60_65": r60_65, "65_69": r65_69, "69": r69, "alerta": alerta,
            "p": (alerta / tot) if tot else 0, "tot": tot}


# ─────────── medidas nuevas v08: Ausentismo diario / Sin descanso / Horas altas / HE Obrero / Rango edades ───────────
def ausentismo_diario_fundo(d, fundo=None):
    """Ausentes por día (Mar-Vie) contra base del Lunes, más % (HTML_KPI02_Asistencia)."""
    b = d if fundo is None else d[d["NombreFundo"] == fundo]
    lunes_rows = b[b["Dia_Semana"].astype(str).str.strip() == "Lunes"]
    if lunes_rows.empty:
        return None
    lunes = lunes_rows["FechaTareo"].min()
    base = b.loc[b["FechaTareo"] == lunes, "DNI_Trabajador"].nunique()
    dias_out = {}
    for i, nombre in enumerate(["Martes", "Mi&eacute;rcoles", "Jueves", "Viernes"], start=1):
        fecha = lunes + pd.Timedelta(days=i)
        presentes = b.loc[b["FechaTareo"] == fecha, "DNI_Trabajador"].nunique()
        ausentes = max(base - presentes, 0)
        dias_out[nombre] = {"fecha": fecha.strftime("%d/%m"), "ausentes": ausentes,
                             "pct": (ausentes / base) if base else 0}
    max_aus = max((v["ausentes"] for v in dias_out.values()), default=0)
    return {"base": base, "dias": dias_out, "max_aus": max_aus,
            "max_pct": (max_aus / base) if base else 0}

def personal_sin_descanso(d, fundo=None):
    """# con cierre de semana completo (7 d&iacute;as tareados) vs total (HTML_KPI10_Personal)."""
    b = d[d["EsUltimo"] == 1]
    if fundo:
        b = b[b["NombreFundo"] == fundo]
    tot = b["DNI_Trabajador"].nunique()
    dias = b.groupby("DNI_Trabajador")["FechaTareo"].nunique()
    con7 = int((dias == 7).sum())
    return {"tot": tot, "con7": con7, "sin7": tot - con7, "pct": (con7 / tot) if tot else 0}

def horas_altas_detalle(d, fundo=None):
    """Trabajadores con 48-60h / 60-72h / >72h en la semana (HTML_KPI18_Horas)."""
    b = d[d["EsUltimo"] == 1]
    if fundo:
        b = b[b["NombreFundo"] == fundo]
    tot = b["DNI_Trabajador"].nunique()
    r48_60 = bucket_horas(b, 48, 60)
    r60_72 = bucket_horas(b, 60, 72)
    rm72 = bucket_horas(b, 72)
    altas = r48_60 + r60_72 + rm72
    return {"tot": tot, "48_60": r48_60, "60_72": r60_72, "m72": rm72,
            "altas": altas, "pct": (altas / tot) if tot else 0}

def he_obrero_area(d, fundo=None):
    """HE (>9.6h) de OBRERO, total y por &aacute;rea (HTML_KPI20_Horas)."""
    b = d[(d["EsUltimo"] == 1) & (d["planilla"].astype(str).str.strip().str.upper() == "OBRERO")]
    if fundo:
        b = b[b["NombreFundo"] == fundo]
    tot_obr = b["DNI_Trabajador"].nunique()
    area = b["arefinal"].astype(str).str.strip()
    mapa = {
        "Cosecha": area == "Cosecha", "Labores": area == "Labores", "Empaque": area == "Empaque",
        "Riego y Fert.": area.isin(["Riego", "Riego Y Fertilizacion"]),
        "Sanidad": area == "Sanidad", "Otros": area == "Otros",
    }
    por_area = {nombre: (b.loc[m & (b["HorasEfectivas"] > UMBRAL_HE), "HorasEfectivas"] - UMBRAL_HE).sum()
                for nombre, m in mapa.items()}
    he_total = (b.loc[b["HorasEfectivas"] > UMBRAL_HE, "HorasEfectivas"] - UMBRAL_HE).sum()
    return {"tot_obr": tot_obr, "he_total": he_total, "area": por_area}

def horas_prom_area(d, fundo=None):
    """Promedio de horas efectivas por trabajador, total y por &aacute;rea (HTML_KPI21_Horas)."""
    b = d if fundo is None else d[d["NombreFundo"] == fundo]
    area = b["arefinal"].astype(str).str.strip()
    mapa = {
        "Cosecha": area == "Cosecha", "Labores": area == "Labores", "Empaque": area == "Empaque",
        "Riego y Fert.": area.isin(["Riego", "Riego Y Fertilizacion"]),
        "Sanidad": area == "Sanidad", "Otros": area == "Otros",
    }
    def prom(sub):
        n = sub["DNI_Trabajador"].nunique()
        return (sub["HorasEfectivas"].sum() / n) if n else 0
    por_area = {nombre: prom(b[m]) for nombre, m in mapa.items()}
    return {"total": prom(b), "area": por_area}

def rango_edades_detalle(d, hoy, fundo=None):
    """Distribuci&oacute;n de edades en 7 rangos: 18-20 ... >69 (HTML_KPI22_Alertas)."""
    b = d[d["EsUltimo"] == 1]
    if fundo:
        b = b[b["NombreFundo"] == fundo]
    tot = b["DNI_Trabajador"].nunique()
    ed = hoy.year - b.groupby("DNI_Trabajador")["FechaNacimiento"].max().dt.year
    ed = ed.dropna()
    out = {}
    for etiq, lo, hi in [("18-20", 18, 20), ("20-30", 20, 30), ("30-40", 30, 40),
                          ("40-50", 40, 50), ("50-60", 50, 60)]:
        n = int(((ed >= lo) & (ed < hi)).sum())
        out[etiq] = {"n": n, "pct": (n / tot) if tot else 0}
    n60_65 = int(((ed >= 60) & (ed <= 65)).sum())
    n65_69 = int(((ed > 65) & (ed <= 69)).sum())
    n69 = int((ed > 69).sum())
    out["60-65"] = {"n": n60_65, "pct": (n60_65 / tot) if tot else 0}
    out["65-69"] = {"n": n65_69, "pct": (n65_69 / tot) if tot else 0}
    out["&gt; 69"] = {"n": n69, "pct": (n69 / tot) if tot else 0}
    return {"tot": tot, "rangos": out}


# ─────────────── cálculo de todos los KPI de una semana ───────────────
def ausentismo_dia_puntual(w, dia, fundo=None):
    """Ausentes de un día puntual vs. base del Lunes de esa semana."""
    b = w if fundo is None else w[w["NombreFundo"] == fundo]
    lunes_rows = b[b["Dia_Semana"].astype(str).str.strip() == "Lunes"]
    if lunes_rows.empty:
        return None
    lunes = lunes_rows["FechaTareo"].min()
    base = b.loc[b["FechaTareo"] == lunes, "DNI_Trabajador"].nunique()
    presentes = b.loc[b["FechaTareo"] == dia, "DNI_Trabajador"].nunique()
    ausentes = max(base - presentes, 0)
    return {"base": base, "ausentes": ausentes, "pct": (ausentes / base) if base else 0}


def kpis_semana(df, sem, hoy):
    w  = df[df["Semana_Correcta"] == sem]
    wa = df[df["Semana_Correcta"] == sem - 1]

    def por_fundo(fn, base):
        return {f: fn(base[base["NombreFundo"] == f]) for f in FUNDOS}

    k = {
        "semana": int(sem),
        "anio": int(w["FechaTareo"].dt.year.max()) if len(w) else hoy.year,
        "fecha_max": w["FechaTareo"].max().strftime("%d/%m/%Y") if len(w) else "-",
        # tarjetas KPI (consolidado)
        "asistencia": asistencia_maxima(w),
        "rotacion": rotacion(df, sem),
        "supervisores": w["Supervisor"].nunique(),
        "extranjeros": w.loc[w["Nacionalidad"].astype(str).str.upper() == "EXTRANJERA",
                             "DNI_Trabajador"].nunique(),
        "ausentismo": pct_ausentismo(w),
        "horas_prom": horas_promedio(w),
    }

    # sección Área y Cultivo
    k["area"] = {f: asist_area_cultivo(w, f) for f in FUNDOS}
    k["area"]["Con"] = asist_area_cultivo(w)
    k["area_max"] = {f: asistencia_maxima(w[w["NombreFundo"] == f]) for f in FUNDOS}
    k["area_max"]["Con"] = k["asistencia"]
    k["area_ant"] = {f: wa.loc[wa["NombreFundo"] == f, "DNI_Trabajador"].nunique() for f in FUNDOS}
    k["area_ant"]["Con"] = wa["DNI_Trabajador"].nunique()

    # sección Información por Fundo (actual y semana anterior)
    def bloque(base, con_extra=True):
        b = {
            "asist":  {f: asistencia_maxima(base[base["NombreFundo"] == f]) for f in FUNDOS},
            "sup":    {f: base.loc[base["NombreFundo"] == f, "Supervisor"].nunique() for f in FUNDOS},
            "ext":    {f: base.loc[(base["NombreFundo"] == f)
                                   & (base["Nacionalidad"].astype(str).str.upper() == "EXTRANJERA"),
                                   "DNI_Trabajador"].nunique() for f in FUNDOS},
            "aus":    por_fundo(pct_ausentismo, base),
            "gen":    por_fundo(genero_m, base),
            "hrs":    por_fundo(horas_promedio, base),
            "he":     por_fundo(horas_extras, base),
            "jorsup": por_fundo(jor_sup, base),
        }
        b["asist"]["Con"] = asistencia_maxima(base)
        b["sup"]["Con"]   = base["Supervisor"].nunique()
        b["ext"]["Con"]   = base.loc[base["Nacionalidad"].astype(str).str.upper() == "EXTRANJERA",
                                     "DNI_Trabajador"].nunique()
        b["aus"]["Con"]   = pct_ausentismo(base)
        b["gen"]["Con"]   = genero_m(base)
        b["hrs"]["Con"]   = horas_promedio(base)
        b["he"]["Con"]    = horas_extras(base)
        b["jorsup"]["Con"] = jor_sup(base)
        return b

    k["info"]  = bloque(w)
    k["infoA"] = bloque(wa)
    k["info"]["rot"]   = {f: rotacion(df, sem, f) for f in FUNDOS};     k["info"]["rot"]["Con"]   = rotacion(df, sem)
    k["infoA"]["rot"]  = {f: rotacion(df, sem - 1, f) for f in FUNDOS}; k["infoA"]["rot"]["Con"]  = rotacion(df, sem - 1)
    k["info"]["ant01"] = {f: antiguedad_0_1(w[w["NombreFundo"] == f], hoy) for f in FUNDOS}
    k["info"]["ant01"]["Con"] = antiguedad_0_1(w, hoy)
    k["info"]["cierre"] = por_fundo(pct_cierre, w); k["info"]["cierre"]["Con"] = pct_cierre(w)
    k["info"]["m60"] = {f: mayor_60h(w[w["NombreFundo"] == f]) for f in FUNDOS}
    k["info"]["m60"]["Con"] = sum(v or 0 for v in k["info"]["m60"].values()) or None

    # gráficos
    k["chart"] = {
        "asist":   [asistencia_maxima(w[w["NombreFundo"] == f]) or 0 for f in FUNDOS],
        "hefect":  [w.loc[w["NombreFundo"] == f, "HorasEfectivas"].sum() for f in FUNDOS],
        "rot":     [(rotacion(df, sem, f) or 0) * 100 for f in FUNDOS],
        "hprom":   [hrs_prom_chart(w, f) or 0 for f in FUNDOS],
        "genM":    [], "genF": [],
    }
    for f in FUNDOS:
        m, fe = genero_chart(w, f)
        k["chart"]["genM"].append((m or 0) * 100)
        k["chart"]["genF"].append((fe or 0) * 100)

    # horas laborales - rango de edades
    k["hl"] = {}
    for etiq, (lo, hi) in {"48_60": (48, 60), "60_72": (60, 72), "m72": (72, None)}.items():
        k["hl"][etiq] = {f: bucket_horas(w[w["NombreFundo"] == f], lo, hi) for f in FUNDOS}
        k["hl"][etiq]["Con"] = bucket_horas(w, lo, hi)
    for etiq, (lo, hi) in {"e60_65": (60, 65), "e65_69": (65, 69), "e69": (69, None)}.items():
        k["hl"][etiq] = {f: bucket_edad(w[w["NombreFundo"] == f], hoy, lo, hi) for f in FUNDOS}
        k["hl"][etiq]["Con"] = bucket_edad(w, hoy, lo, hi)

    # v02: Grupo Asistencia (detalle diario L-D por fundo)
    k["asist_diaria"] = {f: asistencia_diaria_fundo(w, f) for f in FUNDOS}

    # v02: Grupo Ratios (supervisor/área + antigüedad)
    k["ratio_area"] = {f: supervisor_por_area(w, f) for f in FUNDOS}
    k["antiguedad"] = {f: antiguedad_rangos(w, hoy, f) for f in FUNDOS}
    k["antiguedad"]["Con"] = antiguedad_rangos(w, hoy)

    # v02: Grupo Alertas & Demografia
    k["genero_fundo"] = {f: composicion_genero(w, f) for f in FUNDOS}
    k["genero_fundo"]["Con"] = composicion_genero(w)
    k["jubilacion"] = {f: alertas_jubilacion(w, hoy, f) for f in FUNDOS}
    k["jubilacion"]["Con"] = alertas_jubilacion(w, hoy)

    # v08: Ausentismo Diario (Grupo Asistencia)
    k["ausentismo_diario"] = {f: ausentismo_diario_fundo(w, f) for f in FUNDOS}
    k["ausentismo_diario"]["Con"] = ausentismo_diario_fundo(w)

    # v08: Sin Descanso / Horas Altas / HE Obrero / Horas Promedio (Grupo Ratios)
    k["sin_descanso"] = {f: personal_sin_descanso(w, f) for f in FUNDOS}
    k["sin_descanso"]["Con"] = personal_sin_descanso(w)
    k["horas_altas"] = {f: horas_altas_detalle(w, f) for f in FUNDOS}
    k["horas_altas"]["Con"] = horas_altas_detalle(w)
    k["he_obrero"] = {f: he_obrero_area(w, f) for f in FUNDOS}
    k["he_obrero"]["Con"] = he_obrero_area(w)
    k["horas_prom_detalle"] = {f: horas_prom_area(w, f) for f in FUNDOS}
    k["horas_prom_detalle"]["Con"] = horas_prom_area(w)

    # v08: Rango de Edades (Grupo Alertas)
    k["rango_edades"] = {f: rango_edades_detalle(w, hoy, f) for f in FUNDOS}
    k["rango_edades"]["Con"] = rango_edades_detalle(w, hoy)

    # v04: snapshot por día individual, para el filtro de "Fecha" (como en Power BI)
    k["dias"] = {}
    for dia in sorted(w["FechaTareo"].dropna().unique()):
        wd = w[w["FechaTareo"] == dia]
        iso = pd.Timestamp(dia).strftime("%Y-%m-%d")
        area_dia = {f: asist_area_cultivo(wd, f) for f in FUNDOS}
        area_dia["Con"] = asist_area_cultivo(wd)
        asis_fundo = {f: wd.loc[(wd["NombreFundo"] == f) & (wd["EsUltimo"] == 1), "DNI_Trabajador"].nunique() for f in FUNDOS}
        asis_fundo["Con"] = wd.loc[wd["EsUltimo"] == 1, "DNI_Trabajador"].nunique()
        # v09: secciones reactivas al filtro de Fecha en Asistencia/Ratios/Alertas
        k["dias"][iso] = {
            "label": pd.Timestamp(dia).strftime("%d/%m/%Y"),
            "asistencia": asis_fundo["Con"],
            "supervisores": wd["Supervisor"].nunique(),
            "extranjeros": wd.loc[wd["Nacionalidad"].astype(str).str.upper() == "EXTRANJERA", "DNI_Trabajador"].nunique(),
            "horas_prom": horas_promedio(wd),
            "area": area_dia,
            "area_asis": asis_fundo,
            "ausentismo_pt": {**{f: ausentismo_dia_puntual(w, dia, f) for f in FUNDOS}, "Con": ausentismo_dia_puntual(w, dia)},
            "sin_descanso": {**{f: personal_sin_descanso(wd, f) for f in FUNDOS}, "Con": personal_sin_descanso(wd)},
            "horas_altas": {**{f: horas_altas_detalle(wd, f) for f in FUNDOS}, "Con": horas_altas_detalle(wd)},
            "he_obrero": {**{f: he_obrero_area(wd, f) for f in FUNDOS}, "Con": he_obrero_area(wd)},
            "horas_prom_detalle": {**{f: horas_prom_area(wd, f) for f in FUNDOS}, "Con": horas_prom_area(wd)},
            "rango_edades": {**{f: rango_edades_detalle(wd, hoy, f) for f in FUNDOS}, "Con": rango_edades_detalle(wd, hoy)},
            "ratio_area": {f: supervisor_por_area(wd, f) for f in FUNDOS},
            "antiguedad": {**{f: antiguedad_rangos(wd, hoy, f) for f in FUNDOS}, "Con": antiguedad_rangos(wd, hoy)},
            "genero_fundo": {**{f: composicion_genero(wd, f) for f in FUNDOS}, "Con": composicion_genero(wd)},
            "jubilacion": {**{f: alertas_jubilacion(wd, hoy, f) for f in FUNDOS}, "Con": alertas_jubilacion(wd, hoy)},
        }
    return k


# ─────────────────────── plantillas HTML ───────────────────────
CSS_HERO = """
:root{--verde:#1B4332;--verde2:#2D6A4F;--dorado:#C9A227;--tinta:#232B26;--gris:#6F7871;--linea:#E3E6DE;--papel:#F4F5F1}
*{box-sizing:border-box}
table{font-variant-numeric:tabular-nums}
.gold-rule{height:2px;width:72px;background:linear-gradient(90deg,var(--dorado),#E4C766);margin:10px auto 8px;border-radius:2px}
.eyebrow{font-size:.68em;letter-spacing:.28em;text-transform:uppercase;color:var(--dorado);font-weight:700;text-align:center}
"""

def html_hero(k):
    return f"""
<div style='background:linear-gradient(160deg,#FFFFFF 60%,#F7F8F4);padding:26px 32px 22px;border-radius:16px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 14px 40px rgba(20,30,25,.08);position:relative;overflow:hidden;border:1px solid var(--linea)'>
<div style='position:absolute;inset:0 auto 0 0;width:5px;background:linear-gradient(180deg,var(--verde),var(--verde2))'></div>
<div class='eyebrow'>Hass Per&uacute; &middot; Recursos Humanos</div>
<div style='font-size:1.7em;font-weight:800;letter-spacing:.5px;color:var(--verde);text-align:center;margin-top:4px'>Dashboard de Indicadores</div>
<div class='gold-rule'></div>
<div style='text-align:center;font-size:.8em;color:var(--gris);letter-spacing:.14em;text-transform:uppercase'>
An&aacute;lisis integral de KPIs &nbsp;&middot;&nbsp; Semana <span class='js-sem'></span> &nbsp;&middot;&nbsp; <span class='js-anio'></span> &nbsp;&middot;&nbsp; Actualizado <span class='js-fecha'></span></div>
</div>"""

def tarjeta_kpi(titulo, span_id):
    return (f"<div class='kpi-card' style='background:#fff;padding:18px 16px 16px;border-radius:12px;text-align:center;"
            f"border:1px solid var(--linea);box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05);position:relative;overflow:hidden'>"
            f"<div style='position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--verde),var(--verde2))'></div>"
            f"<div style='color:var(--gris);font-size:.72em;text-transform:uppercase;letter-spacing:.12em;font-weight:700;margin-bottom:7px'>{titulo}</div>"
            f"<div style='font-size:1.9em;font-weight:700;color:var(--verde);font-variant-numeric:tabular-nums;line-height:1.1' id='{span_id}'></div></div>")

def _tabla_area_html(a, titulo, color, borde, fondo, tot_col, header_txt):
    filas = ""
    for i, area in enumerate(["Cosecha", "Empaque", "Labores", "Riego y Fert.", "Sanidad", "Otros"]):
        bg = ";background:#FAFAF7" if i % 2 else ""
        v = a[area]
        filas += (f"<tr style='{bg[1:]}'><td style='padding:3px 4px;font-weight:500'>{area}</td>"
                  f"<td style='padding:3px 4px;text-align:center;font-weight:600'>{fmt_miles(v['ara']) if v['ara'] else ''}</td>"
                  f"<td style='padding:3px 4px;text-align:center;font-weight:600'>{fmt_miles(v['pal']) if v['pal'] else ''}</td>"
                  f"<td style='padding:3px 4px;text-align:center;font-weight:600'>{fmt_miles(v['con']) if v['con'] else ''}</td></tr>")
    t = a["Total"]
    filas += (f"<tr style='border-top:2px solid #E3E6DE;font-weight:700;background:#F1F5EE'>"
              f"<td style='padding:4px'>Total</td>"
              f"<td style='padding:4px;text-align:center;color:#1B4332'>{fmt_miles(t['ara'])}</td>"
              f"<td style='padding:4px;text-align:center;color:#B3453C'>{fmt_miles(t['pal'])}</td>"
              f"<td style='padding:4px;text-align:center;color:#5B6660'>{fmt_miles(t['con'])}</td></tr>")
    return f"""
<div style='background:{fondo};border:1px solid #E3E6DE;border-radius:10px;padding:12px;border-top:4px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)'>
<div style='margin-bottom:6px'><span style='font-size:14px;font-weight:600;color:{color}'>&#128107; {titulo}</span></div>
<div style='font-size:11px;color:#98A099;margin-bottom:8px'>{header_txt}</div>
<table style='width:100%;border-collapse:collapse;font-size:11px'>
<thead><tr style='background:#ECEEE8'>
<th style='padding:4px;text-align:left;color:#5B6660'></th>
<th style='padding:4px;text-align:center;color:#1B4332'>&#129744; Ar&aacute;ndano</th>
<th style='padding:4px;text-align:center;color:#B3453C'>&#129361; Palta</th>
<th style='padding:4px;text-align:center;color:#5B6660'>&#128202; {tot_col}</th></tr></thead>
<tbody>{filas}</tbody></table></div>"""

def tabla_area(k, fkey, titulo, color, borde, es_con=False):
    a = k["area"][fkey]
    mx = fmt_miles(k["area_max"][fkey])
    an = fmt_miles(k["area_ant"][fkey])
    fondo = "#FAFAF7" if es_con else "#fff"
    tot_col = "Total General" if es_con else "Consolidado"
    header_txt = f"Asistencia M&aacute;xima: <span style='font-weight:700;color:#161D19'>{mx}</span> | Sem Ant: <span style='font-weight:700;color:#161D19'>{an}</span>"
    return _tabla_area_html(a, titulo, color, borde, fondo, tot_col, header_txt)

def tabla_area_dia(dia, fkey, titulo, color, borde, es_con=False):
    a = dia["area"][fkey]
    asis = fmt_miles(dia["area_asis"][fkey])
    fondo = "#FAFAF7" if es_con else "#fff"
    tot_col = "Total General" if es_con else "Consolidado"
    header_txt = f"Asistencia del d&iacute;a: <span style='font-weight:700;color:#161D19'>{asis}</span>"
    return _tabla_area_html(a, titulo, color, borde, fondo, tot_col, header_txt)

def tabla_info(k, fkey, titulo, borde):
    i, ia = k["info"], k["infoA"]
    la, lc = f"Sem. {k['semana']-1}", f"Sem. {k['semana']}"
    def celda(dic_a, dic, fmt):
        va = dic_a.get(fkey) if dic_a else None
        v  = dic.get(fkey)
        return guion_si_vacio(fmt(va), va), guion_si_vacio(fmt(v), v)
    filas_def = [
        ("Asistencia",           *celda(ia["asist"],  i["asist"],  fmt_miles)),
        ("Rotaci&oacute;n",      *celda(ia["rot"],    i["rot"],    lambda v: fmt_pct(v))),
        ("Supervisores",         *celda(ia["sup"],    i["sup"],    fmt_miles)),
        ("Jornales/Supervisor",  *celda(ia["jorsup"], i["jorsup"], fmt_dec)),
        ("Extranjeros",          *celda(ia["ext"],    i["ext"],    fmt_miles)),
        ("Ausentismo",           *celda(ia["aus"],    i["aus"],    lambda v: fmt_pct(v))),
        ("G&eacute;nero M%",     *celda(ia["gen"],    i["gen"],    lambda v: fmt_pct(v))),
        ("Antiguedad [0-1 A&ntilde;o]", "-", guion_si_vacio(fmt_miles(i["ant01"].get(fkey)), i["ant01"].get(fkey))),
        ("Horas Promedio",       *(x + ("h" if x != "-" else "") for x in celda(ia["hrs"], i["hrs"], fmt_dec))),
        ("Cierre Semana (L-D)",  "-", guion_si_vacio(fmt_pct(i["cierre"].get(fkey)), i["cierre"].get(fkey))),
        ("&gt; 60 Horas",        "-", guion_si_vacio(fmt_miles(i["m60"].get(fkey)), i["m60"].get(fkey))),
        ("Total Horas Extras",   *celda(ia["he"],     i["he"],     fmt_dec)),
    ]
    filas = ""
    for n, (nombre, va, v) in enumerate(filas_def):
        bg = "background:#FAFAF7" if n % 2 else ""
        filas += (f"<tr style='border-bottom:1px solid #ECEEE8;{bg}'>"
                  f"<td style='padding:8px;color:#000'>{nombre}</td>"
                  f"<td style='padding:8px;text-align:right;color:#000;font-weight:500'>{va}</td>"
                  f"<td style='padding:8px;text-align:right;color:#000;font-weight:500'>{v}</td></tr>")
    icono = "&#128202;" if fkey == "Con" else "&#129361;&#129744;"
    return f"""
<div style='background:#fff;padding:20px;border-radius:12px;border-left:5px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06)'>
<div style='font-size:1.2em;font-weight:700;color:#000;margin-bottom:15px'>{icono} {titulo}</div>
<table style='width:100%;border-collapse:collapse;font-size:.9em'>
<tr style='background:#FAFAF7'><th style='text-align:left;padding:8px;font-weight:600;font-size:1.1em'>Indicador</th>
<th style='text-align:right;padding:8px;font-weight:600'>{la}</th>
<th style='text-align:right;padding:8px;font-weight:600'>{lc}</th></tr>
{filas}</table></div>"""

def tabla_horas_laborales(k):
    filas_def = [
        ("Trabajadores 48-60h / Semana", "48_60"),
        ("Trabajadores 60-72h / Semana", "60_72"),
        ("Trabajadores &gt; 72h / Semana", "m72"),
        ("Trabajadores 60-65 A&ntilde;os", "e60_65"),
        ("Trabajadores 65-69 A&ntilde;os", "e65_69"),
        ("Trabajadores &gt; 69 A&ntilde;os", "e69"),
    ]
    filas = ""
    for n, (nombre, key) in enumerate(filas_def):
        d = k["hl"][key]
        bg = "background:#FAFAF7" if n % 2 else ""
        filas += (f"<tr style='border-bottom:1px solid #ECEEE8;{bg}'>"
                  f"<td style='padding:12px 20px'>{nombre}</td>"
                  + "".join(f"<td style='padding:12px 20px;text-align:center;font-weight:600'>{fmt_miles(d[f])}</td>" for f in FUNDOS)
                  + f"<td style='padding:12px 20px;text-align:center;font-weight:700'>{fmt_miles(d['Con'])}</td></tr>")
    return f"""
<div style='color:var(--verde);font-size:1em;font-weight:700;letter-spacing:.14em;text-transform:uppercase;padding-bottom:10px;border-bottom:1px solid var(--linea);margin:30px 0 18px'>&#9200; Detalles de Horas Laborales - Rango de Edades</div>
<div style='background:white;border-radius:12px;overflow-x:auto;box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06)'>
<table style='width:100%;border-collapse:collapse;font-size:.95em;min-width:700px'>
<thead><tr style='background:#1B4332;color:#fff;font-size:.82em;text-transform:uppercase;letter-spacing:.08em'>
<th style='text-align:left;padding:14px 20px;font-weight:600'>M&eacute;trica</th>
<th style='text-align:center;padding:14px 20px'>Fundo HP</th>
<th style='text-align:center;padding:14px 20px'>Fundo BH</th>
<th style='text-align:center;padding:14px 20px'>Fundo Cao</th>
<th style='text-align:center;padding:14px 20px'>Consolidado</th></tr></thead>
<tbody>{filas}</tbody></table></div>"""

def titulo_pestana(texto, sid):
    return f"""<div style='background:#FFFFFF;padding:18px 26px;border-radius:14px;border:1px solid var(--linea);box-shadow:0 1px 2px rgba(20,30,25,.04),0 10px 30px rgba(20,30,25,.06);border-left:4px solid var(--dorado);margin-bottom:18px'>
<div style='font-size:1.35em;font-weight:700;color:var(--verde);letter-spacing:.3px'>{texto}</div>
<div style='font-size:.74em;color:var(--gris);letter-spacing:.14em;text-transform:uppercase;margin-top:5px'>
An&aacute;lisis por KPI &nbsp;&middot;&nbsp; Semana <span class='js-sem'></span> &nbsp;&middot;&nbsp; <span class='js-anio'></span>
&nbsp;&middot;&nbsp; Actualizado <span class='js-fecha'></span></div></div>"""

def _celda_dia(serie, i, color_base, weight_base, es_total=False):
    """Replica el semáforo DAX: Mar-Vie en rojo+negrita si cae vs. el día anterior."""
    v = serie[i]
    if 1 <= i <= 4:  # Martes..Viernes
        if v < serie[i - 1]:
            color, weight = "#B3453C", "bold"
        else:
            color, weight = color_base, weight_base
    elif i in (5, 6):  # Sábado, Domingo
        color, weight = ("#666", "normal") if not es_total else (color_base, weight_base)
    else:  # Lunes
        color, weight = color_base, weight_base
    txt = fmt_miles(v) if (v or es_total) else ""
    return f"<td style='padding:5px 6px;text-align:center;color:{color};font-weight:{weight}'>{txt}</td>"

def tabla_asistencia_diaria(k, fkey, titulo, borde):
    d = k["asist_diaria"][fkey]
    if d is None:
        return f"<div style='background:#fff;border-radius:10px;padding:16px;color:#999'>Sin datos para {titulo}</div>"
    fechas = d["fechas"]
    heads = "".join(f"<th id='header_asist_{fkey}_day{j}' style='padding:10px 6px;text-align:center;font-weight:600'>{n}<br><span style='font-size:10px;font-weight:400'>{f}</span></th>"
                     for j, (n, f) in enumerate(zip(["Lun", "Mar", "Mi&eacute;", "Jue", "Vie", "S&aacute;b", "Dom"], fechas)))
    filas_html = ""
    for i, (nombre, serie) in enumerate(d["filas"].items()):
        bg = "background:#FAFAF7" if i % 2 else ""
        celdas = "".join(_celda_dia(serie, j, "#232B26", "normal") for j in range(7))
        lv = serie[:5]
        celdas += (f"<td style='padding:5px 6px;text-align:center;font-weight:600;color:{borde};background:#F1F5EE'>{fmt_miles(sum(serie))}</td>"
                   f"<td style='padding:5px 6px;text-align:center;font-weight:700;color:{borde};background:#F1F5EE'>{fmt_miles(max(lv))}</td>"
                   f"<td style='padding:5px 6px;text-align:center;font-weight:700;color:{borde};background:#F1F5EE'>{fmt_miles(min(lv))}</td>")
        filas_html += f"<tr style='border-bottom:1px solid #ECEEE8;{bg}'><td style='padding:5px 8px;font-weight:500'>{nombre}</td>{celdas}</tr>"
    tot = d["total"]; lv_tot = tot[:5]
    celdas_tot = "".join(_celda_dia(tot, j, "#1B4332", "600", es_total=True) for j in range(7))
    filas_html += (f"<tr style='background:#EDF2EC'><td style='padding:6px 8px;font-weight:600;color:{borde}'>Total Fundo</td>{celdas_tot}"
                   + f"<td style='padding:6px 6px;text-align:center;font-weight:700;color:{borde};background:#D8E4D6'>{fmt_miles(sum(tot))}</td>"
                   + f"<td style='padding:6px 6px;text-align:center;font-weight:700;color:{borde};background:#D8E4D6'>{fmt_miles(max(lv_tot))}</td>"
                   + f"<td style='padding:6px 6px;text-align:center;font-weight:700;color:{borde};background:#D8E4D6'>{fmt_miles(min(lv_tot))}</td></tr>")
    return f"""
<div style='background:linear-gradient(135deg,#1B4332,#2D6A4F);padding:12px 20px;border-radius:14px 14px 0 0;border-bottom:2px solid #C9A227'>
<span style='color:#fff;font-weight:600;font-size:15px;letter-spacing:.06em'>Asistencia Diaria - Detalle {titulo}</span></div>
<div style='background:#fff;border:1px solid #E3E6DE;border-radius:0 0 14px 14px;padding:16px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05);overflow-x:auto;margin-bottom:20px'>
<table style='width:100%;border-collapse:collapse;font-size:12px;min-width:900px'>
<thead><tr style='background:#1B4332;color:#fff;font-size:.86em;letter-spacing:.04em'><th style='padding:10px 8px;text-align:left'>Grupo</th>{heads}
<th style='padding:10px 6px;background:#6F7871'>Total Semana</th><th style='padding:10px 6px;background:#6F7871'>M&aacute;x.(L-V)</th><th style='padding:10px 6px;background:#6F7871'>M&iacute;n.(L-V)</th></tr></thead>
<tbody>{filas_html}</tbody></table></div>
<div style='font-size:.75em;color:#98A099;margin:-14px 0 20px 4px'>&#128308; Rojo/negrita: ca&iacute;da respecto al d&iacute;a anterior (Mar-Vie)</div>"""

def tarjeta_ausentismo_diario(k, fkey, titulo, borde):
    a = k["ausentismo_diario"][fkey]
    if a is None:
        return f"<div style='background:#fff;border-radius:10px;padding:16px;color:#999'>Sin datos para {titulo}</div>"
    filas = ""
    for nombre, v in a["dias"].items():
        color = "#B3453C" if v["ausentes"] > 0 else "#2D6A4F"
        filas += (f"<div style='display:flex;justify-content:space-between;font-size:.85em;margin:5px 0'>"
                  f"<span>{nombre} <span style='color:#98A099;font-size:.85em'>({v['fecha']})</span></span>"
                  f"<span style='font-weight:600;color:{color}'>{fmt_miles(v['ausentes'])} ({fmt_pct(v['pct'], 1)})</span></div>")
    return f"""
<div style='background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:16px;border-top:4px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)'>
<div style='font-size:15px;font-weight:600;color:{borde};margin-bottom:8px'>{titulo}</div>
<div style='font-size:.8em;color:#6F7871;margin-bottom:8px'>Base (Lunes): <span style='font-weight:700;color:#232B26'>{fmt_miles(a['base'])}</span></div>
<div style='border-top:1px solid #ECEEE8;padding-top:8px'>{filas}</div>
<div style='border-top:1px solid #ECEEE8;margin-top:6px;padding-top:6px;display:flex;justify-content:space-between;font-weight:700'>
<span>M&aacute;x. Ausentes</span><span style='color:{borde}'>{fmt_miles(a['max_aus'])} ({fmt_pct(a['max_pct'], 1)})</span></div>
</div>"""

def bloque_asistencia(k):
    sid = f"w{k['semana']}"
    body = "".join(tabla_asistencia_diaria(k, f, f"Fundo {ETIQ_FUNDO[f]}", c)
                    for f, c in zip(FUNDOS, ["#B3453C", "#31678F", "#1B4332"]))
    aus_titulo = ("<div style='background:linear-gradient(135deg,#1B4332,#2D6A4F);padding:12px 20px;border-radius:14px 14px 0 0;border-bottom:2px solid #C9A227;margin-top:10px'>"
                  "<span style='color:#fff;font-weight:600;font-size:15px;letter-spacing:.06em'>Ausentismo Diario</span></div>")
    aus_cards = "".join(tarjeta_ausentismo_diario(k, f, f"Fundo {ETIQ_FUNDO[f]}" if f != "Con" else "Consolidado", c)
                         for f, c in zip(FUNDOS + ["Con"], ["#B3453C", "#31678F", "#1B4332", "#5B6660"]))
    aus_body = (f"{aus_titulo}<div id='grid_aus_{sid}' style='background:#FAFAF7;border-radius:0 0 14px 14px;padding:16px;"
                f"display:grid;grid-template-columns:repeat(4,1fr);gap:12px'>{aus_cards}</div>")
    return f"<div class='tab-content' data-tab='asistencia' data-sem='{k['semana']}' style='display:none'>{titulo_pestana('Detalle Asistencia &mdash; RRHH', sid)}{body}{aus_body}</div>"

def tabla_ratio_area(k, fkey, titulo, borde):
    r = k["ratio_area"][fkey]
    filas = ""
    for i, area in enumerate(AREAS_RATIO):
        v = r.get(area, {"sup": 0, "trab": 0, "ratio": 0, "horas": 0, "extra": 0})
        bg = "background:#FAFAF7" if i % 2 else ""
        filas += (f"<tr style='{bg}'><td style='padding:6px 8px'>{area}</td>"
                  f"<td style='padding:6px 8px;text-align:center'>{fmt_miles(v['sup'])}</td>"
                  f"<td style='padding:6px 8px;text-align:center'>{fmt_miles(v['trab'])}</td>"
                  f"<td style='padding:6px 8px;text-align:center;font-weight:600;color:{borde}'>{fmt_dec(v['ratio'])}</td>"
                  f"<td style='padding:6px 8px;text-align:center'>{fmt_dec(v['horas'])}</td>"
                  f"<td style='padding:6px 8px;text-align:center'>{fmt_dec(v['extra'])}</td></tr>")
    return f"""
<div style='background:#fff;padding:16px;border-radius:12px;border-left:5px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06);margin-bottom:16px'>
<div style='font-size:1.1em;font-weight:700;margin-bottom:10px'>&#128101; # de Supervisores / Por Labor - {titulo}</div>
<table style='width:100%;border-collapse:collapse;font-size:.85em'>
<tr style='background:#FAFAF7;font-weight:600'><td style='padding:6px 8px'>Labor</td><td style='padding:6px 8px;text-align:center'>Supervisores</td>
<td style='padding:6px 8px;text-align:center'>Trabajadores</td><td style='padding:6px 8px;text-align:center'>Ratio</td>
<td style='padding:6px 8px;text-align:center'>Horas</td><td style='padding:6px 8px;text-align:center'>H. Extra</td></tr>
{filas}</table></div>"""

def tabla_antiguedad(k, fkey, titulo, borde):
    a = k["antiguedad"][fkey]
    filas = "".join(f"<div style='margin-bottom:8px'><div style='display:flex;justify-content:space-between;font-size:.85em;margin-bottom:2px'>"
                     f"<span>{r}</span><span style='font-weight:600'>{fmt_miles(a[r]['n'])} ({fmt_pct(a[r]['pct'])})</span></div>"
                     f"<div style='background:#ECEEE8;border-radius:4px;height:8px'><div style='background:{borde};width:{a[r]['pct']*100:.0f}%;height:8px;border-radius:4px'></div></div></div>"
                     for r in RANGOS_ANTIG)
    return f"""
<div style='background:#fff;padding:16px;border-radius:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06);margin-bottom:16px'>
<div style='font-size:1.05em;font-weight:700;margin-bottom:12px'>{titulo}</div>{filas}</div>"""

def tarjeta_sin_descanso(k, fkey, titulo, borde):
    s = k["sin_descanso"][fkey]
    return f"""
<div style='background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:16px;border-top:4px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)'>
<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
<span style='font-size:15px;font-weight:500;color:{borde}'>{titulo}</span>
<span style='font-size:12px;padding:2px 7px;border-radius:8px;background:#F8F1F0;color:{borde};border:0.5px solid {borde}'>{fmt_pct(s['pct'], 1)}</span></div>
<div style='font-size:28px;font-weight:500;color:#161D19;margin-bottom:2px'>{fmt_miles(s['con7'])}</div>
<div style='font-size:13px;color:#98A099;margin-bottom:10px'>Con cierre de semana (7 d&iacute;as)</div>
<div style='border-top:1px solid #ECEEE8;padding-top:8px;display:flex;flex-direction:column;gap:5px;font-size:13px'>
<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>Total</span><span style='font-weight:600;color:#232B26'>{fmt_miles(s['tot'])}</span></div>
<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>Sin cierre</span><span style='font-weight:600;color:#232B26'>{fmt_miles(s['sin7'])}</span></div>
</div></div>"""

def tarjeta_horas_altas(k, fkey, titulo, borde):
    h = k["horas_altas"][fkey]
    return f"""
<div style='background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:16px;border-top:4px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)'>
<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
<span style='font-size:15px;font-weight:500;color:{borde}'>{titulo}</span>
<span style='font-size:12px;padding:2px 7px;border-radius:8px;background:#F8F1F0;color:{borde};border:0.5px solid {borde}'>{fmt_pct(h['pct'], 1)}</span></div>
<div style='font-size:28px;font-weight:500;color:#161D19;margin-bottom:2px'>{fmt_miles(h['altas'])}</div>
<div style='font-size:13px;color:#98A099;margin-bottom:10px'>Con horas altas</div>
<div style='border-top:1px solid #ECEEE8;padding-top:8px;display:flex;flex-direction:column;gap:5px;font-size:13px'>
<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>Total</span><span style='font-weight:600;color:#232B26'>{fmt_miles(h['tot'])}</span></div>
<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>[48-60 hrs]</span><span style='font-weight:600;color:#B07D2B'>{fmt_miles(h['48_60'])}</span></div>
<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>[60-72 hrs]</span><span style='font-weight:600;color:#B3453C'>{fmt_miles(h['60_72'])}</span></div>
<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>[&gt; 72 hrs]</span><span style='font-weight:600;color:#7A3327'>{fmt_miles(h['m72'])}</span></div>
</div></div>"""

def tarjeta_he_obrero(k, fkey, titulo, borde):
    h = k["he_obrero"][fkey]
    filas = "".join(f"<tr style='border-bottom:1px solid #ECEEE8'><td style='padding:3px 6px;color:#232B26;font-weight:600'>{a}</td>"
                     f"<td style='padding:3px 6px;text-align:right;color:#5B6660'>{fmt_dec(v)}</td></tr>"
                     for a, v in h["area"].items())
    return f"""
<div style='background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:16px;border-top:4px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05);margin-bottom:12px'>
<div style='font-size:15px;font-weight:500;color:{borde};margin-bottom:8px'>&#128337; {titulo}</div>
<div style='font-size:28px;font-weight:500;color:#161D19;margin-bottom:2px'>{fmt_dec(h['he_total'])}</div>
<div style='font-size:13px;color:#98A099;margin-bottom:10px'>Total HE Obrero</div>
<div style='border-top:1px solid #ECEEE8;padding-top:8px;font-size:13px'><span style='color:#6F7871'># Obreros</span> <span style='font-weight:600'>{fmt_miles(h['tot_obr'])}</span></div>
</div>
<div style='background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)'>
<div style='font-size:11px;font-weight:700;color:#98A099;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px'>HE por &Aacute;rea</div>
<table style='width:100%;border-collapse:collapse;font-size:12px'>{filas}</table></div>"""

def tarjeta_horas_prom_detalle(k, fkey, titulo, borde):
    h = k["horas_prom_detalle"][fkey]
    filas = "".join(f"<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>{a}</span>"
                     f"<span style='font-weight:600;color:#232B26'>{fmt_dec(v)}</span></div>"
                     for a, v in h["area"].items())
    return f"""
<div style='background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:16px;border-top:4px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)'>
<div style='font-size:15px;font-weight:500;color:{borde};margin-bottom:8px'>&#128337; {titulo}</div>
<div style='font-size:28px;font-weight:500;color:#161D19;margin-bottom:2px'>{fmt_dec(h['total'])}</div>
<div style='font-size:13px;color:#98A099;margin-bottom:10px'>Horas promedio</div>
<div style='border-top:1px solid #ECEEE8;padding-top:8px;display:flex;flex-direction:column;gap:4px;font-size:12px'>{filas}</div>
</div>"""

def bloque_ratios(k):
    sid = f"w{k['semana']}"
    ratios = "".join(tabla_ratio_area(k, f, f"Fundo {ETIQ_FUNDO[f]}", c) for f, c in zip(FUNDOS, ["#B3453C", "#31678F", "#1B4332"]))
    antig_titulo = ("<div style='color:var(--verde);font-size:.95em;font-weight:700;letter-spacing:.14em;text-transform:uppercase;padding-bottom:9px;border-bottom:1px solid var(--linea);margin:26px 0 16px'>"
                     "# de Personal por Antig&uuml;edad</div>")
    antig = f"<div id='grid_antig_{sid}' class='grid-info' style='display:grid;grid-template-columns:repeat(4,1fr);gap:16px'>" + "".join(
        tabla_antiguedad(k, f, f"Fundo {ETIQ_FUNDO[f]}" if f != "Con" else "Consolidado", c)
        for f, c in zip(FUNDOS + ["Con"], ["#B3453C", "#31678F", "#1B4332", "#5B6660"])) + "</div>"

    colores4 = ["#B3453C", "#31678F", "#1B4332", "#5B6660"]
    nombres4 = [f"Fundo {ETIQ_FUNDO[f]}" for f in FUNDOS] + ["Consolidado"]

    def seccion(titulo, tarjeta_fn, grid_id, cols=4):
        cards = "".join(tarjeta_fn(k, f, n, c) for f, n, c in zip(FUNDOS + ["Con"], nombres4, colores4))
        tit = (f"<div style='background:linear-gradient(135deg,#1B4332,#2D6A4F);padding:12px 20px;border-radius:14px 14px 0 0;border-bottom:2px solid #C9A227;margin-top:10px'>"
               f"<span style='color:#fff;font-weight:600;font-size:15px;letter-spacing:.06em'>{titulo}</span></div>")
        return (f"{tit}<div id='{grid_id}' style='background:#FAFAF7;border-radius:0 0 14px 14px;padding:16px;"
                f"display:grid;grid-template-columns:repeat({cols},1fr);gap:12px'>{cards}</div>")

    sin_descanso_html = seccion("# de Personal Sin Descanso (Lunes - Domingo)", tarjeta_sin_descanso, f"grid_sd_{sid}")
    horas_altas_html = seccion("Detalle de Horas Altas - # de Personas", tarjeta_horas_altas, f"grid_ha_{sid}")
    he_obrero_html = seccion("HE Obrero", tarjeta_he_obrero, f"grid_he_{sid}")
    horas_prom_html = seccion("Horas Promedio", tarjeta_horas_prom_detalle, f"grid_hp_{sid}")

    return (f"<div class='tab-content' data-tab='ratios' data-sem='{k['semana']}' style='display:none'>{titulo_pestana('Detalle Ratios &mdash; RRHH', sid)}"
            f"<div id='grid_ra_{sid}' class='grid-info' style='display:grid;grid-template-columns:repeat(3,1fr);gap:16px'>{ratios}</div>{antig_titulo}{antig}"
            f"{sin_descanso_html}{horas_altas_html}{he_obrero_html}{horas_prom_html}</div>")

def tabla_composicion(k, fkey, titulo, borde):
    g = k["genero_fundo"][fkey]
    return f"""
<div style='background:#fff;padding:16px;border-radius:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06)'>
<div style='font-size:1.05em;font-weight:700;margin-bottom:10px;color:{borde}'>{titulo}</div>
<div style='display:flex;justify-content:space-between;font-size:.9em;margin-bottom:6px'><span>&#9794; Hombres</span><span style='font-weight:600'>{fmt_miles(g['m'])} ({fmt_pct(g['pm'])})</span></div>
<div style='background:#ECEEE8;border-radius:4px;height:10px;margin-bottom:10px'><div style='background:#31678F;width:{g['pm']*100:.0f}%;height:10px;border-radius:4px'></div></div>
<div style='display:flex;justify-content:space-between;font-size:.9em;margin-bottom:6px'><span>&#9792; Mujeres</span><span style='font-weight:600'>{fmt_miles(g['f'])} ({fmt_pct(g['pf'])})</span></div>
<div style='background:#ECEEE8;border-radius:4px;height:10px'><div style='background:#B3453C;width:{g['pf']*100:.0f}%;height:10px;border-radius:4px'></div></div>
</div>"""

def tabla_jubilacion(k, fkey, titulo, borde):
    j = k["jubilacion"][fkey]
    filas = "".join(f"<tr><td style='padding:5px 8px'>{et}</td><td style='padding:5px 8px;text-align:center'>{fmt_miles(j[key])}</td></tr>"
                     for et, key in [("60-65 a&ntilde;os", "60_65"), ("65-69 a&ntilde;os", "65_69"), ("&gt; 69 a&ntilde;os", "69")])
    return f"""
<div style='background:#fff;padding:16px;border-radius:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06);border-top:4px solid {borde}'>
<div style='font-size:1.05em;font-weight:700;margin-bottom:8px'>{titulo}</div>
<table style='width:100%;border-collapse:collapse;font-size:.85em'>{filas}
<tr style='background:#F8F1F0;font-weight:700'><td style='padding:6px 8px;color:#B3453C'>Total Alerta</td><td style='padding:6px 8px;text-align:center;color:#B3453C'>{fmt_miles(j['alerta'])} ({fmt_pct(j['p'])})</td></tr>
</table></div>"""

def tarjeta_rango_edades(k, fkey, titulo, borde):
    r = k["rango_edades"][fkey]
    filas = "".join(f"<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>{etiq} a&ntilde;os</span>"
                     f"<span style='font-weight:600;color:#232B26'>{fmt_miles(v['n'])} <span style='color:#98A099'>({fmt_pct(v['pct'], 1)})</span></span></div>"
                     for etiq, v in r["rangos"].items())
    return f"""
<div style='background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:16px;border-top:4px solid {borde};box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)'>
<div style='font-size:14px;font-weight:600;color:{borde};margin-bottom:12px'>{titulo}</div>
<div style='display:flex;flex-direction:column;gap:6px;font-size:12px'>{filas}
<div style='border-top:1px solid #ECEEE8;padding-top:6px;display:flex;justify-content:space-between;font-weight:700'><span style='color:#232B26'>Total</span><span style='color:{borde}'>{fmt_miles(r['tot'])}</span></div>
</div></div>"""

def bloque_alertas(k):
    sid = f"w{k['semana']}"
    colores = ["#B3453C", "#31678F", "#1B4332", "#5B6660"]
    comp_titulo = "<div style='color:var(--verde);font-size:.95em;font-weight:700;letter-spacing:.14em;text-transform:uppercase;padding-bottom:9px;border-bottom:1px solid var(--linea);margin:26px 0 16px'>Composici&oacute;n del Personal</div>"
    comp = f"<div id='grid_comp_{sid}' class='grid-info' style='display:grid;grid-template-columns:repeat(4,1fr);gap:16px'>" + "".join(
        tabla_composicion(k, f, f"Fundo {ETIQ_FUNDO[f]}" if f != "Con" else "Consolidado", c) for f, c in zip(FUNDOS + ["Con"], colores)) + "</div>"
    jub_titulo = "<div style='color:var(--verde);font-size:.95em;font-weight:700;letter-spacing:.14em;text-transform:uppercase;padding-bottom:9px;border-bottom:1px solid var(--linea);margin:26px 0 16px'>Alertas de Jubilaci&oacute;n</div>"
    jub = f"<div id='grid_jub_{sid}' class='grid-info' style='display:grid;grid-template-columns:repeat(4,1fr);gap:16px'>" + "".join(
        tabla_jubilacion(k, f, f"Fundo {ETIQ_FUNDO[f]}" if f != "Con" else "Consolidado", c) for f, c in zip(FUNDOS + ["Con"], colores)) + "</div>"
    edad_titulo = ("<div style='background:linear-gradient(135deg,#1B4332,#2D6A4F);padding:12px 20px;border-radius:14px 14px 0 0;border-bottom:2px solid #C9A227;margin-top:20px'>"
                   "<span style='color:#fff;font-weight:600;font-size:15px;letter-spacing:.06em'>Rango de Edades</span></div>")
    edad = (f"<div id='grid_edad_{sid}' style='background:#FAFAF7;border-radius:0 0 14px 14px;padding:16px;display:grid;grid-template-columns:repeat(4,1fr);gap:12px'>"
            + "".join(tarjeta_rango_edades(k, f, f"Fundo {ETIQ_FUNDO[f]}" if f != "Con" else "Consolidado", c)
                       for f, c in zip(FUNDOS + ["Con"], colores)) + "</div>")
    return (f"<div class='tab-content' data-tab='alertas' data-sem='{k['semana']}' style='display:none'>"
            f"{titulo_pestana('Detalle Alerta &amp; Demograf&iacute;a &mdash; RRHH', sid)}{comp_titulo}{comp}{jub_titulo}{jub}{edad_titulo}{edad}</div>")


def seccion_graficos(sid):
    def card(titulo, cid):
        return (f"<div style='background:white;padding:20px;border-radius:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06);height:280px'>"
                f"<div style='font-size:.95em;font-weight:600;color:#1B4332;margin-bottom:15px'>{titulo}</div>"
                f"<div style='position:relative;height:210px'><canvas id='{cid}_{sid}'></canvas></div></div>")
    return f"""
<div style='color:var(--verde);font-size:1em;font-weight:700;letter-spacing:.14em;text-transform:uppercase;padding-bottom:10px;border-bottom:1px solid var(--linea);margin:30px 0 18px'>&#128200; Visualizaci&oacute;n de Datos</div>
<div class='grid-charts' style='display:grid;grid-template-columns:repeat(4,1fr);gap:20px'>
{card("Asistencia Por Fundo","chartAsist")}{card("Rotaci&oacute;n De Personal","chartRot")}
{card("Horas Promedio Por Fundo","chartHrs")}{card("G&eacute;nero Por Fundo","chartGen")}</div>"""


def bloque_semana(k):
    """Todo el contenido de una semana (oculto/visible por el selector)."""
    sid = f"w{k['semana']}"
    kpis = "".join([
        tarjeta_kpi("Total Asistencia", f"kpi_asis_{sid}"),
        tarjeta_kpi("Rotaci&oacute;n Prom", f"kpi_rot_{sid}"),
        tarjeta_kpi("Supervisores", f"kpi_sup_{sid}"),
        tarjeta_kpi("Extranjeros", f"kpi_ext_{sid}"),
        tarjeta_kpi("Ausentismo", f"kpi_aus_{sid}"),
        tarjeta_kpi("Horas Prom", f"kpi_hrs_{sid}"),
    ])
    areas_todas = (tabla_area(k, "Hp", "Fundo HP", "#B3453C", "#B3453C")
             + tabla_area(k, "Bh", "Fundo BH", "#31678F", "#31678F")
             + tabla_area(k, "Cao", "Fundo CAO", "#1B4332", "#1B4332")
             + tabla_area(k, "Con", "Consolidado General", "#5B6660", "#5B6660", es_con=True))
    variantes_fecha = f"<div class='fecha-variant' data-fecha='todas' style='display:contents'>{areas_todas}</div>"
    for iso, dia in k["dias"].items():
        areas_dia = (tabla_area_dia(dia, "Hp", "Fundo HP", "#B3453C", "#B3453C")
                     + tabla_area_dia(dia, "Bh", "Fundo BH", "#31678F", "#31678F")
                     + tabla_area_dia(dia, "Cao", "Fundo CAO", "#1B4332", "#1B4332")
                     + tabla_area_dia(dia, "Con", "Consolidado General", "#5B6660", "#5B6660", es_con=True))
        variantes_fecha += f"<div class='fecha-variant' data-fecha='{iso}' style='display:none'>{areas_dia}</div>"
    infos = (tabla_info(k, "Hp", "Fundo HP", "#B3453C") + tabla_info(k, "Bh", "Fundo BH", "#31678F")
             + tabla_info(k, "Cao", "Fundo Cao", "#2D6A4F") + tabla_info(k, "Con", "Consolidado", "#5B6660"))
    return f"""
<div class='tab-content' data-tab='resumen' id='bloque_{sid}' data-sem='{k["semana"]}' style='display:none'>
<div class='grid-kpi' style='display:grid;grid-template-columns:repeat(6,1fr);gap:15px;margin:20px 0'>{kpis}</div>
<div style='background:linear-gradient(135deg,#1B4332,#2D6A4F);padding:12px 20px;border-radius:14px 14px 0 0;border-bottom:2px solid #C9A227'>
<span style='color:#fff;font-weight:600;font-size:15px;letter-spacing:.06em'>Asistencia por &Aacute;rea y Cultivo</span></div>
<div class='grid-areas-wrap' style='background:#FAFAF7;border-radius:0 0 14px 14px;padding:16px'>
<div class='grid-areas' style='display:grid;grid-template-columns:repeat(4,1fr);gap:12px'>{variantes_fecha}</div></div>
<div style='color:var(--verde);font-size:1em;font-weight:700;letter-spacing:.14em;text-transform:uppercase;padding-bottom:10px;border-bottom:1px solid var(--linea);margin:30px 0 18px'>&#127970; Informaci&oacute;n Por Fundo</div>
<div class='grid-info' style='display:grid;grid-template-columns:repeat(4,1fr);gap:20px'>{infos}</div>
{seccion_graficos("w"+str(k["semana"]))}
{tabla_horas_laborales(k)}
</div>"""


def _fmt_dia_extra(dia):
    """Formatea las secciones día-a-día (Ratios/Alertas/Ausentismo) para consumo en JS."""
    def f_ausentismo_pt(v):
        if v is None:
            return None
        return {"base": fmt_miles(v["base"]), "ausentes": fmt_miles(v["ausentes"]), "pct": fmt_pct(v["pct"], 1)}

    def f_sin_descanso(v):
        return {"tot": fmt_miles(v["tot"]), "con7": fmt_miles(v["con7"]), "sin7": fmt_miles(v["sin7"]), "pct": fmt_pct(v["pct"], 1)}

    def f_horas_altas(v):
        return {"tot": fmt_miles(v["tot"]), "48_60": fmt_miles(v["48_60"]), "60_72": fmt_miles(v["60_72"]),
                "m72": fmt_miles(v["m72"]), "altas": fmt_miles(v["altas"]), "pct": fmt_pct(v["pct"], 1)}

    def f_he_obrero(v):
        return {"tot_obr": fmt_miles(v["tot_obr"]), "he_total": fmt_dec(v["he_total"]),
                "area": {a: fmt_dec(h) for a, h in v["area"].items()}}

    def f_horas_prom(v):
        return {"total": fmt_dec(v["total"]), "area": {a: fmt_dec(h) for a, h in v["area"].items()}}

    def f_rango_edades(v):
        return {"tot": fmt_miles(v["tot"]),
                "rangos": {etiq: {"n": fmt_miles(r["n"]), "pct": fmt_pct(r["pct"], 1)} for etiq, r in v["rangos"].items()}}

    def f_ratio_area(v):
        return {a: {"sup": fmt_miles(r["sup"]), "trab": fmt_miles(r["trab"]), "ratio": fmt_dec(r["ratio"]),
                     "horas": fmt_dec(r["horas"]), "extra": fmt_dec(r["extra"])} for a, r in v.items()}

    def f_antiguedad(v):
        return {r: {"n": fmt_miles(v[r]["n"]), "pct": fmt_pct(v[r]["pct"]), "pct_num": round(v[r]["pct"] * 100, 1)}
                for r in v}

    def f_genero(v):
        return {"m": fmt_miles(v["m"]), "f": fmt_miles(v["f"]), "pm": fmt_pct(v["pm"]), "pf": fmt_pct(v["pf"]),
                "pm_num": round(v["pm"] * 100, 1), "pf_num": round(v["pf"] * 100, 1), "tot": fmt_miles(v["tot"])}

    def f_jubilacion(v):
        return {"60_65": fmt_miles(v["60_65"]), "65_69": fmt_miles(v["65_69"]), "69": fmt_miles(v["69"]),
                "alerta": fmt_miles(v["alerta"]), "p": fmt_pct(v["p"]), "tot": fmt_miles(v["tot"])}

    return {
        "ausentismo_pt": {f: f_ausentismo_pt(v) for f, v in dia["ausentismo_pt"].items()},
        "sin_descanso": {f: f_sin_descanso(v) for f, v in dia["sin_descanso"].items()},
        "horas_altas": {f: f_horas_altas(v) for f, v in dia["horas_altas"].items()},
        "he_obrero": {f: f_he_obrero(v) for f, v in dia["he_obrero"].items()},
        "horas_prom_detalle": {f: f_horas_prom(v) for f, v in dia["horas_prom_detalle"].items()},
        "rango_edades": {f: f_rango_edades(v) for f, v in dia["rango_edades"].items()},
        "ratio_area": {f: f_ratio_area(v) for f, v in dia["ratio_area"].items()},
        "antiguedad": {f: f_antiguedad(v) for f, v in dia["antiguedad"].items()},
        "genero_fundo": {f: f_genero(v) for f, v in dia["genero_fundo"].items()},
        "jubilacion": {f: f_jubilacion(v) for f, v in dia["jubilacion"].items()},
    }


def generar_html(lista_kpis, hoy):
    semanas = sorted(x["semana"] for x in lista_kpis)
    default = semanas[-1]
    bloques = "".join(
        bloque_semana(k) + bloque_asistencia(k) + bloque_ratios(k) + bloque_alertas(k)
        for k in lista_kpis
    )

    meta = {str(k["semana"]): {
        "anio": k["anio"], "fecha": k["fecha_max"],
        "asis": fmt_miles(k["asistencia"]),
        "rot": fmt_pct(k["rotacion"], 1) if k["rotacion"] is not None else "0,0%",
        "sup": fmt_miles(k["supervisores"]),
        "ext": fmt_miles(k["extranjeros"]),
        "aus": fmt_pct(k["ausentismo"], 1) if k["ausentismo"] is not None else "0,0%",
        "hrs": (fmt_dec(k["horas_prom"]) + "h") if k["horas_prom"] is not None else "-",
        "chart": {c: [round(float(v), 2) for v in k["chart"][c]] for c in k["chart"]},
        "dias": {iso: {
            "label": dia["label"],
            "day_index": pd.Timestamp(iso).weekday(),  # 0=Lun, ..., 6=Dom
            "asis": fmt_miles(dia["asistencia"]),
            "sup": fmt_miles(dia["supervisores"]),
            "ext": fmt_miles(dia["extranjeros"]),
            "hrs": (fmt_dec(dia["horas_prom"]) + "h") if dia["horas_prom"] is not None else "-",
            "extra": _fmt_dia_extra(dia),
        } for iso, dia in k["dias"].items()},
    } for k in lista_kpis}

    opciones = "".join(f"<option value='{s}'{' selected' if s == default else ''}>{s}</option>" for s in semanas)
    labels_js = json.dumps([ETIQ_CHART[f] for f in FUNDOS])

    return f"""<!DOCTYPE html>
<html lang='es'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Dash Indicadores RRHH</title>
<link href='https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap' rel='stylesheet'>
<script src='https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js'></script>
<script src='https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js'></script>
<style>
{CSS_HERO}
body{{font-family:'Inter','Segoe UI',sans-serif;background:radial-gradient(1100px 480px at 82% -8%,#EDF0E8 0%,rgba(237,240,232,0) 60%),var(--papel);color:var(--tinta);margin:0;padding:20px 22px;-webkit-font-smoothing:antialiased}}
select{{font-family:inherit;font-size:.92em;font-weight:500;color:var(--tinta);padding:8px 30px 8px 12px;border-radius:9px;border:1px solid var(--linea);background-color:#fff;cursor:pointer;appearance:none;-webkit-appearance:none;background-image:url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%236F7871' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 11px center;transition:border-color .15s,box-shadow .15s}}
select:hover{{border-color:#C4CBBD}}
select:focus{{outline:none;border-color:var(--verde2);box-shadow:0 0 0 3px rgba(45,106,79,.12)}}
.tab-btn{{font-family:inherit;font-size:.88em;font-weight:600;letter-spacing:.02em;padding:9px 18px;border:none;border-radius:9px;background:transparent;color:var(--gris);cursor:pointer;transition:color .15s,background .15s}}
.tab-btn:hover{{color:var(--verde)}}
.tab-btn.active{{background:var(--verde);color:#fff;box-shadow:0 2px 8px rgba(27,67,50,.25)}}
.kpi-card{{transition:transform .18s ease,box-shadow .18s ease}}
.kpi-card:hover{{transform:translateY(-2px);box-shadow:0 2px 4px rgba(20,30,25,.05),0 12px 28px rgba(20,30,25,.09)}}
@media(max-width:1100px){{.grid-kpi{{grid-template-columns:repeat(3,1fr)!important}}.grid-areas,.grid-info,.grid-charts{{grid-template-columns:repeat(2,1fr)!important}}}}
@media(max-width:640px){{.grid-kpi{{grid-template-columns:repeat(2,1fr)!important}}.grid-areas,.grid-info,.grid-charts{{grid-template-columns:1fr!important}}.tab-btn{{padding:8px 10px;font-size:.8em}}}}
</style>
</head>
<body>
{html_hero(None)}
<div style='background:#fff;padding:13px 18px;border-radius:12px;border:1px solid var(--linea);margin:16px 0 14px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05);display:flex;gap:22px;align-items:center;flex-wrap:wrap'>
<label style='font-weight:600;font-size:.85em;color:var(--gris);text-transform:uppercase;letter-spacing:.08em'>Semana&nbsp;&nbsp;<select id='selSemana'>{opciones}</select></label>
<label style='font-weight:600;font-size:.85em;color:var(--gris);text-transform:uppercase;letter-spacing:.08em'>Fecha&nbsp;&nbsp;<select id='selFecha'><option value='todas'>Todas</option></select></label>
<span style='color:var(--gris);font-size:.82em;margin-left:auto'>&Uacute;ltima actualizaci&oacute;n del archivo: <span style='font-weight:600;color:var(--tinta)'>{hoy.strftime("%d/%m/%Y")}</span></span>
</div>
<div style='display:inline-flex;gap:4px;flex-wrap:wrap;background:#fff;border:1px solid var(--linea);border-radius:12px;padding:4px;box-shadow:0 1px 2px rgba(20,30,25,.04)'>
<button class='tab-btn active' data-tab='resumen'>Resumen</button>
<button class='tab-btn' data-tab='asistencia'>Grupo Asistencia</button>
<button class='tab-btn' data-tab='ratios'>Grupo Ratios</button>
<button class='tab-btn' data-tab='alertas'>Grupo Alertas &amp; Demografia</button>
</div>
{bloques}
<script>
var META = {json.dumps(meta, ensure_ascii=False)};
var LABELS = {labels_js};
var creados = {{}};

function crearCharts(sem) {{
  if (creados[sem] || typeof Chart === 'undefined') return;
  var m = META[sem].chart;
  var f = 'Inter,Segoe UI,sans-serif';
  try {{ if (!Chart.registry.plugins.get('datalabels')) Chart.register(ChartDataLabels); }} catch(e) {{}}
  Chart.defaults.color = '#6F7871';
  Chart.defaults.font.family = f;
  var base = {{responsive:true, maintainAspectRatio:false,
    plugins:{{legend:{{labels:{{font:{{family:f,size:10}}}}}}, datalabels:{{font:{{family:f,size:10,weight:'600'}}}}}}}};

  new Chart(document.getElementById('chartAsist_w'+sem), {{
    data: {{ labels: LABELS, datasets: [
      {{type:'line', label:'Horas Efectivas', data:m.hefect, borderColor:'#C9A227', backgroundColor:'#C9A227', borderWidth:2, pointRadius:3,
        yAxisID:'y2', tension:.3, datalabels:{{display:false}}}},
      {{type:'bar', label:'Asistencia', data:m.asist,
        backgroundColor:['#B3453C','#31678F','#1B4332'], borderRadius:6, maxBarThickness:52,
        datalabels:{{anchor:'center',align:'center',color:'#fff'}}}} ]}},
    options: Object.assign({{}}, base, {{scales:{{y:{{beginAtZero:true}}, y2:{{position:'right',grid:{{display:false}},ticks:{{display:false}}}}}}}})
  }});
  new Chart(document.getElementById('chartRot_w'+sem), {{
    type:'bar',
    data: {{ labels: LABELS, datasets:[{{label:'Rotaci\\u00f3n %', data:m.rot,
      backgroundColor:['#B3453C','#31678F','#1B4332'], borderRadius:6, maxBarThickness:52,
      datalabels:{{anchor:'end',align:'top',formatter:function(v){{return v.toFixed(1)+'%';}}}}}}]}},
    options: Object.assign({{}}, base, {{plugins:Object.assign({{}},base.plugins,{{legend:{{display:false}}}}),
      scales:{{y:{{beginAtZero:true,ticks:{{callback:function(v){{return v+'%';}}}}}}}}}})
  }});
  new Chart(document.getElementById('chartHrs_w'+sem), {{
    type:'bar',
    data: {{ labels: LABELS, datasets:[{{label:'Horas', data:m.hprom,
      backgroundColor:['#B3453C','#31678F','#1B4332'], borderRadius:6, maxBarThickness:52,
      datalabels:{{anchor:'center',align:'center',color:'#fff',formatter:function(v){{return v.toFixed(1)+'h';}}}}}}]}},
    options: Object.assign({{}}, base, {{plugins:Object.assign({{}},base.plugins,{{legend:{{display:false}}}}),
      scales:{{y:{{beginAtZero:true}}}}}})
  }});
  new Chart(document.getElementById('chartGen_w'+sem), {{
    type:'bar',
    data: {{ labels: LABELS, datasets:[
      {{label:'Hombres', data:m.genM, backgroundColor:'#31678F', borderRadius:4,
        datalabels:{{color:'#fff',formatter:function(v){{return v.toFixed(1)+'%';}}}}}},
      {{label:'Mujeres', data:m.genF, backgroundColor:'#C0798B', borderRadius:4,
        datalabels:{{color:'#fff',formatter:function(v){{return v.toFixed(1)+'%';}}}}}} ]}},
    options: Object.assign({{}}, base, {{scales:{{x:{{stacked:true}},y:{{stacked:true,max:100,
      ticks:{{callback:function(v){{return v+'%';}}}}}}}}}})
  }});
  creados[sem] = true;
}}

var tabActual = 'resumen';
var semActual = '{default}';
var fechaActual = 'todas';

function refrescarVisibilidad() {{
  document.querySelectorAll('.tab-content').forEach(function(b) {{
    b.style.display = (b.dataset.tab === tabActual && b.dataset.sem === String(semActual)) ? 'block' : 'none';
  }});
  var variantes = document.querySelectorAll('.fecha-variant');
  variantes.forEach(function(b) {{
    var debe_mostrar = (b.dataset.fecha === fechaActual);
    b.style.display = debe_mostrar ? 'contents' : 'none';
  }});
}}

function actualizarKpis(sem, fecha) {{
  var m = META[sem];
  var vals = (fecha === 'todas') ? m : m.dias[fecha];
  if (!vals) return;
  document.querySelectorAll('.js-sem').forEach(function(e){{e.textContent = sem;}});
  document.querySelectorAll('.js-anio').forEach(function(e){{e.textContent = m.anio;}});
  document.querySelectorAll('.js-fecha').forEach(function(e){{e.textContent = (fecha === 'todas') ? m.fecha : vals.label;}});
  var kAsis=document.getElementById('kpi_asis_w'+sem);
  if (kAsis) {{
    kAsis.textContent = vals.asis;
    document.getElementById('kpi_sup_w'+sem).textContent  = vals.sup;
    document.getElementById('kpi_ext_w'+sem).textContent  = vals.ext;
    document.getElementById('kpi_hrs_w'+sem).textContent  = vals.hrs;
    // Rotación y Ausentismo son medidas semanales; no aplican a un día puntual
    document.getElementById('kpi_rot_w'+sem).textContent  = (fecha === 'todas') ? m.rot : '-';
    document.getElementById('kpi_aus_w'+sem).textContent  = (fecha === 'todas') ? m.aus : '-';
  }}
}}

function poblarSelectFecha(sem) {{
  var sel = document.getElementById('selFecha');
  var dias = META[sem].dias;
  sel.innerHTML = "<option value='todas'>Todas</option>";
  Object.keys(dias).sort().forEach(function(iso) {{
    var opt = document.createElement('option');
    opt.value = iso; opt.textContent = dias[iso].label;
    sel.appendChild(opt);
  }});
  sel.value = 'todas';
  fechaActual = 'todas';
  destacarDiaEnAsistencia(sem, 'todas');  // Limpiar destacado
  actualizarGridsExtra(sem, 'todas');  // Restaurar grids a la vista semanal completa
}}

function mostrarSemana(sem) {{
  semActual = sem;
  poblarSelectFecha(sem);
  actualizarKpis(sem, 'todas');
  refrescarVisibilidad();
  if (tabActual === 'resumen') crearCharts(sem);
}}

function destacarDiaEnAsistencia(sem, fecha) {{
  // Limpia el destacado anterior
  document.querySelectorAll('[id^="header_asist_"]').forEach(function(h) {{
    h.style.background = '';
    h.style.fontWeight = 'normal';
  }});
  
  // Si es "todas", no destaca nada
  if (fecha === 'todas') return;
  
  // Obtén el day_index de META
  var m = META[sem];
  if (!m || !m.dias || !m.dias[fecha]) return;
  var dayIdx = m.dias[fecha].day_index;
  
  // Destaca los headers del día elegido en todas las tablas de asistencia
  ['Hp', 'Bh', 'Cao'].forEach(function(fundo) {{
    var el = document.getElementById('header_asist_'+fundo+'_day'+dayIdx);
    if (el) {{
      el.style.background = '#FFF3CD';  // Amarillo suave
      el.style.fontWeight = '700';
    }}
  }});
}}

var GRIDS_ORIG = {{}};
document.querySelectorAll('[id^="grid_"]').forEach(function(el){{ GRIDS_ORIG[el.id] = el.innerHTML; }});

var FKEYS4 = ['Hp','Bh','Cao','Con'];
var NOMBRES4 = ['Fundo HP','Fundo BH','Fundo Cao','Consolidado'];
var COLORS4 = ['#B3453C','#31678F','#1B4332','#5B6660'];
var AREAS6 = ['Cosecha','Labores','Empaque','Riego y Fert.','Sanidad','Otros'];
var RANGOS_ANTIG_JS = ['[0 - 1 a\\u00f1o]','[1 - 3 a\\u00f1os]','[3 - 5 a\\u00f1os]','[> 5 a\\u00f1os]'];

function cardWrap(borde, inner, extraStyle) {{
  return "<div style=\\"background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:16px;border-top:4px solid "+borde+";box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)"+(extraStyle||'')+"\\">"+inner+"</div>";
}}

function buildSimpleCards(dataDict, bigKey, bigLabel, rows) {{
  var html = '';
  for (var i=0;i<4;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=COLORS4[i];
    var v = dataDict[fkey];
    if (!v) {{ html += cardWrap(borde, "<div style='color:#999'>Sin datos</div>"); continue; }}
    var rowsHtml = rows.map(function(r) {{
      return "<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>"+r[0]+"</span><span style='font-weight:600;color:#232B26'>"+v[r[1]]+"</span></div>";
    }}).join('');
    var big = bigKey ? ("<div style='font-size:28px;font-weight:500;color:#161D19;margin-bottom:2px'>"+v[bigKey]+"</div><div style='font-size:13px;color:#98A099;margin-bottom:10px'>"+bigLabel+"</div>") : '';
    var inner = "<div style='font-size:15px;font-weight:500;color:"+borde+";margin-bottom:8px'>"+titulo+"</div>"+big+
      "<div style='border-top:1px solid #ECEEE8;padding-top:8px;display:flex;flex-direction:column;gap:5px;font-size:13px'>"+rowsHtml+"</div>";
    html += cardWrap(borde, inner);
  }}
  return html;
}}

function buildAusentismoDia(dia) {{
  var html = '';
  for (var i=0;i<4;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=COLORS4[i];
    var a = dia.extra.ausentismo_pt[fkey];
    if (!a) {{ html += cardWrap(borde, "<div style='color:#999'>Sin datos</div>"); continue; }}
    var colorAus = (parseInt(a.ausentes) > 0) ? '#B3453C' : '#2D6A4F';
    var inner = "<div style='font-size:15px;font-weight:600;color:"+borde+";margin-bottom:8px'>"+titulo+"</div>"+
      "<div style='font-size:.8em;color:#6F7871;margin-bottom:8px'>Base (Lunes): <span style='font-weight:700;color:#232B26'>"+a.base+"</span></div>"+
      "<div style='border-top:1px solid #ECEEE8;padding-top:8px;display:flex;justify-content:space-between;font-weight:700'>"+
      "<span>Ausentes este d\\u00eda</span><span style='color:"+colorAus+"'>"+a.ausentes+" ("+a.pct+")</span></div>";
    html += cardWrap(borde, inner);
  }}
  return html;
}}

function buildHeObrero(dia) {{
  var html = '';
  for (var i=0;i<4;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=COLORS4[i];
    var h = dia.extra.he_obrero[fkey];
    var filas = AREAS6.map(function(a) {{
      return "<tr style='border-bottom:1px solid #ECEEE8'><td style='padding:3px 6px;color:#232B26;font-weight:600'>"+a+"</td><td style='padding:3px 6px;text-align:right;color:#5B6660'>"+(h.area[a]||'0,0')+"</td></tr>";
    }}).join('');
    var top = cardWrap(borde,
      "<div style='font-size:15px;font-weight:500;color:"+borde+";margin-bottom:8px'>&#128337; "+titulo+"</div>"+
      "<div style='font-size:28px;font-weight:500;color:#161D19;margin-bottom:2px'>"+h.he_total+"</div>"+
      "<div style='font-size:13px;color:#98A099;margin-bottom:10px'>Total HE Obrero</div>"+
      "<div style='border-top:1px solid #ECEEE8;padding-top:8px;font-size:13px'><span style='color:#6F7871'># Obreros</span> <span style='font-weight:600'>"+h.tot_obr+"</span></div>",
      ";margin-bottom:12px");
    var tabla = "<div style=\\"background:#fff;border:1px solid #E3E6DE;border-radius:10px;padding:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 6px 18px rgba(20,30,25,.05)\\">"+
      "<div style='font-size:11px;font-weight:700;color:#98A099;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px'>HE por &Aacute;rea</div>"+
      "<table style='width:100%;border-collapse:collapse;font-size:12px'>"+filas+"</table></div>";
    html += top + tabla;
  }}
  return html;
}}

function buildHorasPromReal(dia) {{
  var html = '';
  for (var i=0;i<4;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=COLORS4[i];
    var h = dia.extra.horas_prom_detalle[fkey];
    var filas = AREAS6.map(function(a) {{
      return "<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>"+a+"</span><span style='font-weight:600;color:#232B26'>"+(h.area[a]||'0,0')+"</span></div>";
    }}).join('');
    var inner = "<div style='font-size:15px;font-weight:500;color:"+borde+";margin-bottom:8px'>&#128337; "+titulo+"</div>"+
      "<div style='font-size:28px;font-weight:500;color:#161D19;margin-bottom:2px'>"+h.total+"</div>"+
      "<div style='font-size:13px;color:#98A099;margin-bottom:10px'>Horas promedio</div>"+
      "<div style='border-top:1px solid #ECEEE8;padding-top:8px;display:flex;flex-direction:column;gap:4px;font-size:12px'>"+filas+"</div>";
    html += cardWrap(borde, inner);
  }}
  return html;
}}

function buildRangoEdades(dia) {{
  var html = '';
  for (var i=0;i<4;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=COLORS4[i];
    var r = dia.extra.rango_edades[fkey];
    var filas = Object.keys(r.rangos).map(function(etiq) {{
      var v = r.rangos[etiq];
      return "<div style='display:flex;justify-content:space-between'><span style='color:#6F7871'>"+etiq+" a\\u00f1os</span><span style='font-weight:600;color:#232B26'>"+v.n+" <span style='color:#98A099'>("+v.pct+")</span></span></div>";
    }}).join('');
    var inner = "<div style='font-size:14px;font-weight:600;color:"+borde+";margin-bottom:12px'>"+titulo+"</div>"+
      "<div style='display:flex;flex-direction:column;gap:6px;font-size:12px'>"+filas+
      "<div style='border-top:1px solid #ECEEE8;padding-top:6px;display:flex;justify-content:space-between;font-weight:700'><span style='color:#232B26'>Total</span><span style='color:"+borde+"'>"+r.tot+"</span></div></div>";
    html += cardWrap(borde, inner);
  }}
  return html;
}}

function buildRatioArea(dia) {{
  var html = '';
  var colores3 = ['#B3453C','#31678F','#1B4332'];
  for (var i=0;i<3;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=colores3[i];
    var r = dia.extra.ratio_area[fkey];
    var filas = AREAS6.map(function(a, idx) {{
      var v = r[a] || {{sup:'0',trab:'0',ratio:'0,0',horas:'0,0',extra:'0,0'}};
      var bg = (idx % 2) ? "background:#FAFAF7" : "";
      return "<tr style='"+bg+"'><td style='padding:6px 8px'>"+a+"</td><td style='padding:6px 8px;text-align:center'>"+v.sup+"</td>"+
        "<td style='padding:6px 8px;text-align:center'>"+v.trab+"</td><td style='padding:6px 8px;text-align:center;font-weight:600;color:"+borde+"'>"+v.ratio+"</td>"+
        "<td style='padding:6px 8px;text-align:center'>"+v.horas+"</td><td style='padding:6px 8px;text-align:center'>"+v.extra+"</td></tr>";
    }}).join('');
    html += "<div style=\\"background:#fff;padding:16px;border-radius:12px;border-left:5px solid "+borde+";box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06);margin-bottom:16px\\">"+
      "<div style='font-size:1.1em;font-weight:700;margin-bottom:10px'>&#128101; # de Supervisores / Por Labor - "+titulo+"</div>"+
      "<table style='width:100%;border-collapse:collapse;font-size:.85em'>"+
      "<tr style='background:#FAFAF7;font-weight:600'><td style='padding:6px 8px'>Labor</td><td style='padding:6px 8px;text-align:center'>Supervisores</td>"+
      "<td style='padding:6px 8px;text-align:center'>Trabajadores</td><td style='padding:6px 8px;text-align:center'>Ratio</td>"+
      "<td style='padding:6px 8px;text-align:center'>Horas</td><td style='padding:6px 8px;text-align:center'>H. Extra</td></tr>"+filas+"</table></div>";
  }}
  return html;
}}

function buildAntiguedad(dia) {{
  var html = '';
  for (var i=0;i<4;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=COLORS4[i];
    var a = dia.extra.antiguedad[fkey];
    var filas = RANGOS_ANTIG_JS.map(function(r) {{
      var v = a[r] || {{n:'0',pct:'0,0%',pct_num:0}};
      return "<div style='margin-bottom:8px'><div style='display:flex;justify-content:space-between;font-size:.85em;margin-bottom:2px'>"+
        "<span>"+r+"</span><span style='font-weight:600'>"+v.n+" ("+v.pct+")</span></div>"+
        "<div style='background:#ECEEE8;border-radius:4px;height:8px'><div style='background:"+borde+";width:"+v.pct_num+"%;height:8px;border-radius:4px'></div></div></div>";
    }}).join('');
    html += "<div style=\\"background:#fff;padding:16px;border-radius:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06);margin-bottom:16px\\">"+
      "<div style='font-size:1.05em;font-weight:700;margin-bottom:12px'>"+titulo+"</div>"+filas+"</div>";
  }}
  return html;
}}

function buildComposicion(dia) {{
  var html = '';
  for (var i=0;i<4;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=COLORS4[i];
    var g = dia.extra.genero_fundo[fkey];
    html += "<div style=\\"background:#fff;padding:16px;border-radius:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06)\\">"+
      "<div style='font-size:1.05em;font-weight:700;margin-bottom:10px;color:"+borde+"'>"+titulo+"</div>"+
      "<div style='display:flex;justify-content:space-between;font-size:.9em;margin-bottom:6px'><span>&#9794; Hombres</span><span style='font-weight:600'>"+g.m+" ("+g.pm+")</span></div>"+
      "<div style='background:#ECEEE8;border-radius:4px;height:10px;margin-bottom:10px'><div style='background:#31678F;width:"+g.pm_num+"%;height:10px;border-radius:4px'></div></div>"+
      "<div style='display:flex;justify-content:space-between;font-size:.9em;margin-bottom:6px'><span>&#9792; Mujeres</span><span style='font-weight:600'>"+g.f+" ("+g.pf+")</span></div>"+
      "<div style='background:#ECEEE8;border-radius:4px;height:10px'><div style='background:#B3453C;width:"+g.pf_num+"%;height:10px;border-radius:4px'></div></div></div>";
  }}
  return html;
}}

function buildJubilacion(dia) {{
  var html = '';
  var etiquetas = [['60-65 a\\u00f1os','60_65'],['65-69 a\\u00f1os','65_69'],['&gt; 69 a\\u00f1os','69']];
  for (var i=0;i<4;i++) {{
    var fkey=FKEYS4[i], titulo=NOMBRES4[i], borde=COLORS4[i];
    var j = dia.extra.jubilacion[fkey];
    var filas = etiquetas.map(function(e) {{
      return "<tr><td style='padding:5px 8px'>"+e[0]+"</td><td style='padding:5px 8px;text-align:center'>"+j[e[1]]+"</td></tr>";
    }}).join('');
    html += "<div style=\\"background:#fff;padding:16px;border-radius:12px;box-shadow:0 1px 2px rgba(20,30,25,.04),0 8px 24px rgba(20,30,25,.06);border-top:4px solid "+borde+"\\">"+
      "<div style='font-size:1.05em;font-weight:700;margin-bottom:8px'>"+titulo+"</div>"+
      "<table style='width:100%;border-collapse:collapse;font-size:.85em'>"+filas+
      "<tr style='background:#F8F1F0;font-weight:700'><td style='padding:6px 8px;color:#B3453C'>Total Alerta</td><td style='padding:6px 8px;text-align:center;color:#B3453C'>"+j.alerta+" ("+j.p+")</td></tr></table></div>";
  }}
  return html;
}}

function actualizarGridsExtra(sem, fecha) {{
  var m = META[sem];
  if (!m) return;
  var sid = 'w'+sem;
  var ids = ['grid_aus_','grid_sd_','grid_ha_','grid_he_','grid_hp_','grid_edad_','grid_ra_','grid_antig_','grid_comp_','grid_jub_'];
  if (fecha === 'todas') {{
    ids.forEach(function(pref) {{
      var el = document.getElementById(pref+sid);
      if (el && GRIDS_ORIG[pref+sid] !== undefined) el.innerHTML = GRIDS_ORIG[pref+sid];
    }});
    return;
  }}
  var dia = m.dias[fecha];
  if (!dia || !dia.extra) return;
  var setters = {{
    grid_aus_: buildAusentismoDia, grid_sd_: function(d) {{ return buildSimpleCards(d.extra.sin_descanso, 'con7', 'Con cierre de semana (7 d\\u00edas)',
      [['Total','tot'],['Sin cierre','sin7'],['%','pct']]); }},
    grid_ha_: function(d) {{ return buildSimpleCards(d.extra.horas_altas, 'altas', 'Con horas altas',
      [['Total','tot'],['48-60 hrs','48_60'],['60-72 hrs','60_72'],['&gt; 72 hrs','m72'],['%','pct']]); }},
    grid_he_: buildHeObrero, grid_hp_: buildHorasPromReal, grid_edad_: buildRangoEdades,
    grid_ra_: buildRatioArea, grid_antig_: buildAntiguedad, grid_comp_: buildComposicion, grid_jub_: buildJubilacion,
  }};
  ids.forEach(function(pref) {{
    var el = document.getElementById(pref+sid);
    if (!el) return;
    el.innerHTML = setters[pref](dia);
  }});
}}

function mostrarFecha(fecha) {{
  fechaActual = fecha;
  actualizarKpis(semActual, fecha);
  destacarDiaEnAsistencia(semActual, fecha);
  actualizarGridsExtra(semActual, fecha);
  refrescarVisibilidad();
}}

function mostrarTab(tab) {{
  tabActual = tab;
  document.querySelectorAll('.tab-btn').forEach(function(b) {{
    b.classList.toggle('active', b.dataset.tab === tab);
  }});
  refrescarVisibilidad();
  if (tab === 'resumen') crearCharts(semActual);
}}

document.getElementById('selSemana').addEventListener('change', function() {{ mostrarSemana(this.value); }});
document.getElementById('selFecha').addEventListener('change', function() {{ mostrarFecha(this.value); }});
document.querySelectorAll('.tab-btn').forEach(function(b) {{
  b.addEventListener('click', function() {{ mostrarTab(b.dataset.tab); }});
}});
mostrarSemana('{default}');
mostrarTab('resumen');
</script>
</body>
</html>"""


# ────────────────────────────── main ──────────────────────────────
def main():
    hoy = date.today()
    print("Leyendo Data Mes...")
    df = cargar_data_mes(CARPETA_DATA_MES)
    print(f"  {len(df):,} filas de tareo")
    print("Leyendo Data Personal...")
    dp = cargar_personal(ARCHIVO_PERSONAL, HOJA_PERSONAL)
    print(f"  {len(dp):,} trabajadores")
    df = preparar(df, dp)

    semanas = sorted(int(s) for s in df["Semana_Correcta"].dropna().unique())
    print(f"Semanas detectadas: {semanas}")
    kpis = [kpis_semana(df, s, hoy) for s in semanas]

    html = generar_html(kpis, hoy)
    Path(ARCHIVO_SALIDA).write_text(html, encoding="utf-8")
    print(f"Generado: {ARCHIVO_SALIDA} ({len(html)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
