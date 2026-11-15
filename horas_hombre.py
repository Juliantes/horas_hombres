from __future__ import annotations
import hashlib, os, sqlite3
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st

from sap_import import ImportResult, extraer_tabla

DATA_DIR = Path(os.environ.get("HH_DATA_DIR", Path(__file__).resolve().parent))
DB = str(DATA_DIR / "horas_hombre.db")
XLS_MAESTRO = str(DATA_DIR / "confirmaciones.xlsx")
MESES = [f"{m:02d}" for m in range(1, 13)]
COLS = ["id","orden","operacion","ubicacion","puesto","dur_real","trbjo_real",
        "horas_intervalo","inicio","fin","clase_act","fe_contab","archivo","cargado","row_hash"]
CRITERIOS = {"fin":"Mes de fin real","inicio":"Mes de inicio real",
             "contab":"Mes de Fe.contab.","prorrateo":"Prorrateo por días"}
GRUPO_SIN = "(sin grupo)"

# ---------------------------------------------------------------- base
def _migrar_disponibilidad(c: sqlite3.Connection) -> None:
    cols = {row[1] for row in c.execute("PRAGMA table_info(disponibilidad)").fetchall()}
    if not cols:
        c.execute("""CREATE TABLE disponibilidad(
            periodo TEXT NOT NULL, grupo TEXT NOT NULL DEFAULT '',
            horas REAL, objetivo REAL, nota TEXT, actualizado TEXT,
            PRIMARY KEY (periodo, grupo))""")
        return
    if "grupo" in cols:
        return
    c.execute("""CREATE TABLE disponibilidad_new(
        periodo TEXT NOT NULL, grupo TEXT NOT NULL DEFAULT '',
        horas REAL, objetivo REAL, nota TEXT, actualizado TEXT,
        PRIMARY KEY (periodo, grupo))""")
    c.execute(
        "INSERT INTO disponibilidad_new (periodo, grupo, horas, objetivo, nota, actualizado) "
        "SELECT periodo, '', horas, objetivo, nota, actualizado FROM disponibilidad"
    )
    c.execute("DROP TABLE disponibilidad")
    c.execute("ALTER TABLE disponibilidad_new RENAME TO disponibilidad")

@st.cache_resource
def conectar():
    c = sqlite3.connect(DB, check_same_thread=False)
    c.execute("""CREATE TABLE IF NOT EXISTS confirmaciones(
        id TEXT PRIMARY KEY, orden TEXT, operacion TEXT, ubicacion TEXT, puesto TEXT,
        dur_real REAL, trbjo_real REAL, horas_intervalo REAL, inicio TEXT, fin TEXT,
        clase_act TEXT, archivo TEXT, cargado TEXT, row_hash TEXT)""")
    for ddl in ("ALTER TABLE confirmaciones ADD COLUMN fe_contab TEXT",):
        try: c.execute(ddl)
        except sqlite3.OperationalError: pass
    c.execute("""CREATE TABLE IF NOT EXISTS disponibilidad(
        periodo TEXT NOT NULL, grupo TEXT NOT NULL DEFAULT '',
        horas REAL, objetivo REAL, nota TEXT, actualizado TEXT,
        PRIMARY KEY (periodo, grupo))""")
    _migrar_disponibilidad(c)
    c.commit(); return c

def leer_sql(sql, params=()):
    return pd.read_sql_query(sql, conectar(), params=params)

def normalizar_periodo(val) -> str | None:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return None
    if len(s) >= 7 and s[4:5] == "-":
        try:
            datetime.strptime(s[:7], "%Y-%m")
            return s[:7]
        except ValueError:
            pass
    dt = pd.to_datetime(val, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%Y-%m")
    return None

def normalizar_grupo(val) -> str | None:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return None
    return s

def puestos_en_confirmaciones() -> list[str]:
    df = leer_sql(
        "SELECT DISTINCT puesto FROM confirmaciones "
        "WHERE puesto IS NOT NULL AND TRIM(puesto) <> '' ORDER BY puesto"
    )
    return df["puesto"].tolist()

def guardar_disponibilidad(df: pd.DataFrame) -> tuple[int, int, list[str]]:
    """Upsert por (periodo, grupo). Devuelve (nuevas, actualizadas, errores)."""
    c = conectar()
    ahora = datetime.now().isoformat(timespec="seconds")
    sql = (
        "INSERT INTO disponibilidad VALUES (?,?,?,?,?,?) "
        "ON CONFLICT(periodo, grupo) DO UPDATE SET "
        "horas=excluded.horas, objetivo=excluded.objetivo, nota=excluded.nota, "
        "actualizado=excluded.actualizado"
    )
    nuevas = actualizadas = 0
    errores: list[str] = []
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        periodo = normalizar_periodo(r.get("periodo"))
        grupo = normalizar_grupo(r.get("grupo"))
        if periodo is None:
            if pd.notna(r.get("periodo")) and str(r.get("periodo")).strip():
                errores.append(f"Fila {i}: periodo inválido «{r.get('periodo')}» (usar YYYY-MM)")
            continue
        if grupo is None:
            errores.append(f"Fila {i} ({periodo}): falta grupo técnico")
            continue
        if pd.isna(r.get("horas")):
            errores.append(f"Fila {i} ({periodo} · {grupo}): faltan horas disponibles")
            continue
        prev = c.execute(
            "SELECT 1 FROM disponibilidad WHERE periodo=? AND grupo=?", (periodo, grupo)
        ).fetchone()
        horas = float(r["horas"])
        obj = 85.0 if pd.isna(r.get("objetivo")) else float(r["objetivo"])
        nota = "" if pd.isna(r.get("nota")) else str(r["nota"])
        c.execute(sql, (periodo, grupo, horas, obj, nota, ahora))
        if prev:
            actualizadas += 1
        else:
            nuevas += 1
    c.commit()
    return nuevas, actualizadas, errores

def leer_disponibilidad_archivo(raw: bytes, nombre: str) -> pd.DataFrame:
    low = nombre.lower()
    buf = BytesIO(raw)
    if low.endswith(".xlsx"):
        src = pd.read_excel(buf, engine="openpyxl")
    elif low.endswith(".xls"):
        src = pd.read_excel(buf, engine="xlrd")
    else:
        src = pd.read_csv(buf)
    alias = {
        "mes": "periodo", "año-mes": "periodo", "anio-mes": "periodo", "year-month": "periodo",
        "horas disponibles": "horas", "horas_hombre": "horas", "disponibles": "horas",
        "objetivo %": "objetivo", "objetivo_pct": "objetivo", "meta": "objetivo",
        "puesto": "grupo", "ptotrreal": "grupo", "grupo técnico": "grupo", "grupo tecnico": "grupo",
    }
    cols = {c: alias.get(c.strip().lower(), c.strip().lower()) for c in src.columns}
    out = src.rename(columns=cols)
    for req in ("periodo", "grupo", "horas"):
        if req not in out.columns:
            raise ValueError(
                f"Falta la columna «{req}». Esperado: periodo, grupo, horas; "
                "opcional objetivo, nota"
            )
    if "objetivo" not in out.columns:
        out["objetivo"] = 85.0
    if "nota" not in out.columns:
        out["nota"] = ""
    return out[["periodo", "grupo", "horas", "objetivo", "nota"]]

# ---------------------------------------------------------------- persistencia
def guardar(df: pd.DataFrame):
    c = conectar(); n = a = i = 0
    ahora = datetime.now().isoformat(timespec="seconds")
    sql = (f"INSERT INTO confirmaciones ({','.join(COLS)}) VALUES ({','.join('?'*len(COLS))}) "
           f"ON CONFLICT(id) DO UPDATE SET " +
           ",".join(f"{k}=excluded.{k}" for k in COLS if k != "id"))
    for _, r in df.iterrows():
        base = {k: r.get(k) for k in COLS if k not in ("cargado","row_hash","archivo")}
        base["inicio"] = str(r["inicio"]) if r.get("inicio") else None
        base["fin"] = str(r["fin"]) if r.get("fin") else None
        rh = hashlib.md5("|".join(str(base[k]) for k in COLS
                                  if k not in ("id","cargado","row_hash","archivo")).encode()).hexdigest()
        vals = {**base, "id": str(r["id"]) if r.get("id") else "H"+rh,
                "archivo": r.get("archivo"), "cargado": ahora, "row_hash": rh}
        prev = c.execute("SELECT row_hash FROM confirmaciones WHERE id=?", (vals["id"],)).fetchone()
        if prev is None: n += 1
        elif prev[0] != rh: a += 1
        else: i += 1
        if prev is None or prev[0] != rh:
            c.execute(sql, [vals[k] for k in COLS])
    c.commit(); return n, a, i

_COLS_CONF_EXPORT = (
    ("id", "Nº confirmación"),
    ("orden", "Orden"),
    ("operacion", "Operación"),
    ("ubicacion", "Ubicación técnica"),
    ("puesto", "Puesto"),
    ("dur_real", "DurReal"),
    ("trbjo_real", "Trbjo real"),
    ("horas_intervalo", "Horas intervalo"),
    ("inicio", "Inicio"),
    ("fin", "Fin"),
    ("clase_act", "Clase actividad"),
    ("fe_contab", "Fe.contab."),
    ("archivo", "Archivo origen"),
    ("cargado", "Cargado"),
)


def _confirmaciones_para_export() -> pd.DataFrame:
    keys = [k for k, _ in _COLS_CONF_EXPORT]
    df = leer_sql(
        f"SELECT {','.join(keys)} FROM confirmaciones "
        "ORDER BY fin, orden, operacion, id"
    )
    for c in ("inicio", "fin", "fe_contab", "cargado"):
        df[c] = pd.to_datetime(df[c], errors="coerce")
    return df.rename(columns=dict(_COLS_CONF_EXPORT))


def _tabla_resumen_grupo(res: pd.DataFrame, grupo: str) -> pd.DataFrame:
    cols = (
        "Periodo", "Horas disponibles", "Horas usadas",
        "Esperado cumplimiento %", "Cumplimiento real %",
    )
    g = resumen_por_grupo(res, grupo)
    if g.empty:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame({
        "Periodo": g.index.astype(str),
        "Horas disponibles": g["horas"].values,
        "Horas usadas": g["trabajadas"].values,
        "Esperado cumplimiento %": g["objetivo"].values,
        "Cumplimiento real %": g["cumplimiento"].values,
    })


def _grafico_resumen(
    ws,
    n_filas: int,
    fila_datos: int,
    *,
    titulo: str,
    col_grafico: str = "G",
) -> None:
    if n_filas < 1:
        return
    from openpyxl.chart import BarChart, Reference
    from openpyxl.utils import column_index_from_string

    fila_hdr = fila_datos - 1
    fila_fin = fila_datos + n_filas - 1
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.grouping = "clustered"
    chart.title = titulo
    chart.y_axis.title = "Horas"
    horas = Reference(ws, min_col=2, min_row=fila_hdr, max_col=3, max_row=fila_fin)
    cats = Reference(ws, min_col=1, min_row=fila_datos, max_row=fila_fin)
    chart.add_data(horas, titles_from_data=True)
    chart.set_categories(cats)
    chart2 = BarChart()
    pct = Reference(ws, min_col=4, min_row=fila_hdr, max_row=fila_fin)
    chart2.add_data(pct, titles_from_data=True)
    chart2.set_categories(cats)
    chart.y_axis.axId = 100
    chart2.y_axis.axId = 200
    chart2.y_axis.crosses = "max"
    chart += chart2
    col_i = column_index_from_string(col_grafico)
    ws.add_chart(chart, f"{col_grafico}{fila_hdr}")


def exportar_maestro(ruta: str | Path | None = None) -> Path:
    dest = Path(ruta or XLS_MAESTRO).resolve()
    conf = _confirmaciones_para_export()
    res = agrupar_mensual(leer_sql("SELECT * FROM confirmaciones"), "trbjo_real", "fin", True)
    with pd.ExcelWriter(dest, engine="openpyxl") as writer:
        conf.to_excel(writer, sheet_name="Confirmaciones", index=False)
        ws = writer.book.create_sheet("Resumen", 1)
        ws["A1"] = "Horas hombre = Trbjo real · imputación: mes de fin real · por grupo técnico"
        fila = 3
        for grupo in grupos_en_resumen(res):
            bloque = _tabla_resumen_grupo(res, grupo)
            ws.cell(fila, 1, f"Grupo: {grupo}")
            fila += 1
            if bloque.empty:
                fila += 1
                continue
            for j, col in enumerate(bloque.columns, start=1):
                ws.cell(fila, j, col)
            fila_datos = fila + 1
            for i, row in enumerate(bloque.itertuples(index=False), start=fila_datos):
                for j, val in enumerate(row, start=1):
                    ws.cell(i, j, val)
            _grafico_resumen(
                ws,
                len(bloque),
                fila_datos,
                titulo=f"{grupo}: horas y cumplimiento",
                col_grafico="G",
            )
            fila = fila_datos + len(bloque) + 3
    return dest

# ---------------------------------------------------------------- imputación mensual
COL_HORAS = {"Trbjo real = horas hombre (personas × horas)": "trbjo_real",
             "DurReal = duración de la operación": "dur_real"}

def imputar_horas_por_fecha(
    df: pd.DataFrame, col: str, criterio="fin", dias_habiles=True,
) -> pd.DataFrame:
    """Una fila por día de imputación con «fecha» YYYY-MM-DD y «horas_imp»."""
    d = df[df[col].notna()].copy()
    if criterio in ("inicio", "fin", "contab"):
        campo = {"inicio": "inicio", "fin": "fin", "contab": "fe_contab"}[criterio]
        dt = pd.to_datetime(d[campo], errors="coerce")
        d["fecha"] = dt.dt.strftime("%Y-%m-%d")
        d["horas_imp"] = d[col].astype(float)
        return d[d["fecha"].notna()].reset_index(drop=True)
    regs = []
    for _, r in d.iterrows():
        i = pd.to_datetime(r["inicio"], errors="coerce")
        f = pd.to_datetime(r["fin"], errors="coerce")
        if pd.isna(f):
            continue
        if pd.isna(i) or f < i:
            i = f
        dias = pd.date_range(i.normalize(), f.normalize(), freq="D")
        if dias_habiles:
            hab = dias[dias.dayofweek < 5]
            dias = hab if len(hab) else dias
        n_dias = len(dias)
        if n_dias == 0:
            continue
        por_dia = float(r[col]) / n_dias
        for dia in dias:
            regs.append({**r.to_dict(), "fecha": dia.strftime("%Y-%m-%d"), "horas_imp": por_dia})
    return pd.DataFrame(regs)

def imputar_horas(df: pd.DataFrame, col: str, criterio="fin", dias_habiles=True) -> pd.DataFrame:
    d = imputar_horas_por_fecha(df, col, criterio, dias_habiles)
    if d.empty:
        return d
    d = d.copy()
    d["periodo"] = d["fecha"].str[:7]
    return d.drop(columns=["fecha"])

def _grupo_desde_puesto(puesto) -> str:
    if puesto is None or (isinstance(puesto, float) and np.isnan(puesto)):
        return GRUPO_SIN
    s = str(puesto).strip()
    return s if s else GRUPO_SIN

def grupos_en_resumen(res: pd.DataFrame) -> list[str]:
    if res.empty:
        return []
    if isinstance(res.index, pd.MultiIndex):
        return sorted(res.index.get_level_values("grupo").unique())
    return [GRUPO_SIN]

def _periodo_ym(periodo) -> str:
    if isinstance(periodo, tuple):
        periodo = periodo[0]
    return str(periodo).strip()[:7]

def resumen_por_grupo_diario(res: pd.DataFrame, grupo: str) -> pd.DataFrame:
    if res.empty:
        return res
    if not isinstance(res.index, pd.MultiIndex):
        return res.sort_index()
    g = res.xs(grupo, level="grupo", drop_level=True)
    if isinstance(g, pd.Series):
        fecha = str(g.name)
        return pd.DataFrame([g.to_dict()], index=pd.Index([fecha], name="fecha"))
    return g.sort_index()

def _vista_por_periodo(res: pd.DataFrame) -> pd.DataFrame:
    """Índice YYYY-MM por fila (gráficos mensuales)."""
    if res.empty:
        return res
    out = res.copy()
    if isinstance(out.index, pd.MultiIndex) and "periodo" in out.index.names:
        out.index = out.index.get_level_values("periodo")
    else:
        out.index = [_periodo_ym(ix) for ix in out.index]
    out.index.name = "periodo"
    if out.index.duplicated().any():
        num = out.select_dtypes(include="number").columns
        out = out.groupby(level=0)[list(num)].sum()
        hp = out["horas"].where(out["horas"] > 0)
        out["cumplimiento"] = 100 * out["trabajadas"] / hp
    return out.sort_index()

def resumen_por_grupo(res: pd.DataFrame, grupo: str) -> pd.DataFrame:
    if res.empty:
        return res
    if not isinstance(res.index, pd.MultiIndex):
        return _vista_por_periodo(res)
    g = res.xs(grupo, level="grupo", drop_level=True)
    if isinstance(g, pd.Series):
        periodo = _periodo_ym(g.name)
        fila = g.to_dict()
        return pd.DataFrame([fila], index=pd.Index([periodo], name="periodo"))
    return _vista_por_periodo(g)

def agrupar_mensual(df, col, criterio="fin", dias_habiles=True):
    d = imputar_horas(df, col, criterio, dias_habiles)
    if d.empty:
        return pd.DataFrame()
    d["grupo"] = d["puesto"].apply(_grupo_desde_puesto)
    por_ot = d.groupby(["periodo", "grupo", "orden"], as_index=False)["horas_imp"].sum()
    por_ot = por_ot[por_ot["horas_imp"] > 0]
    g = por_ot.groupby(["periodo", "grupo"])["horas_imp"].agg(
        ordenes="count", hs_ot_min="min", hs_ot_max="max", hs_ot_prom="mean", trabajadas="sum"
    )
    g["confirmaciones"] = d.groupby(["periodo", "grupo"])["id"].nunique()
    disp = leer_sql("SELECT periodo, grupo, horas, objetivo FROM disponibilidad")
    if disp.empty:
        disp = pd.DataFrame(columns=["periodo", "grupo", "horas", "objetivo"])
    disp = disp.set_index(["periodo", "grupo"])
    g = g.join(disp, how="left")
    g["horas"] = g["horas"].fillna(0.0)
    g["objetivo"] = g["objetivo"].fillna(85.0)
    horas_pos = g["horas"].where(g["horas"] > 0)
    g["cumplimiento"] = 100 * g["trabajadas"] / horas_pos
    return g.sort_index()

def agrupar_diario(df, col, criterio="fin", dias_habiles=True):
    d = imputar_horas_por_fecha(df, col, criterio, dias_habiles)
    if d.empty:
        return pd.DataFrame()
    d["grupo"] = d["puesto"].apply(_grupo_desde_puesto)
    g = d.groupby(["fecha", "grupo"]).agg(
        trabajadas=("horas_imp", "sum"),
        confirmaciones=("id", "nunique"),
        ordenes=("orden", "nunique"),
    )
    return g[g["trabajadas"] > 0].sort_index()

def diagnostico(df):
    d = df.copy()
    d["inicio_dt"] = pd.to_datetime(d["inicio"], errors="coerce")
    d["fin_dt"] = pd.to_datetime(d["fin"], errors="coerce")
    d["dias_span"] = (d["fin_dt"].dt.normalize() - d["inicio_dt"].dt.normalize()).dt.days + 1
    d["dotacion"] = np.where(d["horas_intervalo"] > 0,
                             d["trbjo_real"] / d["horas_intervalo"], np.nan)
    d["cruza_mes"] = (d["inicio_dt"].dt.to_period("M") != d["fin_dt"].dt.to_period("M"))
    raras = d[(d["dotacion"] > 25) | (d["dotacion"] < 0.1)].copy()
    raras["trbjo/10"] = raras["trbjo_real"] / 10
    multi = d[d["dias_span"] > 1]
    return d, raras, multi

# ---------------------------------------------------------------- gráficos / PDF
def _mpl():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    return plt, PdfPages

def fig_mensual(res, titulo="Horas hombre disponibles vs. trabajadas en OT"):
    plt, _ = _mpl()
    res = _vista_por_periodo(res)
    x = np.arange(len(res))
    disp = res["horas"].values
    trab = res["trabajadas"].values
    fig, ax = plt.subplots(figsize=(11, 5.5), dpi=150)
    ax.bar(x - 0.21, disp, 0.42, label="Hs hombre disponibles", color="#9bb7d4")
    ax.bar(x + 0.21, trab, 0.42, label="Hs hombre trabajadas en OT", color="#2f6f9f")
    ax.plot(x, disp * res["objetivo"].values / 100, "o--", color="#c0392b", lw=2, label="Objetivo (%)")
    for i, (d_, t_) in enumerate(zip(disp, trab)):
        ax.annotate(
            f"{100 * t_ / d_:.0f}%" if d_ else "s/d",
            (i, max(d_, t_) * 1.03),
            ha="center",
            fontsize=9,
        )
    ax.set_xticks(x)
    ax.set_xticklabels(list(res.index), rotation=45, ha="right")
    ax.set_ylabel("Horas hombre")
    ax.grid(axis="y", alpha=0.3)
    ax.legend()
    ax.set_title(titulo)
    fig.tight_layout()
    return fig

def fig_horas_dia(res_diario: pd.DataFrame, titulo: str = "Trabajo real por día"):
    plt, _ = _mpl()
    d = res_diario.sort_index()
    if d.empty:
        fig, ax = plt.subplots(figsize=(11, 4), dpi=150)
        ax.set_title(f"{titulo} — sin días con horas")
        ax.axis("off")
        fig.tight_layout()
        return fig
    x = np.arange(len(d))
    trab = d["trabajadas"].values
    ancho = max(11.0, min(0.28 * len(d) + 5, 48))
    fig, ax = plt.subplots(figsize=(ancho, 5.5), dpi=150)
    ax.bar(x, trab, 0.72, label="Suma horas del día (imputadas)", color="#2f6f9f")
    if len(d) <= 45:
        for i, t_ in enumerate(trab):
            ax.annotate(f"{t_:.0f}", (i, t_ * 1.02), ha="center", fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels(list(d.index), rotation=45, ha="right",
                       fontsize=8 if len(d) > 20 else 10)
    ax.set_ylabel("Horas hombre")
    ax.grid(axis="y", alpha=0.3)
    ax.legend()
    ax.set_title(titulo)
    fig.tight_layout()
    return fig

def generar_pdf(res, res_diario, ruta, criterio, col):
    plt, PdfPages = _mpl()
    with PdfPages(ruta) as pdf:
        for grupo in grupos_en_resumen(res):
            rg = resumen_por_grupo(res, grupo)
            rd = resumen_por_grupo_diario(res_diario, grupo)
            fig = plt.figure(figsize=(8.27, 11.69))
            fig.subplots_adjust(0, 0, 1, 1)
            fig.text(0.08, 0.94, f"Grupo técnico: {grupo}",
                     fontsize=17, fontweight="bold")
            fig.text(0.08, 0.91, f"Generado {datetime.now():%d/%m/%Y %H:%M} · "
                                 f"Horas hombre: {col} · Imputación: {CRITERIOS[criterio]}",
                     fontsize=9, color="#666")
            y = 0.85
            td, tt = rg["horas"].sum(), rg["trabajadas"].sum()
            lineas = [
                f"Disponibles (total): {td:,.1f} h",
                f"Trabajadas en OT: {tt:,.1f} h",
                f"Cumplimiento: {100*tt/td:.1f}%" if td else "Cumplimiento: s/d",
                f"OT: {int(rg['ordenes'].sum())} · Confirmaciones: {int(rg['confirmaciones'].sum())}",
                "",
                "Detalle por mes:",
            ]
            for p, r in rg.iterrows():
                cum = r["cumplimiento"]
                cum_s = f"{cum:>5.1f}%" if pd.notna(cum) else "  s/d"
                lineas.append(
                    f"  {p}  disp {r['horas']:>7.1f} | trab {r['trabajadas']:>7.1f} | "
                    f"{cum_s} | OT {int(r['ordenes']):>3} | "
                    f"hs/OT mín {r['hs_ot_min']:.1f} · máx {r['hs_ot_max']:.1f} · "
                    f"prom {r['hs_ot_prom']:.1f}"
                )
            for ln in lineas:
                fig.text(0.08, y, ln, fontsize=10, family="DejaVu Sans")
                y -= 0.028
            pdf.savefig(fig)
            plt.close(fig)
            pdf.savefig(
                fig_mensual(rg, titulo=f"{grupo} — totales mensuales")
            )
            pdf.savefig(
                fig_horas_dia(rd, titulo=f"{grupo} — trabajo real por día")
            )
        plt.close("all")
    if isinstance(ruta, BytesIO):
        ruta.seek(0)
        return ruta.getvalue()
    return ruta

def _mostrar_panel_grupo(
    rg, rd: pd.DataFrame, criterio: str, titulo_criterio: str,
) -> None:
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Disponibles", f"{rg['horas'].sum():,.1f} h")
    k2.metric("Trabajadas", f"{rg['trabajadas'].sum():,.1f} h")
    tot = rg["horas"].sum()
    k3.metric("Cumplimiento", f"{100*rg['trabajadas'].sum()/tot:.1f}%" if tot else "s/d")
    k4.metric("OT", int(rg["ordenes"].sum()))
    k5.metric("Confirmaciones", int(rg["confirmaciones"].sum()))

    st.subheader("Totales mensuales")
    st.pyplot(
        fig_mensual(rg, f"{titulo_criterio}"),
        use_container_width=True,
    )
    st.subheader("Trabajo real por día")
    st.caption(
        f"Solo días con horas > 0. Suma imputada al calendario según «{CRITERIOS[criterio]}» "
        "(con prorrateo, el Trbjo real de tareas multi-día se reparte entre los días del intervalo)."
    )
    st.pyplot(fig_horas_dia(rd, titulo=titulo_criterio), use_container_width=True)
    if not rd.empty:
        st.subheader("Detalle por día")
        det = rd.rename(columns={
            "trabajadas": "Trabajadas", "confirmaciones": "Confirm.", "ordenes": "OT",
        })
        st.dataframe(
            det[["Trabajadas", "Confirm.", "OT"]],
            use_container_width=True,
        )
    st.subheader("Detalle mensual")
    mostrar = _vista_por_periodo(rg).rename(columns={
        "horas": "Disponibles", "trabajadas": "Trabajadas", "objetivo": "Objetivo %",
        "cumplimiento": "% Real", "ordenes": "OT", "confirmaciones": "Confirm.",
        "hs_ot_min": "Mín hs/OT", "hs_ot_max": "Máx hs/OT", "hs_ot_prom": "Prom hs/OT",
    })
    st.dataframe(
        mostrar[[
            "Disponibles", "Trabajadas", "Objetivo %", "% Real", "OT", "Confirm.",
            "Mín hs/OT", "Máx hs/OT", "Prom hs/OT",
        ]],
        use_container_width=True,
    )

# ---------------------------------------------------------------- UI
st.set_page_config(page_title="Horas hombre vs OT", layout="wide")
st.sidebar.title("Horas hombre")
pagina = st.sidebar.radio("Pantalla", ["1 · Importar", "2 · Horas disponibles",
                                       "3 · Panel y gráficos", "4 · Exportar PDF",
                                       "5 · Diagnóstico"])

if pagina.startswith("1"):
    st.header("Importar archivos de SAP")
    st.caption(
        "Podés subir el mismo archivo varias veces o archivos viejos: se deduplica por "
        "Nº de confirmación (Notific. en lista dinámica, ClOrd en exports viejos)."
    )
    ups = st.file_uploader("Archivos .xls / .xlsx / .csv / .txt", accept_multiple_files=True)
    parsed: list[tuple[str, ImportResult]] = []
    if ups:
        for u in ups:
            try:
                parsed.append((u.name, extraer_tabla(u.getvalue(), u.name)))
            except Exception as e:
                st.error(f"{u.name}: {e}")
        for nombre, res in parsed:
            with st.expander(f"Preview · {nombre}", expanded=len(parsed) == 1):
                st.caption(f"Adaptador: **{res.adapter}** · confianza layout "
                           f"{res.metrics.get('confianza_layout', 0):.0%}")
                if res.mapping:
                    st.json(res.mapping)
                c1, c2, c3 = st.columns(3)
                c1.metric("Filas", res.metrics.get("filas", len(res.df)))
                c2.metric("% con inicio/fin", f"{res.metrics.get('pct_inicio', 0)}%")
                c3.metric("IDs duplicados", res.metrics.get("ids_duplicados", 0))
                for ln in res.log:
                    st.warning(ln)
                st.dataframe(res.df.head(30), use_container_width=True)
        if parsed and st.button("Confirmar carga", type="primary"):
            for nombre, res in parsed:
                n, a, i = guardar(res.df)
                st.success(f"**{nombre}** — {len(res.df)} filas · {n} nuevas · {a} actualizadas · "
                           f"{i} ya existían")
            maestro = exportar_maestro()
            n_bd = int(leer_sql("SELECT COUNT(*) c FROM confirmaciones")["c"][0])
            st.info(f"Excel maestro actualizado: **{maestro.name}** ({n_bd} confirmaciones, "
                    "hojas Confirmaciones + Resumen con gráfico).")
    n_conf = int(leer_sql("SELECT COUNT(*) c FROM confirmaciones")["c"][0])
    st.metric("Confirmaciones en base", n_conf)
    if n_conf:
        maestro_path = Path(XLS_MAESTRO)
        if not maestro_path.is_file():
            exportar_maestro()
        st.download_button(
            "Descargar Excel maestro",
            maestro_path.read_bytes(),
            file_name=maestro_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

elif pagina.startswith("2"):
    st.header("Horas hombre disponibles por mes y grupo técnico")
    disp_bd = leer_sql(
        "SELECT periodo, grupo, horas, objetivo, nota FROM disponibilidad "
        "ORDER BY periodo, grupo"
    )
    n_legacy = int(leer_sql("SELECT COUNT(*) c FROM disponibilidad WHERE grupo=''")["c"][0])
    if n_legacy:
        st.warning(
            f"Hay **{n_legacy}** registro(s) con grupo vacío (migrados de una versión anterior). "
            "No se vinculan a GSS-CIVI ni a otros grupos: recargalos con el grupo técnico correcto."
        )
    tab_tabla, tab_uno, tab_archivo = st.tabs(["Carga masiva (tabla)", "Un mes", "Importar archivo"])
    puestos = puestos_en_confirmaciones()

    with tab_tabla:
        st.caption(
            "Editá las filas, agregá meses con **+** o pegá desde Excel. "
            "Periodo **YYYY-MM** y **grupo técnico** (PtoTrReal, ej. GSS-CIVI). "
            "La clave es periodo + grupo."
        )
        base = disp_bd if not disp_bd.empty else pd.DataFrame(
            columns=["periodo", "grupo", "horas", "objetivo", "nota"])
        editada = st.data_editor(
            base,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "periodo": st.column_config.TextColumn("Periodo (YYYY-MM)", width="small", required=True),
                "grupo": st.column_config.TextColumn("Grupo técnico", width="medium", required=True),
                "horas": st.column_config.NumberColumn("Horas disponibles", min_value=0.0, step=10.0,
                                                         format="%.1f"),
                "objetivo": st.column_config.NumberColumn("Objetivo %", min_value=0.0, max_value=200.0,
                                                          step=5.0, default=85.0),
                "nota": st.column_config.TextColumn("Nota", width="medium"),
            },
            key="disp_editor",
        )
        if st.button("Guardar tabla", type="primary", key="disp_save_tabla"):
            n, a, err = guardar_disponibilidad(editada)
            if err:
                for e in err:
                    st.error(e)
            if n or a:
                st.success(f"Guardado: {n} nuevos · {a} actualizados.")
                st.rerun()
            elif not err:
                st.info("No hay filas con periodo y horas para guardar.")

    with tab_uno:
        with st.form("alta"):
            c1, c2, c3, c4 = st.columns(4)
            anio = c1.number_input("Año", 2020, 2100, datetime.now().year)
            mes = c2.selectbox("Mes", MESES, index=datetime.now().month-1)
            horas = c3.number_input("Horas hombre disponibles", 0.0, 200000.0, 0.0, step=10.0)
            obj = c4.number_input("Objetivo (%)", 0.0, 200.0, 85.0, step=5.0)
            if puestos:
                grupo_lista = st.selectbox("Grupo en confirmaciones", puestos)
            else:
                grupo_lista = ""
            grupo_otro = st.text_input("Grupo técnico (si no está en la lista)", "")
            nota = st.text_input("Nota (opcional)")
            grupo_final = grupo_otro.strip() or grupo_lista
            if st.form_submit_button("Guardar", type="primary"):
                n, a, err = guardar_disponibilidad(pd.DataFrame([{
                    "periodo": f"{anio}-{mes}",
                    "grupo": grupo_final,
                    "horas": horas,
                    "objetivo": obj,
                    "nota": nota,
                }]))
                if err:
                    st.error(err[0])
                else:
                    st.success("Guardado.")
                    st.rerun()

    with tab_archivo:
        st.caption(
            "CSV o Excel con **periodo**, **grupo**, **horas**; opcional **objetivo**, **nota**."
        )
        plantilla = pd.DataFrame([
            {"periodo": "2025-01", "grupo": "GSS-CIVI", "horas": 12000.0, "objetivo": 85.0, "nota": ""},
            {"periodo": "2025-02", "grupo": "GSS-CIVI", "horas": 11800.0, "objetivo": 85.0, "nota": ""},
        ])
        st.download_button(
            "Descargar plantilla CSV",
            plantilla.to_csv(index=False).encode("utf-8"),
            file_name="disponibilidad_plantilla.csv",
            mime="text/csv",
        )
        up = st.file_uploader("Archivo .csv / .xlsx", type=["csv", "xlsx", "xls"], key="disp_file")
        if up:
            try:
                prev = leer_disponibilidad_archivo(up.getvalue(), up.name)
                st.dataframe(prev, use_container_width=True)
                if st.button("Confirmar importación", type="primary", key="disp_import"):
                    n, a, err = guardar_disponibilidad(prev)
                    for e in err:
                        st.error(e)
                    st.success(f"**{up.name}** — {n} nuevos · {a} actualizados · "
                               f"{len(err)} filas omitidas.")
                    st.rerun()
            except Exception as e:
                st.error(str(e))

    st.subheader("Registro en base")
    st.dataframe(
        leer_sql("SELECT * FROM disponibilidad ORDER BY periodo, grupo"),
        use_container_width=True,
    )

elif pagina.startswith("5"):
    st.header("Diagnóstico de datos")
    df = leer_sql("SELECT * FROM confirmaciones")
    if df.empty: st.info("Sin datos."); st.stop()
    d, raras, multi = diagnostico(df)
    c1, c2, c3 = st.columns(3)
    c1.metric("Tareas de más de 1 día", len(multi))
    c2.metric("Horas hombre en tareas multi-día", f"{multi['trbjo_real'].sum():,.1f} h")
    c3.metric("Filas con dotación sospechosa", len(raras))
    st.subheader("Dotación implícita (Trbjo real ÷ horas de la ventana)")
    st.caption("Menos de 0,1 o más de 25 personas por operación: revisar. La columna "
               "'trbjo/10' muestra la hipótesis de decimal perdido — no se aplica sola.")
    st.dataframe(raras[["id","orden","operacion","trbjo_real","horas_intervalo","dotacion",
                        "trbjo/10","inicio","fin"]], use_container_width=True)
    st.subheader("Tareas que cruzan mes")
    if not multi.empty:
        st.dataframe(multi[["id","orden","trbjo_real","inicio","fin","dias_span"]]
                     .assign(cruza_mes=multi["cruza_mes"]), use_container_width=True)

else:
    df = leer_sql("SELECT * FROM confirmaciones")
    if df.empty: st.info("Todavía no importaste archivos."); st.stop()
    col_label = st.sidebar.selectbox("Horas hombre = ", list(COL_HORAS))
    criterio = st.sidebar.selectbox("Imputar al mes por", list(CRITERIOS),
                                    format_func=lambda k: CRITERIOS[k])
    dias_habiles = st.sidebar.checkbox("Prorratear solo días hábiles", True,
                                       disabled=criterio != "prorrateo")
    col = COL_HORAS[col_label]
    res = agrupar_mensual(df, col, criterio, dias_habiles)
    res_diario = agrupar_diario(df, col, criterio, dias_habiles)
    if res.empty:
        st.warning("Sin horas para agregar con ese criterio.")
        st.stop()
    grupos = grupos_en_resumen(res)
    if not grupos:
        st.warning("Sin grupos técnicos en los datos.")
        st.stop()

    titulo_crit = f"Disponibles vs. trabajadas — {CRITERIOS[criterio]}"
    tabs = st.tabs([str(g) for g in grupos])
    for tab, grupo in zip(tabs, grupos):
        with tab:
            rg = resumen_por_grupo(res, grupo)
            rd = resumen_por_grupo_diario(res_diario, grupo)
            _mostrar_panel_grupo(rg, rd, criterio, f"{grupo} — {titulo_crit}")
    if criterio == "prorrateo":
        st.caption("Con prorrateo, el mín/máx/prom por OT de cada mes es parcial para las OT "
                   "que cruzan mes.")

    if pagina.startswith("4"):
        st.header("Exportar PDF")
        st.caption("Un PDF con una sección y gráficos por cada grupo técnico.")
        pdf_bytes = generar_pdf(res, res_diario, BytesIO(), criterio, col_label)
        st.download_button("Descargar PDF", pdf_bytes,
                           file_name=f"horas_hombre_{datetime.now():%Y%m}.pdf",
                           mime="application/pdf", type="primary")