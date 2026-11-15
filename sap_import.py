from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, time as dtime, timedelta
from typing import Protocol

import numpy as np
import pandas as pd

CANON_COLS = (
    "id", "orden", "operacion", "ubicacion", "puesto", "dur_real", "trbjo_real",
    "h_ini", "f_ini", "h_fin", "f_fin", "clase_act", "fe_contab",
)

TIPOS = {
    "id": "num", "orden": "num", "operacion": "num", "dur_real": "num", "trbjo_real": "num",
    "h_ini": "hora", "h_fin": "hora", "f_ini": "fecha", "f_fin": "fecha",
}

ALIAS_BASE = {
    "orden": "orden",
    "operacion": "op",
    "ubicacion": "ubicaciontecnica",
    "puesto": "ptotrreal",
    "dur_real": "durreal",
    "trbjo_real": "trbjoreal",
    "h_ini": "hinreal",
    "f_ini": "feinreal",
    "h_fin": "finreal",
    "f_fin": "fecfinreal",
    "clase_act": "clactre",
    "fe_contab": "fecontab",
}

ALIAS_LISTA = {**ALIAS_BASE, "id": "notific"}
ALIAS_CLORD = {**ALIAS_BASE, "id": "clord"}


@dataclass
class ImportResult:
    df: pd.DataFrame
    log: list[str] = field(default_factory=list)
    adapter: str = ""
    mapping: dict[str, str] = field(default_factory=dict)
    metrics: dict[str, float | int] = field(default_factory=dict)


def _norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"[\s\.\-_]+", "", s).lower()


def a_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return None if pd.isna(v) else float(v)
    s = str(v).strip().replace("\xa0", "").replace(" ", "")
    if s in ("", "nan", "None", "-"):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def a_hora(v):
    if isinstance(v, dtime):
        return v
    if isinstance(v, datetime):
        return v.time()
    s = str(v).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    return None


def _parse_dash_date(s: str, order: str) -> date | None:
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{2,4})$", s)
    if not m:
        return None
    a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    if order == "mdy":
        month, day = a, b
    else:
        day, month = a, b
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    try:
        return date(y, month, day)
    except ValueError:
        return None


def a_fecha(v, order: str = "mdy"):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d.%m.%y", "%m-%d-%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    d = _parse_dash_date(s, order)
    if d:
        return d
    if re.match(r"^\d+(\.\d+)?$", s):
        try:
            serial = float(s)
            if 30000 < serial < 60000:
                return (datetime(1899, 12, 30) + timedelta(days=serial)).date()
        except ValueError:
            pass
    return None


def detectar_orden_fecha(muestras: list[str]) -> str:
    mdy_ok = dmy_ok = 0
    for s in muestras:
        s = str(s).strip()
        if not re.match(r"^\d{1,2}-\d{1,2}-\d{2,4}$", s):
            continue
        p1, p2 = int(s.split("-")[0]), int(s.split("-")[1])
        if p1 > 12:
            dmy_ok += 1
        if p2 > 12:
            mdy_ok += 1
        if _parse_dash_date(s, "mdy"):
            mdy_ok += 1
        if _parse_dash_date(s, "dmy"):
            dmy_ok += 1
    return "mdy" if mdy_ok >= dmy_ok else "dmy"


def _texto_a_df(txt: str, sep: str) -> pd.DataFrame:
    filas = []
    for line in txt.splitlines():
        line = line.rstrip("\r")
        if not line.strip():
            continue
        filas.append(line.split(sep))
    if not filas:
        return pd.DataFrame()
    n = max(len(f) for f in filas)
    filas = [f + [None] * (n - len(f)) for f in filas]
    return pd.DataFrame(filas, dtype=object)


def leer_crudo(data: bytes) -> pd.DataFrame:
    if data[:4] == b"PK\x03\x04" or data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return pd.read_excel(io.BytesIO(data), header=None, dtype=str)
    txt = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            txt = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if txt is None:
        raise ValueError("No pude decodificar el archivo como texto.")
    if txt.lstrip()[:1] == "<":
        return max(pd.read_html(io.StringIO(txt), header=None), key=len).astype(str)
    sep = (
        "\t"
        if txt.count("\t") >= max(txt.count(";"), txt.count(","))
        else (";" if txt.count(";") > txt.count(",") else ",")
    )
    return _texto_a_df(txt, sep)


def _find_header_row(raw: pd.DataFrame, required: set[str]) -> int | None:
    for i in range(min(80, len(raw))):
        vals = {_norm(v) for v in raw.iloc[i].tolist() if not pd.isna(v) and str(v).strip()}
        if required.issubset(vals):
            return i
    return None


def _header_labels(header: list) -> list[str]:
    return [str(c).strip() if not pd.isna(c) else "" for c in header]


def _mapear(header, alias: dict[str, str]) -> dict[str, int]:
    m = {}
    for i, c in enumerate(header):
        if pd.isna(c):
            continue
        n = _norm(c)
        for canon, key in alias.items():
            if n == key:
                m.setdefault(canon, i)
    return m


def _puntaje(datos: pd.DataFrame, mapa: dict[str, int], date_order: str = "mdy") -> float:
    ok = tot = 0
    parse_fecha = lambda v: a_fecha(v, date_order)
    for canon, esp in TIPOS.items():
        i = mapa.get(canon)
        if i is None or i >= datos.shape[1]:
            continue
        for v in datos[i].dropna().astype(str).head(200):
            tot += 1
            fn = {"num": a_float, "fecha": parse_fecha, "hora": a_hora}[esp]
            if fn(v) is not None:
                ok += 1
    return ok / max(tot, 1)


def _slice_datos(raw: pd.DataFrame, idx: int) -> tuple[list, pd.DataFrame]:
    header = raw.iloc[idx].tolist()
    datos = raw.iloc[idx + 1 :].reset_index(drop=True)
    datos.columns = range(datos.shape[1])
    return header, datos


def _mapping_labels(header: list, mapa: dict[str, int]) -> dict[str, str]:
    labels = _header_labels(header)
    return {canon: labels[i] if i < len(labels) else "" for canon, i in mapa.items()}


def _ratio_numeric(datos: pd.DataFrame, col: int | None, limit: int = 100) -> float:
    if col is None or col >= datos.shape[1]:
        return 0.0
    vals = datos[col].dropna().astype(str).head(limit)
    if vals.empty:
        return 0.0
    ok = sum(1 for v in vals if a_float(v) is not None and re.match(r"^\d", v.strip()))
    return ok / len(vals)


def _ratio_alpha(datos: pd.DataFrame, col: int | None, limit: int = 100) -> float:
    if col is None or col >= datos.shape[1]:
        return 0.0
    vals = datos[col].dropna().astype(str).head(limit)
    if vals.empty:
        return 0.0
    ok = sum(1 for v in vals if re.search(r"[A-Za-z]", v))
    return ok / len(vals)


def _muestras_fecha(datos: pd.DataFrame, mapa: dict[str, int]) -> list[str]:
    out = []
    for key in ("f_ini", "f_fin", "fe_contab"):
        i = mapa.get(key)
        if i is None or i >= datos.shape[1]:
            continue
        out.extend(datos[i].dropna().astype(str).head(80).tolist())
    return out


def _build_out(datos: pd.DataFrame, mapa: dict[str, int], nombre: str) -> pd.DataFrame:
    out = pd.DataFrame({k: datos[i] for k, i in mapa.items() if i < datos.shape[1]})
    out["archivo"] = nombre
    return out


def limpiar(df: pd.DataFrame, date_order: str = "mdy") -> pd.DataFrame:
    parse_f = lambda v: a_fecha(v, date_order)
    for c in ("id", "orden", "operacion", "ubicacion", "puesto", "clase_act"):
        if c not in df:
            df[c] = None
        df[c] = df[c].apply(lambda v: None if pd.isna(v) else str(v).strip())
    for c in ("dur_real", "trbjo_real"):
        df[c] = df[c].apply(a_float) if c in df else np.nan
    fi = df["f_ini"].apply(parse_f) if "f_ini" in df else pd.Series([None] * len(df))
    ff = df["f_fin"].apply(parse_f) if "f_fin" in df else pd.Series([None] * len(df))
    hi = df["h_ini"].apply(a_hora) if "h_ini" in df else pd.Series([None] * len(df))
    hf = df["h_fin"].apply(a_hora) if "h_fin" in df else pd.Series([None] * len(df))
    ini, fin, horas = [], [], []
    for a, b, c, d in zip(fi, ff, hi, hf):
        x = datetime.combine(a, c) if (a and c) else None
        y = datetime.combine(b or a, d) if ((b or a) and d) else None
        if x and y and y < x:
            y += timedelta(days=1)
        ini.append(x)
        fin.append(y)
        horas.append((y - x).total_seconds() / 3600 if (x and y) else None)
    df["inicio"], df["fin"] = ini, fin
    df["horas_intervalo"] = horas
    df["fe_contab"] = df["fe_contab"].apply(parse_f) if "fe_contab" in df else None
    df = df[df["orden"].notna() & (df["orden"] != "")]
    return df.reset_index(drop=True)


def _metricas(df: pd.DataFrame, log: list[str]) -> dict[str, float | int]:
    n = len(df)
    con_inicio = int(df["inicio"].notna().sum()) if "inicio" in df else 0
    con_horas = int(df["horas_intervalo"].notna().sum()) if "horas_intervalo" in df else 0
    dup = int(df["id"].duplicated().sum()) if "id" in df and n else 0
    if dup:
        log.append(f"⚠️ {dup} ID duplicados en el archivo (mismo Nº de confirmación repetido).")
    return {
        "filas": n,
        "pct_inicio": round(100 * con_inicio / max(n, 1), 1),
        "pct_horas_intervalo": round(100 * con_horas / max(n, 1), 1),
        "ids_duplicados": dup,
    }


class SapLayoutAdapter(Protocol):
    name: str

    def detect(self, raw: pd.DataFrame) -> float: ...

    def extract(self, raw: pd.DataFrame, nombre: str) -> tuple[pd.DataFrame, list[str], dict[str, str], str]: ...


class ListaDinamicaAdapter:
    name = "ListaDinamica"

    def detect(self, raw: pd.DataFrame) -> float:
        idx = _find_header_row(raw, {"notific", "clord", "durreal"})
        if idx is None:
            return 0.0
        header, datos = _slice_datos(raw, idx)
        mapa = _mapear(header, ALIAS_LISTA)
        if "id" not in mapa:
            return 0.0
        id_ok = _ratio_numeric(datos, mapa.get("id"))
        clord_col = None
        for i, c in enumerate(header):
            if _norm(c) == "clord":
                clord_col = i
                break
        clord_alpha = _ratio_alpha(datos, clord_col)
        score = 0.35 + 0.35 * id_ok + 0.3 * clord_alpha
        score += 0.1 * _puntaje(datos, mapa, "mdy")
        return min(score, 1.0)

    def extract(self, raw: pd.DataFrame, nombre: str):
        log: list[str] = []
        idx = _find_header_row(raw, {"notific", "clord", "durreal"})
        if idx is None:
            raise ValueError("No encontré encabezado de lista dinámica (Notific. / ClOrd / DurReal).")
        header, datos = _slice_datos(raw, idx)
        mapa = _mapear(header, ALIAS_LISTA)
        date_order = detectar_orden_fecha(_muestras_fecha(datos, mapa))
        base = _puntaje(datos, mapa, date_order)
        if base < 0.85:
            log.append(f"⚠️ Confianza del mapeo: {base:.0%}. Revisá la vista previa.")
        log.append(
            f"Layout: lista dinámica SAP · ID desde Notific. · fechas interpretadas como "
            f"{'mes-día' if date_order == 'mdy' else 'día-mes'}."
        )
        out = _build_out(datos, mapa, nombre)
        mapping = _mapping_labels(header, mapa)
        return limpiar(out, date_order), log, mapping, date_order


class ConfirmacionesClOrdAdapter:
    name = "ConfirmacionesClOrd"

    def detect(self, raw: pd.DataFrame) -> float:
        idx = _find_header_row(raw, {"clord", "durreal"})
        if idx is None:
            return 0.0
        header, datos = _slice_datos(raw, idx)
        mapa = _mapear(header, ALIAS_CLORD)
        if "id" not in mapa:
            return 0.0
        if _find_header_row(raw, {"notific", "clord", "durreal"}) is not None:
            clord_col = mapa.get("id")
            if _ratio_alpha(datos, clord_col) > 0.5:
                return 0.15
        id_ok = _ratio_numeric(datos, mapa.get("id"))
        score = 0.4 + 0.4 * id_ok + 0.2 * _puntaje(datos, mapa, "dmy")
        return min(score, 1.0)

    def extract(self, raw: pd.DataFrame, nombre: str):
        log: list[str] = []
        idx = _find_header_row(raw, {"clord", "durreal"})
        if idx is None:
            raise ValueError("No encontré la fila de encabezado (ClOrd / DurReal).")
        header, datos = _slice_datos(raw, idx)
        mapa = _mapear(header, ALIAS_CLORD)
        date_order = detectar_orden_fecha(_muestras_fecha(datos, mapa)) or "dmy"
        base = _puntaje(datos, mapa, date_order)
        if base < 0.85:
            i0 = mapa.get("id", 0)
            mapa2 = {k: (v + 1 if v > i0 else v) for k, v in mapa.items()}
            if _puntaje(datos, mapa2, date_order) > base:
                log.append("⚠️ Encabezado desalineado (falta una columna después de ClOrd). Corregido.")
                mapa, base = mapa2, _puntaje(datos, mapa2, date_order)
        if base < 0.85:
            log.append(f"⚠️ Confianza del mapeo: {base:.0%}. Revisá la vista previa.")
        log.append("Layout: confirmaciones · ID desde ClOrd.")
        out = _build_out(datos, mapa, nombre)
        mapping = _mapping_labels(header, mapa)
        return limpiar(out, date_order), log, mapping, date_order


_ADAPTERS: list[SapLayoutAdapter] = [
    ListaDinamicaAdapter(),
    ConfirmacionesClOrdAdapter(),
]


def extraer_tabla(data: bytes, nombre: str) -> ImportResult:
    log: list[str] = []
    raw = leer_crudo(data)
    if raw.empty:
        raise ValueError("El archivo no tiene filas legibles.")
    ranked = sorted(((a, a.detect(raw)) for a in _ADAPTERS), key=lambda x: x[1], reverse=True)
    adapter, score = ranked[0]
    if score < 0.2:
        raise ValueError("No reconocí un layout SAP de confirmaciones en este archivo.")
    if ranked[1][1] > 0.35 and abs(ranked[0][1] - ranked[1][1]) < 0.08:
        log.append(
            f"⚠️ Layout ambiguo ({ranked[0][0].name} {ranked[0][1]:.0%} vs "
            f"{ranked[1][0].name} {ranked[1][1]:.0%}). Usando {adapter.name}."
        )
    df, adapter_log, mapping, _ = adapter.extract(raw, nombre)
    log.extend(adapter_log)
    metrics = _metricas(df, log)
    metrics["confianza_layout"] = round(score, 2)
    return ImportResult(df=df, log=log, adapter=adapter.name, mapping=mapping, metrics=metrics)