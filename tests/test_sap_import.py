from pathlib import Path

import pandas as pd

from sap_import import (
    ImportResult,
    a_float,
    a_fecha,
    extraer_tabla,
    leer_crudo,
)

LISTA_FIXTURE = """\ufeff13.11.2026  Salida dinámica de lista  1


,,S,Notific.,ClOrd,Orden,DurReal,Trbjo real,Creado el,Fe.contab.,PtoTrReal,Stat.sist.,Nº pers.,Nombre,DivP,Op.,Ubicación técnica,Fin,an.,Equipo,H.in.real,Fe.in.real,Fin real,FecFinReal,ClActRe,Txt.notif.

,X,,90000001,SDMN,700001797672,8,160,11-13-26,08-10-26,GSS-CIVI,NOTI,,,,10,2200-IN-QUA-QUA01-APRO,X,,,8:00,08-10-26,16:00,08-10-26,22GSCI,
,X,,90000002,SDMN,700001797672,00.5,1,11-13-26,08-10-26,GSS-CIVI,NOTI,,,,20,2200-IN-QUA-QUA01-APRO,X,,,9:00,08-10-26,9:30,08-10-26,22GSCI,
,X,,90000001,SDMN,700001897777,1,1,11-13-26,08-10-26,GSS-CIVI,NOTI,,,,10,2200-IN-TEST,X,,,8:00,11-13-26,9:00,11-13-26,22GSCI,
"""

LEGACY_TSV = """junk
junk2
ClOrd\tOrden\tOp\tDurReal\tTrbjo real\tFeInReal\tHInReal\tFeFinReal\tHFinReal\tFe.contab.
1001\tORD1\t10\t8,5\t17\t14.11.2025\t08:00:00\t14.11.2025\t16:30:00\t14.11.2025
"""


def test_lista_dinamica_id_and_dates():
    res = extraer_tabla(LISTA_FIXTURE.encode("utf-8"), "horas.xls")
    assert isinstance(res, ImportResult)
    assert res.adapter == "ListaDinamica"
    assert res.mapping["id"] == "Notific."
    assert len(res.df) == 3
    assert res.df.iloc[0]["id"] == "90000001"
    assert res.df.iloc[0]["inicio"] is not None
    assert res.df.iloc[0]["inicio"].date().isoformat() == "2026-08-10"
    assert res.df.iloc[2]["inicio"].date().isoformat() == "2026-11-13"
    assert any("duplicados" in ln.lower() for ln in res.log)


def test_legacy_clord_tsv():
    res = extraer_tabla(LEGACY_TSV.encode("utf-8"), "legacy.tsv")
    assert res.adapter == "ConfirmacionesClOrd"
    assert len(res.df) == 1
    assert res.df.iloc[0]["id"] == "1001"
    assert res.df.iloc[0]["dur_real"] == 8.5


def test_fake_xls_is_text_not_excel():
    data = LISTA_FIXTURE.encode("utf-8")
    assert data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    raw = leer_crudo(data)
    assert len(raw) >= 5


def test_durreal_leading_zeros():
    assert a_float("00.5") == 0.5
    res = extraer_tabla(LISTA_FIXTURE.encode("utf-8"), "t.xls")
    half = res.df[res.df["operacion"] == "20"].iloc[0]
    assert half["dur_real"] == 0.5


def test_fecha_mdy_dash():
    assert a_fecha("11-13-26", "mdy").isoformat() == "2026-11-13"
    assert a_fecha("08-10-26", "mdy").isoformat() == "2026-08-10"


def test_real_sample_file_if_present():
    """Opcional: corre si el usuario tiene el export en Downloads."""
    path = Path("/Users/julian/Downloads/horas.xls")
    if not path.is_file():
        return
    res = extraer_tabla(path.read_bytes(), path.name)
    assert res.adapter == "ListaDinamica"
    assert len(res.df) >= 15
    assert res.metrics["pct_inicio"] > 90
    assert str(res.df.iloc[0]["id"]).isdigit()