from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

import horas_hombre as hh
from sap_import import extraer_tabla

LISTA_FIXTURE = """\ufeff13.11.2026  Salida dinámica de lista  1


,,S,Notific.,ClOrd,Orden,DurReal,Trbjo real,Creado el,Fe.contab.,PtoTrReal,Stat.sist.,Nº pers.,Nombre,DivP,Op.,Ubicación técnica,Fin,an.,Equipo,H.in.real,Fe.in.real,Fin real,FecFinReal,ClActRe,Txt.notif.

,X,,90000001,SDMN,700001797672,8,160,11-13-26,08-10-26,GSS-CIVI,NOTI,,,,10,2200-IN-QUA-QUA01-APRO,X,,,8:00,08-10-26,16:00,08-10-26,22GSCI,
,X,,90000002,SDMN,700001797672,00.5,1,11-13-26,08-10-26,GSS-CIVI,NOTI,,,,20,2200-IN-QUA-QUA01-APRO,X,,,9:00,08-10-26,9:30,08-10-26,22GSCI,
,X,,90000001,SDMN,700001897777,1,1,11-13-26,08-10-26,GSS-CIVI,NOTI,,,,10,2200-IN-TEST,X,,,8:00,11-13-26,9:00,11-13-26,22GSCI,
"""

LISTA_DOS_GRUPOS = """\ufeff13.11.2026  Salida dinámica de lista  1


,,S,Notific.,ClOrd,Orden,DurReal,Trbjo real,Creado el,Fe.contab.,PtoTrReal,Stat.sist.,Nº pers.,Nombre,DivP,Op.,Ubicación técnica,Fin,an.,Equipo,H.in.real,Fe.in.real,Fin real,FecFinReal,ClActRe,Txt.notif.

,X,,90000001,SDMN,700001797672,8,100,11-13-26,08-10-26,GSS-CIVI,NOTI,,,,10,2200-IN-QUA-QUA01-APRO,X,,,8:00,08-10-26,16:00,08-10-26,22GSCI,
,X,,90000002,SDMN,700001797672,8,50,11-13-26,08-10-26,GSS-ELEC,NOTI,,,,20,2200-IN-QUA-QUA01-APRO,X,,,8:00,08-10-26,16:00,08-10-26,22GSCI,
"""


@pytest.fixture
def isolated_db(tmp_path, monkeypatch):
    db = tmp_path / "t.db"
    xls = tmp_path / "confirmaciones.xlsx"
    monkeypatch.setattr(hh, "DB", str(db))
    monkeypatch.setattr(hh, "XLS_MAESTRO", str(xls))
    hh.conectar.clear()
    yield xls
    hh.conectar.clear()


def test_maestro_sin_duplicados_y_grafico(isolated_db):
    res = extraer_tabla(LISTA_FIXTURE.encode("utf-8"), "horas.xls")
    hh.guardar(res.df)
    hh.guardar(res.df)
    path = hh.exportar_maestro()

    wb = load_workbook(path)
    assert wb.sheetnames == ["Confirmaciones", "Resumen"]
    conf = pd.read_excel(path, sheet_name="Confirmaciones")
    assert len(conf) == 2
    assert conf["Nº confirmación"].duplicated().sum() == 0
    assert "row_hash" not in conf.columns

    ws = wb["Resumen"]
    assert ws["A3"].value == "Grupo: GSS-CIVI"
    assert len(ws._charts) == 1
    ch = ws._charts[0]
    n_series = sum(len(c.series) for c in ch._charts)
    assert n_series == 3


def test_disponibilidad_mismo_mes_dos_grupos(isolated_db):
    n, a, err = hh.guardar_disponibilidad(
        pd.DataFrame([
            {"periodo": "2026-08", "grupo": "GSS-CIVI", "horas": 1000.0, "objetivo": 85.0, "nota": ""},
            {"periodo": "2026-08", "grupo": "GSS-ELEC", "horas": 400.0, "objetivo": 90.0, "nota": ""},
        ])
    )
    assert not err
    assert n == 2
    rows = hh.leer_sql("SELECT * FROM disponibilidad")
    assert len(rows) == 2
    por_grupo = rows.set_index("grupo")["horas"]
    assert por_grupo["GSS-CIVI"] == 1000.0
    assert por_grupo["GSS-ELEC"] == 400.0


def test_agrupar_cumplimiento_distinto_por_grupo(isolated_db):
    res = extraer_tabla(LISTA_DOS_GRUPOS.encode("utf-8"), "t.xls")
    hh.guardar(res.df)
    hh.guardar_disponibilidad(
        pd.DataFrame([
            {"periodo": "2026-08", "grupo": "GSS-CIVI", "horas": 100.0, "objetivo": 85.0, "nota": ""},
            {"periodo": "2026-08", "grupo": "GSS-ELEC", "horas": 50.0, "objetivo": 85.0, "nota": ""},
        ])
    )
    agg = hh.agrupar_mensual(hh.leer_sql("SELECT * FROM confirmaciones"), "trbjo_real", "fin", True)
    civi = hh.resumen_por_grupo(agg, "GSS-CIVI").loc["2026-08"]
    elec = hh.resumen_por_grupo(agg, "GSS-ELEC").loc["2026-08"]
    assert civi["trabajadas"] == 100.0
    assert elec["trabajadas"] == 50.0
    assert civi["cumplimiento"] == 100.0
    assert elec["cumplimiento"] == 100.0


def test_maestro_dos_grupos_dos_graficos(isolated_db):
    res = extraer_tabla(LISTA_DOS_GRUPOS.encode("utf-8"), "t.xls")
    hh.guardar(res.df)
    path = hh.exportar_maestro()
    wb = load_workbook(path)
    ws = wb["Resumen"]
    valores = [c.value for row in ws.iter_rows(max_col=1) for c in row if c.value]
    assert "Grupo: GSS-CIVI" in valores
    assert "Grupo: GSS-ELEC" in valores
    assert len(ws._charts) == 2