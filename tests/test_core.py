# -*- coding: utf-8 -*-
"""Pruebas del nucleo (sin red) de Campus Centinela.

Validan: login (campos del formulario real), normalizacion de RUT,
clasificacion por prefijo, y el consumo de la API JSON de cursos
(categorias reales, fechas de vigencia, firma digital, filtros y Excel).

Ejecutar:  python tests/test_core.py
"""
import sys
import tempfile
import datetime as dt
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import campus_centinela as cc

BASE = "https://www.campuscentinela.cl"
cfg = cc.Config("21080196-0", "x", "", "salida", BASE)


# --------------------------------------------------------------------------- #
# 1) Clasificacion por prefijo (subcategoria)
# --------------------------------------------------------------------------- #
CASOS = [
    ("DI Ambientes Respetuosos - Ley Karin", "DI"),
    ("DM Camion de Extraccion KOMATSU 980", "DM"),
    ("EDC Trabajos en Altura", "EDC"),
    ("MA Cambio Climatico", "MA"),
    ("OM Geotecnia Operativa Basica", "OM"),
    ("Om Procedimiento Comunicaciones", "OM"),
    ("OPC Manejo de Acido Sulfurico", "OPC"),
    ("TRANSV Primeros Auxilios", "TRANSV"),
    ("Transv Art 360", "TRANSV"),
    ("MUELLE Conduccion Cuesta Michilla", "MUELLE"),
    ("Induccion Medio Ambiente Proyecto Nueva Centinela", "Otros"),
]


def test_clasificacion():
    for nombre, sub in CASOS:
        got = cc.subcategoria_de(nombre)
        assert got == sub, f"'{nombre}' -> {got}, esperado {sub}"
    print(f"[OK] subcategoria por prefijo: {len(CASOS)} casos")


# --------------------------------------------------------------------------- #
# 2) Login: solo radios marcados; opcion nacional/extranjero
# --------------------------------------------------------------------------- #
FORM_LOGIN = """
<form action="https://www.campuscentinela.cl/ingresar?task=user.login" method="post" class="form-validate">
  <label><input type="radio" name="opcion" value="nacional" checked> NACIONAL</label>
  <label><input type="radio" name="opcion" value="extranjero"> EXTRANJERO</label>
  <input type="text" name="username" value="">
  <input type="password" name="password" value="">
  <input type="checkbox" name="remember" value="yes">
  <input type="hidden" name="return" value="aHR0cHM=">
  <input type="hidden" name="20825ccc03ed06f429d141237fc54fa7" value="1">
</form>
"""


def test_campos_login():
    from bs4 import BeautifulSoup
    form = BeautifulSoup(FORM_LOGIN, "html.parser").find("form")
    datos = cc._campos_formulario(form, cc.Config("21.080.196-0", "clave", "", ".", BASE))
    assert datos["opcion"] == "nacional"
    assert datos["username"] == "21080196-0"
    assert datos["password"] == "clave"
    assert datos["remember"] == "yes"
    # el token CSRF (campo oculto) se reenvia
    assert datos.get("20825ccc03ed06f429d141237fc54fa7") == "1"
    datos2 = cc._campos_formulario(form, cc.Config("X", "y", "", ".", BASE, extranjero=True))
    assert datos2["opcion"] == "extranjero"
    print("[OK] login: opcion correcta y token CSRF reenviado")


def test_normalizar_rut():
    assert cc.normalizar_rut(" 21.080.196-0 ") == "21080196-0"
    assert cc.normalizar_rut("12.345.678-k") == "12345678-K"
    c = cc.Config("21.080.196-0", "x", "", "salida", BASE)
    assert c.usuario == "21080196-0" and c.rut == "21080196-0"
    print("[OK] normalizacion de RUT")


# --------------------------------------------------------------------------- #
# 3) Calculo de vencimiento (replica la formula de la plataforma)
# --------------------------------------------------------------------------- #
def test_vencimiento():
    f = dt.date(2024, 3, 15)
    assert cc._calcular_vencimiento("50", "1 año", f) == (dt.date(2025, 3, 15), False)
    assert cc._calcular_vencimiento("50", "2 años", f) == (dt.date(2026, 3, 15), False)
    assert cc._calcular_vencimiento("50", "6 meses", f) == (dt.date(2024, 9, 15), False)
    assert cc._calcular_vencimiento("50", "Indefinida", f) == (None, False)
    # curso de vigencia fija -> no vence
    assert cc._calcular_vencimiento("1411", "1 año", f) == (None, True)
    # la vigencia puede venir como FECHA (caso aulas) -> esa es la expiracion
    assert cc._calcular_vencimiento("801", "2026-12-31", None) == (dt.date(2026, 12, 31), False)
    assert cc._calcular_vencimiento("801", "31-12-2026", None) == (dt.date(2026, 12, 31), False)
    assert cc._parse_fecha_iso("2024-03-15 10:20:30") == dt.date(2024, 3, 15)
    print("[OK] vencimiento: años/meses/indefinida/fija/FECHA y parseo")


# --------------------------------------------------------------------------- #
# 4) Consumo de la API JSON (con sesion simulada)
# --------------------------------------------------------------------------- #
JSON_CURSOS = """{"success":true,"data":[
 {"curso_id":"134","cat_id":"10","curso_cat":"OPERACIONES PLANTA CONCENTRADORA Y/O CATODO",
  "curso_nombre":"OPC Manejo de Acido Sulfurico","estado_curso":"APROBADO","notobt":"100",
  "porcentaje_avance":"100","vigencia":"1 año","fecha_nota":"2024-03-15 00:00:00",
  "certificado_directo":"https://www.campuscentinela.cl/lms/certificados/2_descarga_certificado/certificado.php?id=134&rut=21080196-0",
  "firma_digital":"0","s":"0","d":"0"},
 {"curso_id":"223","cat_id":"5","curso_cat":"SEGURIDAD Y SALUD OCUPACIONAL",
  "curso_nombre":"TRANSV Primeros Auxilios","estado_curso":"APROBADO","notobt":"90",
  "porcentaje_avance":"100","vigencia":"2 años","fecha_nota":"2023-01-10",
  "certificado_directo":"","firma_digital":"191780","s":"0","d":"0"},
 {"curso_id":"999","cat_id":"7","curso_cat":"SUSTENTABILIDAD Y MEDIO AMBIENTE",
  "curso_nombre":"MA Cambio Climatico","estado_curso":"PENDIENTE","notobt":"0",
  "porcentaje_avance":"20","vigencia":"6 meses","fecha_nota":"","firma_digital":"0","s":"0","d":"0"},
 {"curso_id":"77","cat_id":"28","curso_cat":"EXCLUIDA","curso_nombre":"XX Excluido",
  "estado_curso":"APROBADO","vigencia":"1 año","fecha_nota":"2024-01-01","s":"0","d":"0"},
 {"curso_id":"2594","cat_id":"9","curso_cat":"OTRA","curso_nombre":"Oculto",
  "estado_curso":"APROBADO","s":"0","d":"0"},
 {"curso_id":"1411","cat_id":"3","curso_cat":"MUELLE","curso_nombre":"MUELLE Conduccion Cuesta Michilla",
  "estado_curso":"APROBADO","notobt":"100","porcentaje_avance":"100","vigencia":"1 año",
  "fecha_nota":"2020-01-01","certificado_directo":"","firma_digital":"0","s":"0","d":"0"},
 {"curso_id":"801","cat_id":"4","curso_cat":"OPERACIONES PLANTA CONCENTRADORA Y/O CATODO",
  "curso_nombre":"AULA Reglamento General de Aislacion y Bloqueo","estado_curso":"APROBADO",
  "notobt":"100","porcentaje_avance":"100","vigencia":"2026-12-31","fecha_nota":"2026-08-27",
  "certificado_directo":"","firma_digital":"0","s":"0","d":"0"}
]}"""


class _FakeResp:
    def __init__(self, text):
        self.text = text
        self.headers = {"Content-Type": "application/json"}
        self.content = text.encode()
    def raise_for_status(self):
        pass


class _FakeSession:
    def __init__(self, text):
        self._text = text
    def get(self, url, **kw):
        return _FakeResp(self._text)


def test_obtener_cursos():
    cursos = cc.obtener_cursos(_FakeSession(JSON_CURSOS), cfg, diag_dir=None)
    por_id = {c.curso_id: c for c in cursos}
    # Se excluyen cat_id 28 y el curso 2594
    assert set(por_id) == {"134", "223", "999", "1411", "801"}, set(por_id)

    # 801: la vigencia venia como FECHA -> esa es la expiracion (bug corregido)
    c801 = por_id["801"]
    assert c801.fecha_vencimiento == dt.date(2026, 12, 31), c801.fecha_vencimiento
    assert c801.subcategoria in ("Operaciones Planta Concentradora y-o Catodo",)
    assert c801.situacion() in ("VIGENTE", "POR_VENCER", "VENCIDO")

    c134 = por_id["134"]
    assert c134.categoria == "OPERACIONES PLANTA CONCENTRADORA Y/O CATODO"
    assert c134.subcategoria == "OPC"
    assert "certificado.php?id=134" in c134.cert_url
    assert c134.fecha_emision == dt.date(2024, 3, 15)
    assert c134.fecha_vencimiento == dt.date(2025, 3, 15)
    assert c134.situacion() == "VENCIDO"            # vencio en 2025
    assert c134.aprobado

    c223 = por_id["223"]
    assert c223.subcategoria == "TRANSV"
    assert "certificado.php?id=223" in c223.cert_url  # construida (venia vacia)
    assert c223.firma_url.endswith("IdDoc=191780")

    c999 = por_id["999"]
    assert c999.situacion() == "PENDIENTE" and not c999.aprobado

    c1411 = por_id["1411"]
    assert c1411.vigencia_fija and c1411.situacion() == "VIGENCIA_FIJA"
    assert c1411.fecha_vencimiento is None

    # carpetas de 2 niveles: categoria / subcategoria
    assert str(c134.carpeta_relativa()).replace("\\", "/") == \
        "OPERACIONES PLANTA CONCENTRADORA Y-O CATODO/OPC"
    print("[OK] API JSON: 6 -> 4 cursos (filtros), categorias/fechas/firma/vigencia fija")


def test_excel():
    cursos = cc.obtener_cursos(_FakeSession(JSON_CURSOS), cfg, diag_dir=None)
    out = Path(tempfile.gettempdir()) / "test_centinela.xlsx"
    if out.exists():
        out.unlink()
    conteo = cc.generar_excel(cursos, out, nombre="Jonathan Veliz Soza",
                              rut="21080196-0", umbral=60)
    assert out.exists() and out.stat().st_size > 0
    from openpyxl import load_workbook
    wb = load_workbook(out)
    assert "Panel" in wb.sheetnames and "Cursos" in wb.sheetnames
    assert wb.sheetnames[0] == "Panel"

    ws = wb["Cursos"]
    enc = [c.value for c in ws[1]]
    assert enc[0] == "Categoria"
    assert "Dias restantes" in enc and "Situacion" in enc and "Expiracion" in enc
    assert "Fuente vig." in enc

    # Debe existir al menos una formula dinamica (Dias restantes = Exp - HOY())
    formulas_dias = [ws.cell(r, 6).value for r in range(2, ws.max_row + 1)]
    assert any(str(v).startswith("=IF(") and "TODAY()" in str(v) for v in formulas_dias), \
        "falta la formula dinamica de dias restantes"
    # Situacion dinamica referencia la CELDA del umbral (C8, no la etiqueta B8)
    formulas_sit = [ws.cell(r, 7).value for r in range(2, ws.max_row + 1)]
    assert any("Panel!$C$8" in str(v) for v in formulas_sit), \
        "la situacion debe referenciar Panel!$C$8 (la celda con el numero)"
    assert not any("Panel!$B$8" in str(v) for v in formulas_sit), \
        "no debe referenciar B8 (ahi esta la etiqueta de texto)"

    panel = wb["Panel"]
    assert panel["C8"].value == 60                      # umbral editable
    assert panel["C5"].value == "Jonathan Veliz Soza"   # nombre
    # KPIs dinamicos con COUNTIF sobre la hoja Cursos
    valores_panel = [panel.cell(r, 3).value for r in range(11, 20)]
    assert any("COUNTIF(Cursos" in str(v) for v in valores_panel), \
        "faltan los KPIs dinamicos en el Panel"

    assert conteo.get("VENCIDO", 0) >= 2
    print(f"[OK] excel: Panel+Cursos, formulas dinamicas y umbral; conteo={dict(conteo)}")


if __name__ == "__main__":
    fallos = 0
    for fn in (test_clasificacion, test_campos_login, test_normalizar_rut,
               test_vencimiento, test_obtener_cursos, test_excel):
        try:
            fn()
        except AssertionError as e:
            fallos += 1
            print(f"[FALLO] {fn.__name__}:\n{e}")
    if fallos:
        print(f"\n{fallos} prueba(s) fallaron.")
        sys.exit(1)
    print("\nTODAS LAS PRUEBAS PASARON")
