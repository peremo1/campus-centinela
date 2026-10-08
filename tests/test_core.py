# -*- coding: utf-8 -*-
"""Pruebas del nucleo con nombres reales de cursos de Campus Centinela.

No tocan la red: validan clasificacion (modalidad + subcategoria), extraccion
de ids, construccion de la URL de certificado y generacion del Excel.

Ejecutar:  python tests/test_core.py
"""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import campus_centinela as cc

cfg = cc.Config(usuario="21080196-0", contrasena="x", rut="21080196-0",
                carpeta_salida="salida", url_base="https://www.campuscentinela.cl")

# --------------------------------------------------------------------------- #
# 1) Clasificacion: casos reales (modalidad, subcategoria esperada)
# --------------------------------------------------------------------------- #
CASOS = [
    # Aulas Virtuales
    ("Aula IRL Mina General", "Aulas Virtuales", "Seguridad y Salud Ocupacional"),
    ("Aula IRL Muelle", "Aulas Virtuales", "Seguridad y Salud Ocupacional"),
    ("Aula IRL SSO Persona Nueva", "Aulas Virtuales", "Seguridad y Salud Ocupacional"),
    ("Aula IRL Mantenimiento Mina", "Aulas Virtuales", "Seguridad y Salud Ocupacional"),
    ("AULA IRL Depósito de Relaves Espesados", "Aulas Virtuales", "Seguridad y Salud Ocupacional"),
    ("Aula IRL Autonomía", "Aulas Virtuales", "Seguridad y Salud Ocupacional"),
    ("AULA Reglamento General de Aislación y Bloqueo", "Aulas Virtuales",
     "Operaciones Planta Concentradora y-o Catodo"),
    ("AULA Operador/a de Aislamiento", "Aulas Virtuales",
     "Operaciones Planta Concentradora y-o Catodo"),
    ("AULA Reglamento de Aislación y Bloqueo para Equipos Móviles", "Aulas Virtuales",
     "Mantenimiento Mina"),
    # E-Learning por prefijo
    ("DI Ambientes Respetuosos - Ley Karin", "E-Learning", "DI"),
    ("DM Camión de Extracción KOMATSU 980 E-5 AHT", "E-Learning", "DM"),
    ("EDC Trabajos en Altura", "E-Learning", "EDC"),
    ("EDC Prevención de la Silicosis", "E-Learning", "EDC"),
    ("EO Tengo Una Idea", "E-Learning", "EO"),
    ("FIN Administradores de Contrato en la Práctica", "E-Learning", "FIN"),
    ("MA Cambio Climático", "E-Learning", "MA"),
    ("MANT Bloqueo Digital Biométrico", "E-Learning", "MANT"),
    ("MUELLE Conducción Cuesta Michilla", "E-Learning", "MUELLE"),
    ("OM Geotecnia Operativa Básica", "E-Learning", "OM"),
    ("Om Procedimiento Comunicaciones", "E-Learning", "OM"),          # mixed case
    ("OPC Manejo de Ácido Sulfúrico", "E-Learning", "OPC"),
    ("TRANSV Primeros Auxilios", "E-Learning", "TRANSV"),
    ("Transv Art 360", "E-Learning", "TRANSV"),                       # mixed case
    ("Transv Recss", "E-Learning", "TRANSV"),
    # Sin prefijo reconocible
    ("Inducción Medio Ambiente Proyecto Nueva Centinela", "E-Learning", "Otros"),
    ("Estandar por Contingencia de Gases Nitrosos", "E-Learning", "Otros"),
]


def test_clasificacion():
    errores = []
    for nombre, mod_esp, sub_esp in CASOS:
        mod, sub = cc.clasificar_curso(nombre)
        if (mod, sub) != (mod_esp, sub_esp):
            errores.append(f"  '{nombre}' -> ({mod}, {sub})  esperado ({mod_esp}, {sub_esp})")
    assert not errores, "Clasificaciones incorrectas:\n" + "\n".join(errores)
    print(f"[OK] clasificacion: {len(CASOS)} casos correctos")


# --------------------------------------------------------------------------- #
# 2) Extraccion de ids + construccion de URL de certificado
# --------------------------------------------------------------------------- #
HTML_PROGRESO = """
<h3>Mis cursos</h3>
<table>
 <tr><th>Curso</th><th>Emisión</th><th>Vence</th><th>Certificado</th></tr>
 <tr><td>TRANSV Primeros Auxilios</td><td>10-01-2024</td><td>10-01-2026</td>
     <td><a href="/lms/certificados/2_descarga_certificado/certificado.php?id=223&amp;rut=21080196-0">Descargar</a></td></tr>
 <tr><td>OPC Manejo de Ácido Sulfúrico</td><td>05-05-2025</td><td>05-05-2027</td>
     <td><a href="https://www.campuscentinela.cl/lms/certificados/2_descarga_certificado/certificado.php?id=134&rut=21080196-0">Descargar</a></td></tr>
 <tr><td>MA Cambio Climático</td><td>01-02-2023</td><td>01-02-2025</td>
     <td><button onclick="location.href='/lms//auth/joomdle/land.php?username=21080196-0&token=TOKENFALSO&mtype=course&id=178&use_wrapper=1&Itemid=405'">Entrar</button></td></tr>
</table>
"""


def test_extraccion():
    cursos = cc.extraer_cursos(HTML_PROGRESO, cfg)
    por_id = {c.curso_id: c for c in cursos}
    assert set(por_id) == {"223", "134", "178"}, f"ids: {set(por_id)}"

    # 223: enlace certificado.php directo
    c223 = por_id["223"]
    assert (c223.modalidad, c223.subcategoria) == ("E-Learning", "TRANSV")
    assert "certificado.php?id=223" in c223.cert_url

    # 134: enlace certificado.php absoluto
    c134 = por_id["134"]
    assert (c134.modalidad, c134.subcategoria) == ("E-Learning", "OPC")
    assert "certificado.php?id=134" in c134.cert_url

    # 178: SOLO habia land.php -> la URL de certificado se CONSTRUYE desde el id
    c178 = por_id["178"]
    assert (c178.modalidad, c178.subcategoria) == ("E-Learning", "MA")
    assert c178.cert_url == cfg.url_certificado("178"), c178.cert_url
    assert "certificado.php?id=178" in c178.cert_url
    assert "rut=21080196-0" in c178.cert_url

    print("[OK] extraccion: 3 cursos, ids correctos, URL de certificado construida "
          "desde land.php cuando no hay enlace directo")


# --------------------------------------------------------------------------- #
# 3) Fechas / situacion y estructura de carpetas
# --------------------------------------------------------------------------- #
def test_fechas_y_carpetas():
    cursos = cc.extraer_cursos(HTML_PROGRESO, cfg)
    por_id = {c.curso_id: c for c in cursos}
    assert por_id["223"].fecha_emision.isoformat() == "2024-01-10"
    assert por_id["223"].fecha_vencimiento.isoformat() == "2026-01-10"
    # 134 vence 2027 -> vigente ; 178 vence 2025 -> vencido
    assert por_id["134"].situacion() == "VIGENTE", por_id["134"].situacion()
    assert por_id["178"].situacion() == "VENCIDO", por_id["178"].situacion()
    # Carpeta relativa de 2 niveles
    assert str(por_id["223"].carpeta_relativa()).replace("\\", "/") == "E-Learning/TRANSV"
    assert por_id["223"].nombre_archivo().endswith("_id223.pdf")
    print("[OK] fechas, situacion y carpetas de 2 niveles correctas")


# --------------------------------------------------------------------------- #
# 4) Generacion de Excel
# --------------------------------------------------------------------------- #
def test_excel():
    cursos = cc.extraer_cursos(HTML_PROGRESO, cfg)
    out = Path(tempfile.gettempdir()) / "test_centinela.xlsx"
    if out.exists():
        out.unlink()
    conteo = cc.generar_excel(cursos, out)
    assert out.exists() and out.stat().st_size > 0
    assert conteo["VENCIDO"] >= 1 and conteo["VIGENTE"] >= 1
    # Verificar que se puede reabrir y tiene las columnas nuevas
    from openpyxl import load_workbook
    wb = load_workbook(out)
    ws = wb["Certificados"]
    encabezados = [c.value for c in ws[1]]
    assert encabezados[0] == "Modalidad" and encabezados[1] == "Subcategoria"
    print(f"[OK] excel generado y reabierto ({out.stat().st_size} bytes); "
          f"conteo={conteo}")


if __name__ == "__main__":
    fallos = 0
    for fn in (test_clasificacion, test_extraccion, test_fechas_y_carpetas, test_excel):
        try:
            fn()
        except AssertionError as e:
            fallos += 1
            print(f"[FALLO] {fn.__name__}:\n{e}")
    if fallos:
        print(f"\n{fallos} prueba(s) fallaron.")
        sys.exit(1)
    print("\nTODAS LAS PRUEBAS PASARON")
