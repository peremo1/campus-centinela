#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Campus Centinela - Nucleo de descarga y seguimiento de certificados
===================================================================

Inicia sesion en https://www.campuscentinela.cl (Joomla + Joomdle/Moodle),
lee la pagina de progreso de cursos, descarga los certificados PDF, los
ordena en carpetas por categoria (nivel 1) y subcategoria (nivel 2) y genera
un Excel con fechas de emision y vencimiento.

NO hay API publica: la autenticacion es por formulario con token CSRF de
Joomla y las paginas se renderizan en el servidor. Se trabaja con una sesion
(cookie) como lo haria un navegador.

Categorias:
  Nivel 1 (modalidad):  "Aulas Virtuales"  |  "E-Learning"
  Nivel 2 (subcategoria):
    - E-Learning: el prefijo del curso (TRANSV, OM, EDC, DI, DM, MA, MANT,
      MUELLE, OPC, EO, FIN, ...). Si no hay prefijo reconocible -> "Otros".
    - Aulas Virtuales: el area (Seguridad y Salud Ocupacional / Operaciones
      Planta Concentradora y-o Catodo / Mantenimiento Mina), deducida del
      nombre del aula.

El id del certificado (certificado.php?id=NNN) es el id de curso de Moodle,
el mismo que aparece en los enlaces land.php?...&id=NNN. La herramienta
extrae ese id de cada curso y arma la URL del certificado.

Este archivo funciona como modulo (lo usa la interfaz grafica) y como CLI.
Uso CLI:
    python campus_centinela.py --dump        (primera vez: diagnostico)
    python campus_centinela.py                (descarga + Excel)
"""

import argparse
import configparser
import datetime as dt
import getpass
import logging
import os
import re
import sys
import time
import unicodedata
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("Faltan dependencias. Ejecuta:  pip install -r requirements.txt")
    raise

# --------------------------------------------------------------------------- #
# Constantes
# --------------------------------------------------------------------------- #

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
TIMEOUT = 40
RUTA_LOGIN = "/ingresar"
RUTA_PORTADA = "/portada"
RUTA_PROGRESO = "/progreso-cursos"
# Ruta del generador de certificados (segun el ejemplo del usuario).
PLANTILLA_CERT = "/lms/certificados/2_descarga_certificado/certificado.php?id={id}&rut={rut}"

RE_FECHA = re.compile(r"\b(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})\b")
# Cualquier URL que contenga certificado.php
RE_CERT_URL = re.compile(r"[^\s\"'<>]*certificado\.php\?[^\s\"'<>]*", re.IGNORECASE)
# id dentro de una querystring
RE_ID = re.compile(r"[?&]id=(\d+)")
# Atributos donde puede venir un enlace a curso/certificado
ATRIBUTOS_URL = ("href", "onclick", "data-href", "data-url", "formaction", "action")
# Prefijos conocidos de E-Learning (subcategorias). Se comparan en mayuscula.
PREFIJOS_CONOCIDOS = {
    "DI", "DM", "EDC", "EO", "FIN", "MA", "MANT", "MUELLE", "OM", "OPC",
    "TRANSV", "OPERACIONES", "TRANSVERSAL",
}

log = logging.getLogger("campus")


# --------------------------------------------------------------------------- #
# Utilidades de texto
# --------------------------------------------------------------------------- #

def _sin_acentos(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texto or "")
                    if unicodedata.category(c) != "Mn")


def limpiar_nombre(texto: str, maximo: int = 120) -> str:
    """Nombre seguro para archivo/carpeta (conserva acentos, quita ilegales)."""
    texto = re.sub(r"\s+", " ", (texto or "").strip())
    texto = re.sub(r'[\\/:*?"<>|]+', "-", texto)
    texto = texto.strip(" .-")
    return (texto or "sin_nombre")[:maximo]


def normalizar_rut(valor: str) -> str:
    """Limpia un RUT/usuario: quita puntos y espacios y pasa a mayusculas.
    Conserva el guion. Ej: ' 21.080.196-0 ' -> '21080196-0'."""
    return (valor or "").strip().replace(".", "").replace(" ", "").upper()


def todas_las_fechas(texto: str):
    fechas = []
    for m in RE_FECHA.finditer(texto or ""):
        d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a < 100:
            a += 2000
        try:
            fechas.append(dt.date(a, mes, d))
        except ValueError:
            pass
    return fechas


# --------------------------------------------------------------------------- #
# Clasificacion en categoria (nivel 1) y subcategoria (nivel 2)
# --------------------------------------------------------------------------- #

def _area_aula(nombre: str) -> str:
    """Deduce el area de un Aula Virtual a partir de su nombre."""
    n = _sin_acentos(nombre).lower()
    if "equipos moviles" in n or "equipos movil" in n:
        return "Mantenimiento Mina"
    if "operador" in n or "reglamento general de aislacion" in n:
        return "Operaciones Planta Concentradora y-o Catodo"
    # El resto de las aulas IRL del catalogo son de SSO.
    return "Seguridad y Salud Ocupacional"


def clasificar_curso(nombre: str):
    """Devuelve (modalidad, subcategoria) para un curso.

    modalidad: "Aulas Virtuales" o "E-Learning".
    subcategoria: area (aulas) o prefijo (e-learning) o "Otros".
    """
    n = (nombre or "").strip()
    low = _sin_acentos(n).lower()

    # Aulas Virtuales: el nombre empieza con "aula"
    if low.startswith("aula"):
        return "Aulas Virtuales", _area_aula(n)

    # E-Learning: subcategoria = prefijo (primer token)
    tokens = n.split()
    primer = tokens[0] if tokens else ""
    pref = _sin_acentos(primer).upper().strip(".")
    if pref in PREFIJOS_CONOCIDOS:
        return "E-Learning", pref
    # Cualquier token inicial en MAYUSCULAS de 2 a 7 letras se trata como codigo
    if primer.isupper() and primer.isalpha() and 2 <= len(primer) <= 7:
        return "E-Learning", pref
    return "E-Learning", "Otros"


# --------------------------------------------------------------------------- #
# Configuracion
# --------------------------------------------------------------------------- #

class Config:
    def __init__(self, usuario, contrasena, rut, carpeta_salida, url_base):
        self.usuario = normalizar_rut(usuario)
        self.contrasena = contrasena
        self.rut = normalizar_rut(rut) or self.usuario
        self.carpeta_salida = carpeta_salida
        self.url_base = url_base.rstrip("/")

    def url_certificado(self, curso_id: str) -> str:
        return self.url_base + PLANTILLA_CERT.format(id=curso_id, rut=self.rut)


def cargar_config(ruta_ini="config.ini", pedir_interactivo=False,
                  usuario=None, contrasena=None, rut=None,
                  carpeta=None, url_base=None) -> Config:
    """Carga credenciales. Prioridad: argumentos > entorno > config.ini > prompt."""
    usuario = usuario or os.environ.get("CC_USUARIO", "")
    contrasena = contrasena or os.environ.get("CC_CONTRASENA", "")
    rut = rut or os.environ.get("CC_RUT", "")
    carpeta = carpeta or "salida"
    url_base = url_base or "https://www.campuscentinela.cl"

    p = Path(ruta_ini)
    if p.exists():
        cp = configparser.ConfigParser()
        cp.read(p, encoding="utf-8")
        usuario = usuario or cp.get("credenciales", "usuario", fallback="").strip()
        contrasena = contrasena or cp.get("credenciales", "contrasena", fallback="").strip()
        rut = rut or cp.get("credenciales", "rut", fallback="").strip()
        carpeta = cp.get("opciones", "carpeta_salida", fallback=carpeta).strip() or carpeta
        url_base = cp.get("opciones", "url_base", fallback=url_base).strip() or url_base

    if pedir_interactivo:
        if not usuario:
            usuario = input("RUT / usuario: ").strip()
        if not contrasena:
            contrasena = getpass.getpass("Contrasena: ")

    if not usuario or not contrasena:
        raise ValueError(
            "Faltan credenciales. Completa config.ini, define CC_USUARIO/"
            "CC_CONTRASENA, o usa --interactivo."
        )
    if contrasena.upper() == "CAMBIAME":
        raise ValueError("La contrasena en config.ini sigue en 'CAMBIAME'.")
    return Config(usuario, contrasena, rut, carpeta, url_base)


# --------------------------------------------------------------------------- #
# Sesion y login (Joomla com_users)
# --------------------------------------------------------------------------- #

def crear_sesion() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "es-CL,es;q=0.9"})
    return s


def _buscar_formulario_login(soup):
    campo = soup.find("input", {"name": "username"}) or soup.find("input", {"name": "password"})
    if campo:
        form = campo.find_parent("form")
        if form:
            return form
    for form in soup.find_all("form"):
        if form.find("input", {"type": "password"}):
            return form
    return None


def _intentar_login(session: requests.Session, cfg: Config) -> bool:
    """Un intento de login. Devuelve True si quedo autenticado.

    Reenvia TODOS los campos ocultos del formulario real (token CSRF de Joomla
    incluido) y verifica consultando /progreso-cursos.
    """
    url_login = cfg.url_base + RUTA_LOGIN
    r = session.get(url_login, timeout=TIMEOUT)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    form = _buscar_formulario_login(soup)
    if form is None:
        raise RuntimeError("No se encontro el formulario de login. Revisa la URL base "
                           "o usa Diagnostico para guardar el HTML.")
    action = urljoin(url_login, form.get("action") or url_login)
    datos = {}
    for inp in form.find_all(["input", "button"]):
        nombre = inp.get("name")
        if nombre:
            datos[nombre] = inp.get("value", "")
    datos["username"] = cfg.usuario
    datos["password"] = cfg.contrasena
    datos["remember"] = "yes"
    resp = session.post(action, data=datos, timeout=TIMEOUT, allow_redirects=True,
                        headers={"Referer": url_login})
    resp.raise_for_status()
    verif = session.get(cfg.url_base + RUTA_PROGRESO, timeout=TIMEOUT)
    return not _parece_pagina_login(verif.text)


def login(session: requests.Session, cfg: Config, intentos: int = 4) -> None:
    """Login con reintentos.

    El sitio rechaza de forma intermitente el primer POST (se observa que un
    segundo intento entra). Por eso se reintenta con la sesion limpia hasta
    'intentos' veces antes de darse por vencido.
    """
    log.info("Abriendo pagina de login...")
    for intento in range(1, intentos + 1):
        try:
            if _intentar_login(session, cfg):
                log.info("Sesion iniciada correctamente.")
                return
            log.warning("La plataforma no acepto el acceso (intento %d/%d). "
                        "Reintentando...", intento, intentos)
        except requests.RequestException as e:
            log.warning("Problema de red en el login (intento %d/%d): %s",
                        intento, intentos, e)
        session.cookies.clear()
        time.sleep(min(2 * intento, 6))
    raise RuntimeError(
        "No se pudo iniciar sesion tras %d intentos. Revisa el RUT y la contrasena. "
        "Si usas documento extranjero, verifica que esa sea la credencial correcta."
        % intentos)


def _parece_pagina_login(html: str) -> bool:
    if not html:
        return True
    return (('name="password"' in html or "name='password'" in html) and
            ('name="username"' in html or "name='username'" in html))


# --------------------------------------------------------------------------- #
# Extraccion de cursos / certificados
# --------------------------------------------------------------------------- #

class Curso:
    def __init__(self, curso_id, rut, nombre, modalidad, subcategoria,
                 estado, fecha_emision, fecha_vencimiento, cert_url, texto_crudo):
        self.curso_id = curso_id
        self.rut = rut
        self.nombre = nombre
        self.modalidad = modalidad
        self.subcategoria = subcategoria
        self.estado = estado
        self.fecha_emision = fecha_emision
        self.fecha_vencimiento = fecha_vencimiento
        self.cert_url = cert_url
        self.texto_crudo = texto_crudo
        self.archivo = None

    def dias_para_vencer(self):
        if not self.fecha_vencimiento:
            return None
        return (self.fecha_vencimiento - dt.date.today()).days

    def situacion(self):
        d = self.dias_para_vencer()
        if d is None:
            return "SIN_FECHA"
        if d < 0:
            return "VENCIDO"
        if d <= 60:
            return "POR_VENCER"
        return "VIGENTE"

    def carpeta_relativa(self):
        return Path(limpiar_nombre(self.modalidad, 40)) / limpiar_nombre(self.subcategoria, 50)

    def nombre_archivo(self):
        return f"{limpiar_nombre(self.nombre, 90)}_id{self.curso_id}.pdf"


def _fila_contenedora(enlace):
    tr = enlace.find_parent("tr")
    if tr:
        return tr
    nodo = enlace
    for _ in range(6):
        padre = nodo.find_parent(["div", "li", "article"])
        if padre is None:
            break
        if len(padre.get_text(" ", strip=True)) > 25:
            return padre
        nodo = padre
    return enlace.parent or enlace


def _nombre_curso(fila, enlace, curso_id):
    # 1) Celdas de tabla: la mas larga que no sea solo una fecha.
    celdas = fila.find_all(["td", "th"])
    if celdas:
        textos = [c.get_text(" ", strip=True) for c in celdas]
        textos = [t for t in textos if t and not RE_FECHA.fullmatch(t)
                  and "descarg" not in t.lower() and t.lower() not in ("pdf", "entrar")]
        if textos:
            return max(textos, key=len)
    # 2) Encabezado o enlace dentro de la tarjeta.
    enc = fila.find(["h1", "h2", "h3", "h4", "h5", "strong", "b", "a"])
    if enc:
        t = enc.get_text(" ", strip=True)
        if t and "certificado.php" not in t:
            return t
    # 3) Texto del propio enlace.
    t = enlace.get_text(" ", strip=True)
    if t and t.lower() not in ("descargar", "pdf", "entrar", "entrar (ya inscrito)"):
        return t
    return f"Curso {curso_id}"


def _recolectar_ids(soup):
    """Devuelve dict id -> (tag, cert_url_o_None).

    Busca certificado.php (preferente) y, en su defecto, cualquier enlace con
    id de curso (land.php, course/view.php, mtype=course), ya que el id es el
    mismo id de curso de Moodle y sirve para construir la URL del certificado.
    """
    directos = {}   # id -> (tag, cert_url)
    indirectos = {}  # id -> tag
    for tag in soup.find_all(True):
        for attr in ATRIBUTOS_URL:
            val = tag.get(attr)
            if not val:
                continue
            val = val.replace("&amp;", "&")
            if "certificado.php" in val.lower():
                m = RE_CERT_URL.search(val)
                mid = RE_ID.search(val)
                if mid:
                    directos.setdefault(mid.group(1),
                                        (tag, m.group(0) if m else None))
                break
            if ("land.php" in val.lower() or "course/view.php" in val.lower()
                    or "mtype=course" in val.lower()):
                mid = RE_ID.search(val)
                if mid:
                    indirectos.setdefault(mid.group(1), tag)
                break
    # Combinar: los directos mandan.
    resultado = {}
    for cid, (tag, cert) in directos.items():
        resultado[cid] = (tag, cert)
    for cid, tag in indirectos.items():
        resultado.setdefault(cid, (tag, None))
    return resultado


def extraer_cursos(html: str, cfg: Config):
    """Extrae la lista de cursos con certificado desde el HTML de progreso."""
    soup = BeautifulSoup(html, "html.parser")
    ids = _recolectar_ids(soup)
    cursos = []
    for curso_id, (tag, cert_url) in ids.items():
        fila = _fila_contenedora(tag)
        texto = fila.get_text(" | ", strip=True)
        nombre = _nombre_curso(fila, tag, curso_id)
        modalidad, subcat = clasificar_curso(nombre)
        fechas = todas_las_fechas(texto)
        f_emision = fechas[0] if len(fechas) >= 1 else None
        f_venc = fechas[1] if len(fechas) >= 2 else None
        estado = ""
        for palabra in ("vencido", "vigente", "por vencer", "aprobado", "pendiente"):
            if palabra in texto.lower():
                estado = palabra
                break
        url = urljoin(cfg.url_base, cert_url) if cert_url else cfg.url_certificado(curso_id)
        cursos.append(Curso(curso_id, cfg.rut, nombre, modalidad, subcat, estado,
                            f_emision, f_venc, url, texto))

    cursos.sort(key=lambda c: (c.modalidad, c.subcategoria, c.nombre.lower()))
    log.info("Se detectaron %d cursos con id de certificado.", len(cursos))
    if not cursos:
        log.warning("No se detectaron ids de curso/certificado. El HTML puede ser "
                    "distinto o cargarse por JavaScript. Usa --dump y revisa "
                    "diagnostico/progreso-cursos.html.")
    return cursos


# --------------------------------------------------------------------------- #
# Descarga
# --------------------------------------------------------------------------- #

def descargar_certificados(session, cfg, cursos, carpeta_base: Path, progreso_cb=None):
    carpeta_base.mkdir(parents=True, exist_ok=True)
    ok = 0
    total = len(cursos)
    for i, curso in enumerate(cursos, 1):
        destino_dir = carpeta_base / curso.carpeta_relativa()
        destino_dir.mkdir(parents=True, exist_ok=True)
        destino = destino_dir / curso.nombre_archivo()
        try:
            r = session.get(curso.cert_url, timeout=TIMEOUT)
            r.raise_for_status()
            contenido = r.content
            ctype = r.headers.get("Content-Type", "")
            if b"%PDF" not in contenido[:1024] and "pdf" not in ctype.lower():
                log.warning("  [%d/%d] id=%s no devolvio PDF (%s). Probablemente el "
                            "curso no tiene certificado emitido. Se omite.",
                            i, total, curso.curso_id, ctype or "sin content-type")
            else:
                destino.write_bytes(contenido)
                curso.archivo = str(destino)
                ok += 1
                log.info("  [%d/%d] OK  %s", i, total, destino.name)
        except requests.RequestException as e:
            log.error("  [%d/%d] Error id=%s: %s", i, total, curso.curso_id, e)
        if progreso_cb:
            progreso_cb(i, total)
    log.info("Descargados %d de %d certificados.", ok, total)
    return ok


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #

def generar_excel(cursos, ruta_xlsx: Path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Certificados"
    encabezados = ["Modalidad", "Subcategoria", "Curso", "Estado", "Situacion",
                   "Fecha emision", "Fecha vencimiento", "Dias para vencer",
                   "ID curso", "RUT", "Archivo PDF", "URL certificado"]
    ws.append(encabezados)

    cab_fill = PatternFill("solid", fgColor="0B6E7A")
    cab_font = Font(bold=True, color="FFFFFF")
    for col in range(1, len(encabezados) + 1):
        c = ws.cell(row=1, column=col)
        c.fill = cab_fill
        c.font = cab_font
        c.alignment = Alignment(vertical="center")

    rojo = PatternFill("solid", fgColor="F4CCCC")
    amarillo = PatternFill("solid", fgColor="FCE8B2")
    verde = PatternFill("solid", fgColor="D9EAD3")

    def clave(c):
        d = c.dias_para_vencer()
        return (c.modalidad, c.subcategoria, d if d is not None else 10**6)

    for curso in sorted(cursos, key=clave):
        sit = curso.situacion()
        ws.append([
            curso.modalidad, curso.subcategoria, curso.nombre, curso.estado, sit,
            curso.fecha_emision.strftime("%d-%m-%Y") if curso.fecha_emision else "",
            curso.fecha_vencimiento.strftime("%d-%m-%Y") if curso.fecha_vencimiento else "",
            curso.dias_para_vencer() if curso.dias_para_vencer() is not None else "",
            curso.curso_id, curso.rut, curso.archivo or "", curso.cert_url,
        ])
        relleno = {"VENCIDO": rojo, "POR_VENCER": amarillo, "VIGENTE": verde}.get(sit)
        if relleno:
            for col in range(1, len(encabezados) + 1):
                ws.cell(row=ws.max_row, column=col).fill = relleno

    for i, ancho in enumerate([16, 22, 50, 12, 12, 14, 16, 15, 9, 13, 42, 50], 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(encabezados))}{ws.max_row}"

    resumen = wb.create_sheet("Resumen")
    conteo = {"VENCIDO": 0, "POR_VENCER": 0, "VIGENTE": 0, "SIN_FECHA": 0}
    por_modalidad = {}
    for c in cursos:
        conteo[c.situacion()] += 1
        por_modalidad[c.modalidad] = por_modalidad.get(c.modalidad, 0) + 1
    resumen.append(["Generado", dt.datetime.now().strftime("%d-%m-%Y %H:%M")])
    resumen.append(["Total cursos", len(cursos)])
    resumen.append([])
    resumen.append(["Situacion", "Cantidad"])
    for k in ("VENCIDO", "POR_VENCER", "VIGENTE", "SIN_FECHA"):
        resumen.append([k, conteo[k]])
    resumen.append([])
    resumen.append(["Modalidad", "Cantidad"])
    for k, v in sorted(por_modalidad.items()):
        resumen.append([k, v])
    resumen.column_dimensions["A"].width = 24
    resumen.column_dimensions["B"].width = 22
    for celda in ("A4", "B4", "A10", "B10"):
        resumen[celda].font = Font(bold=True)

    ruta_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta_xlsx)
    log.info("Excel generado: %s", ruta_xlsx)
    log.info("Vencidos: %d | Por vencer (<=60d): %d | Vigentes: %d | Sin fecha: %d",
             conteo["VENCIDO"], conteo["POR_VENCER"], conteo["VIGENTE"], conteo["SIN_FECHA"])
    return conteo


# --------------------------------------------------------------------------- #
# Diagnostico
# --------------------------------------------------------------------------- #

def guardar_diagnostico(session, cfg, carpeta: Path):
    carpeta.mkdir(parents=True, exist_ok=True)
    for ruta, nombre in ((RUTA_PORTADA, "portada.html"),
                         (RUTA_PROGRESO, "progreso-cursos.html")):
        try:
            r = session.get(cfg.url_base + ruta, timeout=TIMEOUT)
            (carpeta / nombre).write_text(r.text, encoding="utf-8")
            log.info("Guardado %s (%d bytes)", nombre, len(r.text))
        except requests.RequestException as e:
            log.error("No se pudo guardar %s: %s", nombre, e)


# --------------------------------------------------------------------------- #
# Orquestador (lo usan CLI y GUI)
# --------------------------------------------------------------------------- #

def ejecutar(cfg: Config, dump=False, sin_descarga=False, progreso_cb=None):
    """Flujo completo: login -> progreso -> (descarga) -> Excel.
    Devuelve (cursos, ruta_excel, conteo)."""
    session = crear_sesion()
    login(session, cfg)
    base = Path(cfg.carpeta_salida)
    try:
        base.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        raise RuntimeError(
            "No se pudo crear la carpeta de salida '%s' (%s). Elige una carpeta "
            "con permisos de escritura, por ejemplo en Documentos." % (base, e))
    if dump:
        guardar_diagnostico(session, cfg, base / "diagnostico")
    log.info("Leyendo pagina de progreso de cursos...")
    r = session.get(cfg.url_base + RUTA_PROGRESO, timeout=TIMEOUT)
    r.raise_for_status()
    cursos = extraer_cursos(r.text, cfg)
    if not cursos:
        return [], None, {}
    if not sin_descarga:
        descargar_certificados(session, cfg, cursos, base / "certificados", progreso_cb)
    ruta_xlsx = base / "certificados_centinela.xlsx"
    conteo = generar_excel(cursos, ruta_xlsx)
    return cursos, ruta_xlsx, conteo


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Descarga certificados de Campus Centinela y genera un Excel.")
    parser.add_argument("--config", default="config.ini")
    parser.add_argument("--interactivo", action="store_true")
    parser.add_argument("--dump", action="store_true",
                        help="Guarda el HTML de portada y progreso para diagnostico")
    parser.add_argument("--sin-descarga", action="store_true",
                        help="Solo genera el Excel")
    parser.add_argument("--salida", default=None)
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-7s %(message)s", datefmt="%H:%M:%S")

    try:
        cfg = cargar_config(args.config, args.interactivo, carpeta=args.salida)
    except ValueError as e:
        log.error("%s", e)
        return 2

    try:
        cursos, ruta, _ = ejecutar(cfg, dump=args.dump, sin_descarga=args.sin_descarga)
    except Exception as e:
        log.error("Fallo la ejecucion: %s", e)
        return 2
    if not cursos:
        log.error("No hay cursos que procesar. Usa --dump y revisa el HTML.")
        return 1
    log.info("Listo. Resultados en la carpeta '%s'.", cfg.carpeta_salida)
    return 0


if __name__ == "__main__":
    sys.exit(main())
