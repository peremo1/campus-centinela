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
import calendar
import configparser
import datetime as dt
import getpass
import json
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
# API JSON interna que usa la pagina de progreso para traer los cursos.
RUTA_JSON = ("/index.php?option=com_ajax&module=customphp&"
             "method=consultorcursosSegundaVersion&format=json&rut={rut}")
RUTA_FIRMA = "/lms/firmadigital/documento.php?IdDoc={id}"

# Cursos de "vigencia fija" (no vencen) - tal cual los marca la plataforma.
VIGENCIA_FIJA = {65, 68, 69, 70, 73, 75, 76, 79, 121, 207, 293, 294, 325, 1398,
                 1399, 1411, 2180, 2181, 2457, 2619, 2621, 2623, 2625, 2626,
                 2628, 2629, 2630, 2631, 2633, 2637}
# Categorias que la plataforma excluye de "Mis Cursos Generales".
EXCLUIR_CAT = {28, 38, 39, 40, 41, 63, 64, 65, 87, 88, 89, 92, 93, 94, 96, 97,
               99, 100, 101, 102, 103, 104, 109, 27, 67, 52, 69, 86, 105, 106,
               107, 108, 111, 138, 139, 140, 141}
# Cursos individuales que la plataforma oculta.
SKIP_CURSO = {2594}

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


def subcategoria_de(nombre: str) -> str:
    """Subcategoria (prefijo) de un curso: TRANSV, OM, EDC, MA, ... u 'Otros'."""
    return clasificar_curso(nombre)[1]


# --------------------------------------------------------------------------- #
# Ayudantes para la API JSON (fechas, vigencia)
# --------------------------------------------------------------------------- #

def _num_inicial(texto: str) -> int:
    m = re.search(r"\d+", texto or "")
    return int(m.group()) if m else 0


def _parse_fecha_iso(valor) -> "dt.date|None":
    """Parsea 'YYYY-MM-DD[ HH:MM:SS]' o 'DD-MM-YYYY' a date."""
    s = str(valor or "").strip()
    if not s or s in ("0000-00-00", "0000-00-00 00:00:00"):
        return None
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", s)
    if m:
        try:
            return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def _add_years(d: dt.date, n: int) -> dt.date:
    try:
        return d.replace(year=d.year + n)
    except ValueError:  # 29-feb
        return d.replace(year=d.year + n, day=28)


def _add_months(d: dt.date, n: int) -> dt.date:
    total = d.month - 1 + n
    y = d.year + total // 12
    m = total % 12 + 1
    day = min(d.day, calendar.monthrange(y, m)[1])
    return dt.date(y, m, day)


def _calcular_vencimiento(curso_id, vigencia, fecha_nota):
    """Replica la logica de la plataforma: emision (fecha_nota) + vigencia.

    Devuelve (fecha_vencimiento|None, vigencia_fija:bool).
    """
    try:
        cid = int(str(curso_id).strip())
    except (TypeError, ValueError):
        cid = -1
    if cid in VIGENCIA_FIJA:
        return None, True
    if not fecha_nota:
        return None, False
    v = _sin_acentos(str(vigencia or "")).lower()
    if not v or "indefinid" in v:
        return None, False
    n = _num_inicial(v)
    if n <= 0:
        return None, False
    if "ano" in v or "año" in str(vigencia or "").lower():
        return _add_years(fecha_nota, n), False
    if "mes" in v:
        return _add_months(fecha_nota, n), False
    return None, False


def _desempaquetar_json(texto: str):
    """Obtiene la lista de cursos desde la respuesta de com_ajax.

    Formatos posibles: {"data":[...]}, {"data":"<json>"}, [...], [[...]].
    """
    raw = json.loads(texto)
    data = raw.get("data", raw) if isinstance(raw, dict) else raw
    if isinstance(data, str):
        data = json.loads(data)
    if (isinstance(data, list) and len(data) == 1
            and isinstance(data[0], (list, str))):
        inner = data[0]
        if isinstance(inner, str):
            inner = json.loads(inner)
        if isinstance(inner, list):
            data = inner
    if not isinstance(data, list):
        raise RuntimeError("La respuesta de cursos no tiene el formato esperado.")
    return data


# --------------------------------------------------------------------------- #
# Configuracion
# --------------------------------------------------------------------------- #

class Config:
    def __init__(self, usuario, contrasena, rut, carpeta_salida, url_base,
                 extranjero=False):
        self.usuario = normalizar_rut(usuario)
        self.contrasena = contrasena
        self.rut = normalizar_rut(rut) or self.usuario
        self.carpeta_salida = carpeta_salida
        self.url_base = url_base.rstrip("/")
        # El formulario tiene un radio name="opcion": nacional | extranjero.
        self.opcion = "extranjero" if extranjero else "nacional"

    def url_certificado(self, curso_id: str) -> str:
        return self.url_base + PLANTILLA_CERT.format(id=curso_id, rut=self.rut)


def cargar_config(ruta_ini="config.ini", pedir_interactivo=False,
                  usuario=None, contrasena=None, rut=None,
                  carpeta=None, url_base=None, extranjero=False) -> Config:
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
    if not extranjero and p.exists():
        cp = configparser.ConfigParser()
        cp.read(p, encoding="utf-8")
        extranjero = cp.getboolean("credenciales", "extranjero", fallback=False)
    return Config(usuario, contrasena, rut, carpeta, url_base, extranjero=extranjero)


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


def _campos_formulario(form, cfg: Config) -> dict:
    """Arma los datos del POST desde el formulario real.

    Clave: para radios/checkbox se incluye SOLO el que esta marcado (checked).
    El formulario tiene dos radios name="opcion" (nacional/extranjero); sin este
    filtro se enviaba el ultimo (extranjero) por error. Luego se fija 'opcion'
    segun la eleccion del usuario.
    """
    datos = {}
    for inp in form.find_all("input"):
        nombre = inp.get("name")
        if not nombre:
            continue
        tipo = (inp.get("type") or "text").lower()
        if tipo in ("radio", "checkbox") and inp.get("checked") is None:
            continue  # ignorar los no marcados
        datos[nombre] = inp.get("value", "")
    datos["username"] = cfg.usuario
    datos["password"] = cfg.contrasena
    datos["remember"] = "yes"
    datos["opcion"] = cfg.opcion
    return datos


def _intentar_login(session: requests.Session, cfg: Config):
    """Un intento de login. Devuelve (exito, html_login, html_respuesta)."""
    url_login = cfg.url_base + RUTA_LOGIN
    r = session.get(url_login, timeout=TIMEOUT)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    form = _buscar_formulario_login(soup)
    if form is None:
        raise RuntimeError("No se encontro el formulario de login. Revisa la URL base "
                           "o usa Diagnostico para guardar el HTML.")
    action = urljoin(url_login, form.get("action") or url_login)
    datos = _campos_formulario(form, cfg)
    hay_token = any(re.fullmatch(r"[0-9a-f]{32}", k) for k in datos)
    log.debug("Login action=%s | campos=%s | token=%s | opcion=%s", action,
              ",".join(sorted(k for k in datos if k != "password")), hay_token, cfg.opcion)
    resp = session.post(action, data=datos, timeout=TIMEOUT, allow_redirects=True,
                        headers={"Referer": url_login})
    resp.raise_for_status()
    verif = session.get(cfg.url_base + RUTA_PROGRESO, timeout=TIMEOUT)
    return (not _parece_pagina_login(verif.text)), r.text, resp.text


def login(session: requests.Session, cfg: Config, intentos: int = 4,
          diag_dir=None) -> None:
    """Login con reintentos. Si falla, guarda el detalle para diagnostico."""
    log.info("Abriendo pagina de login (opcion: %s)...", cfg.opcion)
    ultimo = None
    for intento in range(1, intentos + 1):
        try:
            exito, html_login, html_resp = _intentar_login(session, cfg)
            ultimo = (html_login, html_resp)
            if exito:
                log.info("Sesion iniciada correctamente.")
                return
            log.warning("La plataforma no acepto el acceso (intento %d/%d). "
                        "Reintentando...", intento, intentos)
        except requests.RequestException as e:
            log.warning("Problema de red en el login (intento %d/%d): %s",
                        intento, intentos, e)
        session.cookies.clear()
        time.sleep(min(2 * intento, 6))

    if diag_dir and ultimo:
        try:
            d = Path(diag_dir)
            d.mkdir(parents=True, exist_ok=True)
            (d / "login_pagina.html").write_text(ultimo[0], encoding="utf-8")
            (d / "login_respuesta.html").write_text(ultimo[1], encoding="utf-8")
            log.info("Guarde el detalle del login en %s para diagnostico.", d)
        except Exception as e:
            log.warning("No pude guardar el diagnostico de login: %s", e)
    raise RuntimeError(
        "No se pudo iniciar sesion tras %d intentos. Verifica el RUT y la contrasena; "
        "si tu documento es extranjero, marca EXTRANJERO. Se guardo el detalle del "
        "intento en la carpeta de salida (diagnostico) para revisar." % intentos)


def _parece_pagina_login(html: str) -> bool:
    if not html:
        return True
    return (('name="password"' in html or "name='password'" in html) and
            ('name="username"' in html or "name='username'" in html))


# --------------------------------------------------------------------------- #
# Extraccion de cursos / certificados
# --------------------------------------------------------------------------- #

class Curso:
    def __init__(self, curso_id, rut, nombre, categoria, subcategoria, estado,
                 nota, avance, vigencia, fecha_emision, fecha_vencimiento,
                 cert_url, firma_url="", vigencia_fija=False):
        self.curso_id = str(curso_id)
        self.rut = rut
        self.nombre = nombre
        self.categoria = categoria          # curso_cat real de la plataforma
        self.subcategoria = subcategoria    # prefijo (EDC, OM, MA, ...)
        self.estado = (estado or "").upper()  # APROBADO / PENDIENTE / REPROBADO
        self.nota = nota
        self.avance = avance
        self.vigencia = vigencia
        self.fecha_emision = fecha_emision
        self.fecha_vencimiento = fecha_vencimiento
        self.cert_url = cert_url
        self.firma_url = firma_url
        self.vigencia_fija = vigencia_fija
        self.archivo = None

    @property
    def aprobado(self):
        return self.estado == "APROBADO"

    def dias_para_vencer(self):
        if not self.fecha_vencimiento:
            return None
        return (self.fecha_vencimiento - dt.date.today()).days

    def situacion(self):
        if self.vigencia_fija:
            return "VIGENCIA_FIJA"
        if self.estado and self.estado != "APROBADO":
            return self.estado  # PENDIENTE / REPROBADO / VENCIDO
        d = self.dias_para_vencer()
        if d is None:
            return "VIGENTE" if self.estado == "APROBADO" else "SIN_FECHA"
        if d < 0:
            return "VENCIDO"
        if d <= 60:
            return "POR_VENCER"
        return "VIGENTE"

    def carpeta_relativa(self):
        return (Path(limpiar_nombre(self.categoria, 50)) /
                limpiar_nombre(self.subcategoria, 40))

    def nombre_archivo(self):
        return f"{limpiar_nombre(self.nombre, 90)}_id{self.curso_id}.pdf"


def obtener_cursos(session, cfg: Config, diag_dir=None):
    """Obtiene los cursos desde la API JSON interna de la plataforma.

    La pagina /progreso-cursos no trae los cursos en el HTML: los carga por
    AJAX desde el metodo 'consultorcursosSegundaVersion'. Consumimos ese JSON
    directamente, que trae id, categoria, nombre, estado, nota, vigencia,
    fecha_nota (emision) y certificado_directo.
    """
    url = cfg.url_base + RUTA_JSON.format(rut=cfg.rut)
    log.info("Consultando cursos (API interna)...")
    r = session.get(url, timeout=TIMEOUT,
                    headers={"X-Requested-With": "XMLHttpRequest",
                             "Referer": cfg.url_base + RUTA_PROGRESO})
    r.raise_for_status()
    if diag_dir:
        try:
            Path(diag_dir).mkdir(parents=True, exist_ok=True)
            (Path(diag_dir) / "cursos.json").write_text(r.text, encoding="utf-8")
        except Exception:
            pass
    try:
        data = _desempaquetar_json(r.text)
    except Exception as e:
        raise RuntimeError("No se pudieron leer los cursos (%s). Se guardo la "
                           "respuesta en diagnostico/cursos.json para revisar." % e)

    cursos = []
    for a in data:
        if not isinstance(a, dict):
            continue
        if str(a.get("s")) == "1" or str(a.get("d")) == "1":
            continue  # marcado como eliminado
        cat_id = _num_inicial(str(a.get("cat_id", "0")))
        if cat_id in EXCLUIR_CAT:
            continue
        curso_id = str(a.get("curso_id", "")).strip()
        try:
            if int(curso_id) in SKIP_CURSO:
                continue
        except ValueError:
            pass

        nombre = (a.get("curso_nombre") or "").strip() or ("Curso " + curso_id)
        categoria = (a.get("curso_cat") or "").strip() or "Sin categoria"
        subcat = subcategoria_de(nombre)
        estado = (a.get("estado_curso") or "").strip().upper()
        vigencia = (a.get("vigencia") or "").strip()
        fecha_emision = _parse_fecha_iso(a.get("fecha_nota"))
        fecha_venc, fija = _calcular_vencimiento(curso_id, vigencia, fecha_emision)

        cert = (a.get("certificado_directo") or "").strip().replace("&amp;", "&")
        if cert:
            cert = urljoin(cfg.url_base + "/", cert)
        elif curso_id:
            cert = cfg.url_certificado(curso_id)
        firma = str(a.get("firma_digital", "0")).strip()
        firma_url = (cfg.url_base + RUTA_FIRMA.format(id=firma)
                     if firma and firma != "0" else "")

        cursos.append(Curso(
            curso_id, cfg.rut, nombre, categoria, subcat, estado,
            a.get("notobt", ""), a.get("porcentaje_avance", ""), vigencia,
            fecha_emision, fecha_venc, cert, firma_url, fija))

    cursos.sort(key=lambda c: (c.categoria.lower(), c.subcategoria, c.nombre.lower()))
    log.info("Se obtuvieron %d cursos (%d aprobados con certificado).",
             len(cursos), sum(1 for c in cursos if c.aprobado))
    if not cursos:
        log.warning("La API no devolvio cursos. Revisa diagnostico/cursos.json.")
    return cursos


def obtener_nombre(session, cfg: Config) -> str:
    """Nombre del trabajador logeado (del saludo de /progreso-cursos)."""
    try:
        r = session.get(cfg.url_base + RUTA_PROGRESO, timeout=TIMEOUT)
        soup = BeautifulSoup(r.text, "html.parser")
        saludo = soup.select_one(".saludo")
        if saludo:
            for d in saludo.select(".color-1"):
                t = d.get_text(strip=True)
                if t and "hola" not in _sin_acentos(t).lower():
                    return t
    except Exception:
        pass
    return cfg.rut


# --------------------------------------------------------------------------- #
# Descarga
# --------------------------------------------------------------------------- #

def descargar_certificados(session, cfg, cursos, carpeta_base: Path, progreso_cb=None):
    carpeta_base.mkdir(parents=True, exist_ok=True)
    ok = 0
    total = len(cursos)
    for i, curso in enumerate(cursos, 1):
        if not curso.aprobado:
            if progreso_cb:
                progreso_cb(i, total)
            continue  # sin certificado emitido (pendiente/reprobado)
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

def _situacion_snapshot(c, hoy, umbral):
    """Situacion calculada en Python (foto del momento) para KPIs/logs."""
    if not c.aprobado:
        return (c.estado or "PENDIENTE")
    if c.fecha_vencimiento is None:   # Indefinida o vigencia fija
        return "INDEFINIDO"
    d = (c.fecha_vencimiento - hoy).days
    if d < 0:
        return "VENCIDO"
    if d <= umbral:
        return "POR VENCER"
    return "VIGENTE"


def generar_excel(cursos, ruta_xlsx: Path, nombre="", rut="", umbral=60):
    """Genera el Excel con dos hojas:

    - "Panel": datos del trabajador, parametro editable (dias de aviso) y KPIs
      que se recalculan solos al abrir (COUNTIF sobre la hoja de detalle).
    - "Cursos": tabla por categoria. 'Dias restantes' = Expiracion - HOY() (lee
      el reloj del PC) y 'Situacion' se colorea sola segun el umbral del Panel.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import FormulaRule

    TEAL = "0B6E7A"
    COLORES = {"VIGENTE": "D9EAD3", "POR VENCER": "FCE8B2", "VENCIDO": "F4CCCC",
               "INDEFINIDO": "E8EEF0", "PENDIENTE": "EDEDED", "REPROBADO": "F4CCCC"}
    cab_fill = PatternFill("solid", fgColor=TEAL)
    cab_font = Font(bold=True, color="FFFFFF")
    link_font = Font(color="0563C1", underline="single")
    borde = Border(*(Side(style="thin", color="D0D7DA"),) * 4)

    wb = Workbook()
    panel = wb.active
    panel.title = "Panel"
    ws = wb.create_sheet("Cursos")

    # -------------------- Hoja "Cursos" (detalle) --------------------
    enc = ["Categoria", "Subcategoria", "Curso", "Fecha de nota", "Expiracion",
           "Dias restantes", "Situacion", "Nota %", "Avance %", "Vigencia",
           "Estado", "Archivo PDF", "URL certificado"]
    ws.append(enc)
    for col in range(1, len(enc) + 1):
        c = ws.cell(row=1, column=col)
        c.fill = cab_fill
        c.font = cab_font
        c.alignment = Alignment(vertical="center", horizontal="center")

    datos = sorted(cursos, key=lambda c: (c.categoria.lower(), c.subcategoria,
                                          c.nombre.lower()))
    UMBRAL = "Panel!$C$8"   # celda con el numero de dias de aviso (editable)
    hoy = dt.date.today()
    snap = {}
    r = 1
    for c in datos:
        r += 1
        tiene_exp = c.aprobado and (c.fecha_vencimiento is not None)
        ws.cell(r, 1, c.categoria)
        ws.cell(r, 2, c.subcategoria)
        ws.cell(r, 3, c.nombre)
        if c.fecha_emision:
            cell = ws.cell(r, 4, c.fecha_emision)
            cell.number_format = "DD-MM-YYYY"
        if tiene_exp:
            e = ws.cell(r, 5, c.fecha_vencimiento)
            e.number_format = "DD-MM-YYYY"
            d = ws.cell(r, 6)
            d.value = f'=IF(E{r}="","",E{r}-TODAY())'
            d.number_format = "0"
            d.alignment = Alignment(horizontal="center")
            ws.cell(r, 7).value = (
                f'=IF(E{r}="","INDEFINIDO",IF(F{r}<0,"VENCIDO",'
                f'IF(F{r}<={UMBRAL},"POR VENCER","VIGENTE")))')
        else:
            ws.cell(r, 7, c.estado if not c.aprobado else "INDEFINIDO")
        ws.cell(r, 8, c.nota)
        ws.cell(r, 9, c.avance)
        ws.cell(r, 10, c.vigencia)
        ws.cell(r, 11, c.estado)
        # Archivo PDF local -> hipervinculo para abrirlo
        cl = ws.cell(r, 12, c.archivo or "")
        if c.archivo:
            cl.hyperlink = c.archivo
            cl.font = link_font
        # URL del certificado -> hipervinculo clickable
        cm = ws.cell(r, 13, "Descargar certificado" if c.aprobado else "")
        if c.aprobado and c.cert_url:
            cm.hyperlink = c.cert_url
            cm.font = link_font
        for col in range(1, len(enc) + 1):
            ws.cell(r, col).border = borde
        s = _situacion_snapshot(c, hoy, umbral)
        snap[s] = snap.get(s, 0) + 1
    nfilas = r

    anchos = [34, 13, 46, 13, 12, 13, 13, 7, 8, 13, 12, 36, 44]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(enc))}{nfilas}"

    # Colores dinamicos por situacion (columna G) sobre toda la fila.
    rango = f"A2:{get_column_letter(len(enc))}{nfilas}"
    for etiqueta, color in COLORES.items():
        ws.conditional_formatting.add(rango, FormulaRule(
            formula=[f'$G2="{etiqueta}"'], stopIfTrue=False,
            fill=PatternFill("solid", fgColor=color)))

    # Refuerzo: pinta la casilla "Dias restantes" (F) por su propio numero
    # (rojo si vencio, amarillo si quedan <= umbral, verde si quedan mas).
    fcol = f"F2:F{nfilas}"
    ws.conditional_formatting.add(fcol, FormulaRule(
        formula=['AND($F2<>"",$F2<0)'], stopIfTrue=True,
        fill=PatternFill("solid", fgColor="F4CCCC"),
        font=Font(bold=True, color="9C2A2A")))
    ws.conditional_formatting.add(fcol, FormulaRule(
        formula=[f'AND($F2<>"",$F2<={UMBRAL})'], stopIfTrue=True,
        fill=PatternFill("solid", fgColor="FCE8B2"),
        font=Font(bold=True, color="8A6D1A")))
    ws.conditional_formatting.add(fcol, FormulaRule(
        formula=[f'AND($F2<>"",$F2>{UMBRAL})'], stopIfTrue=True,
        fill=PatternFill("solid", fgColor="D9EAD3"),
        font=Font(bold=True, color="2E6B23")))

    # -------------------- Hoja "Panel" --------------------
    panel.sheet_view.showGridLines = False
    panel.column_dimensions["A"].width = 3
    panel.column_dimensions["B"].width = 30
    panel.column_dimensions["C"].width = 16

    panel.merge_cells("B2:C2")
    t = panel["B2"]
    t.value = "CAMPUS CENTINELA"
    t.font = Font(bold=True, size=18, color="FFFFFF")
    t.alignment = Alignment(vertical="center", horizontal="center")
    panel["B2"].fill = cab_fill
    panel["C2"].fill = cab_fill
    panel.merge_cells("B3:C3")
    panel["B3"] = "Control de certificaciones"
    panel["B3"].font = Font(italic=True, color="FFFFFF")
    panel["B3"].alignment = Alignment(horizontal="center")
    panel["B3"].fill = PatternFill("solid", fgColor="095761")
    panel["C3"].fill = PatternFill("solid", fgColor="095761")
    panel.row_dimensions[2].height = 28

    def etiqueta(celda, texto):
        panel[celda] = texto
        panel[celda].font = Font(bold=True)

    etiqueta("B5", "Trabajador"); panel["C5"] = nombre or "-"
    etiqueta("B6", "RUT"); panel["C6"] = rut or "-"
    etiqueta("B7", "Generado"); panel["C7"] = dt.datetime.now().strftime("%d-%m-%Y %H:%M")

    etiqueta("B8", "Dias de aviso (por vencer)")
    pbox = panel["C8"]
    pbox.value = int(umbral)
    pbox.font = Font(bold=True, color="0B6E7A")
    pbox.fill = PatternFill("solid", fgColor="FFF4CC")
    pbox.alignment = Alignment(horizontal="center")
    panel["B9"] = "(edita este numero y los colores se recalculan)"
    panel["B9"].font = Font(italic=True, size=9, color="888888")

    # KPIs dinamicos
    kpis = [
        ("Vigentes", "VIGENTE", "D9EAD3"),
        ("Por vencer", "POR VENCER", "FCE8B2"),
        ("Vencidos", "VENCIDO", "F4CCCC"),
        ("Indefinidos / fija", "INDEFINIDO", "E8EEF0"),
        ("Pendientes", "PENDIENTE", "EDEDED"),
    ]
    panel["B11"] = "RESUMEN (se actualiza al abrir)"
    panel["B11"].font = Font(bold=True, size=12, color="0B6E7A")
    fila = 12
    rango_sit = f"Cursos!$G$2:$G${nfilas}"
    for texto, clave, color in kpis:
        panel[f"B{fila}"] = texto
        panel[f"B{fila}"].fill = PatternFill("solid", fgColor=color)
        panel[f"B{fila}"].border = borde
        cc = panel[f"C{fila}"]
        cc.value = f'=COUNTIF({rango_sit},"{clave}")'
        cc.font = Font(bold=True, size=12)
        cc.alignment = Alignment(horizontal="center")
        cc.fill = PatternFill("solid", fgColor=color)
        cc.border = borde
        fila += 1
    panel[f"B{fila}"] = "Total cursos"
    panel[f"B{fila}"].font = Font(bold=True)
    panel[f"C{fila}"].value = f"=COUNTA(Cursos!$C$2:$C${nfilas})"
    panel[f"C{fila}"].font = Font(bold=True)
    panel[f"C{fila}"].alignment = Alignment(horizontal="center")

    # Totales por categoria (dinamicos)
    fila += 2
    panel[f"B{fila}"] = "CURSOS POR CATEGORIA"
    panel[f"B{fila}"].font = Font(bold=True, size=12, color="0B6E7A")
    fila += 1
    for cat in sorted({c.categoria for c in cursos}):
        panel[f"B{fila}"] = cat
        panel[f"B{fila}"].border = borde
        cell = panel[f"C{fila}"]
        cat_esc = cat.replace('"', '""')
        cell.value = f'=COUNTIF(Cursos!$A$2:$A${nfilas},"{cat_esc}")'
        cell.alignment = Alignment(horizontal="center")
        cell.border = borde
        fila += 1

    wb.active = wb.sheetnames.index("Panel")

    ruta_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta_xlsx)
    log.info("Excel generado: %s", ruta_xlsx)
    log.info("Vigentes: %d | Por vencer: %d | Vencidos: %d | Indefinidos: %d",
             snap.get("VIGENTE", 0), snap.get("POR VENCER", 0),
             snap.get("VENCIDO", 0), snap.get("INDEFINIDO", 0))
    return {"VIGENTE": snap.get("VIGENTE", 0), "POR_VENCER": snap.get("POR VENCER", 0),
            "VENCIDO": snap.get("VENCIDO", 0), "INDEFINIDO": snap.get("INDEFINIDO", 0),
            "PENDIENTE": snap.get("PENDIENTE", 0)}


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
    base = Path(cfg.carpeta_salida)
    try:
        base.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        raise RuntimeError(
            "No se pudo crear la carpeta de salida '%s' (%s). Elige una carpeta "
            "con permisos de escritura, por ejemplo en Documentos." % (base, e))
    login(session, cfg, diag_dir=base / "diagnostico")
    if dump:
        guardar_diagnostico(session, cfg, base / "diagnostico")
    cursos = obtener_cursos(session, cfg, diag_dir=base / "diagnostico")
    if not cursos:
        return [], None, {}
    nombre = obtener_nombre(session, cfg)
    if not sin_descarga:
        descargar_certificados(session, cfg, cursos, base / "certificados", progreso_cb)
    ruta_xlsx = base / "certificados_centinela.xlsx"
    conteo = generar_excel(cursos, ruta_xlsx, nombre=nombre, rut=cfg.rut)
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
    parser.add_argument("--extranjero", action="store_true",
                        help="Usar la opcion EXTRANJERO en el login (por defecto NACIONAL)")
    parser.add_argument("--salida", default=None)
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-7s %(message)s", datefmt="%H:%M:%S")

    try:
        cfg = cargar_config(args.config, args.interactivo, carpeta=args.salida,
                            extranjero=args.extranjero)
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
