#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Campus Centinela - Descarga y seguimiento de certificados
==========================================================

Inicia sesion en https://www.campuscentinela.cl (plataforma Joomla),
lee la pagina de progreso de cursos, descarga los certificados en PDF,
los ordena por categoria y genera un Excel con las fechas de emision y
vencimiento para controlar cursos pendientes o vencidos.

NO hay API publica en la plataforma: la autenticacion es por formulario
con token CSRF de Joomla y las paginas se renderizan en el servidor, por
lo que se trabaja con una sesion (cookie) como lo haria un navegador.

Uso basico:
    1. pip install -r requirements.txt
    2. copiar config.example.ini a config.ini y completar credenciales
    3. python campus_centinela.py --dump        (primera vez: diagnostico)
    4. python campus_centinela.py                (descarga + Excel)

Ver README.md para el detalle.
"""

import argparse
import configparser
import datetime as dt
import getpass
import logging
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("Faltan dependencias. Ejecuta:  pip install -r requirements.txt")
    sys.exit(1)


# --------------------------------------------------------------------------- #
# Configuracion general
# --------------------------------------------------------------------------- #

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
TIMEOUT = 40  # segundos por peticion
RUTA_LOGIN = "/ingresar"
RUTA_PORTADA = "/portada"
RUTA_PROGRESO = "/progreso-cursos"

# Patron de la URL de certificado: certificado.php?id=NNN&rut=XXXX
RE_CERT = re.compile(r"certificado\.php\?[^\"'\s<>]*", re.IGNORECASE)
# Fechas chilenas: dd-mm-yyyy  o  dd/mm/yyyy
RE_FECHA = re.compile(r"\b(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})\b")

log = logging.getLogger("campus")


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #

def configurar_logging(verboso: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verboso else logging.INFO,
        format="%(asctime)s  %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
    )


def limpiar_nombre(texto: str, maximo: int = 120) -> str:
    """Convierte un texto en un nombre de archivo/carpeta seguro."""
    texto = (texto or "").strip()
    texto = re.sub(r"\s+", " ", texto)
    texto = re.sub(r'[\\/:*?"<>|]+', "-", texto)
    texto = texto.strip(" .-")
    if not texto:
        texto = "sin_nombre"
    return texto[:maximo]


def parsear_fecha(texto: str):
    """Devuelve un datetime.date desde el primer patron de fecha encontrado."""
    m = RE_FECHA.search(texto or "")
    if not m:
        return None
    d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if a < 100:  # anio de 2 digitos
        a += 2000
    try:
        return dt.date(a, mes, d)
    except ValueError:
        return None


def todas_las_fechas(texto: str):
    """Lista de fechas (date) encontradas en el texto, en orden de aparicion."""
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
# Carga de configuracion
# --------------------------------------------------------------------------- #

class Config:
    def __init__(self, usuario, contrasena, rut, carpeta_salida, url_base):
        self.usuario = usuario
        self.contrasena = contrasena
        self.rut = rut or usuario
        self.carpeta_salida = carpeta_salida
        self.url_base = url_base.rstrip("/")


def cargar_config(ruta_ini: str, pedir_interactivo: bool) -> Config:
    """Carga credenciales desde config.ini, variables de entorno o teclado.

    Prioridad: variables de entorno > config.ini > pregunta interactiva.
    Variables: CC_USUARIO, CC_CONTRASENA, CC_RUT.
    """
    usuario = os.environ.get("CC_USUARIO", "")
    contrasena = os.environ.get("CC_CONTRASENA", "")
    rut = os.environ.get("CC_RUT", "")
    carpeta = "salida"
    url_base = "https://www.campuscentinela.cl"

    p = Path(ruta_ini)
    if p.exists():
        cp = configparser.ConfigParser()
        cp.read(p, encoding="utf-8")
        usuario = usuario or cp.get("credenciales", "usuario", fallback="").strip()
        contrasena = contrasena or cp.get("credenciales", "contrasena", fallback="").strip()
        rut = rut or cp.get("credenciales", "rut", fallback="").strip()
        carpeta = cp.get("opciones", "carpeta_salida", fallback=carpeta).strip() or carpeta
        url_base = cp.get("opciones", "url_base", fallback=url_base).strip() or url_base
    elif not (usuario and contrasena):
        log.warning("No existe %s. Usa config.example.ini como plantilla.", ruta_ini)

    if pedir_interactivo:
        if not usuario:
            usuario = input("RUT / usuario: ").strip()
        if not contrasena:
            contrasena = getpass.getpass("Contrasena: ")

    if not usuario or not contrasena:
        raise SystemExit(
            "Faltan credenciales. Completa config.ini, define las variables "
            "CC_USUARIO/CC_CONTRASENA, o ejecuta con --interactivo."
        )
    if contrasena.upper() == "CAMBIAME":
        raise SystemExit("La contrasena en config.ini sigue en 'CAMBIAME'. Editala.")

    return Config(usuario, contrasena, rut, carpeta, url_base)


# --------------------------------------------------------------------------- #
# Login (Joomla com_users)
# --------------------------------------------------------------------------- #

def crear_sesion() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "es-CL,es;q=0.9"})
    return s


def _buscar_formulario_login(soup: BeautifulSoup):
    """Encuentra el <form> que contiene los campos username/password."""
    campo = soup.find("input", {"name": "username"}) or soup.find(
        "input", {"name": "password"}
    )
    if campo:
        form = campo.find_parent("form")
        if form:
            return form
    # Fallback: primer form con un input type=password
    for form in soup.find_all("form"):
        if form.find("input", {"type": "password"}):
            return form
    return None


def login(session: requests.Session, cfg: Config) -> None:
    """Inicia sesion reenviando TODOS los campos ocultos del formulario real.

    Esto cubre el token CSRF dinamico de Joomla y los campos option/task/return
    sin tener que adivinarlos: se leen directamente del HTML del login.
    """
    url_login = cfg.url_base + RUTA_LOGIN
    log.info("Abriendo pagina de login...")
    r = session.get(url_login, timeout=TIMEOUT)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    form = _buscar_formulario_login(soup)
    if form is None:
        raise RuntimeError(
            "No se encontro el formulario de login. Revisa la URL base o "
            "ejecuta con --dump para guardar el HTML y diagnosticar."
        )

    action = urljoin(url_login, form.get("action") or url_login)
    datos = {}
    for inp in form.find_all(["input", "button"]):
        nombre = inp.get("name")
        if not nombre:
            continue
        datos[nombre] = inp.get("value", "")

    # Completar credenciales
    datos["username"] = cfg.usuario
    datos["password"] = cfg.contrasena
    datos["remember"] = "yes"

    campos_log = ", ".join(sorted(k for k in datos if k != "password"))
    log.debug("POST login a %s | campos: %s", action, campos_log)

    resp = session.post(action, data=datos, timeout=TIMEOUT, allow_redirects=True)
    resp.raise_for_status()

    # Verificacion: pedir /progreso-cursos y confirmar que NO nos devuelve el login
    verif = session.get(cfg.url_base + RUTA_PROGRESO, timeout=TIMEOUT)
    if _parece_pagina_login(verif.text):
        raise RuntimeError(
            "El login no fue exitoso (la plataforma devolvio el formulario de "
            "acceso). Revisa usuario/contrasena. Si son correctos, ejecuta con "
            "--dump para inspeccionar el HTML."
        )
    log.info("Sesion iniciada correctamente.")


def _parece_pagina_login(html: str) -> bool:
    """True si el HTML parece la pagina de login (no autenticado)."""
    if not html:
        return True
    tiene_pass = 'name="password"' in html or "name='password'" in html
    tiene_user = 'name="username"' in html or "name='username'" in html
    return tiene_pass and tiene_user


# --------------------------------------------------------------------------- #
# Lectura de la pagina de progreso de cursos
# --------------------------------------------------------------------------- #

class Certificado:
    def __init__(self, cert_id, rut, nombre, categoria, estado,
                 fecha_emision, fecha_vencimiento, url, texto_crudo):
        self.cert_id = cert_id
        self.rut = rut
        self.nombre = nombre
        self.categoria = categoria
        self.estado = estado
        self.fecha_emision = fecha_emision
        self.fecha_vencimiento = fecha_vencimiento
        self.url = url
        self.texto_crudo = texto_crudo
        self.archivo = None  # ruta al PDF una vez descargado

    def dias_para_vencer(self):
        if not self.fecha_vencimiento:
            return None
        return (self.fecha_vencimiento - dt.date.today()).days

    def situacion(self):
        """Clasificacion util: VENCIDO / POR_VENCER / VIGENTE / SIN_FECHA."""
        d = self.dias_para_vencer()
        if d is None:
            return "SIN_FECHA"
        if d < 0:
            return "VENCIDO"
        if d <= 60:
            return "POR_VENCER"
        return "VIGENTE"


def _categoria_desde_contexto(nodo) -> str:
    """Busca la categoria antes del nodo: el encabezado de seccion (h1..h5) o
    una fila de grupo (un <tr> con una sola celda), lo que este mas cerca.

    NO se usan los <th> de cabecera porque esos son titulos de columna
    ("Curso", "Vence", ...), no categorias.
    """
    for prev in nodo.find_all_previous(["h1", "h2", "h3", "h4", "h5", "tr"]):
        if prev.name == "tr":
            celdas = prev.find_all(["td", "th"], recursive=False)
            # Fila de grupo = una sola celda con texto (separador de categoria).
            if len(celdas) == 1:
                texto = celdas[0].get_text(" ", strip=True)
                if texto and len(texto) <= 80:
                    return texto
            # Fila normal de datos (varias celdas): se ignora y se sigue atras.
            continue
        texto = prev.get_text(" ", strip=True)
        if texto and len(texto) <= 80:
            return texto
    return "Sin categoria"


def _fila_contenedora(enlace):
    """Devuelve la fila/tarjeta que contiene al enlace del certificado."""
    tr = enlace.find_parent("tr")
    if tr:
        return tr
    # tarjetas: subir hasta un div con varias lineas de texto
    nodo = enlace
    for _ in range(6):
        padre = nodo.find_parent(["div", "li", "article"])
        if padre is None:
            break
        if len(padre.get_text(" ", strip=True)) > 25:
            return padre
        nodo = padre
    return enlace.parent or enlace


def extraer_certificados(html: str, cfg: Config):
    """Extrae la lista de certificados anclando en los enlaces certificado.php.

    Estrategia robusta: la pagina de progreso DEBE contener los enlaces de
    descarga (es donde el usuario los baja). Por cada enlace se toma su fila
    contenedora y de ahi el nombre del curso, la categoria y las fechas.
    """
    soup = BeautifulSoup(html, "html.parser")
    certificados = []
    vistos = set()

    # 1) Enlaces <a>, <button>, o cualquier atributo que apunte a certificado.php
    candidatos = []
    for tag in soup.find_all(True):
        for attr in ("href", "onclick", "data-href", "data-url", "formaction"):
            val = tag.get(attr)
            if val and "certificado.php" in val.lower():
                m = RE_CERT.search(val)
                if m:
                    candidatos.append((tag, m.group(0)))
                    break

    for tag, trozo_url in candidatos:
        url = urljoin(cfg.url_base + RUTA_PROGRESO, trozo_url.replace("&amp;", "&"))
        qs = parse_qs(urlparse(url).query)
        cert_id = (qs.get("id", [""])[0]).strip()
        rut = (qs.get("rut", [cfg.rut])[0]).strip() or cfg.rut
        clave = (cert_id, rut)
        if not cert_id or clave in vistos:
            continue
        vistos.add(clave)

        fila = _fila_contenedora(tag)
        texto = fila.get_text(" | ", strip=True)

        # Nombre del curso: celda mas larga de la fila, o el texto del enlace
        nombre = ""
        celdas = fila.find_all(["td", "th"])
        if celdas:
            textos = [c.get_text(" ", strip=True) for c in celdas]
            textos = [t for t in textos if t and not RE_FECHA.fullmatch(t)]
            if textos:
                nombre = max(textos, key=len)
        if not nombre:
            nombre = tag.get_text(" ", strip=True) or f"Certificado {cert_id}"

        categoria = _categoria_desde_contexto(fila)
        fechas = todas_las_fechas(texto)
        fecha_emision = fechas[0] if len(fechas) >= 1 else None
        fecha_vencimiento = fechas[1] if len(fechas) >= 2 else None

        estado = ""
        for palabra in ("vencido", "vigente", "por vencer", "aprobado", "pendiente"):
            if palabra in texto.lower():
                estado = palabra
                break

        certificados.append(Certificado(
            cert_id=cert_id, rut=rut, nombre=nombre, categoria=categoria,
            estado=estado, fecha_emision=fecha_emision,
            fecha_vencimiento=fecha_vencimiento, url=url, texto_crudo=texto,
        ))

    log.info("Se encontraron %d certificados en la pagina de progreso.",
             len(certificados))
    if not certificados:
        log.warning(
            "No se detectaron enlaces 'certificado.php'. Puede que el HTML sea "
            "distinto al esperado o que se carguen por JavaScript. Ejecuta con "
            "--dump y revisa diagnostico/progreso-cursos.html."
        )
    return certificados


# --------------------------------------------------------------------------- #
# Descarga de certificados
# --------------------------------------------------------------------------- #

def descargar_certificados(session, cfg, certificados, carpeta_base: Path):
    carpeta_base.mkdir(parents=True, exist_ok=True)
    ok = 0
    for i, cert in enumerate(certificados, 1):
        carpeta_cat = carpeta_base / limpiar_nombre(cert.categoria, 60)
        carpeta_cat.mkdir(parents=True, exist_ok=True)
        nombre_pdf = f"{limpiar_nombre(cert.nombre, 90)}_id{cert.cert_id}.pdf"
        destino = carpeta_cat / nombre_pdf
        try:
            r = session.get(cert.url, timeout=TIMEOUT, stream=True)
            r.raise_for_status()
            contenido = r.content
            ctype = r.headers.get("Content-Type", "")
            if b"%PDF" not in contenido[:1024] and "pdf" not in ctype.lower():
                log.warning("  [%d/%d] id=%s no devolvio un PDF (Content-Type: %s). "
                            "Se omite.", i, len(certificados), cert.cert_id, ctype)
                continue
            destino.write_bytes(contenido)
            cert.archivo = str(destino)
            ok += 1
            log.info("  [%d/%d] OK  %s", i, len(certificados), destino.name)
        except requests.RequestException as e:
            log.error("  [%d/%d] Error id=%s: %s", i, len(certificados),
                      cert.cert_id, e)
    log.info("Descargados %d de %d certificados.", ok, len(certificados))
    return ok


# --------------------------------------------------------------------------- #
# Reporte Excel
# --------------------------------------------------------------------------- #

def generar_excel(certificados, ruta_xlsx: Path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Certificados"

    encabezados = [
        "Categoria", "Curso", "Estado plataforma", "Situacion",
        "Fecha emision", "Fecha vencimiento", "Dias para vencer",
        "ID", "RUT", "Archivo PDF", "URL",
    ]
    ws.append(encabezados)

    cabecera_fill = PatternFill("solid", fgColor="0B6E7A")
    cabecera_font = Font(bold=True, color="FFFFFF")
    for col, _ in enumerate(encabezados, 1):
        c = ws.cell(row=1, column=col)
        c.fill = cabecera_fill
        c.font = cabecera_font
        c.alignment = Alignment(vertical="center")

    rojo = PatternFill("solid", fgColor="F4CCCC")       # vencido
    amarillo = PatternFill("solid", fgColor="FCE8B2")    # por vencer
    verde = PatternFill("solid", fgColor="D9EAD3")       # vigente

    # Ordenar: primero por categoria, luego por fecha de vencimiento mas cercana
    def clave_orden(c):
        d = c.dias_para_vencer()
        return (c.categoria.lower(), d if d is not None else 10**6)

    for cert in sorted(certificados, key=clave_orden):
        situacion = cert.situacion()
        fila = [
            cert.categoria,
            cert.nombre,
            cert.estado,
            situacion,
            cert.fecha_emision.strftime("%d-%m-%Y") if cert.fecha_emision else "",
            cert.fecha_vencimiento.strftime("%d-%m-%Y") if cert.fecha_vencimiento else "",
            cert.dias_para_vencer() if cert.dias_para_vencer() is not None else "",
            cert.cert_id,
            cert.rut,
            cert.archivo or "",
            cert.url,
        ]
        ws.append(fila)
        fila_idx = ws.max_row
        relleno = {"VENCIDO": rojo, "POR_VENCER": amarillo, "VIGENTE": verde}.get(situacion)
        if relleno:
            for col in range(1, len(encabezados) + 1):
                ws.cell(row=fila_idx, column=col).fill = relleno

    # Ancho de columnas y filtros
    anchos = [26, 50, 16, 13, 14, 16, 15, 7, 13, 40, 50]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(encabezados))}{ws.max_row}"

    # Hoja resumen
    resumen = wb.create_sheet("Resumen")
    conteo = {"VENCIDO": 0, "POR_VENCER": 0, "VIGENTE": 0, "SIN_FECHA": 0}
    for c in certificados:
        conteo[c.situacion()] += 1
    resumen.append(["Generado", dt.datetime.now().strftime("%d-%m-%Y %H:%M")])
    resumen.append(["Total certificados", len(certificados)])
    resumen.append([])
    resumen.append(["Situacion", "Cantidad"])
    for k in ("VENCIDO", "POR_VENCER", "VIGENTE", "SIN_FECHA"):
        resumen.append([k, conteo[k]])
    resumen.column_dimensions["A"].width = 22
    resumen.column_dimensions["B"].width = 22
    resumen["A4"].font = Font(bold=True)
    resumen["B4"].font = Font(bold=True)

    ruta_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta_xlsx)
    log.info("Excel generado: %s", ruta_xlsx)
    log.info("Resumen -> Vencidos: %d | Por vencer (<=60d): %d | Vigentes: %d | Sin fecha: %d",
             conteo["VENCIDO"], conteo["POR_VENCER"], conteo["VIGENTE"], conteo["SIN_FECHA"])


# --------------------------------------------------------------------------- #
# Diagnostico (dump de HTML)
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
    log.info("Revisa la carpeta '%s' y comparte progreso-cursos.html si el "
             "parser no reconocio los certificados.", carpeta)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Descarga certificados de Campus Centinela y genera un Excel "
                    "con fechas de emision y vencimiento."
    )
    parser.add_argument("--config", default="config.ini",
                        help="Ruta al archivo de configuracion (def: config.ini)")
    parser.add_argument("--interactivo", action="store_true",
                        help="Pedir usuario/contrasena por teclado si faltan")
    parser.add_argument("--dump", action="store_true",
                        help="Guardar el HTML de portada y progreso-cursos para diagnostico")
    parser.add_argument("--sin-descarga", action="store_true",
                        help="No descargar PDFs; solo generar el Excel")
    parser.add_argument("--salida", default=None,
                        help="Carpeta de salida (sobrescribe la del config)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Modo detallado")
    args = parser.parse_args(argv)

    configurar_logging(args.verbose)
    cfg = cargar_config(args.config, args.interactivo)
    if args.salida:
        cfg.carpeta_salida = args.salida

    session = crear_sesion()
    try:
        login(session, cfg)
    except Exception as e:
        log.error("Fallo el login: %s", e)
        if args.dump:
            try:
                Path("diagnostico").mkdir(exist_ok=True)
                r = session.get(cfg.url_base + RUTA_LOGIN, timeout=TIMEOUT)
                Path("diagnostico/login.html").write_text(r.text, encoding="utf-8")
                log.info("Guardado diagnostico/login.html")
            except Exception:
                pass
        return 2

    base = Path(cfg.carpeta_salida)

    if args.dump:
        guardar_diagnostico(session, cfg, Path("diagnostico"))

    log.info("Leyendo pagina de progreso de cursos...")
    r = session.get(cfg.url_base + RUTA_PROGRESO, timeout=TIMEOUT)
    r.raise_for_status()
    certificados = extraer_certificados(r.text, cfg)

    if not certificados:
        log.error("No hay certificados que procesar. Ejecuta con --dump y revisa el HTML.")
        return 1

    if not args.sin_descarga:
        log.info("Descargando certificados en '%s'...", base / "certificados")
        descargar_certificados(session, cfg, certificados, base / "certificados")

    ruta_xlsx = base / "certificados_centinela.xlsx"
    generar_excel(certificados, ruta_xlsx)

    log.info("Listo. Resultados en la carpeta '%s'.", base)
    return 0


if __name__ == "__main__":
    sys.exit(main())
