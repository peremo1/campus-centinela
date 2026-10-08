#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Campus Centinela - Interfaz Flet (estilo app)
=============================================

Interfaz moderna (Material, basada en Flet) para descargar los certificados y
generar el Excel. Reutiliza el motor de campus_centinela.py.

Es la base del ejecutable: se empaqueta con  flet pack  (ver workflow / README).

Ejecutar desde fuente:
    pip install -r requirements.txt -r requirements-gui.txt
    python campus_centinela_flet.py
"""

import configparser
import logging
import os
import sys
import threading
from pathlib import Path

import flet as ft

import campus_centinela as cc

TEAL = "#0B6E7A"
TEAL_OSC = "#095761"
FONDO_LOG = "#0f1b1d"
TEXTO_LOG = "#d6e6e8"


def _carpeta_default() -> str:
    docs = Path.home() / "Documents"
    base = docs if docs.exists() else Path.home()
    return str(base / "CampusCentinela")


def _ruta_config() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "config.ini"
    return Path(__file__).resolve().parent / "config.ini"


def main(page: ft.Page):
    page.title = "Campus Centinela - Certificados"
    page.theme_mode = ft.ThemeMode.LIGHT
    try:
        page.theme = ft.Theme(color_scheme_seed=TEAL)
    except Exception:
        pass
    page.padding = 0
    page.bgcolor = "#FFFFFF"
    for ancho, alto in ((900, 720),):
        try:
            page.window.width, page.window.height = ancho, alto
            page.window.min_width, page.window.min_height = 760, 620
        except Exception:
            pass
        try:
            page.window_width, page.window_height = ancho, alto
            page.window_min_width, page.window_min_height = 760, 620
        except Exception:
            pass

    # ---------------- controles ----------------
    tf_user = ft.TextField(label="RUT / usuario", hint_text="21080196-0",
                           prefix_text="", border_radius=10)
    tf_pass = ft.TextField(label="Contrasena", password=True,
                           can_reveal_password=True, border_radius=10)
    rg_opcion = ft.RadioGroup(value="nacional", content=ft.Row(spacing=16, controls=[
        ft.Radio(value="nacional", label="NACIONAL"),
        ft.Radio(value="extranjero", label="EXTRANJERO"),
    ]))
    tf_folder = ft.TextField(label="Carpeta de salida", value=_carpeta_default(),
                             expand=True, border_radius=10)
    sw_guardar = ft.Switch(label="Recordar datos en este equipo", value=True)

    progress = ft.ProgressBar(value=0, color=TEAL, bgcolor="#dfe7e8", border_radius=6)
    status = ft.Text("Listo.", size=12, color="#555")
    log_list = ft.ListView(expand=True, auto_scroll=True, spacing=1, padding=10)

    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)

    # ---------------- logging hacia la lista ----------------
    def log_linea(txt: str):
        log_list.controls.append(
            ft.Text(txt, size=12, selectable=True, color=TEXTO_LOG,
                    font_family="Consolas"))
        try:
            page.update()
        except Exception:
            pass

    class ListHandler(logging.Handler):
        def emit(self, record):
            log_linea(self.format(record))

    handler = ListHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
    logger = logging.getLogger("campus")
    logger.setLevel(logging.INFO)
    logger.handlers = [handler]

    estado = {"trabajando": False}

    # ---------------- config ----------------
    def cargar_config():
        ruta = _ruta_config()
        if not ruta.exists():
            return
        cp = configparser.ConfigParser()
        try:
            cp.read(ruta, encoding="utf-8")
            tf_user.value = cp.get("credenciales", "usuario", fallback="")
            tf_pass.value = cp.get("credenciales", "contrasena", fallback="")
            tf_folder.value = cp.get("opciones", "carpeta_salida",
                                     fallback=_carpeta_default())
            log_linea("Datos cargados desde config.ini")
        except Exception as e:
            log_linea("No se pudo leer config.ini: %s" % e)

    def guardar_config(cfg):
        cp = configparser.ConfigParser()
        cp["credenciales"] = {"usuario": cfg.usuario, "contrasena": cfg.contrasena}
        cp["opciones"] = {"carpeta_salida": cfg.carpeta_salida, "url_base": cfg.url_base}
        try:
            with open(_ruta_config(), "w", encoding="utf-8") as f:
                cp.write(f)
        except Exception as e:
            log_linea("No se pudo guardar config.ini: %s" % e)

    # ---------------- acciones ----------------
    def aviso(msg, error=False):
        status.value = msg
        try:
            page.open(ft.SnackBar(ft.Text(msg),
                                  bgcolor="#B00020" if error else TEAL))
        except Exception:
            try:
                page.snack_bar = ft.SnackBar(ft.Text(msg))
                page.snack_bar.open = True
            except Exception:
                pass
        page.update()

    def set_busy(b):
        estado["trabajando"] = b
        for btn in (btn_desc, btn_excel, btn_diag):
            btn.disabled = b
        progress.value = None if b else 0
        status.value = "Trabajando..." if b else status.value
        page.update()

    def construir_cfg():
        usuario = (tf_user.value or "").strip()
        contrasena = tf_pass.value or ""
        if not usuario or not contrasena:
            aviso("Ingresa el RUT/usuario y la contrasena.", error=True)
            return None
        return cc.Config(usuario=usuario, contrasena=contrasena, rut="",
                         carpeta_salida=(tf_folder.value or "").strip() or _carpeta_default(),
                         url_base="https://www.campuscentinela.cl",
                         extranjero=(rg_opcion.value == "extranjero"))

    def trabajo(cfg, dump, sin_descarga):
        def progreso_cb(i, total):
            progress.value = (i / total) if total else None
            try:
                page.update()
            except Exception:
                pass
        try:
            cursos, ruta, conteo = cc.ejecutar(cfg, dump=dump,
                                               sin_descarga=sin_descarga,
                                               progreso_cb=progreso_cb)
            if not cursos:
                aviso("No se detectaron certificados. Usa 'Diagnostico' y revisa "
                      "la carpeta diagnostico.", error=True)
            else:
                msg = ("Listo. %d cursos | Vencidos: %d | Por vencer: %d | Vigentes: %d"
                       % (len(cursos), conteo.get("VENCIDO", 0),
                          conteo.get("POR_VENCER", 0), conteo.get("VIGENTE", 0)))
                aviso(msg)
        except Exception as e:
            aviso("Error: %s" % e, error=True)
        finally:
            progress.value = 0
            set_busy(False)

    def lanzar(dump, sin_descarga):
        if estado["trabajando"]:
            return
        cfg = construir_cfg()
        if cfg is None:
            return
        if sw_guardar.value:
            guardar_config(cfg)
        set_busy(True)
        threading.Thread(target=trabajo, args=(cfg, dump, sin_descarga),
                         daemon=True).start()

    def abrir_carpeta(e=None):
        carpeta = Path((tf_folder.value or "").strip() or _carpeta_default())
        try:
            carpeta.mkdir(parents=True, exist_ok=True)
            ruta = str(carpeta.resolve())
            if sys.platform.startswith("win"):
                os.startfile(ruta)  # noqa
            elif sys.platform == "darwin":
                os.system('open "%s"' % ruta)
            else:
                os.system('xdg-open "%s"' % ruta)
        except Exception as ex:
            aviso("No se pudo abrir la carpeta: %s" % ex, error=True)

    def on_pick(e: ft.FilePickerResultEvent):
        if e.path:
            tf_folder.value = e.path
            page.update()
    file_picker.on_result = on_pick

    def elegir_carpeta(e=None):
        try:
            file_picker.get_directory_path(dialog_title="Carpeta de salida")
        except Exception:
            aviso("No se pudo abrir el selector de carpetas.", error=True)

    # ---------------- botones ----------------
    btn_desc = ft.FilledButton("Descargar certificados + Excel",
                               on_click=lambda e: lanzar(False, False),
                               style=ft.ButtonStyle(bgcolor=TEAL, color="#FFFFFF",
                                                    padding=18))
    btn_excel = ft.OutlinedButton("Solo Excel",
                                  on_click=lambda e: lanzar(False, True))
    btn_diag = ft.OutlinedButton("Diagnostico",
                                 on_click=lambda e: lanzar(True, True))
    btn_abrir = ft.TextButton("Abrir carpeta", on_click=abrir_carpeta)

    # ---------------- layout ----------------
    header = ft.Container(
        bgcolor=TEAL, padding=ft.padding.symmetric(20, 24),
        content=ft.Column(spacing=2, controls=[
            ft.Text("Campus Centinela", size=22, weight=ft.FontWeight.BOLD,
                    color="#FFFFFF"),
            ft.Text("Descarga de certificados y control de vencimientos",
                    size=12, color="#d8ecee"),
        ]))

    tarjeta = ft.Container(
        padding=18, border_radius=14, bgcolor="#F4F7F8",
        content=ft.Column(spacing=12, controls=[
            ft.Text("Tipo de documento:", size=12, color="#555"),
            rg_opcion,
            tf_user, tf_pass,
            ft.Row([tf_folder, ft.OutlinedButton("Examinar", on_click=elegir_carpeta)],
                   spacing=8),
            sw_guardar,
        ]))

    acciones = ft.Row(wrap=True, spacing=10,
                      controls=[btn_desc, btn_excel, btn_diag,
                                ft.Container(expand=True), btn_abrir])

    caja_log = ft.Container(
        bgcolor=FONDO_LOG, border_radius=12, expand=True, content=log_list)

    cuerpo = ft.Container(
        padding=20, expand=True,
        content=ft.Column(expand=True, spacing=14, controls=[
            tarjeta, acciones, progress,
            ft.Text("Actividad:", size=12, weight=ft.FontWeight.BOLD),
            caja_log, status,
        ]))

    page.add(ft.Column(expand=True, spacing=0, controls=[header, cuerpo]))
    cargar_config()


if __name__ == "__main__":
    ft.app(target=main)
