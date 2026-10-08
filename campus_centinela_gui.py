#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Campus Centinela - Interfaz grafica
===================================

Ventana sencilla para descargar los certificados y generar el Excel sin usar
la linea de comandos. Usa Tkinter (incluido en Python, sin instalar nada).

Es la base del ejecutable .exe: al compilarla con PyInstaller queda un
programa de un solo archivo con menu/botones.

Ejecutar desde fuente:  python campus_centinela_gui.py
"""

import os
import queue
import sys
import threading
import logging
import configparser
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    import campus_centinela as cc
except Exception as e:  # dependencias faltantes, etc.
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error al iniciar",
                         "No se pudo cargar el nucleo del programa:\n\n%s\n\n"
                         "Si ejecutas desde el codigo fuente, instala las "
                         "dependencias:\n    pip install -r requirements.txt" % e)
    raise

# Paleta (tono de la plataforma)
TEAL = "#0B6E7A"
TEAL_OSCURO = "#095761"
GRIS = "#F2F5F6"
BLANCO = "#FFFFFF"


class ColaLogHandler(logging.Handler):
    """Envia los mensajes de log a una cola para mostrarlos en la ventana."""
    def __init__(self, cola):
        super().__init__()
        self.cola = cola

    def emit(self, record):
        self.cola.put(("log", self.format(record)))


class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=0)
        self.master = master
        self.cola = queue.Queue()
        self.trabajando = False
        self._construir()
        self._configurar_logging()
        self._cargar_config_si_existe()
        self.after(120, self._procesar_cola)

    # ------------------------------------------------------------------ UI
    def _construir(self):
        self.master.title("Campus Centinela - Certificados")
        self.master.geometry("820x620")
        self.master.minsize(720, 560)
        self.master.configure(bg=GRIS)

        estilo = ttk.Style()
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass
        estilo.configure("TFrame", background=GRIS)
        estilo.configure("TLabel", background=GRIS, font=("Segoe UI", 10))
        estilo.configure("Header.TLabel", background=TEAL, foreground=BLANCO,
                         font=("Segoe UI", 16, "bold"))
        estilo.configure("Sub.TLabel", background=TEAL, foreground=BLANCO,
                         font=("Segoe UI", 9))
        estilo.configure("TButton", font=("Segoe UI", 10, "bold"), padding=8)
        estilo.configure("Primario.TButton", foreground=BLANCO, background=TEAL)
        estilo.map("Primario.TButton",
                   background=[("active", TEAL_OSCURO), ("disabled", "#9bbdc2")])
        estilo.configure("TEntry", padding=5)
        estilo.configure("TCheckbutton", background=GRIS, font=("Segoe UI", 9))
        self.pack(fill="both", expand=True)

        # Encabezado
        cab = tk.Frame(self, bg=TEAL)
        cab.pack(fill="x")
        ttk.Label(cab, text="  Campus Centinela", style="Header.TLabel").pack(
            anchor="w", padx=16, pady=(14, 0))
        ttk.Label(cab, text="  Descarga de certificados y control de vencimientos",
                  style="Sub.TLabel").pack(anchor="w", padx=16, pady=(0, 14))

        cuerpo = ttk.Frame(self, padding=16)
        cuerpo.pack(fill="both", expand=True)

        # --- Credenciales ---
        form = ttk.LabelFrame(cuerpo, text=" Datos de acceso ", padding=12)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="RUT / usuario:").grid(row=0, column=0, sticky="w", pady=4)
        self.var_usuario = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_usuario).grid(
            row=0, column=1, columnspan=2, sticky="ew", pady=4, padx=(8, 0))

        ttk.Label(form, text="Contrasena:").grid(row=1, column=0, sticky="w", pady=4)
        self.var_pass = tk.StringVar()
        self.entry_pass = ttk.Entry(form, textvariable=self.var_pass, show="•")
        self.entry_pass.grid(row=1, column=1, sticky="ew", pady=4, padx=(8, 0))
        self.var_ver = tk.BooleanVar(value=False)
        ttk.Checkbutton(form, text="Ver", variable=self.var_ver,
                        command=self._toggle_pass).grid(row=1, column=2, padx=(8, 0))

        ttk.Label(form, text="RUT en certificado:").grid(row=2, column=0, sticky="w", pady=4)
        self.var_rut = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_rut).grid(
            row=2, column=1, columnspan=2, sticky="ew", pady=4, padx=(8, 0))
        ttk.Label(form, text="(dejar vacio = usar el mismo RUT de arriba)",
                  foreground="#666").grid(row=3, column=1, columnspan=2, sticky="w")

        ttk.Label(form, text="Carpeta de salida:").grid(row=4, column=0, sticky="w", pady=4)
        self.var_salida = tk.StringVar(value="salida")
        ttk.Entry(form, textvariable=self.var_salida).grid(
            row=4, column=1, sticky="ew", pady=4, padx=(8, 0))
        ttk.Button(form, text="Examinar...", command=self._elegir_carpeta,
                   style="TButton").grid(row=4, column=2, padx=(8, 0))

        self.var_guardar = tk.BooleanVar(value=True)
        ttk.Checkbutton(form, text="Recordar datos en config.ini (no incluye el .exe)",
                        variable=self.var_guardar).grid(
            row=5, column=1, columnspan=2, sticky="w", pady=(6, 0))

        # --- Botones de accion ---
        acciones = ttk.Frame(cuerpo)
        acciones.pack(fill="x", pady=(12, 8))
        self.btn_descargar = ttk.Button(
            acciones, text="Descargar certificados + Excel",
            style="Primario.TButton", command=self._accion_completa)
        self.btn_descargar.pack(side="left")
        self.btn_excel = ttk.Button(
            acciones, text="Solo Excel (sin PDFs)", command=self._accion_solo_excel)
        self.btn_excel.pack(side="left", padx=(8, 0))
        self.btn_dump = ttk.Button(
            acciones, text="Diagnostico", command=self._accion_dump)
        self.btn_dump.pack(side="left", padx=(8, 0))
        self.btn_abrir = ttk.Button(
            acciones, text="Abrir carpeta", command=self._abrir_carpeta)
        self.btn_abrir.pack(side="right")

        # --- Progreso ---
        self.progreso = ttk.Progressbar(cuerpo, mode="determinate")
        self.progreso.pack(fill="x", pady=(4, 8))

        # --- Log ---
        ttk.Label(cuerpo, text="Actividad:").pack(anchor="w")
        self.txt = scrolledtext.ScrolledText(cuerpo, height=14, wrap="word",
                                             font=("Consolas", 9), state="disabled",
                                             bg="#0f1b1d", fg="#d6e6e8")
        self.txt.pack(fill="both", expand=True, pady=(2, 0))

        self.estado = ttk.Label(cuerpo, text="Listo.", foreground="#444")
        self.estado.pack(anchor="w", pady=(6, 0))

    def _toggle_pass(self):
        self.entry_pass.configure(show="" if self.var_ver.get() else "•")

    def _elegir_carpeta(self):
        d = filedialog.askdirectory(title="Carpeta de salida")
        if d:
            self.var_salida.set(d)

    # ------------------------------------------------------------- logging
    def _configurar_logging(self):
        handler = ColaLogHandler(self.cola)
        handler.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
        logger = logging.getLogger("campus")
        logger.setLevel(logging.INFO)
        logger.handlers = [handler]

    def _log(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    # ------------------------------------------------------------- config
    def _ruta_config(self):
        # Junto al ejecutable/script.
        base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        if getattr(sys, "frozen", False):
            base = Path(sys.executable).resolve().parent
        return base / "config.ini"

    def _cargar_config_si_existe(self):
        ruta = self._ruta_config()
        if not ruta.exists():
            return
        cp = configparser.ConfigParser()
        try:
            cp.read(ruta, encoding="utf-8")
            self.var_usuario.set(cp.get("credenciales", "usuario", fallback=""))
            self.var_pass.set(cp.get("credenciales", "contrasena", fallback=""))
            self.var_rut.set(cp.get("credenciales", "rut", fallback=""))
            self.var_salida.set(cp.get("opciones", "carpeta_salida", fallback="salida"))
            self._log("Datos cargados desde config.ini")
        except Exception as e:
            self._log("No se pudo leer config.ini: %s" % e)

    def _guardar_config(self, cfg):
        cp = configparser.ConfigParser()
        cp["credenciales"] = {"usuario": cfg.usuario, "contrasena": cfg.contrasena,
                              "rut": self.var_rut.get().strip()}
        cp["opciones"] = {"carpeta_salida": cfg.carpeta_salida, "url_base": cfg.url_base}
        try:
            with open(self._ruta_config(), "w", encoding="utf-8") as f:
                cp.write(f)
        except Exception as e:
            self._log("No se pudo guardar config.ini: %s" % e)

    # ------------------------------------------------------------- acciones
    def _construir_cfg(self):
        usuario = self.var_usuario.get().strip()
        contrasena = self.var_pass.get()
        if not usuario or not contrasena:
            messagebox.showwarning("Faltan datos",
                                   "Ingresa el RUT/usuario y la contrasena.")
            return None
        return cc.Config(usuario=usuario, contrasena=contrasena,
                         rut=self.var_rut.get().strip(),
                         carpeta_salida=self.var_salida.get().strip() or "salida",
                         url_base="https://www.campuscentinela.cl")

    def _accion_completa(self):
        self._lanzar(dump=False, sin_descarga=False)

    def _accion_solo_excel(self):
        self._lanzar(dump=False, sin_descarga=True)

    def _accion_dump(self):
        self._lanzar(dump=True, sin_descarga=True)

    def _lanzar(self, dump, sin_descarga):
        if self.trabajando:
            return
        cfg = self._construir_cfg()
        if cfg is None:
            return
        if self.var_guardar.get():
            self._guardar_config(cfg)
        self.trabajando = True
        self._set_botones(False)
        self.progreso.configure(mode="indeterminate")
        self.progreso.start(12)
        self.estado.configure(text="Trabajando...")
        hilo = threading.Thread(target=self._trabajo, args=(cfg, dump, sin_descarga),
                                daemon=True)
        hilo.start()

    def _trabajo(self, cfg, dump, sin_descarga):
        def progreso_cb(i, total):
            self.cola.put(("progreso", i, total))
        try:
            cursos, ruta, conteo = cc.ejecutar(cfg, dump=dump, sin_descarga=sin_descarga,
                                               progreso_cb=progreso_cb)
            if not cursos:
                self.cola.put(("fin", False,
                               "No se detectaron certificados. Usa 'Diagnostico' "
                               "y revisa la carpeta 'diagnostico'."))
            else:
                resumen = ("Listo. %d cursos. Vencidos: %d | Por vencer: %d | "
                           "Vigentes: %d." % (len(cursos), conteo.get("VENCIDO", 0),
                                              conteo.get("POR_VENCER", 0),
                                              conteo.get("VIGENTE", 0)))
                self.cola.put(("fin", True, resumen))
        except Exception as e:
            self.cola.put(("fin", False, "Error: %s" % e))

    def _set_botones(self, activo):
        estado = "normal" if activo else "disabled"
        for b in (self.btn_descargar, self.btn_excel, self.btn_dump):
            b.configure(state=estado)

    def _abrir_carpeta(self):
        carpeta = Path(self.var_salida.get().strip() or "salida")
        carpeta.mkdir(parents=True, exist_ok=True)
        ruta = str(carpeta.resolve())
        try:
            if sys.platform.startswith("win"):
                os.startfile(ruta)  # noqa
            elif sys.platform == "darwin":
                os.system('open "%s"' % ruta)
            else:
                os.system('xdg-open "%s"' % ruta)
        except Exception as e:
            messagebox.showinfo("Carpeta", "Carpeta de salida:\n%s\n\n(%s)" % (ruta, e))

    # ------------------------------------------------------------- cola
    def _procesar_cola(self):
        try:
            while True:
                item = self.cola.get_nowait()
                tipo = item[0]
                if tipo == "log":
                    self._log(item[1])
                elif tipo == "progreso":
                    i, total = item[1], item[2]
                    if self.progreso["mode"] != "determinate":
                        self.progreso.stop()
                        self.progreso.configure(mode="determinate", maximum=total)
                    self.progreso["value"] = i
                elif tipo == "fin":
                    ok, msg = item[1], item[2]
                    self.progreso.stop()
                    self.progreso.configure(mode="determinate")
                    self.progreso["value"] = self.progreso["maximum"] if ok else 0
                    self.trabajando = False
                    self._set_botones(True)
                    self.estado.configure(text=msg)
                    self._log(msg)
                    if ok:
                        messagebox.showinfo("Completado", msg)
                    else:
                        messagebox.showwarning("Aviso", msg)
        except queue.Empty:
            pass
        self.after(120, self._procesar_cola)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
