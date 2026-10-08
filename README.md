# Campus Centinela — Descarga y seguimiento de certificados

Herramienta para **iniciar sesión en [campuscentinela.cl](https://www.campuscentinela.cl)**,
descargar los certificados de los cursos, **ordenarlos por categoría y subcategoría** y
generar un **Excel con las fechas de emisión y vencimiento**, para controlar cursos
pendientes o vencidos.

Incluye **interfaz gráfica** (botones, sin línea de comandos) y se puede obtener como
**ejecutable `.exe` de Windows**.

---

## ¿Tiene API la plataforma?

**No.** El sitio corre sobre **Joomla + Joomdle/Moodle**. El login es el módulo estándar
`com_users` con **token CSRF** dinámico. Los certificados los genera un **script PHP**
(`certificado.php?id=...&rut=...`). El `id` del certificado es el **id de curso de
Moodle**, el mismo que aparece en los enlaces `land.php?...&id=...` del catálogo. Por eso
la herramienta extrae ese `id` de cada curso y arma la URL del certificado.

El enfoque correcto es mantener una **sesión (cookie)** e interactuar con las páginas
como un navegador. Es lo que hace esta herramienta.

---

## Categorías

Los certificados se guardan en **dos niveles de carpetas**:

```
salida/certificados/
├── Aulas Virtuales/
│   ├── Seguridad y Salud Ocupacional/
│   ├── Operaciones Planta Concentradora y-o Catodo/
│   └── Mantenimiento Mina/
└── E-Learning/
    ├── TRANSV/        (Transversales)
    ├── OM/            (Operación Mina)
    ├── EDC/           (Estándares de Control)
    ├── MA/            (Medio Ambiente)
    ├── DI/  DM/  MANT/  MUELLE/  OPC/  EO/  FIN/
    └── Otros/         (sin prefijo reconocible)
```

- **Nivel 1 (modalidad):** `Aulas Virtuales` o `E-Learning`, según el nombre del curso.
- **Nivel 2 (subcategoría):**
  - E-Learning → el **prefijo** del curso (TRANSV, OM, EDC, DI, DM, MA, MANT, MUELLE,
    OPC, EO, FIN…). Si no hay prefijo → `Otros`.
  - Aulas Virtuales → el **área** (SSO / Operaciones Planta / Mantenimiento Mina).

Las mismas columnas (Modalidad y Subcategoría) aparecen en el Excel, con filtros.

---

## Opción A — Ejecutable `.exe` (recomendado si Python te da problemas)

No necesitas instalar Python. Hay dos formas de conseguir el `.exe`:

### A.1 Descargarlo ya compilado (GitHub Actions)

Cada vez que se actualiza el código, GitHub compila el `.exe` automáticamente en un
servidor Windows:

1. Entra al repositorio en GitHub → pestaña **Actions**.
2. Abre la ejecución más reciente de **“Compilar EXE Windows”** (con ✓ verde).
3. En **Artifacts**, descarga **`CampusCentinela-windows`** y descomprímelo.
4. Ejecuta **`CampusCentinela.exe`**.

### A.2 Compilarlo tú en tu PC con Windows

Con Python instalado, haz doble clic en **`build_exe.bat`**. Al terminar, el ejecutable
queda en `dist\CampusCentinela.exe`.

> El `.exe` **no incluye** tus credenciales. Al abrirlo ingresas RUT y contraseña; puedes
> marcar “Recordar datos” para guardarlos en un `config.ini` junto al ejecutable.

---

## Opción B — Interfaz gráfica desde el código

Interfaz **Flet** (estilo app, moderna):

```bash
pip install -r requirements.txt -r requirements-gui.txt
python campus_centinela_flet.py
```

Se abre una ventana: ingresas **RUT y contraseña** y usas los botones
(**Descargar certificados + Excel**, **Solo Excel**, **Diagnóstico**, **Abrir carpeta**).

> También existe `campus_centinela_gui.py` (Tkinter, sin dependencias extra) como
> alternativa de respaldo; el ejecutable oficial usa Flet.

---

## Opción C — Línea de comandos

```bash
pip install -r requirements.txt
cp config.example.ini config.ini        # completar RUT + contraseña
python campus_centinela.py --dump        # 1ª vez: diagnóstico
python campus_centinela.py               # descarga + Excel
```

Credenciales: en `config.ini` (ignorado por git), por variables de entorno
`CC_USUARIO`/`CC_CONTRASENA`/`CC_RUT`, o con `--interactivo`.

| Opción            | Qué hace                                               |
|-------------------|--------------------------------------------------------|
| `--dump`          | Guarda `diagnostico/portada.html` y `progreso-cursos.html` |
| `--sin-descarga`  | Solo genera el Excel                                   |
| `--salida CARPETA`| Cambia la carpeta de resultados                        |
| `--interactivo`   | Pide usuario/contraseña por teclado                    |
| `-v`              | Modo detallado                                         |

---

## El Excel

Hoja **“Certificados”**: Modalidad, Subcategoría, Curso, Estado, Situación, Fecha de
emisión, Fecha de vencimiento, Días para vencer, ID curso, RUT, Archivo PDF, URL.
Colores: 🔴 vencido · 🟡 por vencer (≤ 60 días) · 🟢 vigente.
Hoja **“Resumen”**: totales por situación y por modalidad.

---

## Pruebas

Las pruebas (sin red) validan la clasificación con nombres reales de cursos, la
extracción de `id`, la construcción de la URL de certificado y la generación del Excel:

```bash
python tests/test_core.py
```

---

## Si el diagnóstico no reconoce los certificados

La página `/progreso-cursos` podría tener una estructura distinta a la supuesta. En ese
caso:

1. Botón **Diagnóstico** (o `python campus_centinela.py --dump`).
2. Revisa / comparte `diagnostico/progreso-cursos.html` para afinar el parser
   (`extraer_cursos` en `campus_centinela.py`).

---

## Nota de seguridad

- Usa la herramienta solo con cuentas que estés autorizado/a a gestionar.
- `config.ini` está excluido del repositorio para no subir credenciales.
- Si una contraseña o un **token de sesión** (los enlaces `land.php?...&token=...`) se
  compartió por chat o captura, conviene **cerrarlo/cambiarla**.
