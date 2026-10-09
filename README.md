# Campus Centinela — Descarga y seguimiento de certificados

Herramienta para **iniciar sesión en [campuscentinela.cl](https://www.campuscentinela.cl)**,
descargar los certificados de los cursos, **ordenarlos por categoría y subcategoría** y
generar un **Excel con las fechas de emisión y vencimiento**, para controlar cursos
pendientes o vencidos.

Incluye **interfaz gráfica** (botones, sin línea de comandos) y se puede obtener como
**ejecutable `.exe` de Windows**.

---

## ¿Cómo obtiene los datos?

El sitio corre sobre **Joomla 3.10 + Joomdle/Moodle**. El login es el módulo estándar
`com_users` (con **token CSRF** y un selector **NACIONAL/EXTRANJERO**). No hay API
pública documentada, pero la página de progreso usa internamente un **endpoint JSON**
(`index.php?option=com_ajax&module=customphp&method=consultorcursosSegundaVersion&format=json&rut=...`)
que devuelve **todos los cursos** con su categoría, estado, nota, avance, **vigencia**,
**fecha de nota (emisión)** y la **URL directa del certificado** (`certificado_directo`).

La herramienta inicia sesión (manteniendo la cookie), **consume ese JSON directamente**
—mucho más robusto que scrapear HTML— y de ahí calcula vencimientos y descarga los
certificados (`certificado.php?id=<id_curso>&rut=<rut>`). También recoge el certificado
de **firma digital** cuando existe.

---

## Categorías

Los certificados se guardan en **dos niveles de carpetas**: la **categoría real** de la
plataforma (nivel 1) y el **prefijo** del curso (nivel 2):

```
salida/certificados/
├── SEGURIDAD Y SALUD OCUPACIONAL/
│   ├── EDC/
│   └── TRANSV/
├── OPERACIONES MINA MANUAL Y-O AUTONOMA/
│   └── OM/
├── OPERACIONES PLANTA CONCENTRADORA Y-O CATODO/
│   └── OPC/
├── SUSTENTABILIDAD Y MEDIO AMBIENTE/
│   └── MA/
├── MANTENIMIENTO MINA/   MUELLE/   DIVERSIDAD E INCLUSION/   ...
```

- **Nivel 1 (categoría):** la categoría real que entrega la plataforma (p. ej.
  `SEGURIDAD Y SALUD OCUPACIONAL`, `OPERACIONES PLANTA CONCENTRADORA Y/O CÁTODO`).
- **Nivel 2 (subcategoría):** el **prefijo** del curso (TRANSV, OM, EDC, MA, DI, DM,
  MANT, MUELLE, OPC, EO, FIN…); si no hay prefijo → `Otros`.

Solo se descargan los certificados de cursos **APROBADOS** (los demás no tienen
certificado emitido), pero todos los cursos aparecen en el Excel.

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

El Excel tiene **dos hojas**:

- **“Panel”** (se abre primero): datos del trabajador (nombre, RUT), un **parámetro
  editable** — *Días de aviso (por vencer)* en la celda **C8**, por defecto 60 — y los
  **KPIs** (Vigentes / Por vencer / Vencidos / Indefinidos / Pendientes) y totales por
  categoría, todos con fórmulas que **se recalculan al abrir**.
- **“Cursos”**: tabla ordenada por categoría con Curso, **Fecha de nota**, **Expiración**,
  **Días restantes** y **Situación**. “Días restantes” es una **fórmula dinámica**
  (`= Expiración − HOY()`) que lee el reloj del PC, y la fila se **colorea sola** según el
  umbral del Panel: 🟢 vigente (> umbral) · 🟡 por vencer (≤ umbral) · 🔴 vencida ·
  ⬜ indefinida/fija. Columnas extra: Nota %, Avance %, Vigencia, Estado, enlaces al
  PDF local (si lo descargaste) y a la URL del certificado online, y **Fuente vig.**

**Segunda verificación con el PDF:** al descargar los certificados, la herramienta
**lee cada PDF** y toma de ahí la fecha de vencimiento real (`VIGENCIA DEL CERTIFICADO`).
Esas filas quedan con *Fuente vig. = PDF (certificado)*; las demás usan la fecha de la
plataforma (*Plataforma*). Además, todo el proceso queda en
`salida/diagnostico/registro.txt` para revisar qué pasó o por qué falló algo.

Como todo es por fórmula, basta **reabrir el Excel** (o cambiar el 60 de la celda C8)
para ver el estado actualizado sin volver a ejecutar el programa.

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
