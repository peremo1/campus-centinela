# Campus Centinela — Descarga y seguimiento de certificados

Herramienta para **iniciar sesión en [campuscentinela.cl](https://www.campuscentinela.cl)**,
descargar los certificados de los cursos, **ordenarlos por categoría** y generar un
**Excel con las fechas de emisión y vencimiento**, para controlar cursos pendientes o
vencidos y no quedar atrasado.

> Pensada para que un/a encargado/a de capacitación o prevención lleve el control de la
> situación de sus cursos (o los del personal autorizado a su cargo), usando las
> credenciales de la cuenta correspondiente.

---

## ¿Tiene API la plataforma?

**No.** El sitio corre sobre **Joomla** (CMS): el login es el módulo estándar
`com_users`, con un **token CSRF** dinámico (el campo oculto de nombre aleatorio con
valor `1`) y un campo `return` en base64. Los certificados se generan con un **script
PHP a medida** (`certificado.php?id=...&rut=...`), no con un API REST. El mismo link
cambia el PDF según el parámetro `id`; el `rut` identifica a la persona.

Por eso el enfoque correcto es **mantener una sesión (cookie)** e interactuar con las
páginas como lo haría un navegador. Eso es justamente lo que hace esta herramienta.

---

## Requisitos

- Python 3.9 o superior
- Conexión a internet con acceso a `campuscentinela.cl`

Instala las dependencias:

```bash
pip install -r requirements.txt
```

---

## Configuración (credenciales fuera del código)

1. Copia la plantilla y edítala:

   ```bash
   cp config.example.ini config.ini
   ```

2. Completa `config.ini` con el **RUT** (usuario) y la **contraseña**:

   ```ini
   [credenciales]
   usuario = 21080196-0
   contrasena = TU_CONTRASENA
   rut =                 ; vacío = usa el mismo 'usuario'

   [opciones]
   carpeta_salida = salida
   url_base = https://www.campuscentinela.cl
   ```

> **`config.ini` nunca se sube al repositorio** (está en `.gitignore`).
> También puedes pasar las credenciales por variables de entorno
> `CC_USUARIO` / `CC_CONTRASENA` / `CC_RUT`, o con `--interactivo` para que te las pida
> por teclado.

---

## Uso

**Primera vez — diagnóstico** (inicia sesión y guarda el HTML real de las páginas para
verificar que todo calza con tu cuenta):

```bash
python campus_centinela.py --dump
```

**Uso normal** (login → descarga certificados → genera Excel):

```bash
python campus_centinela.py
```

Opciones útiles:

| Opción            | Qué hace                                                        |
|-------------------|-----------------------------------------------------------------|
| `--dump`          | Guarda `diagnostico/portada.html` y `progreso-cursos.html`      |
| `--sin-descarga`  | Solo genera el Excel, sin bajar los PDF                          |
| `--salida CARPETA`| Cambia la carpeta de resultados                                 |
| `--interactivo`   | Pide usuario/contraseña por teclado si faltan                   |
| `-v`              | Modo detallado (debug)                                          |

---

## Resultado

```
salida/
├── certificados/
│   ├── Seguridad Minera/
│   │   ├── Induccion Hombre Nuevo_id134.pdf
│   │   └── Manejo de Sustancias Peligrosas_id140.pdf
│   └── Salud Ocupacional/
│       └── Protocolo Silice_id201.pdf
└── certificados_centinela.xlsx
```

El Excel (`certificados_centinela.xlsx`) trae:

- **Hoja "Certificados":** categoría, curso, estado, situación, fecha de emisión,
  fecha de vencimiento, días para vencer, ID, RUT, ruta del PDF y URL. Con filtros
  y colores: 🔴 vencido · 🟡 por vencer (≤ 60 días) · 🟢 vigente.
- **Hoja "Resumen":** totales por situación.

---

## Convertir a `.exe` (opcional, Windows)

```bash
pip install pyinstaller
pyinstaller --onefile campus_centinela.py
```

El ejecutable queda en `dist/`. Debe acompañarse de un `config.ini` en la misma carpeta
(o usar variables de entorno / `--interactivo`).

---

## Si el parser no reconoce los certificados

La página de progreso puede tener una estructura distinta a la esperada (o cargar datos
por JavaScript). En ese caso:

1. Ejecuta `python campus_centinela.py --dump`
2. Revisa `diagnostico/progreso-cursos.html`

Con ese HTML se pueden afinar los selectores del parser (función `extraer_certificados`).

---

## Nota de seguridad

- No guardes la contraseña en claro más de lo necesario; `config.ini` está excluido del
  repositorio justamente por eso.
- Usa esta herramienta solo con cuentas que estés autorizado/a a gestionar.
- Si una contraseña se compartió por chat o captura de pantalla, conviene **cambiarla**.
