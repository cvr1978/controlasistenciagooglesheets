# Sistema de Control de Asistencia con QR — Google Sheets

Sistema web completo para registrar la asistencia de estudiantes escaneando el **código QR de su carnet** con la cámara del dispositivo. Construido sobre **Google Apps Script** y **Google Sheets**.

---

## Arquitectura

```
Google Sheets (base de datos)
    ├── Hoja "Estudiantes"      — Catálogo de estudiantes
    ├── Hoja "Asistencia"       — Registro de entradas con fecha/hora
    └── Hoja "Configuracion"    — Parámetros del sistema

Google Apps Script (backend + frontend)
    ├── Code.gs                 — Lógica del servidor
    ├── Index.html              — Escáner QR (página principal)
    ├── Dashboard.html          — Reportes y estadísticas
    └── Estudiantes.html        — Gestión del catálogo
```

---

## Instalación paso a paso

### 1. Crear el Google Spreadsheet

1. Ve a [sheets.google.com](https://sheets.google.com) y crea un nuevo spreadsheet.
2. Nómbralo, por ejemplo: **"Control de Asistencia"**.

### 2. Crear el proyecto Apps Script

1. En el spreadsheet, ve a **Extensiones → Apps Script**.
2. Se abre el editor de código.
3. **Borra** el contenido del archivo `Code.gs` que aparece por defecto.

### 3. Copiar los archivos

Copia el contenido de cada archivo de este repositorio al editor:

| Archivo del repo | Tipo en Apps Script |
|-----------------|---------------------|
| `Code.gs`       | Script (ya existe, reemplaza su contenido) |
| `Index.html`    | Archivo HTML nuevo → botón **+** → HTML → nombre: `Index` |
| `Dashboard.html`| Archivo HTML nuevo → nombre: `Dashboard` |
| `Estudiantes.html` | Archivo HTML nuevo → nombre: `Estudiantes` |
| `appsscript.json` | Visible en **Configuración del proyecto → Mostrar archivo de manifiesto** |

> **Importante:** Los nombres de los archivos HTML deben ser exactamente `Index`, `Dashboard` y `Estudiantes` (sin extensión en el editor de Apps Script).

### 4. Ajustar la zona horaria

En `appsscript.json`, verifica que `timeZone` coincida con tu país:

| País | Zona horaria |
|------|-------------|
| Colombia | `America/Bogota` |
| México (CDMX) | `America/Mexico_City` |
| Argentina | `America/Argentina/Buenos_Aires` |
| España | `Europe/Madrid` |

### 5. Inicializar el Spreadsheet

1. En el editor, busca la función `inicializarSpreadsheet` y ejecútala pulsando **▶ Ejecutar**.
2. Acepta los permisos que solicite Google.
3. Esto creará automáticamente las tres hojas con sus encabezados y formatos.

### 6. Publicar como Web App

1. En el editor, haz clic en **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web**.
3. Configuración:
   - **Ejecutar como:** Yo (tu cuenta)
   - **Quién tiene acceso:** Cualquier usuario (o "Todos en mi organización")
4. Haz clic en **Implementar**.
5. Copia la **URL de la aplicación web** — esa es la URL que usarás.

---

## Uso del sistema

### Escáner QR (página principal)
- Abre la URL de la web app en cualquier dispositivo con cámara.
- El escáner se inicia automáticamente.
- Apunta el QR del carnet del estudiante a la cámara.
- El sistema registra la asistencia y muestra el nombre del estudiante.
- Si no lee bien, usa el campo de **Registro Manual**.

### Dashboard
- Accede a `[URL]?page=dashboard`
- Muestra KPIs del día: presentes, puntuales, tardíos, ausentes.
- Filtra por rango de fechas.
- Exporta a CSV.

### Gestión de Estudiantes
- Accede a `[URL]?page=estudiantes`
- Agrega, edita o desactiva estudiantes.
- El **ID/Código** debe coincidir exactamente con el contenido del QR del carnet.

---

## QR del carnet

El QR del carnet debe contener **únicamente el ID o código del estudiante** (texto plano, sin espacios extra).

### Opciones para generar QR
- [qr-code-generator.com](https://www.qr-code-generator.com/) — gratuito, en línea
- [goqr.me](https://goqr.me/) — gratuito, en línea
- Librería Python `qrcode` para generación masiva:

```python
import qrcode

estudiantes = [
    ("2024001", "Juan García"),
    ("2024002", "María López"),
]

for codigo, nombre in estudiantes:
    img = qrcode.make(codigo)
    img.save(f"qr_{codigo}_{nombre}.png")
```

Instala con: `pip install qrcode[pil]`

---

## Configuración avanzada

Edita la hoja **"Configuracion"** en el spreadsheet:

| Clave | Descripción | Valor por defecto |
|-------|-------------|------------------|
| `Institucion` | Nombre de la institución | Mi Institución |
| `Materia` | Materia que se está dictando | Materia General |
| `Docente` | Nombre del docente | Nombre del Docente |
| `Tolerancia` | Minutos de tolerancia para marcar "Puntual" | 15 |
| `HoraClaseH` | Hora de inicio de clase (horas) | 7 |
| `HoraClaseM` | Hora de inicio de clase (minutos) | 0 |

---

## Estructura del Spreadsheet

### Hoja "Estudiantes"
| Columna | Descripción |
|---------|-------------|
| ID / Código | Debe coincidir con el contenido del QR |
| Nombre Completo | Nombre del estudiante |
| Programa / Curso | Carrera o curso |
| Email | Correo electrónico (opcional) |
| Estado | `Activo` o `Inactivo` |

### Hoja "Asistencia"
| Columna | Descripción |
|---------|-------------|
| Fecha | `yyyy-MM-dd` |
| Hora | `HH:mm:ss` |
| ID / Código | Código del estudiante |
| Nombre Completo | Nombre registrado |
| Programa / Curso | Programa del estudiante |
| Materia | Materia configurada |
| Observación | `Puntual` o `Tarde` |

---

## Reglas del sistema

- Un estudiante **no puede registrar asistencia dos veces en el mismo día**.
- Si el estado es **Inactivo**, no se registra asistencia.
- Se marca **Tarde** si llega después de: hora de clase + minutos de tolerancia.
- Las filas de asistencia se colorean: **verde** para Puntual, **amarillo** para Tarde.

---

## Menú en Google Sheets

Después de inicializar, aparece el menú **🎓 Asistencia QR** en la barra del spreadsheet con accesos directos a todas las funciones.

---

## Tecnologías usadas

- **Google Apps Script** — backend y hosting
- **Google Sheets** — base de datos
- **html5-qrcode** — librería de escaneo QR con cámara
- HTML5 / CSS3 / JavaScript vanilla
