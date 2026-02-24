# DP Educación — Automatización Google Sheets

Sistema de gestión de inscripciones y salones de clase implementado con **Google Apps Script**.

---

## Estructura de Archivos

| Archivo | Descripción |
|---|---|
| `Code.gs` | **Archivo único** con todo el código (8 secciones) |
| `appsscript.json` | Manifiesto del proyecto Apps Script |
| `.clasp.json` | Configuración para deploy con clasp CLI |

### Secciones dentro de `Code.gs`

| Sección | Contenido |
|---|---|
| 1 · Constantes | Variables globales: columnas, grados, colores, listas, URLs Kobo |
| 2 · Menú | `onOpen()` y menú "DP Educación" (incluye submenú KoboToolbox) |
| 3 · Hoja Interés | `setupHojaInteres()`: encabezados, validaciones, formato |
| 4 · Papelería | Sidebar HTML para multi-selección de documentos |
| 5 · Hojas de Grado | `crearHojaGrado()`: estructura, formato y validaciones |
| 6 · Transferencia | `onEdit()` y lógica de movimiento de alumnos |
| 7 · Triggers | `installTriggers()` / `removeTriggers()` |
| 8 · KoboToolbox | Descarga CSV, filtro educación, importación y sync automático |
| 9 · Seguimiento | Hoja de seguimiento post-graduación (Quinto Bachillerato) |

---

## Integración KoboToolbox

Dos fuentes de datos se importan directamente desde KoboToolbox vía API:

| Fuente | Asset ID | Hoja destino | Frecuencia |
|---|---|---|---|
| Datos actuales (2025+) | `auvEELWQEgiwF54W4pGpV5` | `Kobo: Interés Actual` | Cada hora (automático) |
| Histórico (hasta 2024) | `akz5K2bGfvvisQaE7VaHev` | `Kobo: Histórico` | Una sola vez |

### Pasos de configuración

1. **Token**: menú → 🌐 KoboToolbox → 🔑 Configurar token de API
2. **Ver columnas**: menú → 🌐 KoboToolbox → 🔍 Ver columnas disponibles
3. **Ajustar mapeo**: edita `KOBO_MAP` en la Sección 1 de `Code.gs` con los nombres exactos de los campos
4. **Importar**: menú → 🌐 KoboToolbox → 🔄 Actualizar datos actuales
5. **Histórico**: menú → 🌐 KoboToolbox → 📥 Importar histórico *(solo la 1ª vez)*
6. **Auto-sync**: menú → 🌐 KoboToolbox → 🔁 Sync automático (cada hora)

> **Token API**: se obtiene en `kf.kobotoolbox.org` → ícono de usuario → API Key.
> Se guarda en **Script Properties** (no en el código ni en el repositorio).

---

## Flujo de Trabajo

```
Hoja "Interés"
  │
  │  Llenar datos del estudiante
  │  Seleccionar Acción → "Enviar a: [Grado]"
  │
  ▼
[Trigger onEdit dispara automáticamente]
  │
  ▼
Hoja "[Grado] [Año]"  (ej. "Primero Básico 2026")
  │  ID auto-generado
  │  Datos copiados
  │  Modalidad y Estado con desplegables
  ▼
Fila en Interés marcada como ✅
```

---

## Hoja "Interés" — Columnas

| Columna | Tipo | Descripción |
|---|---|---|
| Nombre Completo | Texto | Nombre del aspirante |
| Edad | Número (5-99) | Edad del estudiante |
| DPI / CUI | Texto | Documento de identidad |
| Último Año Cursado | Desplegable | Último grado aprobado |
| Papelería Faltante | Texto / Multi-select | Documentos pendientes |
| Comentario | Texto | Observaciones generales |
| Acción | **Desplegable** | Seleccionar grado para transferir |

---

## Hojas de Grado — Columnas

| Columna | Tipo | Descripción |
|---|---|---|
| ID | Auto (Gx-NNN) | Identificador único del alumno |
| Nombre Completo | Texto | Copiado desde Interés |
| DPI / CUI | Texto | Copiado desde Interés |
| No. Teléfono | Texto | Se llena en esta hoja |
| Edad | Número | Copiado desde Interés |
| Grado | Desplegable | Pre-llenado según selección |
| Modalidad | Desplegable | Presencial / Semi-presencial |
| Estado | **Desplegable** | Oyente / Deserción / Graduando |

**Formato condicional por estado:**
- 🔴 Deserción → fondo rojo claro
- 🟢 Graduando → fondo verde claro
- 🟡 Oyente → fondo amarillo claro

---

## Grados Disponibles

| # | Nombre | Hoja creada |
|---|---|---|
| 1 | Primera Etapa de Primaria | `Primera Etapa de Primaria 2026` |
| 2 | Segunda Etapa de Primaria | `Segunda Etapa de Primaria 2026` |
| 3 | Primera Etapa de Básicos | `Primera Etapa de Básicos 2026` |
| 4 | Segunda Etapa de Básicos | `Segunda Etapa de Básicos 2026` |
| 5 | Cuarto Bachillerato | `Cuarto Bachillerato 2026` |
| 6 | **Quinto Bachillerato** | `Quinto Bachillerato 2026` ← genera graduados |

Al marcar un alumno de **Quinto Bachillerato** como **Graduando**, se ofrece registrarlo automáticamente en la hoja **Seguimiento Graduados**.

## Hoja de Seguimiento de Graduados

| Columna | Descripción |
|---|---|
| ID | SEG-001, SEG-002… |
| Nombre Completo | Copiado desde la hoja de Quinto Bachillerato |
| DPI / CUI | Copiado automáticamente |
| No. Teléfono | Copiado automáticamente |
| Edad | Copiado automáticamente |
| Año Graduación | Año en que se marcó como Graduando |
| Estado Post-Grad | Empleado / Continúa estudiando / Sin contacto / Emigró |
| Observaciones | Nota con ID de origen |

**Colores por estado:**
- 🟢 Empleado → verde
- 🔵 Continúa estudiando → azul claro
- 🟡 Sin contacto → amarillo
- 🟣 Emigró → morado claro

---

## Instalación Paso a Paso

### 1. Crear el Google Spreadsheet
1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja nueva.

### 2. Abrir el Editor de Apps Script
1. Menú **Extensiones → Apps Script**.
2. Borra el código de `Código.gs` por defecto.

### 3. Copiar el código
Abre el archivo **`Code.gs`** de este repositorio y copia **todo** su contenido
en el editor de Apps Script.

O usa **clasp** (herramienta de línea de comandos):
```bash
npm install -g @google/clasp
clasp login
# Edita .clasp.json con tu scriptId real
clasp push
```

### 4. Guardar y recargar la hoja
Al recargar, aparecerá el menú **DP Educación**.

### 5. Configurar la hoja Interés
```
DP Educación → ⚙️ Configurar hoja Interés
```

### 6. Instalar el trigger automático
```
DP Educación → (Ejecutar desde el editor) → installTriggers()
```
> Solo se hace una vez. Esto activa el proceso automático al cambiar la columna "Acción".

### 7. Crear las hojas de grado
```
DP Educación → 📚 Crear hoja de grado → [seleccionar grado]
```
O bien: **Crear TODOS los grados** para crearlos todos de una vez.

---

## Uso de Papelería Faltante (Multi-selección)

La columna **Papelería Faltante** acepta texto libre separado por coma.

Documentos disponibles:
- Partida de nacimiento
- DPI / CUI
- Constancia de notas
- Foto reciente
- Paz y salvo
- Formulario de inscripción
- Certificado médico

---

## Menú Completo

```
DP Educación
  ├── ⚙️  Configurar hoja Interés
  ├── 🔧  Instalar trigger automático
  ├── 📚 Crear hoja de grado
  │     ├── Primera Etapa de Primaria
  │     ├── Segunda Etapa de Primaria
  │     ├── Primera Etapa de Básicos
  │     ├── Segunda Etapa de Básicos
  │     ├── Cuarto Bachillerato
  │     ├── Quinto Bachillerato
  │     └── ✨ Crear TODOS los grados
  ├── 📋 Seleccionar papelería faltante
  ├── 🔄 Procesar acciones pendientes
  ├── 📊 Ver resumen de alumnos
  ├── 🎓 Configurar hoja Seguimiento
  ├── 🎓 Registrar graduado manualmente
  └── 🌐 KoboToolbox
        ├── 🔑 Configurar token de API
        ├── 🗺️ Ver hoja de mapeo (actuales)
        ├── 🗺️ Ver hoja de mapeo (histórico)
        ├── 🔄 Sync a hoja Interés (solo educación)
        ├── 📥 Importar histórico (una sola vez)
        ├── 🔁 Sync automático (cada hora)
        └── ⛔ Detener sync automático
```
