# DP Educación — Automatización Google Sheets

Sistema de gestión de inscripciones y salones de clase implementado con **Google Apps Script**.

---

## Estructura de Archivos

| Archivo | Descripción |
|---|---|
| `Constantes.gs` | Variables globales: nombres de columnas, listas de grados, colores |
| `Menu.gs` | Menú personalizado "DP Educación" y wrappers de funciones |
| `HojaInteres.gs` | Configuración de la hoja principal "Interés" |
| `HojasGrado.gs` | Creación y formato de las hojas de cada grado |
| `Transferencia.gs` | Lógica de transferencia de estudiantes + resumen |
| `Triggers.gs` | Instalación del trigger `onEdit` instalable |
| `appsscript.json` | Manifiesto del proyecto Apps Script |

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

1. Primero Básico
2. Segundo Básico
3. Tercero Básico
4. Cuarto Bachillerato
5. Quinto Bachillerato
6. Sexto Bachillerato

Cada hoja se crea con el año actual automáticamente:
`Primero Básico 2026`, `Segundo Básico 2026`, etc.

---

## Instalación Paso a Paso

### 1. Crear el Google Spreadsheet
1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja nueva.

### 2. Abrir el Editor de Apps Script
1. Menú **Extensiones → Apps Script**.
2. Borra el código de `Código.gs` por defecto.

### 3. Crear los archivos `.gs`
Crea un archivo por cada `.gs` en este repositorio y copia el contenido.

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
  ├── ⚙️ Configurar hoja Interés
  ├── 📚 Crear hoja de grado
  │     ├── Primero Básico
  │     ├── Segundo Básico
  │     ├── Tercero Básico
  │     ├── Cuarto Bachillerato
  │     ├── Quinto Bachillerato
  │     ├── Sexto Bachillerato
  │     └── Crear TODOS los grados
  ├── 🔄 Procesar acciones pendientes
  └── 📋 Ver resumen de alumnos
```
