// ============================================================
//  TRANSFERENCIA DE ESTUDIANTES: Interés → Hoja de Grado
// ============================================================

/**
 * Trigger onEdit: detecta cambios en la columna "Acción" de la hoja Interés
 * y transfiere automáticamente al estudiante a la hoja de grado correspondiente.
 *
 * Para activar este trigger de forma automática, ejecuta installTriggers().
 */
function onEdit(e) {
  const range = e.range;
  const hoja  = range.getSheet();

  // Solo actuar en la hoja Interés, columna Acción, fila >= 2
  if (hoja.getName() !== HOJA_INTERES)       return;
  if (range.getColumn() !== COL_INTERES.ACCION) return;
  if (range.getRow() < 2)                    return;

  const valor = e.value || '';
  if (!valor || valor === '-- Seleccionar --') return;

  // Extraer el grado del valor "Enviar a: Primero Básico"
  const grado = valor.replace('Enviar a: ', '').trim();
  if (!GRADOS.includes(grado)) return;

  transferirEstudiante(range.getRow(), grado);
}

/**
 * Lee los datos de la fila indicada en la hoja Interés y los copia
 * a la hoja del grado correspondiente.
 *
 * @param {number} fila  - Número de fila en la hoja Interés (base 1)
 * @param {string} grado - Nombre del grado de destino
 */
function transferirEstudiante(fila, grado) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInteres = ss.getSheetByName(HOJA_INTERES);

  // Leer datos del estudiante
  const datos = hojaInteres.getRange(fila, 1, 1, 7).getValues()[0];
  const nombre      = datos[COL_INTERES.NOMBRE - 1];
  const edad        = datos[COL_INTERES.EDAD - 1];
  const dpi         = datos[COL_INTERES.DPI - 1];
  const ultimoAnio  = datos[COL_INTERES.ULTIMO_ANIO - 1];
  const papeleria   = datos[COL_INTERES.PAPELERIA - 1];
  const comentario  = datos[COL_INTERES.COMENTARIO - 1];

  // Validar que al menos haya nombre
  if (!nombre) {
    SpreadsheetApp.getUi().alert('⚠️ La fila seleccionada no tiene nombre. No se realizó la transferencia.');
    // Resetear la celda de acción
    hojaInteres.getRange(fila, COL_INTERES.ACCION).setValue('-- Seleccionar --');
    return;
  }

  // Obtener o crear la hoja de grado
  const nombreHoja  = _nombreHoja(grado);
  let   hojaGrado   = ss.getSheetByName(nombreHoja);

  if (!hojaGrado) {
    hojaGrado = crearHojaGrado(grado);
  }

  // Generar ID único
  const id = _siguienteId(hojaGrado, grado);

  // Encontrar la primera fila vacía (desde fila 3)
  const ultimaFila  = hojaGrado.getLastRow();
  const filaDestino = Math.max(ultimaFila + 1, 3);

  // Escribir datos en la hoja de grado
  hojaGrado.getRange(filaDestino, 1, 1, 8).setValues([[
    id,           // ID
    nombre,       // Nombre Completo
    dpi,          // DPI / CUI
    '',           // Teléfono (no está en la hoja Interés, se llena luego)
    edad,         // Edad
    grado,        // Grado (prellenado)
    'Presencial', // Modalidad (valor por defecto)
    'Oyente'      // Estado (valor por defecto)
  ]]);

  // Formato de la nueva fila
  hojaGrado.getRange(filaDestino, 1, 1, 8)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  // Nombre y DPI alineados a la izquierda
  hojaGrado.getRange(filaDestino, COL_GRADO.NOMBRE)
    .setHorizontalAlignment('left');
  hojaGrado.setRowHeight(filaDestino, 26);

  // Agregar nota con la papelería y comentario
  if (papeleria || comentario) {
    let nota = '';
    if (papeleria)  nota += `Papelería faltante: ${papeleria}\n`;
    if (comentario) nota += `Comentario: ${comentario}`;
    hojaGrado.getRange(filaDestino, COL_GRADO.NOMBRE).setNote(nota.trim());
  }

  // Marcar la fila de Interés como transferida (fondo verde claro)
  hojaInteres.getRange(fila, 1, 1, 7)
    .setBackground('#E8F5E9');

  // Actualizar la celda de acción con un sello de completado
  hojaInteres.getRange(fila, COL_INTERES.ACCION)
    .setValue(`✅ ${grado}`)
    .setDataValidation(null);  // quitar la validación para que no se pueda cambiar

  // Mensaje de confirmación
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `"${nombre}" fue agregado a "${nombreHoja}" con ID ${id}`,
    '✅ Estudiante transferido',
    5
  );
}

// ─────────────────────────────────────────────────────────────
//  PROCESO MASIVO (para ejecutar manualmente si hay pendientes)
// ─────────────────────────────────────────────────────────────

/**
 * Recorre todas las filas de la hoja Interés y procesa las acciones
 * que aún tengan un grado seleccionado (no transferidas todavía).
 * Útil si el trigger onEdit no procesó alguna fila.
 */
function procesarAccionesPendientes() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInteres = ss.getSheetByName(HOJA_INTERES);

  if (!hojaInteres) {
    SpreadsheetApp.getUi().alert('❌ La hoja "Interés" no existe. Ejecuta primero "Configurar hoja Interés".');
    return;
  }

  const ultimaFila = hojaInteres.getLastRow();
  if (ultimaFila < 2) {
    SpreadsheetApp.getUi().alert('No hay filas de datos en la hoja Interés.');
    return;
  }

  const acciones = hojaInteres
    .getRange(2, COL_INTERES.ACCION, ultimaFila - 1, 1)
    .getValues();

  let procesados = 0;

  acciones.forEach((row, idx) => {
    const valor = (row[0] || '').toString().trim();
    if (!valor || valor === '-- Seleccionar --' || valor.startsWith('✅')) return;

    const grado = valor.replace('Enviar a: ', '').trim();
    if (!GRADOS.includes(grado)) return;

    transferirEstudiante(idx + 2, grado);
    procesados++;
  });

  if (procesados === 0) {
    SpreadsheetApp.getUi().alert('No hay acciones pendientes por procesar.');
  } else {
    SpreadsheetApp.getUi().alert(`✅ Se procesaron ${procesados} estudiante(s) correctamente.`);
  }
}

// ─────────────────────────────────────────────────────────────
//  RESUMEN GENERAL
// ─────────────────────────────────────────────────────────────

/**
 * Muestra un cuadro de diálogo con el conteo de alumnos por grado.
 */
function mostrarResumen() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const anio  = new Date().getFullYear();
  let   texto = `📊 RESUMEN DE ALUMNOS — ${anio}\n${'─'.repeat(40)}\n`;

  GRADOS.forEach(grado => {
    const nombreH = _nombreHoja(grado);
    const hoja    = ss.getSheetByName(nombreH);
    if (!hoja) {
      texto += `${grado}: hoja no creada\n`;
      return;
    }

    const ultimaFila = hoja.getLastRow();
    const totalAlumnos = Math.max(ultimaFila - 2, 0);

    // Contar por estado
    let oyentes = 0, deserciones = 0, graduandos = 0;
    if (totalAlumnos > 0) {
      const estados = hoja.getRange(3, COL_GRADO.ESTADO, totalAlumnos, 1).getValues().flat();
      estados.forEach(e => {
        if (e === 'Oyente')    oyentes++;
        if (e === 'Deserción') deserciones++;
        if (e === 'Graduando') graduandos++;
      });
    }

    texto += `\n${grado} (${totalAlumnos} alumnos)\n`;
    texto += `  Oyentes: ${oyentes}  |  Deserciones: ${deserciones}  |  Graduandos: ${graduandos}\n`;
  });

  SpreadsheetApp.getUi().alert('Resumen de Alumnos', texto, SpreadsheetApp.getUi().ButtonSet.OK);
}
