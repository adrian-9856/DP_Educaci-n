// ============================================================
//  CREACIÓN Y GESTIÓN DE HOJAS DE GRADO
// ============================================================

/**
 * Obtiene el nombre de la hoja para un grado en el año actual.
 * Formato: "Primero Básico 2026"
 */
function _nombreHoja(grado) {
  const anio = new Date().getFullYear();
  return `${grado} ${anio}`;
}

/**
 * Crea (o reconfigura) la hoja de un grado específico.
 * Si ya existe la hoja, pregunta si desea reiniciarla.
 *
 * @param {string} grado - Nombre del grado, p.ej. "Primero Básico"
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} La hoja creada/existente.
 */
function crearHojaGrado(grado) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const nombreH   = _nombreHoja(grado);
  let   hoja      = ss.getSheetByName(nombreH);
  const ui        = SpreadsheetApp.getUi();

  if (hoja) {
    const resp = ui.alert(
      'Hoja existente',
      `La hoja "${nombreH}" ya existe.\n¿Deseas reiniciar su formato sin borrar los datos?`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return hoja;
    // Solo reformatear, no borrar datos
    _formatearHojaGrado(hoja, grado);
    ui.alert(`✅ Hoja "${nombreH}" reformateada.`);
    return hoja;
  }

  // Crear hoja nueva al final
  hoja = ss.insertSheet(nombreH);
  _formatearHojaGrado(hoja, grado);

  ui.alert(`✅ Hoja "${nombreH}" creada correctamente.\n\nAño: ${new Date().getFullYear()}`);
  return hoja;
}

/**
 * Aplica encabezados, formato y validaciones a una hoja de grado.
 */
function _formatearHojaGrado(hoja, grado) {
  const anio = new Date().getFullYear();

  // ── Título en fila 1 ─────────────────────────────────────
  hoja.getRange(1, 1, 1, 8).merge()
    .setValue(`${grado.toUpperCase()}  ·  Año ${anio}`)
    .setBackground(COLOR_HEADER_GRADO)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 40);

  // ── Encabezados en fila 2 ────────────────────────────────
  const headers = [
    'ID',
    'Nombre Completo',
    'DPI / CUI',
    'No. Teléfono',
    'Edad',
    'Grado',
    'Modalidad',
    'Estado'
  ];

  const rH = hoja.getRange(2, 1, 1, headers.length);
  rH.setValues([headers]);
  rH.setBackground(COLOR_HEADER_GRADO)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 32);

  hoja.setFrozenRows(2);

  // ── Anchos de columna ────────────────────────────────────
  hoja.setColumnWidth(COL_GRADO.ID,        70);
  hoja.setColumnWidth(COL_GRADO.NOMBRE,   220);
  hoja.setColumnWidth(COL_GRADO.DPI,      140);
  hoja.setColumnWidth(COL_GRADO.TELEFONO, 130);
  hoja.setColumnWidth(COL_GRADO.EDAD,      60);
  hoja.setColumnWidth(COL_GRADO.GRADO,    160);
  hoja.setColumnWidth(COL_GRADO.MODALIDAD,140);
  hoja.setColumnWidth(COL_GRADO.ESTADO,   130);

  // ── Validaciones para filas de datos (3-300) ─────────────
  const maxRows = 300;

  // Grado: lista desplegable (pre-llenado pero editable)
  const validGrado = SpreadsheetApp.newDataValidation()
    .requireValueInList(GRADOS, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona el grado')
    .build();
  hoja.getRange(3, COL_GRADO.GRADO, maxRows, 1).setDataValidation(validGrado);

  // Modalidad: Presencial / Semi-presencial
  const validModalidad = SpreadsheetApp.newDataValidation()
    .requireValueInList(MODALIDADES, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona la modalidad')
    .build();
  hoja.getRange(3, COL_GRADO.MODALIDAD, maxRows, 1).setDataValidation(validModalidad);

  // Estado: Oyente / Deserción / Graduando
  const validEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(ESTADOS, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona el estado del estudiante')
    .build();
  hoja.getRange(3, COL_GRADO.ESTADO, maxRows, 1).setDataValidation(validEstado);

  // Edad: número
  const validEdad = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(5, 99)
    .setAllowInvalid(false)
    .build();
  hoja.getRange(3, COL_GRADO.EDAD, maxRows, 1).setDataValidation(validEdad);

  // ── Formato condicional por estado ──────────────────────
  hoja.clearConditionalFormatRules();
  const estadoRange = hoja.getRange(3, 1, maxRows, 8);

  const reglaDesercion = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$H3="Deserción"`)
    .setBackground('#FFCDD2')   // rojo claro
    .setRanges([estadoRange])
    .build();

  const reglaGraduando = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$H3="Graduando"`)
    .setBackground('#C8E6C9')   // verde claro
    .setRanges([estadoRange])
    .build();

  const reglaOyente = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$H3="Oyente"`)
    .setBackground('#FFF9C4')   // amarillo claro
    .setRanges([estadoRange])
    .build();

  hoja.setConditionalFormatRules([reglaDesercion, reglaGraduando, reglaOyente]);

  // ── Fórmula de ID automático ─────────────────────────────
  // Se usará al insertar filas; el ID se genera en el script de transferencia.
  // Aquí mostramos una nota en la columna ID:
  hoja.getRange(2, COL_GRADO.ID)
    .setNote('ID generado automáticamente al agregar un estudiante.');

  // ── Nota del año en la celda de título ──────────────────
  hoja.getRange(1, 1)
    .setNote(`Hoja creada el: ${new Date().toLocaleDateString('es-GT')}\nAño escolar: ${anio}`);
}

/**
 * Obtiene el próximo ID correlativo dentro de una hoja de grado.
 * Formato: G1-001, G2-001, etc. según el índice del grado.
 */
function _siguienteId(hoja, grado) {
  const indiceGrado = GRADOS.indexOf(grado) + 1;
  const prefijo     = `G${indiceGrado}-`;

  // Contar filas con datos (desde fila 3)
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 3) return `${prefijo}001`;

  const ids = hoja.getRange(3, COL_GRADO.ID, ultimaFila - 2, 1)
    .getValues()
    .flat()
    .filter(v => v !== '');

  if (ids.length === 0) return `${prefijo}001`;

  // Extraer el número más alto
  const numeros = ids
    .map(id => parseInt(id.toString().replace(prefijo, ''), 10))
    .filter(n => !isNaN(n));

  const siguiente = Math.max(...numeros) + 1;
  return `${prefijo}${String(siguiente).padStart(3, '0')}`;
}
