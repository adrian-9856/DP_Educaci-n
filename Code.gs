// ════════════════════════════════════════════════════════════════════════════
//  DP EDUCACIÓN — Automatización Google Sheets
//  Archivo único: todo el flujo en un mismo Spreadsheet con varias pestañas
// ════════════════════════════════════════════════════════════════════════════

// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 1 · CONSTANTES GLOBALES
// ────────────────────────────────────────────────────────────────────────────

const HOJA_INTERES = 'Interés';

// Columnas de la hoja Interés (1-based)
const COL_INTERES = {
  NOMBRE:      1,
  EDAD:        2,
  DPI:         3,
  ULTIMO_ANIO: 4,
  PAPELERIA:   5,
  COMENTARIO:  6,
  ACCION:      7
};

// Columnas de cada hoja de grado (1-based)
const COL_GRADO = {
  ID:        1,
  NOMBRE:    2,
  DPI:       3,
  TELEFONO:  4,
  EDAD:      5,
  GRADO:     6,
  MODALIDAD: 7,
  ESTADO:    8
};

const GRADOS = [
  'Primero Básico',
  'Segundo Básico',
  'Tercero Básico',
  'Cuarto Bachillerato',
  'Quinto Bachillerato',
  'Sexto Bachillerato'
];

const PAPELERIA_OPCIONES = [
  'Partida de nacimiento',
  'DPI / CUI',
  'Constancia de notas',
  'Foto reciente',
  'Paz y salvo',
  'Formulario de inscripción',
  'Certificado médico'
];

const MODALIDADES          = ['Presencial', 'Semi-presencial'];
const ESTADOS              = ['Oyente', 'Deserción', 'Graduando'];
const ACCIONES             = ['-- Seleccionar --', ...GRADOS.map(g => 'Enviar a: ' + g)];
const ULTIMO_ANIO_OPCIONES = [
  'Primaria',
  'Primero Básico',
  'Segundo Básico',
  'Tercero Básico',
  'Cuarto Bachillerato',
  'Quinto Bachillerato'
];

const COLOR_HEADER_INTERES = '#1565C0';
const COLOR_HEADER_GRADO   = '#2E7D32';
const COLOR_FONT_HEADER    = '#FFFFFF';

// ── KoboToolbox ──────────────────────────────────────────────────────────────
// URL pública de exportación CSV (requiere token en Script Properties)
const KOBO_URL_ACTUAL    = 'https://kf.kobotoolbox.org/api/v2/assets/auvEELWQEgiwF54W4pGpV5/export-settings/eseYzEgWw6Tui9y2eppZy3L/data.csv';
const KOBO_URL_HISTORICO = 'https://kf.kobotoolbox.org/api/v2/assets/akz5K2bGfvvisQaE7VaHev/export-settings/esuV4RKqQhYUUaUizfWBP8S/data.csv';

// Nombre de las hojas destino
const HOJA_KOBO_ACTUAL    = 'Kobo: Interés Actual';
const HOJA_KOBO_HISTORICO = 'Kobo: Histórico';

// Colores encabezado Kobo
const COLOR_HEADER_KOBO = '#6A1B9A';

// Clave en Script Properties donde se guarda el token de Kobo
const PROP_KOBO_TOKEN = 'KOBO_API_TOKEN';

// Mapeo de columnas KoboToolbox → columnas de la hoja Interés
// AJUSTA los valores izquierdos (nombre exacto del campo en el CSV de Kobo)
// después de ejecutar "🔍 Ver columnas disponibles"
const KOBO_MAP = {
  NOMBRE:      'nombre_completo',   // ← nombre del campo Kobo
  EDAD:        'edad',
  DPI:         'dpi_cui',
  ULTIMO_ANIO: 'ultimo_anio_cursado',
  PAPELERIA:   'papeleria_faltante',
  COMENTARIO:  'comentario'
};


// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 2 · MENÚ Y PUNTO DE ENTRADA
// ────────────────────────────────────────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('DP Educación')
    .addItem('⚙️  Configurar hoja Interés',      'setupHojaInteres')
    .addItem('🔧  Instalar trigger automático',   'installTriggers')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('📚 Crear hoja de grado')
        .addItem('Primero Básico',           'crearHojaPrimeroBasico')
        .addItem('Segundo Básico',           'crearHojaSegundoBasico')
        .addItem('Tercero Básico',           'crearHojaTerceroBasico')
        .addItem('Cuarto Bachillerato',      'crearHojaCuartoBachillerato')
        .addItem('Quinto Bachillerato',      'crearHojaQuintoBachillerato')
        .addItem('Sexto Bachillerato',       'crearHojaSextoBachillerato')
        .addSeparator()
        .addItem('✨ Crear TODOS los grados', 'crearTodasLasHojas')
    )
    .addSeparator()
    .addItem('📋 Seleccionar papelería faltante', 'abrirSelectorPapeleria')
    .addItem('🔄 Procesar acciones pendientes',   'procesarAccionesPendientes')
    .addItem('📊 Ver resumen de alumnos',          'mostrarResumen')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('🌐 KoboToolbox')
        .addItem('🔑 Configurar token de API',           'koboConfigurarToken')
        .addSeparator()
        .addItem('🔍 Ver columnas disponibles (actual)',  'koboDescubrirColumnasActual')
        .addItem('🔍 Ver columnas disponibles (hist.)',   'koboDescubrirColumnasHistorico')
        .addSeparator()
        .addItem('🔄 Actualizar datos actuales (2025+)',  'koboImportarActual')
        .addItem('📥 Importar histórico (una sola vez)',  'koboImportarHistorico')
        .addSeparator()
        .addItem('🔁 Sync automático (cada hora)',        'koboInstalarTriggerSync')
        .addItem('⛔ Detener sync automático',            'koboEliminarTriggerSync')
    )
    .addToUi();
}

function crearHojaPrimeroBasico()      { crearHojaGrado('Primero Básico'); }
function crearHojaSegundoBasico()      { crearHojaGrado('Segundo Básico'); }
function crearHojaTerceroBasico()      { crearHojaGrado('Tercero Básico'); }
function crearHojaCuartoBachillerato() { crearHojaGrado('Cuarto Bachillerato'); }
function crearHojaQuintoBachillerato() { crearHojaGrado('Quinto Bachillerato'); }
function crearHojaSextoBachillerato()  { crearHojaGrado('Sexto Bachillerato'); }

function crearTodasLasHojas() {
  GRADOS.forEach(g => crearHojaGrado(g, true));
  SpreadsheetApp.getUi().alert('✅ Todas las hojas de grado fueron creadas correctamente.');
}


// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 3 · CONFIGURACIÓN DE LA HOJA "INTERÉS"
// ────────────────────────────────────────────────────────────────────────────

function setupHojaInteres() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   hoja = ss.getSheetByName(HOJA_INTERES);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_INTERES);
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(1);
  }

  hoja.clearFormats();
  hoja.clearConditionalFormatRules();

  // Encabezados fila 1
  const headers = [
    'Nombre Completo', 'Edad', 'DPI / CUI',
    'Último Año Cursado', 'Papelería Faltante', 'Comentario', 'Acción'
  ];
  hoja.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLOR_HEADER_INTERES)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 36);

  // Anchos
  hoja.setColumnWidth(COL_INTERES.NOMBRE,      220);
  hoja.setColumnWidth(COL_INTERES.EDAD,         60);
  hoja.setColumnWidth(COL_INTERES.DPI,         140);
  hoja.setColumnWidth(COL_INTERES.ULTIMO_ANIO, 160);
  hoja.setColumnWidth(COL_INTERES.PAPELERIA,   260);
  hoja.setColumnWidth(COL_INTERES.COMENTARIO,  220);
  hoja.setColumnWidth(COL_INTERES.ACCION,      200);

  const MAX = 200;

  // Validación Edad
  hoja.getRange(2, COL_INTERES.EDAD, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(5, 99).setAllowInvalid(false)
      .setHelpText('Edad entre 5 y 99').build()
  );

  // Validación Último año cursado
  hoja.getRange(2, COL_INTERES.ULTIMO_ANIO, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(ULTIMO_ANIO_OPCIONES, true).setAllowInvalid(false)
      .setHelpText('Último año cursado').build()
  );

  // Validación Acción (grado destino)
  hoja.getRange(2, COL_INTERES.ACCION, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(ACCIONES, true).setAllowInvalid(false)
      .setHelpText('Selecciona el grado al que enviar al estudiante').build()
  );

  // Fondo azul claro en Papelería
  hoja.getRange(2, COL_INTERES.PAPELERIA, MAX, 1).setBackground('#E3F2FD');

  // Formato condicional: resaltar acción pendiente
  hoja.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(G2<>"",G2<>"-- Seleccionar --",LEFT(G2,1)<>"✅")')
      .setBackground('#FFF9C4')
      .setRanges([hoja.getRange(2, COL_INTERES.ACCION, MAX, 1)])
      .build()
  ]);

  // Nota de ayuda en encabezado Papelería
  hoja.getRange(1, COL_INTERES.PAPELERIA).setNote(
    'CAMPO MULTI-SELECCIÓN\n\n' +
    'Selecciona la celda y usa:\n' +
    'DP Educación → Seleccionar papelería faltante\n\n' +
    'O escribe manualmente separando con coma:\n' +
    PAPELERIA_OPCIONES.map((p, i) => (i+1) + '. ' + p).join('\n')
  );

  SpreadsheetApp.getUi().alert(
    '✅ Hoja "Interés" configurada.\n\n' +
    'Pasos siguientes:\n' +
    '1. Crea los salones: menú → 📚 Crear hoja de grado\n' +
    '2. Instala el trigger: menú → 🔧 Instalar trigger automático\n' +
    '3. Llena datos y selecciona Acción para transferir alumnos.'
  );
}


// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 4 · SIDEBAR DE PAPELERÍA (multi-selección via HTML)
// ────────────────────────────────────────────────────────────────────────────

function abrirSelectorPapeleria() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getActiveSheet();
  const celda = hoja.getActiveCell();

  if (hoja.getName() !== HOJA_INTERES) {
    SpreadsheetApp.getUi().alert('Navega a la hoja "Interés" primero.');
    return;
  }
  if (celda.getColumn() !== COL_INTERES.PAPELERIA) {
    SpreadsheetApp.getUi().alert('Selecciona una celda de la columna "Papelería Faltante" (columna E).');
    return;
  }

  const actual       = celda.getValue().toString();
  const seleccionados = actual ? actual.split(',').map(s => s.trim()) : [];

  ss.show(
    HtmlService.createHtmlOutput(_htmlPapeleria(seleccionados))
      .setTitle('Papelería Faltante')
      .setWidth(330)
  );
}

function _htmlPapeleria(seleccionados) {
  const items = PAPELERIA_OPCIONES.map(function(op) {
    const chk = seleccionados.indexOf(op) >= 0 ? 'checked' : '';
    return '<label class="item"><input type="checkbox" value="' + op + '" ' + chk + '> ' + op + '</label>';
  }).join('');

  return '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    '*{box-sizing:border-box}' +
    'body{font-family:Arial,sans-serif;padding:18px;margin:0}' +
    'h3{color:#1565C0;margin:0 0 6px;font-size:15px}' +
    'p{color:#666;font-size:12px;margin:0 0 12px}' +
    '.item{display:block;padding:5px 0;font-size:13px;cursor:pointer}' +
    '.item:hover{color:#1565C0}' +
    '.btns{margin-top:16px;display:flex;gap:8px}' +
    'button{flex:1;padding:9px 0;border:none;border-radius:5px;font-size:13px;cursor:pointer;font-weight:bold}' +
    '.ok{background:#1565C0;color:#fff}.ok:hover{background:#0D47A1}' +
    '.cl{background:#e0e0e0;color:#333}.cl:hover{background:#bdbdbd}' +
    '</style></head><body>' +
    '<h3>📋 Papelería Faltante</h3>' +
    '<p>Marca los documentos que faltan:</p>' +
    items +
    '<div class="btns">' +
    '<button class="ok" onclick="guardar()">✅ Guardar</button>' +
    '<button class="cl" onclick="google.script.host.close()">Cancelar</button>' +
    '</div>' +
    '<script>' +
    'function guardar(){' +
    'var vals=[...document.querySelectorAll("input:checked")].map(function(c){return c.value;});' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();}).guardarPapeleria(vals.join(", "));' +
    '}' +
    '</script></body></html>';
}

function guardarPapeleria(valor) {
  SpreadsheetApp.getActiveSheet().getActiveCell().setValue(valor);
}


// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 5 · CREACIÓN Y FORMATO DE HOJAS DE GRADO
// ────────────────────────────────────────────────────────────────────────────

function _nombreHoja(grado) {
  return grado + ' ' + new Date().getFullYear();
}

function crearHojaGrado(grado, silencioso) {
  if (silencioso === undefined) silencioso = false;
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const nombre = _nombreHoja(grado);
  let   hoja   = ss.getSheetByName(nombre);
  const ui     = SpreadsheetApp.getUi();

  if (hoja) {
    if (!silencioso) {
      const r = ui.alert('Hoja existente',
        '"' + nombre + '" ya existe.\n¿Deseas reformatear sin borrar datos?',
        ui.ButtonSet.YES_NO);
      if (r !== ui.Button.YES) return hoja;
    }
    _formatearHojaGrado(hoja, grado);
    if (!silencioso) ui.alert('✅ "' + nombre + '" reformateada.');
    return hoja;
  }

  hoja = ss.insertSheet(nombre);
  _formatearHojaGrado(hoja, grado);
  if (!silencioso) ui.alert('✅ Hoja "' + nombre + '" creada.\n\nAño: ' + new Date().getFullYear());
  return hoja;
}

function _formatearHojaGrado(hoja, grado) {
  const anio = new Date().getFullYear();
  const MAX  = 300;

  // Fila 1: título
  hoja.getRange(1, 1, 1, 8).merge()
    .setValue(grado.toUpperCase() + '  ·  AÑO ' + anio)
    .setBackground(COLOR_HEADER_GRADO)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 42);
  hoja.getRange(1, 1).setNote('Creada: ' + new Date().toLocaleDateString('es-GT') + '\nAño escolar: ' + anio);

  // Fila 2: encabezados
  hoja.getRange(2, 1, 1, 8)
    .setValues([['ID','Nombre Completo','DPI / CUI','No. Teléfono','Edad','Grado','Modalidad','Estado']])
    .setBackground(COLOR_HEADER_GRADO)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 30);
  hoja.setFrozenRows(2);

  // Anchos
  hoja.setColumnWidth(COL_GRADO.ID,        70);
  hoja.setColumnWidth(COL_GRADO.NOMBRE,   220);
  hoja.setColumnWidth(COL_GRADO.DPI,      140);
  hoja.setColumnWidth(COL_GRADO.TELEFONO, 130);
  hoja.setColumnWidth(COL_GRADO.EDAD,      60);
  hoja.setColumnWidth(COL_GRADO.GRADO,    160);
  hoja.setColumnWidth(COL_GRADO.MODALIDAD,145);
  hoja.setColumnWidth(COL_GRADO.ESTADO,   130);

  // Validaciones
  hoja.getRange(3, COL_GRADO.GRADO, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(GRADOS, true)
      .setAllowInvalid(false).setHelpText('Grado del estudiante').build());

  hoja.getRange(3, COL_GRADO.MODALIDAD, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MODALIDADES, true)
      .setAllowInvalid(false).setHelpText('Presencial o Semi-presencial').build());

  hoja.getRange(3, COL_GRADO.ESTADO, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true)
      .setAllowInvalid(false).setHelpText('Estado del estudiante').build());

  hoja.getRange(3, COL_GRADO.EDAD, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(5, 99)
      .setAllowInvalid(false).build());

  // Formato condicional por estado
  hoja.clearConditionalFormatRules();
  const dr = hoja.getRange(3, 1, MAX, 8);
  hoja.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H3="Deserción"')
      .setBackground('#FFCDD2').setRanges([dr]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H3="Graduando"')
      .setBackground('#C8E6C9').setRanges([dr]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H3="Oyente"')
      .setBackground('#FFF9C4').setRanges([dr]).build()
  ]);

  hoja.getRange(2, COL_GRADO.ID).setNote('ID auto-generado al transferir desde Interés.');
}

function _siguienteId(hoja, grado) {
  const prefijo    = 'G' + (GRADOS.indexOf(grado) + 1) + '-';
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 3) return prefijo + '001';

  const ids = hoja.getRange(3, COL_GRADO.ID, ultimaFila - 2, 1)
    .getValues().flat().filter(function(v){ return v !== ''; });
  if (!ids.length) return prefijo + '001';

  const nums = ids
    .map(function(id){ return parseInt(id.toString().replace(prefijo, ''), 10); })
    .filter(function(n){ return !isNaN(n); });

  const sig = Math.max.apply(null, nums) + 1;
  return prefijo + String(sig).padStart(3, '0');
}


// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 6 · TRANSFERENCIA: Interés → Hoja de Grado
// ────────────────────────────────────────────────────────────────────────────

function onEdit(e) {
  const range = e.range;
  const hoja  = range.getSheet();

  if (hoja.getName()    !== HOJA_INTERES)       return;
  if (range.getColumn() !== COL_INTERES.ACCION) return;
  if (range.getRow()    < 2)                    return;

  const valor = (e.value || '').toString().trim();
  if (!valor || valor === '-- Seleccionar --')  return;

  const grado = valor.replace('Enviar a: ', '').trim();
  if (GRADOS.indexOf(grado) < 0)                return;

  _transferirEstudiante(range.getRow(), grado);
}

function _transferirEstudiante(fila, grado) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInteres = ss.getSheetByName(HOJA_INTERES);
  const datos       = hojaInteres.getRange(fila, 1, 1, 7).getValues()[0];

  const nombre     = datos[COL_INTERES.NOMBRE - 1];
  const edad       = datos[COL_INTERES.EDAD - 1];
  const dpi        = datos[COL_INTERES.DPI - 1];
  const papeleria  = datos[COL_INTERES.PAPELERIA - 1];
  const comentario = datos[COL_INTERES.COMENTARIO - 1];

  if (!nombre) {
    SpreadsheetApp.getUi().alert('⚠️ La fila no tiene nombre. No se realizó la transferencia.');
    hojaInteres.getRange(fila, COL_INTERES.ACCION).setValue('-- Seleccionar --');
    return;
  }

  const nombreHoja = _nombreHoja(grado);
  let   hojaGrado  = ss.getSheetByName(nombreHoja);
  if (!hojaGrado) hojaGrado = crearHojaGrado(grado, true);

  const filaDestino = Math.max(hojaGrado.getLastRow() + 1, 3);
  const id          = _siguienteId(hojaGrado, grado);

  hojaGrado.getRange(filaDestino, 1, 1, 8).setValues([[
    id, nombre, dpi,
    '',           // Teléfono: se completa en la hoja de grado
    edad, grado,
    'Presencial', // Modalidad por defecto
    'Oyente'      // Estado por defecto
  ]]);

  hojaGrado.getRange(filaDestino, 1, 1, 8)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  hojaGrado.getRange(filaDestino, COL_GRADO.NOMBRE).setHorizontalAlignment('left');
  hojaGrado.getRange(filaDestino, COL_GRADO.DPI).setHorizontalAlignment('left');
  hojaGrado.setRowHeight(filaDestino, 26);

  if (papeleria || comentario) {
    const nota = [
      papeleria  ? 'Papelería faltante: ' + papeleria  : '',
      comentario ? 'Comentario: '         + comentario : ''
    ].filter(Boolean).join('\n');
    hojaGrado.getRange(filaDestino, COL_GRADO.NOMBRE).setNote(nota);
  }

  // Marcar fila origen como procesada
  hojaInteres.getRange(fila, 1, 1, 7).setBackground('#E8F5E9');
  hojaInteres.getRange(fila, COL_INTERES.ACCION)
    .setValue('✅ ' + grado)
    .setDataValidation(null);

  ss.toast('"' + nombre + '" → "' + nombreHoja + '" · ID: ' + id, '✅ Estudiante transferido', 5);
}

function procesarAccionesPendientes() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInteres = ss.getSheetByName(HOJA_INTERES);
  const ui          = SpreadsheetApp.getUi();

  if (!hojaInteres) {
    ui.alert('❌ La hoja "Interés" no existe. Ejecuta primero "Configurar hoja Interés".');
    return;
  }

  const ultimaFila = hojaInteres.getLastRow();
  if (ultimaFila < 2) { ui.alert('No hay datos en la hoja Interés.'); return; }

  const acciones = hojaInteres
    .getRange(2, COL_INTERES.ACCION, ultimaFila - 1, 1).getValues();

  let procesados = 0;
  acciones.forEach(function(row, idx) {
    const v = (row[0] || '').toString().trim();
    if (!v || v === '-- Seleccionar --' || v.charAt(0) === '✅') return;
    const grado = v.replace('Enviar a: ', '').trim();
    if (GRADOS.indexOf(grado) < 0) return;
    _transferirEstudiante(idx + 2, grado);
    procesados++;
  });

  ui.alert(procesados === 0
    ? 'No hay acciones pendientes.'
    : '✅ Se procesaron ' + procesados + ' estudiante(s).');
}

function mostrarResumen() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const anio = new Date().getFullYear();
  let texto  = '📊 RESUMEN DE ALUMNOS — ' + anio + '\n' + '─'.repeat(42) + '\n';

  GRADOS.forEach(function(grado) {
    const hoja = ss.getSheetByName(_nombreHoja(grado));
    if (!hoja) { texto += '\n' + grado + ': pestaña no creada\n'; return; }

    const total = Math.max(hoja.getLastRow() - 2, 0);
    let oyentes = 0, deserc = 0, grad = 0;

    if (total > 0) {
      hoja.getRange(3, COL_GRADO.ESTADO, total, 1).getValues().flat().forEach(function(e) {
        if (e === 'Oyente')    oyentes++;
        if (e === 'Deserción') deserc++;
        if (e === 'Graduando') grad++;
      });
    }

    texto += '\n' + grado + '  (' + total + ' alumnos)\n';
    texto += '   Oyentes: ' + oyentes + '  |  Deserciones: ' + deserc + '  |  Graduandos: ' + grad + '\n';
  });

  SpreadsheetApp.getUi().alert('Resumen de Alumnos', texto, SpreadsheetApp.getUi().ButtonSet.OK);
}


// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 7 · TRIGGERS
// ────────────────────────────────────────────────────────────────────────────

function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers()
    .filter(function(t){ return t.getHandlerFunction() === 'onEdit'; })
    .forEach(function(t){ ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();

  SpreadsheetApp.getUi().alert(
    '✅ Trigger instalado.\n\n' +
    'Al seleccionar un grado en la columna "Acción" de la hoja "Interés",\n' +
    'el alumno se transfiere automáticamente a su pestaña.'
  );
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('Todos los triggers eliminados.');
}


// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 8 · INTEGRACIÓN KOBOTOOLBOX
// ────────────────────────────────────────────────────────────────────────────
//
//  Flujo de uso:
//  1. Menú → KoboToolbox → 🔑 Configurar token de API  (una vez)
//  2. Menú → KoboToolbox → 🔍 Ver columnas disponibles (ajusta KOBO_MAP arriba)
//  3. Menú → KoboToolbox → 🔄 Actualizar datos actuales
//     ↳ Descarga el CSV de URL_ACTUAL, reemplaza "Kobo: Interés Actual"
//  4. Menú → KoboToolbox → 📥 Importar histórico       (solo la 1ª vez)
//     ↳ Descarga el CSV de URL_HISTORICO, crea "Kobo: Histórico"
//  5. Opcional: Menú → KoboToolbox → 🔁 Sync automático (cada hora)
//
// ────────────────────────────────────────────────────────────────────────────

// ── 8.1  Gestión del token ───────────────────────────────────────────────────

function koboConfigurarToken() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    '🔑 Token de API — KoboToolbox',
    'Pega aquí tu token (lo encontrarás en kf.kobotoolbox.org → ícono de usuario → API Key).\n\n' +
    'Se guardará de forma segura en Script Properties (no en el código).',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const token = res.getResponseText().trim();
  if (!token) { ui.alert('Token vacío. No se guardó nada.'); return; }

  PropertiesService.getScriptProperties().setProperty(PROP_KOBO_TOKEN, token);
  ui.alert('✅ Token guardado.\n\nYa puedes usar las opciones de importación.');
}

function _koboGetToken() {
  const token = PropertiesService.getScriptProperties().getProperty(PROP_KOBO_TOKEN);
  if (!token) throw new Error(
    'No hay token configurado.\n\n' +
    'Ve a: DP Educación → KoboToolbox → 🔑 Configurar token de API'
  );
  return token;
}


// ── 8.2  Descarga y parseo de CSV ────────────────────────────────────────────

function _koboFetchCsv(url) {
  const token = _koboGetToken();
  const resp  = UrlFetchApp.fetch(url, {
    headers:            { Authorization: 'Token ' + token },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code === 401 || code === 403) throw new Error(
    'Error ' + code + ': token inválido o sin permisos.\n' +
    'Ve a: DP Educación → KoboToolbox → 🔑 Configurar token de API'
  );
  if (code !== 200) throw new Error(
    'Error al descargar el CSV (HTTP ' + code + ').\nURL: ' + url
  );

  return resp.getContentText('UTF-8');
}

/**
 * Parsea un CSV respetando campos entre comillas con comas internas.
 * Devuelve un array 2D: [[col1, col2, ...], [val1, val2, ...], ...]
 */
function _koboParseCsv(text) {
  const rows   = [];
  const lines  = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');

  lines.forEach(function(line) {
    if (line.trim() === '') return;
    const cols = [];
    let cur    = '';
    let inQ    = false;

    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') {
        if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
        else inQ = !inQ;
      } else if (c === ',' && !inQ) {
        cols.push(cur.trim()); cur = '';
      } else {
        cur += c;
      }
    }
    cols.push(cur.trim());
    rows.push(cols);
  });

  return rows;
}


// ── 8.3  Ver columnas disponibles ────────────────────────────────────────────

function koboDescubrirColumnasActual()    { _koboMostrarColumnas(KOBO_URL_ACTUAL,    'Datos actuales (2025+)'); }
function koboDescubrirColumnasHistorico() { _koboMostrarColumnas(KOBO_URL_HISTORICO, 'Histórico'); }

function _koboMostrarColumnas(url, etiqueta) {
  const ui = SpreadsheetApp.getUi();
  try {
    const csv  = _koboFetchCsv(url);
    const rows = _koboParseCsv(csv);
    if (!rows.length) { ui.alert('El CSV está vacío.'); return; }

    const headers = rows[0];
    const total   = rows.length - 1;

    let msg = '📋 COLUMNAS — ' + etiqueta + '\n';
    msg += '(' + total + ' registros, ' + headers.length + ' columnas)\n';
    msg += '─'.repeat(40) + '\n\n';
    headers.forEach(function(h, i) {
      msg += (i + 1) + '. ' + h + '\n';
    });
    msg += '\n─'.repeat(40) + '\n';
    msg += '💡 Copia los nombres exactos en el objeto KOBO_MAP\n';
    msg += '   (Sección 1 de Code.gs) para activar la sincronización.';

    ui.alert('Columnas KoboToolbox', msg, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('❌ Error\n\n' + err.message);
  }
}


// ── 8.4  Importar datos actuales (2025+) → "Kobo: Interés Actual" ────────────

function koboImportarActual() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert(
      '⏳ Descargando datos...',
      'Esto puede tardar unos segundos. Haz clic en OK para continuar.',
      ui.ButtonSet.OK
    );

    const csv  = _koboFetchCsv(KOBO_URL_ACTUAL);
    const rows = _koboParseCsv(csv);

    if (rows.length < 2) {
      ui.alert('El CSV no tiene datos (solo encabezado o está vacío).');
      return;
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   hoja  = ss.getSheetByName(HOJA_KOBO_ACTUAL);
    if (!hoja) {
      hoja = ss.insertSheet(HOJA_KOBO_ACTUAL);
    } else {
      hoja.clearContents();
      hoja.clearFormats();
      hoja.clearConditionalFormatRules();
    }

    // Escribir todo el CSV de una vez (más rápido)
    hoja.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

    // Formato encabezado
    _koboFormatearHoja(hoja, rows[0].length,
      '📥 KOBO: INTERÉS ACTUAL  ·  Actualizado: ' + new Date().toLocaleString('es-GT'));

    // Ordenar por la 1ª columna (normalmente _id o fecha)
    if (rows.length > 2) {
      hoja.getRange(2, 1, rows.length - 1, rows[0].length).sort({ column: 1, ascending: false });
    }

    ss.toast(
      (rows.length - 1) + ' registros importados desde KoboToolbox.',
      '✅ Datos actuales actualizados', 6
    );
    ui.alert(
      '✅ Importación completada\n\n' +
      'Registros: ' + (rows.length - 1) + '\n' +
      'Columnas:  ' + rows[0].length + '\n' +
      'Hoja:      "' + HOJA_KOBO_ACTUAL + '"\n\n' +
      '💡 Si las columnas de KOBO_MAP coinciden, puedes sincronizar\n' +
      'directamente a la hoja "Interés" con "Procesar acciones pendientes".'
    );

  } catch (err) {
    ui.alert('❌ Error al importar\n\n' + err.message);
  }
}


// ── 8.5  Importar histórico (una sola vez) → "Kobo: Histórico" ───────────────

function koboImportarHistorico() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Advertir si ya existe
  const hojaExistente = ss.getSheetByName(HOJA_KOBO_HISTORICO);
  if (hojaExistente) {
    const r = ui.alert(
      'Hoja ya existe',
      '"' + HOJA_KOBO_HISTORICO + '" ya existe con ' +
      Math.max(hojaExistente.getLastRow() - 1, 0) + ' registros.\n\n' +
      '¿Deseas reemplazarla con los datos más recientes?',
      ui.ButtonSet.YES_NO
    );
    if (r !== ui.Button.YES) return;
  }

  try {
    ui.alert(
      '⏳ Descargando histórico...',
      'Los datos históricos pueden ser grandes. Haz clic en OK para continuar.',
      ui.ButtonSet.OK
    );

    const csv  = _koboFetchCsv(KOBO_URL_HISTORICO);
    const rows = _koboParseCsv(csv);

    if (rows.length < 2) {
      ui.alert('El CSV histórico no tiene datos o está vacío.');
      return;
    }

    let hoja = hojaExistente;
    if (!hoja) {
      hoja = ss.insertSheet(HOJA_KOBO_HISTORICO);
    } else {
      hoja.clearContents();
      hoja.clearFormats();
      hoja.clearConditionalFormatRules();
    }

    // Escribir
    hoja.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

    // Formato
    _koboFormatearHoja(hoja, rows[0].length,
      '📚 KOBO: HISTÓRICO  ·  Importado: ' + new Date().toLocaleString('es-GT'));

    // Ordenar descendente por col 1
    if (rows.length > 2) {
      hoja.getRange(2, 1, rows.length - 1, rows[0].length).sort({ column: 1, ascending: false });
    }

    // Proteger: solo lectura para evitar ediciones accidentales
    const proteccion = hoja.protect().setDescription('Datos históricos — solo lectura');
    proteccion.setWarningOnly(true);

    ss.toast(
      (rows.length - 1) + ' registros históricos importados.',
      '✅ Histórico importado', 6
    );
    ui.alert(
      '✅ Histórico importado correctamente\n\n' +
      'Registros: ' + (rows.length - 1) + '\n' +
      'Columnas:  ' + rows[0].length + '\n' +
      'Hoja:      "' + HOJA_KOBO_HISTORICO + '"\n\n' +
      '⚠️ La hoja tiene protección de advertencia para evitar\n' +
      'modificaciones accidentales. Es de solo consulta.'
    );

  } catch (err) {
    ui.alert('❌ Error al importar histórico\n\n' + err.message);
  }
}


// ── 8.6  Formato visual de hojas Kobo ────────────────────────────────────────

function _koboFormatearHoja(hoja, numCols, titulo) {
  const MAX = hoja.getLastRow();

  // Título en fila 1 (encima de los datos)
  hoja.insertRowBefore(1);
  hoja.getRange(1, 1, 1, numCols).merge()
    .setValue(titulo)
    .setBackground(COLOR_HEADER_KOBO)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 36);

  // Encabezados en fila 2
  hoja.getRange(2, 1, 1, numCols)
    .setBackground(COLOR_HEADER_KOBO)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 28);
  hoja.setFrozenRows(2);

  // Alternar colores de fila para legibilidad
  if (MAX > 2) {
    for (let f = 3; f <= MAX + 1; f++) {
      hoja.getRange(f, 1, 1, numCols)
        .setBackground(f % 2 === 0 ? '#F3E5F5' : '#FFFFFF');
    }
  }

  // Auto-ajuste de columnas (máx 300 px)
  hoja.autoResizeColumns(1, numCols);
  for (let c = 1; c <= numCols; c++) {
    if (hoja.getColumnWidth(c) > 300) hoja.setColumnWidth(c, 300);
  }
}


// ── 8.7  Trigger automático de sincronización ─────────────────────────────────

function koboInstalarTriggerSync() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Verificar token antes de instalar
  try { _koboGetToken(); } catch (e) {
    ui.alert('❌ ' + e.message);
    return;
  }

  // Eliminar triggers previos del mismo handler
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'koboImportarActual'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('koboImportarActual')
    .forSpreadsheet(ss)
    .timeBased()
    .everyHours(1)
    .create();

  ui.alert(
    '✅ Sync automático activado\n\n' +
    'La hoja "' + HOJA_KOBO_ACTUAL + '" se actualizará cada hora\n' +
    'automáticamente desde KoboToolbox.\n\n' +
    'Para detenerlo: menú → KoboToolbox → ⛔ Detener sync automático'
  );
}

function koboEliminarTriggerSync() {
  const eliminados = ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'koboImportarActual'; });

  eliminados.forEach(function(t) { ScriptApp.deleteTrigger(t); });

  SpreadsheetApp.getUi().alert(
    eliminados.length > 0
      ? '✅ Sync automático detenido (' + eliminados.length + ' trigger(s) eliminado(s)).'
      : 'No había ningún trigger de sync activo.'
  );
}
