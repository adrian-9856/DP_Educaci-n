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
        .addItem('🔑 Configurar token de API',              'koboConfigurarToken')
        .addSeparator()
        .addItem('🗺️ Ver hoja de mapeo (actuales)',         'koboVerHojaMapeoActual')
        .addItem('🗺️ Ver hoja de mapeo (histórico)',        'koboVerHojaMapeoHistorico')
        .addSeparator()
        .addItem('🔄 Sync a hoja Interés (actuales 2025+)', 'koboSincronizarHojaInteres')
        .addItem('📥 Importar histórico (una sola vez)',     'koboImportarHistorico')
        .addSeparator()
        .addItem('🔁 Sync automático (cada hora)',           'koboInstalarTriggerSync')
        .addItem('⛔ Detener sync automático',               'koboEliminarTriggerSync')
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
//  1. Menú → KoboToolbox → 🔑 Configurar token de API  (una sola vez)
//  2. Menú → KoboToolbox → 🗺️ Ver hoja de mapeo
//     ↳ Crea la hoja "Kobo: Mapeo" con todas las columnas y ejemplos
//     ↳ Ajusta KOBO_MAP en Sección 1 con los nombres exactos que veas ahí
//  3. Menú → KoboToolbox → 🔄 Sync a hoja Interés
//     ↳ Agrega solo registros nuevos (deduplica por DPI)
//  4. Menú → KoboToolbox → 📥 Importar histórico  (una sola vez)
//  5. Opcional: 🔁 Sync automático (cada hora)
//
// ────────────────────────────────────────────────────────────────────────────

// ── 8.1  Gestión del token ───────────────────────────────────────────────────

function koboConfigurarToken() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    '🔑 Token de API — KoboToolbox',
    'Pega aquí tu token.\n(kf.kobotoolbox.org → ícono de usuario → API Key)\n\n' +
    'Se guarda en Script Properties, no en el código.',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const token = res.getResponseText().trim();
  if (!token) { ui.alert('Token vacío. No se guardó nada.'); return; }
  PropertiesService.getScriptProperties().setProperty(PROP_KOBO_TOKEN, token);
  ui.alert('✅ Token guardado.\nYa puedes usar las opciones de importación.');
}

function _koboGetToken() {
  const token = PropertiesService.getScriptProperties().getProperty(PROP_KOBO_TOKEN);
  if (!token) throw new Error(
    'No hay token configurado.\nVe a: DP Educación → 🌐 KoboToolbox → 🔑 Configurar token de API'
  );
  return token;
}


// ── 8.2  Descarga del CSV ─────────────────────────────────────────────────────

function _koboFetchCsv(url) {
  const token = _koboGetToken();
  const resp  = UrlFetchApp.fetch(url, {
    headers:            { Authorization: 'Token ' + token },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code === 401 || code === 403) throw new Error(
    'Error ' + code + ': token inválido o sin permisos.\n' +
    'Ve a: DP Educación → 🌐 KoboToolbox → 🔑 Configurar token de API'
  );
  if (code !== 200) throw new Error('Error al descargar el CSV (HTTP ' + code + ').');
  return resp.getContentText('UTF-8');
}


// ── 8.3  Parser CSV con detección automática de separador ────────────────────

/**
 * Detecta si el CSV usa coma (,) o punto y coma (;) como separador.
 * Compara cuántos de cada tipo hay en la primera línea fuera de comillas.
 */
function _koboDetectarSeparador(primeraLinea) {
  let comas = 0, puntoYComas = 0, inQ = false;
  for (let i = 0; i < primeraLinea.length; i++) {
    const c = primeraLinea[i];
    if (c === '"') { inQ = !inQ; continue; }
    if (inQ) continue;
    if (c === ',')  comas++;
    if (c === ';')  puntoYComas++;
  }
  return puntoYComas >= comas ? ';' : ',';
}

/**
 * Parsea un CSV respetando campos entre comillas.
 * Auto-detecta separador (coma o punto y coma).
 * Devuelve un array 2D: [[col1, col2, ...], [val1, val2, ...], ...]
 */
function _koboParseCsv(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n')
                    .filter(function(l) { return l.trim() !== ''; });
  if (!lines.length) return [];

  const sep  = _koboDetectarSeparador(lines[0]);
  const rows = [];

  lines.forEach(function(line) {
    const cols = [];
    let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') {
        if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
        else inQ = !inQ;
      } else if (c === sep && !inQ) {
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


// ── 8.4  Hoja de mapeo visual ─────────────────────────────────────────────────
//
//  Crea "Kobo: Mapeo" con una fila por cada columna del CSV:
//  Índice | Nombre del campo | Ejemplo 1 | Ejemplo 2 | Ejemplo 3 | Mapeo actual
//  Así puedes identificar los campos y ajustar KOBO_MAP en Sección 1.

function koboVerHojaMapeoActual()    { _koboCrearHojaMapeo(KOBO_URL_ACTUAL,    'actual');    }
function koboVerHojaMapeoHistorico() { _koboCrearHojaMapeo(KOBO_URL_HISTORICO, 'histórico'); }

function _koboCrearHojaMapeo(url, etiqueta) {
  const ui = SpreadsheetApp.getUi();
  try {
    const csv  = _koboFetchCsv(url);
    const rows = _koboParseCsv(csv);
    if (rows.length < 1) { ui.alert('El CSV está vacío.'); return; }

    const headers  = rows[0];
    const muestra1 = rows.length > 1 ? rows[1] : [];
    const muestra2 = rows.length > 2 ? rows[2] : [];
    const muestra3 = rows.length > 3 ? rows[3] : [];

    // Invertir KOBO_MAP para búsqueda rápida: valor → clave
    const mapaInvertido = {};
    Object.keys(KOBO_MAP).forEach(function(k) {
      mapaInvertido[KOBO_MAP[k]] = k;
    });

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const nombre = 'Kobo: Mapeo (' + etiqueta + ')';
    let hoja   = ss.getSheetByName(nombre);
    if (!hoja) hoja = ss.insertSheet(nombre);
    else hoja.clearContents();

    // Fila de título
    hoja.getRange(1, 1, 1, 6).merge()
      .setValue('🗺️ MAPEO DE COLUMNAS KOBO — ' + etiqueta.toUpperCase() +
                '  |  ' + (rows.length - 1) + ' registros  |  ' + headers.length + ' columnas')
      .setBackground(COLOR_HEADER_KOBO)
      .setFontColor(COLOR_FONT_HEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    hoja.setRowHeight(1, 36);

    // Encabezados de la hoja de mapeo
    const encabezadosMapeo = ['#', 'Nombre del campo (Kobo)', 'Ejemplo 1', 'Ejemplo 2', 'Ejemplo 3', 'Mapea a Interés →'];
    hoja.getRange(2, 1, 1, 6).setValues([encabezadosMapeo])
      .setBackground(COLOR_HEADER_KOBO)
      .setFontColor(COLOR_FONT_HEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    hoja.setRowHeight(2, 26);
    hoja.setFrozenRows(2);

    // Una fila por columna del CSV
    const filas = headers.map(function(h, i) {
      const mapeo = mapaInvertido[h] ? '→ ' + mapaInvertido[h] : '';
      return [
        i + 1,
        h,
        muestra1[i] !== undefined ? muestra1[i] : '',
        muestra2[i] !== undefined ? muestra2[i] : '',
        muestra3[i] !== undefined ? muestra3[i] : '',
        mapeo
      ];
    });
    hoja.getRange(3, 1, filas.length, 6).setValues(filas);

    // Resaltar filas que ya están mapeadas
    filas.forEach(function(f, idx) {
      if (f[5]) {
        hoja.getRange(idx + 3, 1, 1, 6).setBackground('#E8F5E9');
      }
    });

    // Anchos
    hoja.setColumnWidth(1, 45);
    hoja.setColumnWidth(2, 280);
    hoja.setColumnWidth(3, 200);
    hoja.setColumnWidth(4, 200);
    hoja.setColumnWidth(5, 200);
    hoja.setColumnWidth(6, 170);

    ss.setActiveSheet(hoja);
    ss.toast(
      'Revisa la columna "Nombre del campo" y actualiza KOBO_MAP en Sección 1.',
      '🗺️ Hoja de mapeo lista', 8
    );
  } catch (err) {
    ui.alert('❌ Error\n\n' + err.message);
  }
}


// ── 8.5  Sincronizar datos actuales → hoja "Interés" ─────────────────────────
//
//  - Descarga el CSV de KOBO_URL_ACTUAL
//  - Para cada fila Kobo, extrae solo los campos de KOBO_MAP
//  - Si el DPI ya existe en la hoja Interés → lo omite (sin duplicados)
//  - Si no existe → agrega la fila al final

function koboSincronizarHojaInteres() {
  const ui = SpreadsheetApp.getUi();
  try {
    const csv  = _koboFetchCsv(KOBO_URL_ACTUAL);
    const rows = _koboParseCsv(csv);

    if (rows.length < 2) {
      ui.alert('El CSV no tiene datos (solo encabezado o está vacío).');
      return;
    }

    const headers = rows[0];

    // Índices de los campos mapeados en el CSV (busca por nombre)
    const idx = {};
    Object.keys(KOBO_MAP).forEach(function(campo) {
      const nombreKobo = KOBO_MAP[campo];
      const pos        = headers.indexOf(nombreKobo);
      idx[campo]       = pos; // -1 si no se encuentra
    });

    // Campos no encontrados
    const noEncontrados = Object.keys(idx).filter(function(k) { return idx[k] < 0; });
    if (noEncontrados.length > 0) {
      const aviso = noEncontrados.map(function(k) {
        return '  • ' + k + ' → buscando "' + KOBO_MAP[k] + '" (no encontrado)';
      }).join('\n');
      const r = ui.alert(
        '⚠️ Campos no encontrados en el CSV',
        'Los siguientes campos de KOBO_MAP no se encontraron en el CSV:\n\n' + aviso +
        '\n\nEjecuta "🗺️ Ver hoja de mapeo" para ver los nombres exactos.\n\n' +
        '¿Continuar de todas formas con los campos que sí se encontraron?',
        ui.ButtonSet.YES_NO
      );
      if (r !== ui.Button.YES) return;
    }

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const hojaInteres = ss.getSheetByName(HOJA_INTERES);
    if (!hojaInteres) {
      ui.alert('❌ La hoja "Interés" no existe. Ejecuta primero "⚙️ Configurar hoja Interés".');
      return;
    }

    // DPIs que ya existen en la hoja Interés (para no duplicar)
    const ultimaFilaI  = hojaInteres.getLastRow();
    const dpisExistentes = new Set();
    if (ultimaFilaI >= 2) {
      hojaInteres.getRange(2, COL_INTERES.DPI, ultimaFilaI - 1, 1)
        .getValues().flat()
        .forEach(function(d) { if (d !== '') dpisExistentes.add(String(d).trim()); });
    }

    // Procesar filas Kobo
    const filasNuevas = [];
    rows.slice(1).forEach(function(row) {
      const dpi = idx.DPI >= 0 ? String(row[idx.DPI] || '').trim() : '';
      if (dpi && dpisExistentes.has(dpi)) return; // ya existe

      const nombre    = idx.NOMBRE      >= 0 ? (row[idx.NOMBRE]      || '') : '';
      const edad      = idx.EDAD        >= 0 ? (row[idx.EDAD]        || '') : '';
      const ultimoA   = idx.ULTIMO_ANIO >= 0 ? (row[idx.ULTIMO_ANIO] || '') : '';
      const papeleria = idx.PAPELERIA   >= 0 ? (row[idx.PAPELERIA]   || '') : '';
      const comentario= idx.COMENTARIO  >= 0 ? (row[idx.COMENTARIO]  || '') : '';

      if (!nombre && !dpi) return; // fila vacía

      filasNuevas.push([
        String(nombre).trim(),
        String(edad).trim(),
        dpi,
        String(ultimoA).trim(),
        String(papeleria).trim(),
        String(comentario).trim(),
        '-- Seleccionar --'
      ]);

      if (dpi) dpisExistentes.add(dpi); // prevenir duplicados dentro del mismo lote
    });

    if (!filasNuevas.length) {
      ui.alert('✅ Sin novedades\n\nTodos los registros de KoboToolbox ya existen en la hoja "Interés".');
      return;
    }

    // Escribir al final de la hoja Interés
    const primeraFilaLibre = Math.max(hojaInteres.getLastRow() + 1, 2);
    hojaInteres.getRange(primeraFilaLibre, 1, filasNuevas.length, 7).setValues(filasNuevas);

    // Aplicar validación de Acción a las nuevas filas
    const validAccion = SpreadsheetApp.newDataValidation()
      .requireValueInList(ACCIONES, true).setAllowInvalid(false).build();
    hojaInteres.getRange(primeraFilaLibre, COL_INTERES.ACCION, filasNuevas.length, 1)
      .setDataValidation(validAccion);

    // Validación Último año
    const validAnio = SpreadsheetApp.newDataValidation()
      .requireValueInList(ULTIMO_ANIO_OPCIONES, true).setAllowInvalid(false).build();
    hojaInteres.getRange(primeraFilaLibre, COL_INTERES.ULTIMO_ANIO, filasNuevas.length, 1)
      .setDataValidation(validAnio);

    ss.setActiveSheet(hojaInteres);
    ss.toast(
      filasNuevas.length + ' nuevos registros agregados desde KoboToolbox.',
      '✅ Sync completado', 7
    );
    ui.alert(
      '✅ Sincronización completada\n\n' +
      'Registros nuevos agregados: ' + filasNuevas.length + '\n' +
      'Registros ya existentes:    ' + (rows.length - 1 - filasNuevas.length) + '\n\n' +
      '💡 Revisa la columna "Acción" para asignar los alumnos nuevos a su grado.'
    );

  } catch (err) {
    ui.alert('❌ Error en sincronización\n\n' + err.message);
  }
}


// ── 8.6  Importar histórico → "Kobo: Histórico" ──────────────────────────────

function koboImportarHistorico() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaExistente = ss.getSheetByName(HOJA_KOBO_HISTORICO);
  if (hojaExistente) {
    const r = ui.alert(
      'Hoja ya existe',
      '"' + HOJA_KOBO_HISTORICO + '" ya existe con ' +
      Math.max(hojaExistente.getLastRow() - 2, 0) + ' registros.\n\n' +
      '¿Reemplazar con los datos más recientes?',
      ui.ButtonSet.YES_NO
    );
    if (r !== ui.Button.YES) return;
  }

  try {
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

    hoja.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    _koboFormatearHojaRaw(hoja, rows[0].length,
      '📚 KOBO: HISTÓRICO  ·  Importado: ' + new Date().toLocaleString('es-GT'));

    if (rows.length > 2) {
      hoja.getRange(3, 1, rows.length - 1, rows[0].length).sort({ column: 1, ascending: false });
    }

    hoja.protect().setDescription('Datos históricos — solo lectura').setWarningOnly(true);

    ss.toast((rows.length - 1) + ' registros históricos importados.', '✅ Histórico listo', 6);
    ui.alert(
      '✅ Histórico importado\n\n' +
      'Registros: ' + (rows.length - 1) + '\n' +
      'Columnas:  ' + rows[0].length + '\n' +
      'Hoja:      "' + HOJA_KOBO_HISTORICO + '"\n\n' +
      '⚠️ Hoja protegida (solo lectura).'
    );

  } catch (err) {
    ui.alert('❌ Error al importar histórico\n\n' + err.message);
  }
}


// ── 8.7  Formato visual de hojas Kobo (datos crudos) ─────────────────────────

function _koboFormatearHojaRaw(hoja, numCols, titulo) {
  hoja.insertRowBefore(1);
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

  hoja.getRange(2, 1, 1, numCols)
    .setBackground(COLOR_HEADER_KOBO)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 26);
  hoja.setFrozenRows(2);

  hoja.autoResizeColumns(1, numCols);
  for (let c = 1; c <= numCols; c++) {
    if (hoja.getColumnWidth(c) > 280) hoja.setColumnWidth(c, 280);
  }
}


// ── 8.8  Trigger automático de sincronización ─────────────────────────────────

function koboInstalarTriggerSync() {
  const ui = SpreadsheetApp.getUi();
  try { _koboGetToken(); } catch (e) { ui.alert('❌ ' + e.message); return; }

  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'koboSincronizarHojaInteres'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('koboSincronizarHojaInteres')
    .timeBased().everyHours(1).create();

  ui.alert(
    '✅ Sync automático activado\n\n' +
    'Cada hora se agregarán automáticamente los registros nuevos de\n' +
    'KoboToolbox a la hoja "Interés".\n\n' +
    'Para detenerlo: menú → 🌐 KoboToolbox → ⛔ Detener sync automático'
  );
}

function koboEliminarTriggerSync() {
  const eliminados = ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'koboSincronizarHojaInteres'; });

  eliminados.forEach(function(t) { ScriptApp.deleteTrigger(t); });

  SpreadsheetApp.getUi().alert(
    eliminados.length > 0
      ? '✅ Sync automático detenido (' + eliminados.length + ' trigger(s) eliminado(s)).'
      : 'No había ningún trigger de sync activo.'
  );
}
