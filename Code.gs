// ════════════════════════════════════════════════════════════════════════════
//  DP EDUCACIÓN — Automatización Google Sheets
//  Archivo único: todo el flujo en un mismo Spreadsheet con varias pestañas
// ════════════════════════════════════════════════════════════════════════════

// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 1 · CONSTANTES GLOBALES
// ────────────────────────────────────────────────────────────────────────────

const HOJA_INTERES      = 'Hoja de Interés';
const HOJA_REFERENCIAS  = 'Referencias a Educación';
const HOJA_LISTA_ESPERA = 'Lista de Espera';

// Año del ciclo escolar activo — cambia aquí cada año
const SCHOOL_YEAR = 2026;

// Columnas de la hoja Interés (1-based)  — 14 columnas totales
const COL_INTERES = {
  CREAMOS_ID:  1,   // Creamos ID (desde KoboToolbox)
  NOMBRE:      2,   // Nombre(s) + Apellido(s)
  NOMBRE_PREF: 3,   // Nombre Preferido
  DPI:         4,   // Número de DPI
  FECHA_NAC:   5,   // Fecha de nacimiento
  EDAD:        6,   // Calculada automáticamente
  GENERO:      7,   // Género
  TELEFONO:    8,   // Número de Teléfono
  ZONA:        9,   // Zona / Colonia
  ULTIMO_ANIO: 10,  // Último nivel de estudios
  GRADO_KOBO:  11,  // Grado asignado por Creamos (desde Kobo)
  PAPELERIA:   12,  // Documentos faltantes (construido automáticamente)
  COMENTARIO:  13,  // Comentarios de papelería
  ACCION:      14   // Desplegable: Enviar a grado
};

// Columnas de "Referencias a Educación" (1-based) — 14 columnas
const COL_REF = {
  CREAMOS_ID:     1,   // Creamos ID
  NOMBRE:         2,   // Nombre Completo
  NOMBRE_PREF:    3,   // Nombre Preferido
  DPI:            4,   // DPI / CUI
  FECHA_NAC:      5,   // Fecha de Nacimiento
  EDAD:           6,   // Edad (calculada)
  GENERO:         7,   // Género
  TELEFONO:       8,   // Teléfono
  ZONA:           9,   // Zona / Colonia
  ULTIMO_ANIO:   10,   // Último Nivel Cursado
  FECHA_REF:     11,   // Fecha de Referencia (de KoboToolbox)
  RESPONSABLE:   12,   // Responsable de la Referencia
  ESTADO_ESTUDIO:13,   // Estado del Estudio
  ACCION:        14    // Acción (→Lista de Espera)
};
// Columnas de "Lista de Espera" (1-based) — 11 columnas (estructura original)
const COL_LISTA = {
  CREAMOS_ID:  1,
  NOMBRE:      2,
  NOMBRE_PREF: 3,
  DPI:         4,
  FECHA_NAC:   5,
  EDAD:        6,
  GENERO:      7,
  TELEFONO:    8,
  ZONA:        9,
  ULTIMO_ANIO: 10,
  ACCION:      11   // Acción → Enviar a grado
};

// Columnas de cada hoja de grado (1-based) — 9 columnas totales
const COL_GRADO = {
  ID:         1,
  CREAMOS_ID: 2,
  NOMBRE:     3,
  DPI:        4,
  TELEFONO:   5,
  EDAD:       6,
  GRADO:      7,
  MODALIDAD:  8,
  ESTADO:     9
};

const GRADOS = [
  'Primera Etapa de Primaria',
  'Segunda Etapa de Primaria',
  'Primera Etapa de Básicos',
  'Segunda Etapa de Básicos',
  'Cuarto Bachillerato',
  'Quinto Bachillerato'
];

// Grado del que salen los graduados que pasan a Seguimiento
const GRADO_GRADUACION = 'Quinto Bachillerato';

const PAPELERIA_OPCIONES = [
  'Convenio',
  'Código Personal del Alumno',
  'Fe de Edad',
  'DPI',
  'Cert. 1° Primaria',
  'Cert. 2° Primaria',
  'Cert. 3° Primaria',
  'Cert. 4° Primaria',
  'Cert. 5° Primaria',
  'Cert. 6° Primaria',
  'Diploma 6° Primaria',
  'Cert. 1° Básico',
  'Cert. 2° Básico',
  'Cert. 3° Básico',
  'Diploma Básico',
  'Cert. 4° Bachillerato'
];

const MODALIDADES    = ['Presencial', 'Semi-presencial'];
// Estados para hojas de grado (excepto Quinto Bachillerato)
const ESTADOS        = ['Inscritx', 'Retiradx', 'Graduadx'];
// Estado especial para Quinto Bachillerato al completar el ciclo
const ESTADOS_QUINTO = ['Inscritx', 'Retiradx', 'Ciclo de Vida Terminado'];
const ACCIONES       = ['-- Seleccionar --', 'Enviar a: Lista de Espera'];
const ULTIMO_ANIO_OPCIONES = [
  'Sin estudios previos',
  'Primera Etapa de Primaria',
  'Segunda Etapa de Primaria',
  'Primera Etapa de Básicos',
  'Segunda Etapa de Básicos',
  'Cuarto Bachillerato'
];

const COLOR_HEADER_INTERES     = '#1565C0';
const COLOR_HEADER_REFERENCIAS = '#6A1B9A';  // morado
const COLOR_HEADER_LISTA       = '#E65100';  // naranja
const COLOR_HEADER_GRADO   = '#2E7D32';
const COLOR_FONT_HEADER    = '#FFFFFF';

// ── KoboToolbox ──────────────────────────────────────────────────────────────
// URL de exportación CSV — datos de educación (requiere token en Script Properties)
// URL nueva — formulario activo (sync automático cada minuto)
const KOBO_URL_ACTUAL = 'https://kf.kobotoolbox.org/api/v2/assets/auvEELWQEgiwF54W4pGpV5/export-settings/esd2gxqN87HPuQDypxFqUNi/data.csv';

// URL histórica — formulario de años anteriores (importación única / manual)
const KOBO_URL_HISTORICO = 'https://kf.kobotoolbox.org/api/v2/assets/akz5K2bGfvvisQaE7VaHev/export-settings/esuV4RKqQhYUUaUizfWBP8S/data.csv';

// URL formulario de Referencias a Educación
const KOBO_URL_REFERENCIAS = 'https://kf.kobotoolbox.org/api/v2/assets/afuD8C8AzoLfd4o5ksTWUw/export-settings/es52swrnjWcz8NnhY5Wyng3/data.csv';

// Colores encabezado Kobo
const COLOR_HEADER_KOBO = '#6A1B9A';

// Clave en Script Properties donde se guarda el token de Kobo
const PROP_KOBO_TOKEN = 'KOBO_API_TOKEN';

// Mapeo de campos KoboToolbox → columnas de la hoja Interés
// Nombres exactos del CSV de KoboToolbox
// ── Mapa de campos Kobo → columnas del CSV ───────────────────────────────────
// Nombres primarios = formulario NUEVO (sin prefijos de grupo).
// Los formularios históricos usan prefijos "Inicio/" y "Educación Extraescolar/..."
// → se manejan vía fallbacks más abajo en _koboSincronizar.
const KOBO_MAP = {
  // ── Datos personales ────────────────────────────────────────────────────
  CREAMOS_ID:     'Creamos ID',
  NOMBRE:         'Nombre(s)',
  APELLIDO:       'Apellido(s)',
  NOMBRE_PREF:    'Nombre Preferido',
  DPI:            'Número de DPI',           // solo formulario histórico
  FECHA_NAC:      'Fecha de nacimiento',     // solo formulario histórico
  EDAD_DIRECTA:   'Edad',                    // formulario nuevo (viene directo)
  GENERO:         'Género',
  AUTODESCRIBE:   '¿Cómo te autodescribes?', // formulario nuevo (complementa Género)
  TELEFONO:       'Número de Teléfono',
  ZONA:           'Zona',
  OTRA_ZONA:      'Otra zona',
  COLONIA:        'Colonia',
  OTRA_COLONIA:   'Otra colonia',
  ULTIMO_ANIO:    '¿Cuál es tu último nivel de estudios terminado?',
  // ── Filtro de programa ──────────────────────────────────────────────────
  // En el nuevo formulario, "¿Qué programas te interesan?" se desglosa en
  // columnas booleanas individuales por programa (1 = marcado, 0/vacío = no).
  PROGRAMAS_EDUC: '¿Qué programas te interesan?/Educación Extraescolar/Alternativa',
  // ── Sección Educación ───────────────────────────────────────────────────
  INSCRIPCION:    '¿Deseas inscribirte en el programa de Educación?',
  GRADO_KOBO:     '¿Qué grado/etapa te toca con Creamos?',
  COMENTARIO:     'Comentarios',             // sin ":" es la columna de Educación
  // ── Identificador único de envío (para deduplicar sin DPI) ──────────────
  UUID:           '_uuid'
};

// Texto clave para búsqueda en campo de texto multi-programa (fallback)
const KOBO_KEYWORD_EDUCACION = 'educacion extraescolar';

// Campos de papelería en Kobo (0 = faltante, 1 = entregado)
// Se construye automáticamente la lista de documentos faltantes
const KOBO_MAP_PAPELERIA = {
  'Convenio':             'Educación Extraescolar/Alternativa/Convenio',
  'Código Personal':      'Educación Extraescolar/Alternativa/Código Personal del Alumno',
  'Fe de Edad':           'Educación Extraescolar/Alternativa/Fe de Edad',
  'DPI':                  'Educación Extraescolar/Alternativa/DPI',
  'Cert. 1° Primaria':    'Educación Extraescolar/Alternativa/Certificado 1ero. Primaria',
  'Cert. 2° Primaria':    'Educación Extraescolar/Alternativa/Certificado 2do. Primaria',
  'Cert. 3° Primaria':    'Educación Extraescolar/Alternativa/Certificado 3ero. Primaria',
  'Cert. 4° Primaria':    'Educación Extraescolar/Alternativa/Certificado 4to. Primaria',
  'Cert. 5° Primaria':    'Educación Extraescolar/Alternativa/Certificado 5to. Primaria',
  'Cert. 6° Primaria':    'Educación Extraescolar/Alternativa/Certificado 6to. Primaria',
  'Diploma 6° Primaria':  'Educación Extraescolar/Alternativa/Diploma 6to. Primaria',
  'Cert. 1° Básico':      'Educación Extraescolar/Alternativa/Certificado 1ero. Básico',
  'Cert. 2° Básico':      'Educación Extraescolar/Alternativa/Certificado 2do. Básico',
  'Cert. 3° Básico':      'Educación Extraescolar/Alternativa/Certificado 3ero. Básico',
  'Diploma Básico':       'Educación Extraescolar/Alternativa/Diploma Básico',
  'Cert. 4° Bachillerato':'Educación Extraescolar/Alternativa/Certificado 4to. Bachillerato'
};

// Mapeo de valores de grado en Kobo → nombres exactos de GRADOS
// Incluye los valores cortos que realmente aparecen en el CSV
const KOBO_GRADO_MAP = {
  // Valores cortos del formulario actual
  'Etapa Post 1':               'Primera Etapa de Primaria',
  ' Etapa Post 1':              'Primera Etapa de Primaria',  // a veces con espacio inicial
  'Etapa Post 2':               'Segunda Etapa de Primaria',
  ' Etapa Post 2':              'Segunda Etapa de Primaria',
  'Básico I':                   'Primera Etapa de Básicos',
  'Basico I':                   'Primera Etapa de Básicos',
  'Básico II':                  'Segunda Etapa de Básicos',
  'Basico II':                  'Segunda Etapa de Básicos',
  '4to Bachillerato':           'Cuarto Bachillerato',
  '5to Bachillerato':           'Quinto Bachillerato',
  // Por si llegan con nombre completo
  'Primera Etapa de Primaria':  'Primera Etapa de Primaria',
  'Segunda Etapa de Primaria':  'Segunda Etapa de Primaria',
  'Primera Etapa de Básicos':   'Primera Etapa de Básicos',
  'Segunda Etapa de Básicos':   'Segunda Etapa de Básicos',
  'Cuarto Bachillerato':        'Cuarto Bachillerato',
  'Quinto Bachillerato':        'Quinto Bachillerato'
};

// ── Hoja de Seguimiento ───────────────────────────────────────────────────────
const HOJA_SEGUIMIENTO    = 'Seguimiento Graduados';
const COLOR_HEADER_SEGUIM = '#4A148C';

// Columnas de la hoja Seguimiento (1-based)
const COL_SEGUIM = {
  ID:          1,
  NOMBRE:      2,
  DPI:         3,
  TELEFONO:    4,
  EDAD:        5,
  ANIO_GRAD:   6,
  ESTADO_POST: 7,
  OBSERVACION: 8
};

const ESTADOS_POST_GRAD = [
  'Seguimiento activo',
  'Empleado',
  'Continúa estudiando',
  'Sin contacto',
  'Emigró'
];


// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 2 · MENÚ Y PUNTO DE ENTRADA
// ────────────────────────────────────────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('DP Educación')
    .addItem('📖  Guía de uso',                    'mostrarGuiaDeUso')
    .addSeparator()
    .addItem('⚙️  Configurar hoja Interés',        'setupHojaInteres')
    .addItem('📥 Configurar Referencias a Educ.',   'setupHojaReferencias')
    .addItem('⏳ Configurar Lista de Espera',        'setupHojaListaEspera')
    .addItem('🔧  Instalar trigger automático',     'installTriggers')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('📚 Crear hoja de grado')
        .addItem('Primera Etapa de Primaria',  'crearHojaPrimariaEtapa1')
        .addItem('Segunda Etapa de Primaria',  'crearHojaPrimariaEtapa2')
        .addItem('Primera Etapa de Básicos',   'crearHojaBasicosEtapa1')
        .addItem('Segunda Etapa de Básicos',   'crearHojaBasicosEtapa2')
        .addItem('Cuarto Bachillerato',         'crearHojaCuartoBachillerato')
        .addItem('Quinto Bachillerato',         'crearHojaQuintoBachillerato')
        .addSeparator()
        .addItem('✨ Crear TODOS los grados',   'crearTodasLasHojas')
    )
    .addSeparator()
    .addItem('📋 Seleccionar papelería faltante',   'abrirSelectorPapeleria')
    .addItem('🔄 Procesar acciones pendientes',     'procesarAccionesPendientes')
    .addItem('📊 Ver resumen de alumnos',            'mostrarResumen')
    .addItem('📅 Cerrar ciclo escolar',              'cerrarCicloEscolar')
    .addSeparator()
    .addItem('🎓 Configurar hoja Seguimiento',       'setupHojaSeguimiento')
    .addItem('🎓 Registrar graduado manualmente',    'registrarGraduadoManual')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('🌐 KoboToolbox')
        .addItem('🔑 Configurar token de API',        'koboConfigurarToken')
        .addSeparator()
        .addItem('📦 Importar datos HISTÓRICOS',           'koboImportarHistorico')
        .addItem('🔄 Sync → hoja Interés (actual)',        'koboSincronizarHojaInteres')
        .addItem('🔄 Sync → Referencias a Educación',      'koboSincronizarReferencias')
        .addItem('🔄 Sync TODO (Interés + Referencias)',   'koboSincronizarTodo')
        .addSeparator()
        .addItem('🔁 Sync automático (cada minuto)',  'koboInstalarTriggerSync')
        .addItem('⛔ Detener sync automático',          'koboEliminarTriggerSync')
    )
    .addSeparator()
    .addItem('🔁 Reiniciar sistema (⚠️ borra todo)', 'reiniciarSistema')
    .addToUi();
}

function crearHojaPrimariaEtapa1()     { crearHojaGrado('Primera Etapa de Primaria'); }
function crearHojaPrimariaEtapa2()     { crearHojaGrado('Segunda Etapa de Primaria'); }
function crearHojaBasicosEtapa1()      { crearHojaGrado('Primera Etapa de Básicos'); }
function crearHojaBasicosEtapa2()      { crearHojaGrado('Segunda Etapa de Básicos'); }
function crearHojaCuartoBachillerato() { crearHojaGrado('Cuarto Bachillerato'); }
function crearHojaQuintoBachillerato() { crearHojaGrado('Quinto Bachillerato'); }

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

  const NUM_COLS = 14;
  const MAX      = 500;

  // Encabezados fila 1 — 14 columnas
  const headers = [
    'Creamos ID',            // 1
    'Nombre Completo',       // 2
    'Nombre Preferido',      // 3
    'DPI / CUI',             // 4
    'Fecha de Nacimiento',   // 5
    'Edad',                  // 6
    'Género',                // 7
    'Teléfono',              // 8
    'Zona / Colonia',        // 9
    'Último Nivel Cursado',  // 10
    'Grado Asignado (Kobo)', // 11
    'Papelería Faltante',    // 12
    'Comentario',            // 13
    'Acción'                 // 14
  ];
  hoja.getRange(1, 1, 1, NUM_COLS)
    .setValues([headers])
    .setBackground(COLOR_HEADER_INTERES)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 36);

  // ── Anchos de columna (diseño compacto y legible) ──────────────────────
  hoja.setColumnWidth(COL_INTERES.CREAMOS_ID,   95);
  hoja.setColumnWidth(COL_INTERES.NOMBRE,       205);
  hoja.setColumnWidth(COL_INTERES.NOMBRE_PREF,  130);
  hoja.setColumnWidth(COL_INTERES.DPI,          115);
  hoja.setColumnWidth(COL_INTERES.FECHA_NAC,    105);
  hoja.setColumnWidth(COL_INTERES.EDAD,          45);
  hoja.setColumnWidth(COL_INTERES.GENERO,        72);
  hoja.setColumnWidth(COL_INTERES.TELEFONO,     105);
  hoja.setColumnWidth(COL_INTERES.ZONA,         160);
  hoja.setColumnWidth(COL_INTERES.ULTIMO_ANIO,  165);
  hoja.setColumnWidth(COL_INTERES.GRADO_KOBO,   175);
  hoja.setColumnWidth(COL_INTERES.PAPELERIA,    245);
  hoja.setColumnWidth(COL_INTERES.COMENTARIO,   180);
  hoja.setColumnWidth(COL_INTERES.ACCION,       185);

  // ── Alturas de fila uniformes (datos compactos) ─────────────────────────
  hoja.setRowHeights(2, MAX, 24);

  // ── Filas alternas para facilitar la lectura ────────────────────────────
  try { hoja.getBandings().forEach(function(b) { b.remove(); }); } catch(e) {}
  hoja.getRange(2, 1, MAX, NUM_COLS)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#FFFFFF')
    .setSecondRowColor('#F5F5F5');

  // ── Alineación y wrap de toda el área de datos ──────────────────────────
  const rangoData = hoja.getRange(2, 1, MAX, NUM_COLS);
  rangoData.setVerticalAlignment('middle')
           .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Columnas cortas/numéricas: centradas
  [COL_INTERES.CREAMOS_ID, COL_INTERES.DPI, COL_INTERES.FECHA_NAC,
   COL_INTERES.EDAD, COL_INTERES.GENERO, COL_INTERES.TELEFONO,
   COL_INTERES.GRADO_KOBO, COL_INTERES.ACCION].forEach(function(c) {
    hoja.getRange(2, c, MAX, 1).setHorizontalAlignment('center');
  });
  // Texto libre: izquierda
  [COL_INTERES.NOMBRE, COL_INTERES.NOMBRE_PREF, COL_INTERES.ZONA,
   COL_INTERES.ULTIMO_ANIO, COL_INTERES.PAPELERIA, COL_INTERES.COMENTARIO].forEach(function(c) {
    hoja.getRange(2, c, MAX, 1).setHorizontalAlignment('left');
  });

  // ── Colores especiales por columna (se aplican encima del banding) ───────
  hoja.getRange(2, COL_INTERES.PAPELERIA,  MAX, 1).setBackground('#E3F2FD'); // azul claro
  hoja.getRange(2, COL_INTERES.GRADO_KOBO, MAX, 1).setBackground('#F3E5F5'); // morado claro
  hoja.getRange(2, COL_INTERES.ACCION,     MAX, 1).setBackground('#FFFDE7'); // amarillo muy claro
  hoja.getRange(2, COL_INTERES.FECHA_NAC,  MAX, 1).setNumberFormat('dd/mm/yyyy');

  // ── Borde exterior de tabla ─────────────────────────────────────────────
  hoja.getRange(1, 1, MAX + 1, NUM_COLS)
    .setBorder(true, true, true, true, null, null,
      '#9E9E9E', SpreadsheetApp.BorderStyle.SOLID);
  // Línea inferior del encabezado más gruesa
  hoja.getRange(1, 1, 1, NUM_COLS)
    .setBorder(null, null, true, null, null, null,
      '#1565C0', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ── Validaciones ────────────────────────────────────────────────────────
  hoja.getRange(2, COL_INTERES.ULTIMO_ANIO, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(ULTIMO_ANIO_OPCIONES, true).setAllowInvalid(true)
      .setHelpText('Último nivel de estudios').build()
  );
  hoja.getRange(2, COL_INTERES.GRADO_KOBO, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(GRADOS, true).setAllowInvalid(true)
      .setHelpText('Grado que le corresponde según Creamos').build()
  );
  // setAllowInvalid(true) evita excepción al escribir desde script
  hoja.getRange(2, COL_INTERES.ACCION, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(ACCIONES, true).setAllowInvalid(true)
      .setHelpText('Selecciona el grado al que enviar al estudiante').build()
  );

  // ── Formato condicional ─────────────────────────────────────────────────
  hoja.setConditionalFormatRules([
    // Fila completa verde = ya procesada (✅)
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LEFT($N2,1)="✅"')
      .setBackground('#E8F5E9').setFontColor('#1B5E20')
      .setRanges([hoja.getRange(2, 1, MAX, NUM_COLS)])
      .build(),
    // Acción pendiente = amarillo en columna Acción
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($N2<>"",($N2<>"-- Seleccionar --"),LEFT($N2,1)<>"✅")')
      .setBackground('#FFF176').setFontColor('#F57F17')
      .setRanges([hoja.getRange(2, COL_INTERES.ACCION, MAX, 1)])
      .build()
  ]);

  // ── Notas de encabezado ─────────────────────────────────────────────────
  hoja.getRange(1, COL_INTERES.PAPELERIA).setNote(
    'Se construye automáticamente desde KoboToolbox.\n\n' +
    'O usa: DP Educación → Seleccionar papelería faltante\n\n' +
    'Documentos:\n' +
    PAPELERIA_OPCIONES.map(function(p, i) { return (i+1) + '. ' + p; }).join('\n')
  );
  hoja.getRange(1, COL_INTERES.GRADO_KOBO).setNote(
    'Grado que le corresponde según evaluación de Creamos.\n' +
    'Se completa automáticamente desde KoboToolbox.\n\n' +
    'Úsalo como referencia para seleccionar la Acción.'
  );
  hoja.getRange(1, COL_INTERES.ACCION).setNote(
    '🟡 Amarillo = acción pendiente de procesar\n' +
    '🟢 Verde = ya transferido a hoja de grado\n\n' +
    'Para transferir: DP Educación → 🔄 Procesar acciones pendientes'
  );

  SpreadsheetApp.getUi().alert(
    '✅ Hoja "Interés" configurada (14 columnas).\n\n' +
    'Pasos siguientes:\n' +
    '1. Crea los salones: menú → 📚 Crear hoja de grado\n' +
    '2. Instala el trigger: menú → 🔧 Instalar trigger automático\n' +
    '3. Sync Kobo: menú → 🌐 KoboToolbox → 🔄 Sync a hoja Interés\n' +
    '4. Revisa la columna "Acción" y transfiere alumnos.'
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

function _nombreHoja(grado, anio) {
  return grado + ' ' + (anio || SCHOOL_YEAR);
}

// Devuelve el siguiente grado en la secuencia, o null si es el último
function _siguienteGrado(grado) {
  const idx = GRADOS.indexOf(grado);
  return idx >= 0 && idx < GRADOS.length - 1 ? GRADOS[idx + 1] : null;
}

function crearHojaGrado(grado, silencioso, anio) {
  if (silencioso === undefined) silencioso = false;
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const nombre = _nombreHoja(grado, anio);  // usa anio si se especifica, si no SCHOOL_YEAR
  let   hoja   = ss.getSheetByName(nombre);
  const ui     = SpreadsheetApp.getUi();

  if (hoja) {
    if (!silencioso) {
      const r = ui.alert('Hoja existente',
        '"' + nombre + '" ya existe.\n¿Deseas reformatear sin borrar datos?',
        ui.ButtonSet.YES_NO);
      if (r !== ui.Button.YES) return hoja;
    }
    _formatearHojaGrado(hoja, grado, anio);
    if (!silencioso) ui.alert('✅ "' + nombre + '" reformateada.');
    return hoja;
  }

  hoja = ss.insertSheet(nombre);
  _formatearHojaGrado(hoja, grado, anio);
  if (!silencioso) ui.alert('✅ Hoja "' + nombre + '" creada.\n\nAño: ' + (anio || SCHOOL_YEAR));
  return hoja;
}

function _formatearHojaGrado(hoja, grado, anio) {
  anio = anio || SCHOOL_YEAR;
  const MAX     = 300;
  const esQuinto = grado === GRADO_GRADUACION;
  const estados  = esQuinto ? ESTADOS_QUINTO : ESTADOS;

  // Fila 1: título
  hoja.getRange(1, 1, 1, 9).merge()
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
  hoja.getRange(2, 1, 1, 9)
    .setValues([['ID','Creamos ID','Nombre Completo','DPI / CUI','No. Teléfono','Edad','Grado','Modalidad','Estado']])
    .setBackground(COLOR_HEADER_GRADO)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 30);
  hoja.setFrozenRows(2);

  // Anchos
  hoja.setColumnWidth(COL_GRADO.ID,         70);
  hoja.setColumnWidth(COL_GRADO.CREAMOS_ID, 110);
  hoja.setColumnWidth(COL_GRADO.NOMBRE,    220);
  hoja.setColumnWidth(COL_GRADO.DPI,       140);
  hoja.setColumnWidth(COL_GRADO.TELEFONO,  130);
  hoja.setColumnWidth(COL_GRADO.EDAD,       60);
  hoja.setColumnWidth(COL_GRADO.GRADO,     160);
  hoja.setColumnWidth(COL_GRADO.MODALIDAD, 145);
  hoja.setColumnWidth(COL_GRADO.ESTADO,    130);

  // Validaciones
  hoja.getRange(3, COL_GRADO.GRADO, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(GRADOS, true)
      .setAllowInvalid(false).setHelpText('Grado del estudiante').build());

  hoja.getRange(3, COL_GRADO.MODALIDAD, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MODALIDADES, true)
      .setAllowInvalid(false).setHelpText('Presencial o Semi-presencial').build());

  hoja.getRange(3, COL_GRADO.ESTADO, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(estados, true)
      .setAllowInvalid(false)
      .setHelpText(esQuinto
        ? 'Inscritx · Retiradx · Ciclo de Vida Terminado'
        : 'Inscritx · Retiradx · Graduadx').build());

  hoja.getRange(3, COL_GRADO.EDAD, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(5, 99)
      .setAllowInvalid(false).build());

  // Formato condicional por estado (colores Salesforce-friendly)
  // ESTADO está en col 9 = I
  hoja.clearConditionalFormatRules();
  const dr = hoja.getRange(3, 1, MAX, 9);
  hoja.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I3="Retiradx"')
      .setBackground('#FFCDD2').setRanges([dr]).build(),       // rojo
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I3="Graduadx"')
      .setBackground('#C8E6C9').setRanges([dr]).build(),       // verde
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I3="Inscritx"')
      .setBackground('#FFF9C4').setRanges([dr]).build(),       // amarillo
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I3="Ciclo de Vida Terminado"')
      .setBackground('#E8EAF6').setRanges([dr]).build()        // azul índigo
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
  const range     = e.range;
  const hoja      = range.getSheet();
  const nombreH   = hoja.getName();
  const col       = range.getColumn();
  const fila      = range.getRow();
  const valor     = (e.value || '').toString().trim();

  // ── Transferir desde Hoja de Interés → Lista de Espera ──────────────────
  if (nombreH === HOJA_INTERES && col === COL_INTERES.ACCION && fila >= 2) {
    if (valor === 'Enviar a: Lista de Espera') _transInteresAListaEspera(fila);
    return;
  }

  // ── Transferir desde Referencias a Educación → Lista de Espera ─────────
  if (nombreH === HOJA_REFERENCIAS && col === COL_REF.ACCION && fila >= 2) {
    if (valor === 'Enviar a: Lista de Espera') _transReferenciasAListaEspera(fila);
    return;
  }

  // ── Transferir desde Lista de Espera → Hoja de Grado ─────────────────────
  if (nombreH === HOJA_LISTA_ESPERA && col === COL_LISTA.ACCION && fila >= 2) {
    if (!valor || valor === '-- Seleccionar --') return;
    const grado = valor.replace('Enviar a: ', '').trim();
    if (GRADOS.indexOf(grado) >= 0) _transListaEsperaAGrado(fila, grado);
    return;
  }

  // ── Detectar cambio de Estado en hojas de grado ──────────────────────────
  if (col === COL_GRADO.ESTADO && fila >= 3) {
    // Ciclo de Vida Terminado → solo en Quinto Bachillerato → Seguimiento
    if (nombreH === _nombreHoja(GRADO_GRADUACION) && valor === 'Ciclo de Vida Terminado') {
      _ofrecerRegistrarGraduado(hoja, fila);
      return;
    }
    // Graduadx en cualquier otro grado → ofrecer avanzar al siguiente
    if (valor === 'Graduadx') {
      GRADOS.forEach(function(g) {
        if (nombreH === _nombreHoja(g) && g !== GRADO_GRADUACION) {
          _ofrecerAvanzarSiguienteGrado(hoja, fila, g);
        }
      });
    }
  }
}

function _transferirEstudiante(fila, grado) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInteres = ss.getSheetByName(HOJA_INTERES);
  // Lee las 14 columnas de la fila
  const datos = hojaInteres.getRange(fila, 1, 1, 14).getValues()[0];

  const creamosId  = datos[COL_INTERES.CREAMOS_ID  - 1];
  const nombre     = datos[COL_INTERES.NOMBRE      - 1];
  const dpi        = datos[COL_INTERES.DPI         - 1];
  const edad       = datos[COL_INTERES.EDAD        - 1];
  const telefono   = datos[COL_INTERES.TELEFONO    - 1];
  const papeleria  = datos[COL_INTERES.PAPELERIA   - 1];
  const comentario = datos[COL_INTERES.COMENTARIO  - 1];

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

  hojaGrado.getRange(filaDestino, 1, 1, 9).setValues([[
    id, creamosId || '', nombre, dpi,
    telefono || '',   // Teléfono copiado desde Interés
    edad, grado,
    'Presencial',     // Modalidad por defecto
    'Inscritx'        // Estado por defecto
  ]]);

  hojaGrado.getRange(filaDestino, 1, 1, 9)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  hojaGrado.getRange(filaDestino, COL_GRADO.NOMBRE).setHorizontalAlignment('left');
  hojaGrado.getRange(filaDestino, COL_GRADO.DPI).setHorizontalAlignment('left');
  hojaGrado.setRowHeight(filaDestino, 26);

  if (papeleria || comentario) {
    hojaGrado.getRange(filaDestino, COL_GRADO.NOMBRE).setNote(
      [papeleria  ? 'Papelería faltante: ' + papeleria  : '',
       comentario ? 'Comentario: '         + comentario : '']
      .filter(Boolean).join('\n')
    );
  }

  // Marcar fila origen como procesada (fondo verde + ✅ en Acción)
  hojaInteres.getRange(fila, 1, 1, 14).setBackground('#E8F5E9');
  hojaInteres.getRange(fila, COL_INTERES.ACCION)
    .setValue('✅ ' + grado)
    .setDataValidation(null);

  ss.toast('"' + nombre + '" → "' + nombreHoja + '" · ID: ' + id, '✅ Estudiante transferido', 5);
}

// ── Ofrecer avanzar al siguiente grado (al marcar Graduadx) ──────────────────
// Al graduarse, el estudiante pasa al SIGUIENTE grado en el AÑO SIGUIENTE.
// La fila actual queda oculta (el registro histórico no se borra).
function _ofrecerAvanzarSiguienteGrado(hojaActual, fila, gradoActual) {
  const siguienteGrado = _siguienteGrado(gradoActual);
  if (!siguienteGrado) return; // no hay siguiente

  const ui          = SpreadsheetApp.getUi();
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const datos       = hojaActual.getRange(fila, 1, 1, 9).getValues()[0];
  const nombre      = datos[COL_GRADO.NOMBRE - 1];
  const anioSig     = SCHOOL_YEAR + 1; // siempre el año siguiente

  const r = ui.alert(
    '🎓 ¿Avanzar al siguiente grado?',
    '"' + nombre + '" fue marcado como Graduadx en\n"' + gradoActual + ' ' + SCHOOL_YEAR + '".\n\n' +
    '¿Crear inscripción en:\n"' + siguienteGrado + ' ' + anioSig + '"?\n\n' +
    '(La fila actual quedará oculta — los datos se conservan)',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) return;

  // Crear la hoja del siguiente grado (año siguiente) si no existe
  const nombreSig = _nombreHoja(siguienteGrado, anioSig);
  let   hojaSig   = ss.getSheetByName(nombreSig);
  if (!hojaSig) hojaSig = crearHojaGrado(siguienteGrado, true, anioSig);

  const filaDestino = Math.max(hojaSig.getLastRow() + 1, 3);
  const id          = _siguienteId(hojaSig, siguienteGrado);

  hojaSig.getRange(filaDestino, 1, 1, 9).setValues([[
    id,
    datos[COL_GRADO.CREAMOS_ID - 1] || '',
    datos[COL_GRADO.NOMBRE     - 1],
    datos[COL_GRADO.DPI        - 1],
    datos[COL_GRADO.TELEFONO   - 1],
    datos[COL_GRADO.EDAD       - 1],
    siguienteGrado,
    datos[COL_GRADO.MODALIDAD  - 1] || 'Presencial',
    'Inscritx'
  ]]);
  hojaSig.getRange(filaDestino, 1, 1, 9)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  hojaSig.getRange(filaDestino, COL_GRADO.NOMBRE).setHorizontalAlignment('left');
  hojaSig.setRowHeight(filaDestino, 26);

  // Ocultar la fila original (historial conservado, vista limpia)
  hojaActual.hideRows(fila);

  ss.toast('"' + nombre + '" → "' + nombreSig + '" · ID: ' + id, '✅ Avanzado', 5);
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
    if (v === 'Enviar a: Lista de Espera') {
      _transInteresAListaEspera(idx + 2);
      procesados++;
    }
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
        if (e === 'Inscritx')                oyentes++;
        if (e === 'Retiradx')                deserc++;
        if (e === 'Graduadx')                grad++;
        if (e === 'Ciclo de Vida Terminado') grad++;
      });
    }

    texto += '\n' + grado + '  (' + total + ' alumnos)\n';
    texto += '   Inscritx: ' + oyentes + '  |  Retiradx: ' + deserc + '  |  Graduadx/CVT: ' + grad + '\n';
  });

  SpreadsheetApp.getUi().alert('Resumen de Alumnos', texto, SpreadsheetApp.getUi().ButtonSet.OK);
}


// ── Cerrar ciclo escolar ──────────────────────────────────────────────────────
//  Para cada hoja de grado del año anterior:
//  - Estudiantes con 'Inscritx' o 'Retiradx' → nueva inscripción en el
//    mismo grado pero con SCHOOL_YEAR y la fila anterior queda oculta
//  - Estudiantes 'Graduadx' / 'Ciclo de Vida Terminado' → ya se procesaron
//    por el trigger onEdit; sus filas quedan ocultas también
function cerrarCicloEscolar() {
  const ui      = SpreadsheetApp.getUi();
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const anioAnt = SCHOOL_YEAR - 1;  // ciclo que se cierra

  const r = ui.alert(
    '📅 Cerrar ciclo ' + anioAnt,
    'Esto procesará todas las hojas de grado del ciclo ' + anioAnt + ':\n\n' +
    '• Estudiantes Inscritx/Retiradx → se crean en la hoja ' + SCHOOL_YEAR + ' del mismo grado\n' +
    '• Filas procesadas quedan ocultas en la hoja ' + anioAnt + '\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) return;

  let totalMovidos = 0, totalHojasArchivadas = 0;

  GRADOS.forEach(function(grado) {
    const nombreAnt = _nombreHoja(grado, anioAnt);
    const hojaAnt   = ss.getSheetByName(nombreAnt);
    if (!hojaAnt) return; // no existía este grado el año pasado

    const ultimaFila = hojaAnt.getLastRow();
    if (ultimaFila < 3) {
      // Hoja vacía → archivar igualmente
      hojaAnt.hideSheet();
      totalHojasArchivadas++;
      return;
    }

    const datos = hojaAnt.getRange(3, 1, ultimaFila - 2, 9).getValues();

    datos.forEach(function(row, i) {
      const estado = String(row[COL_GRADO.ESTADO - 1] || '').trim();
      const nombre = String(row[COL_GRADO.NOMBRE - 1] || '').trim();
      if (!nombre) return;

      // Estudiantes que no completaron el ciclo → mover al mismo grado año nuevo
      if (estado === 'Inscritx' || estado === 'Retiradx') {
        const nombreNvo = _nombreHoja(grado, SCHOOL_YEAR);
        let   hojaNva   = ss.getSheetByName(nombreNvo);
        if (!hojaNva) hojaNva = crearHojaGrado(grado, true);

        const filaDestino = Math.max(hojaNva.getLastRow() + 1, 3);
        const idNvo       = _siguienteId(hojaNva, grado);

        hojaNva.getRange(filaDestino, 1, 1, 9).setValues([[
          idNvo,
          row[COL_GRADO.CREAMOS_ID - 1] || '',
          row[COL_GRADO.NOMBRE     - 1],
          row[COL_GRADO.DPI        - 1],
          row[COL_GRADO.TELEFONO   - 1],
          row[COL_GRADO.EDAD       - 1],
          grado,
          row[COL_GRADO.MODALIDAD  - 1] || 'Presencial',
          'Inscritx'
        ]]);
        hojaNva.getRange(filaDestino, 1, 1, 9)
          .setVerticalAlignment('middle').setHorizontalAlignment('center');
        hojaNva.getRange(filaDestino, COL_GRADO.NOMBRE).setHorizontalAlignment('left');
        hojaNva.setRowHeight(filaDestino, 26);
        totalMovidos++;
      }
      // Nota: filas Graduadx/CVT ya quedaron ocultas por el trigger onEdit
    });

    // Archivar (ocultar tab) la hoja del año anterior — datos conservados
    hojaAnt.hideSheet();
    totalHojasArchivadas++;
  });

  ui.alert(
    '✅ Ciclo ' + anioAnt + ' archivado\n\n' +
    'Estudiantes inscritos en ' + SCHOOL_YEAR + ': ' + totalMovidos + '\n' +
    'Hojas archivadas (ocultas): ' + totalHojasArchivadas + '\n\n' +
    '💡 Las hojas de ' + anioAnt + ' siguen accesibles:\n' +
    '   Clic derecho en una pestaña → "Mostrar hojas".\n\n' +
    '📌 Recuerda actualizar SCHOOL_YEAR a ' + SCHOOL_YEAR +
    '\n   en el código si todavía no lo has hecho.'
  );
}


// ── Reiniciar sistema (eliminar todo y reinstalar) ────────────────────────────
//  Útil para instalar actualizaciones del script o empezar de cero.
//  CUIDADO: elimina permanentemente todos los datos de hojas DP Educación.
function reiniciarSistema() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Confirmación doble por ser operación destructiva
  const r1 = ui.alert(
    '⚠️ REINICIAR SISTEMA',
    'Esta acción eliminará PERMANENTEMENTE:\n\n' +
    '• Hoja de Interés (todos los registros)\n' +
    '• Todas las hojas de grado (todos los años)\n' +
    '• Hoja Seguimiento Graduados\n' +
    '• Hoja Referencias a Educación\n' +
    '• Hoja Lista de Espera\n\n' +
    '⚠️ Los datos NO se pueden recuperar después.\n\n' +
    '¿Estás seguro de que quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (r1 !== ui.Button.YES) { ui.alert('Operación cancelada.'); return; }

  const r2 = ui.prompt(
    '🔐 Confirmación final',
    'Escribe exactamente  REINICIAR  para confirmar:',
    ui.ButtonSet.OK_CANCEL
  );
  if (r2.getSelectedButton() !== ui.Button.OK) { ui.alert('Operación cancelada.'); return; }
  if ((r2.getResponseText() || '').trim() !== 'REINICIAR') {
    ui.alert('❌ Texto incorrecto. Operación cancelada.'); return;
  }

  // 1. Eliminar todos los triggers del proyecto
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // 2. Identificar hojas a eliminar (Interés + grados todos los años + Seguimiento + Referencias + Lista de Espera)
  const hojas       = ss.getSheets();
  const nombresDP   = [HOJA_INTERES, HOJA_SEGUIMIENTO, HOJA_REFERENCIAS, HOJA_LISTA_ESPERA];
  // Incluir cualquier hoja cuyo nombre empiece con un grado conocido
  hojas.forEach(function(h) {
    const n = h.getName();
    GRADOS.forEach(function(g) {
      if (n.startsWith(g)) nombresDP.push(n);
    });
  });
  const nombresUnicos = [...new Set(nombresDP)];

  // Necesitamos conservar al menos una hoja en el spreadsheet
  const hojasTotales     = ss.getSheets().length;
  const hojasAEliminar   = ss.getSheets().filter(function(h) {
    return nombresUnicos.indexOf(h.getName()) >= 0;
  });
  const hojasSobreviven  = hojasTotales - hojasAEliminar.length;

  // Crear hoja temporal SIEMPRE antes de eliminar nada.
  // Google Sheets exige al menos una hoja visible en todo momento.
  const hojaTemporal = ss.insertSheet('_reinstalando_');

  hojasAEliminar.forEach(function(h) {
    try { ss.deleteSheet(h); } catch(e) { /* hoja ya eliminada o protegida */ }
  });

  // Reinstalar todo (crea Interés + grados + Seguimiento + Referencias + Lista de Espera + triggers)
  setupHojaInteres();
  crearTodasLasHojas();
  setupHojaSeguimiento();
  setupHojaReferencias();
  setupHojaListaEspera();
  installTriggers();

  // Ahora que existen las hojas nuevas, eliminar la temporal
  try { ss.deleteSheet(hojaTemporal); } catch(e) { /* ignorar */ }

  ui.alert(
    '✅ Sistema reiniciado\n\n' +
    'Se crearon de nuevo:\n' +
    '• Hoja de Interés\n' +
    '• ' + GRADOS.length + ' hojas de grado (' + SCHOOL_YEAR + ')\n' +
    '• Hoja Seguimiento Graduados\n' +
    '• Hoja Referencias a Educación\n' +
    '• Hoja Lista de Espera\n' +
    '• Triggers automáticos\n\n' +
    '💡 Recuerda configurar el token de KoboToolbox en:\n' +
    '   🌐 KoboToolbox → 🔑 Configurar token de API'
  );
}


// ── Guía de uso ───────────────────────────────────────────────────────────────
function mostrarGuiaDeUso() {
  crearHojaGuia();
}

// ── Hoja visual "📖 Guía de Uso" ─────────────────────────────────────────────
function crearHojaGuia() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const NOMBRE = '📖 Guía de Uso';

  let hoja = ss.getSheetByName(NOMBRE);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE, 0);
  } else {
    hoja.clearContents();
    hoja.clearFormats();
  }

  function titulo(fila, texto, bg, fg) {
    const r = hoja.getRange(fila, 1, 1, 6);
    r.merge().setValue(texto)
      .setBackground(bg || '#1565C0').setFontColor(fg || '#FFFFFF')
      .setFontSize(13).setFontWeight('bold')
      .setVerticalAlignment('middle').setWrap(true);
    hoja.setRowHeight(fila, 32);
  }
  function subtitulo(fila, texto, bg) {
    const r = hoja.getRange(fila, 1, 1, 6);
    r.merge().setValue(texto)
      .setBackground(bg || '#E3F2FD').setFontColor('#0D47A1')
      .setFontSize(11).setFontWeight('bold')
      .setVerticalAlignment('middle').setWrap(true);
    hoja.setRowHeight(fila, 26);
  }
  function fila2(fila, colA, colB, bgA, bgB) {
    hoja.getRange(fila, 1, 1, 2).merge().setValue(colA)
      .setBackground(bgA || '#F5F5F5').setFontColor('#212121')
      .setFontSize(10).setVerticalAlignment('middle').setWrap(true);
    hoja.getRange(fila, 3, 1, 4).merge().setValue(colB)
      .setBackground(bgB || '#FFFFFF').setFontColor('#424242')
      .setFontSize(10).setVerticalAlignment('middle').setWrap(true);
    hoja.setRowHeight(fila, 22);
  }
  function filaC(fila, texto, bg, bold) {
    hoja.getRange(fila, 1, 1, 6).merge().setValue(texto)
      .setBackground(bg || '#FFFFFF').setFontColor('#424242')
      .setFontSize(10).setFontWeight(bold ? 'bold' : 'normal')
      .setVerticalAlignment('middle').setWrap(true);
    hoja.setRowHeight(fila, 20);
  }
  function esp(fila) {
    hoja.getRange(fila, 1, 1, 6).merge().setBackground('#FFFFFF');
    hoja.setRowHeight(fila, 10);
  }

  hoja.setColumnWidth(1, 200);
  hoja.setColumnWidth(2, 110);
  hoja.setColumnWidth(3, 110);
  hoja.setColumnWidth(4, 110);
  hoja.setColumnWidth(5, 110);
  hoja.setColumnWidth(6, 110);

  let f = 1;

  // ── ENCABEZADO ────────────────────────────────────────────────────────────
  titulo(f++, '📖  GUÍA DE USO — SISTEMA DP EDUCACIÓN', '#0D47A1', '#FFFFFF');
  filaC(f++, 'Esta hoja explica para qué sirve cada pestaña, qué hace cada columna y cómo usar el sistema paso a paso.', '#E8EAF6');
  esp(f++);

  // ── SECCIÓN 1: LAS HOJAS ──────────────────────────────────────────────────
  titulo(f++, '🗂️  LAS HOJAS DEL SISTEMA', '#1565C0', '#FFFFFF');
  esp(f++);

  subtitulo(f++, '📋  Hoja de Interés  —  BANDEJA DE ENTRADA', '#E3F2FD');
  filaC(f++, 'Aquí llegan las personas del formulario KoboToolbox que marcaron SÍ en "¿Deseas inscribirte en Educación?". Nadie de otros programas (Inclusión Laboral, etc.) entra. Desde acá se transfieren a su hoja de grado usando la columna "Acción".', '#FAFAFA');
  esp(f++);
  subtitulo(f++, '   Columnas de "Hoja de Interés"', '#EDE7F6');
  fila2(f++, '  Col 1 · ID', 'Número interno. No editar.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 2 · Creamos ID', 'ID de KoboToolbox / Salesforce.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 3 · Nombre', 'Nombre completo del participante.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 4 · DPI', 'DPI. Evita duplicados: si el DPI ya existe no se importa de nuevo.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 5 · Fecha de Nac.', 'Tal como llegó del formulario.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 6 · Edad', 'Calculada automáticamente. No editar.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 7 · Género', 'Género declarado en el formulario.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 8 · Teléfono', 'Número de contacto.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 9 · Zona/Colonia', 'Lugar de residencia.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 10 · Último nivel', 'Último grado completado antes de entrar.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 11 · Grado Kobo', 'Grado sugerido por KoboToolbox según nivel declarado.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 12 · Papelería', 'Documentos faltantes. Se llena con el selector del menú.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  Col 13 · Comentario', 'Notas de seguimiento sobre papelería u otro asunto.', '#F3E5F5', '#FAFAFA');
  fila2(f++, '  ⬅ Col 14 · Acción  (LA MÁS IMPORTANTE)', 'Selecciona "Enviar a: [Grado]" → el participante se mueve automáticamente a esa hoja de grado. Cuando ya fue enviado muestra "✅ [Grado]" en gris.', '#CE93D8', '#F3E5F5');
  esp(f++);

  subtitulo(f++, '📚  Hojas de Grado  (ej. "Primera Etapa de Primaria 2025")', '#E8F5E9');
  filaC(f++, 'Una hoja por cada nivel y año. Al enviar a alguien desde "Hoja de Interés" aparece aquí como "Inscritx". Es el registro activo de cada grado.', '#FAFAFA');
  esp(f++);
  subtitulo(f++, '   Columnas de cada Hoja de Grado', '#DCEDC8');
  fila2(f++, '  Col 1 · ID', 'ID dentro de esta hoja.', '#F1F8E9', '#FAFAFA');
  fila2(f++, '  Col 2 · Creamos ID', 'Mismo ID de Interés.', '#F1F8E9', '#FAFAFA');
  fila2(f++, '  Col 3 · Nombre', 'Nombre completo.', '#F1F8E9', '#FAFAFA');
  fila2(f++, '  Col 4 · DPI', 'Documento de identidad.', '#F1F8E9', '#FAFAFA');
  fila2(f++, '  Col 5 · Teléfono', 'Contacto.', '#F1F8E9', '#FAFAFA');
  fila2(f++, '  Col 6 · Edad', 'Edad del participante.', '#F1F8E9', '#FAFAFA');
  fila2(f++, '  Col 7 · Grado', 'Nombre del nivel (ej. Primera Etapa de Primaria).', '#F1F8E9', '#FAFAFA');
  fila2(f++, '  Col 8 · Modalidad', '"Presencial" o "Semi-presencial". Selección manual.', '#F1F8E9', '#FAFAFA');
  fila2(f++, '  ⬅ Col 9 · Estado  (LA MÁS IMPORTANTE)', '"Inscritx" = activo  |  "Retiradx" = se fue (fila se oculta)  |  "Graduadx" = completó el grado, el sistema pregunta si avanzar al año siguiente  |  "Ciclo de Vida Terminado" (solo Quinto Bach.) = pasa a Seguimiento Graduados.', '#A5D6A7', '#E8F5E9');
  esp(f++);

  subtitulo(f++, '🎓  Seguimiento Graduados', '#FFF3E0');
  filaC(f++, 'Para quienes terminaron Quinto Bachillerato (último grado). Permite dar seguimiento de qué pasó con ellos después: empleo, estudios, etc.', '#FAFAFA');
  esp(f++);
  subtitulo(f++, '   Columnas de "Seguimiento Graduados"', '#FFE0B2');
  fila2(f++, '  Col 1 · ID', 'ID autoasignado.', '#FFF8E1', '#FAFAFA');
  fila2(f++, '  Col 2 · Nombre', 'Nombre completo del graduado.', '#FFF8E1', '#FAFAFA');
  fila2(f++, '  Col 3 · DPI', 'Documento de identidad.', '#FFF8E1', '#FAFAFA');
  fila2(f++, '  Col 4 · Teléfono', 'Contacto.', '#FFF8E1', '#FAFAFA');
  fila2(f++, '  Col 5 · Edad', 'Edad al momento del registro.', '#FFF8E1', '#FAFAFA');
  fila2(f++, '  Col 6 · Año de Graduación', 'Año en que completó Quinto Bachillerato.', '#FFF8E1', '#FAFAFA');
  fila2(f++, '  Col 7 · Estado Post-Grad', '"Seguimiento activo"  "Empleado"  "Continúa estudiando"  "Sin contacto"  "Emigró"', '#FFCC80', '#FFF8E1');
  fila2(f++, '  Col 8 · Observación', 'Notas libres de seguimiento.', '#FFF8E1', '#FAFAFA');
  esp(f++);

  // ── SECCIÓN 2: EL MENÚ ───────────────────────────────────────────────────
  titulo(f++, '🍽️  EL MENÚ "DP Educación" — QUÉ HACE CADA OPCIÓN', '#1B5E20', '#FFFFFF');
  esp(f++);
  fila2(f++, '  ⚙️ Configurar hoja Interés', 'Crea o resetea la Hoja de Interés con encabezados y validaciones. Usar al inicio del año o si la hoja quedó dañada.', '#E3F2FD', '#FAFAFA');
  fila2(f++, '  🔧 Instalar trigger', 'Activa el detector de edición (onEdit). Necesario para que la columna "Acción" transfiera sola. Ejecutar una sola vez.', '#E3F2FD', '#FAFAFA');
  fila2(f++, '  📚 Crear hoja de grado', 'Crea la hoja de un grado para el año actual. "✨ Crear TODOS" crea los 6 grados de golpe.', '#E3F2FD', '#FAFAFA');
  fila2(f++, '  📋 Seleccionar papelería', 'Marca los documentos que le faltan al participante. Se guardan en la columna "Papelería".', '#E3F2FD', '#FAFAFA');
  fila2(f++, '  🔄 Procesar acciones', 'Transfiere en lote a todos los de "Hoja de Interés" que tengan grado seleccionado pero aún no enviados.', '#E3F2FD', '#FAFAFA');
  fila2(f++, '  📊 Ver resumen', 'Muestra conteo de estudiantes por grado y estado.', '#E3F2FD', '#FAFAFA');
  fila2(f++, '  📅 Cerrar ciclo escolar', 'Fin de año: mueve Inscritx/Retiradx al año nuevo y oculta hojas anteriores. Actualiza SCHOOL_YEAR antes de usarlo.', '#E3F2FD', '#FAFAFA');
  fila2(f++, '  🎓 Seguimiento', '"Configurar" crea la hoja. "Registrar graduado manual" agrega a alguien que ya terminó sin estar en el sistema.', '#E3F2FD', '#FAFAFA');
  fila2(f++, '  🔑 KoboToolbox · Token', 'Guarda el token de API. Necesario la primera vez o si el token expira.', '#EDE7F6', '#FAFAFA');
  fila2(f++, '  📦 KoboToolbox · Histórico', 'Importa registros del formulario antiguo. Solo usar una vez para migración inicial.', '#EDE7F6', '#FAFAFA');
  fila2(f++, '  🔄 KoboToolbox · Sync manual', 'Importa los registros nuevos del formulario actual. Solo entran personas con Educación = Sí.', '#EDE7F6', '#FAFAFA');
  fila2(f++, '  🔁 Sync automático (c/min)  ⬅', 'Activa sincronización cada minuto. KoboToolbox se revisa solo sin hacer nada.', '#CE93D8', '#F3E5F5');
  fila2(f++, '  ⛔ Detener sync automático', 'Desactiva el sync cada minuto.', '#EDE7F6', '#FAFAFA');
  esp(f++);

  // ── SECCIÓN 3: FLUJO PASO A PASO ─────────────────────────────────────────
  titulo(f++, '🔵  FLUJO COMPLETO PASO A PASO', '#1565C0', '#FFFFFF');
  esp(f++);

  subtitulo(f++, 'PASO 1 — Configuración inicial (solo una vez al año)', '#E3F2FD');
  filaC(f++, '1a.  Menú → ⚙️ Configurar hoja Interés', '#FAFAFA');
  filaC(f++, '1b.  Menú → 📚 Crear hoja de grado → ✨ Crear TODOS los grados', '#FAFAFA');
  filaC(f++, '1c.  Menú → 🔧 Instalar trigger automático', '#FAFAFA');
  filaC(f++, '1d.  Menú → 🌐 KoboToolbox → 🔑 Configurar token de API', '#FAFAFA');
  filaC(f++, '1e.  Menú → 🌐 KoboToolbox → 🔁 Sync automático (cada minuto)  ← activa la magia', '#FFF9C4');
  esp(f++);

  subtitulo(f++, 'PASO 2 — Llegan inscritos nuevos (automático)', '#E8F5E9');
  filaC(f++, 'El sync automático jala los datos solos cada minuto. También puedes hacerlo manualmente:', '#FAFAFA');
  filaC(f++, '2a.  Menú → 🌐 KoboToolbox → 🔄 Sync → hoja Interés (actual)', '#FAFAFA');
  filaC(f++, '2b.  Solo personas con "¿Deseas inscribirte en Educación? = Sí" aparecen en Hoja de Interés', '#FAFAFA');
  filaC(f++, '2c.  La columna "Grado Kobo" ya sugiere a qué grado va cada quien', '#FAFAFA');
  esp(f++);

  subtitulo(f++, 'PASO 3 — Transferir participantes a su grado', '#FFF3E0');
  filaC(f++, '3a.  En "Hoja de Interés", busca la fila del participante', '#FAFAFA');
  filaC(f++, '3b.  En columna "Acción" (última columna), selecciona "Enviar a: [Grado]"', '#FAFAFA');
  filaC(f++, '3c.  El participante aparece automáticamente en su hoja de grado como "Inscritx"', '#FAFAFA');
  filaC(f++, '3d.  La columna "Acción" cambia a "✅ [Grado]" — ya fue enviado', '#FAFAFA');
  filaC(f++, '  →  También puedes ir a Menú → 🔄 Procesar acciones para enviar a todos de golpe', '#F5F5F5');
  esp(f++);

  subtitulo(f++, 'PASO 4 — Durante el ciclo escolar', '#EDE7F6');
  filaC(f++, '4a.  En la hoja de grado, cambia el "Estado" de cada estudiante según corresponda', '#FAFAFA');
  filaC(f++, '4b.  "Inscritx" = activo  |  "Retiradx" = se fue  |  "Graduadx" = completó el grado', '#FAFAFA');
  filaC(f++, '4c.  Al poner "Graduadx" el sistema pregunta si avanzarlo al mismo grado en el año siguiente', '#FAFAFA');
  filaC(f++, '4d.  En Quinto Bachillerato: "Ciclo de Vida Terminado" → pasa a Seguimiento Graduados', '#FAFAFA');
  esp(f++);

  subtitulo(f++, 'PASO 5 — Fin del ciclo escolar', '#DCEDC8');
  filaC(f++, '5a.  Actualiza SCHOOL_YEAR en el código de Apps Script al año nuevo', '#FAFAFA');
  filaC(f++, '5b.  Menú → 📅 Cerrar ciclo escolar', '#FAFAFA');
  filaC(f++, '5c.  Los "Inscritx"/"Retiradx" pasan a hojas del año nuevo. Los "Graduadx" al siguiente grado.', '#FAFAFA');
  filaC(f++, '5d.  Las hojas del año anterior se ocultan (el historial NO se borra)', '#FAFAFA');
  filaC(f++, '  →  Para ver hojas ocultas: clic derecho en cualquier pestaña → "Mostrar hojas"', '#F5F5F5');
  esp(f++);

  // ── SECCIÓN 4: ESTADOS ───────────────────────────────────────────────────
  titulo(f++, '🟡  ESTADOS Y QUÉ SIGNIFICAN', '#F57F17', '#FFFFFF');
  esp(f++);
  fila2(f++, '  Inscritx',                'Estudiante activo en el grado.',                                                 '#A5D6A7', '#E8F5E9');
  fila2(f++, '  Retiradx',                'Se retiró del programa. La fila se oculta pero no se borra.',                    '#EF9A9A', '#FFEBEE');
  fila2(f++, '  Graduadx',                'Completó el grado. Se pregunta si avanzar al año siguiente.',                   '#81D4FA', '#E1F5FE');
  fila2(f++, '  Ciclo de Vida Terminado', 'Solo Quinto Bachillerato. Completó todo el programa → pasa a Seguimiento.',     '#CE93D8', '#F3E5F5');
  esp(f++);

  // ── SECCIÓN 5: ORDEN GRADOS ──────────────────────────────────────────────
  titulo(f++, '🟢  ORDEN DE LOS GRADOS', '#2E7D32', '#FFFFFF');
  esp(f++);
  fila2(f++, '  1°  Primera Etapa de Primaria',  'Nivel más básico.',                                            '#C8E6C9', '#F1F8E9');
  fila2(f++, '  2°  Segunda Etapa de Primaria', '',                                                              '#C8E6C9', '#F1F8E9');
  fila2(f++, '  3°  Primera Etapa de Básicos',  '',                                                              '#A5D6A7', '#E8F5E9');
  fila2(f++, '  4°  Segunda Etapa de Básicos',  '',                                                              '#A5D6A7', '#E8F5E9');
  fila2(f++, '  5°  Cuarto Bachillerato',        '',                                                              '#81C784', '#E8F5E9');
  fila2(f++, '  6°  Quinto Bachillerato',        'Último grado → al terminar pasa a Seguimiento Graduados.',    '#4CAF50', '#E8F5E9');
  esp(f++);

  // ── SECCIÓN 6: TIPS ──────────────────────────────────────────────────────
  titulo(f++, '💡  TIPS Y PROBLEMAS COMUNES', '#4A148C', '#FFFFFF');
  esp(f++);
  subtitulo(f++, 'Tips generales', '#EDE7F6');
  filaC(f++, '✔  El sync filtra SOLO personas de Educación (marcaron Sí). Nadie de Inclusión Laboral ni otros programas entra.', '#FAFAFA');
  filaC(f++, '✔  No se crean duplicados: el sistema compara por DPI antes de importar.', '#FAFAFA');
  filaC(f++, '✔  Creamos ID viene de KoboToolbox para mantener sincronía con Salesforce.', '#FAFAFA');
  filaC(f++, '✔  Los comentarios de papelería se guardan como nota flotante en la celda del nombre.', '#FAFAFA');
  filaC(f++, '✔  El historial NUNCA se borra, solo se oculta. Siempre puedes ver hojas y filas anteriores.', '#FAFAFA');
  esp(f++);
  subtitulo(f++, 'Problemas comunes y soluciones', '#FFCDD2');
  fila2(f++, '  No llegan datos del sync',       'Revisar: 1) Token de API configurado  2) Sync automático activo  3) Conexión a internet', '#FFEBEE', '#FAFAFA');
  fila2(f++, '  La "Acción" no transfiere solo', 'Reinstalar el trigger: Menú → 🔧 Instalar trigger automático', '#FFEBEE', '#FAFAFA');
  fila2(f++, '  Hoja de grado no existe',         'Crearla: Menú → 📚 Crear hoja de grado → elegir el grado', '#FFEBEE', '#FAFAFA');
  fila2(f++, '  Sistema dañado o raro',           'Menú → 🔁 Reiniciar sistema  (PRECAUCIÓN: borra y recrea todo)', '#FFCDD2', '#FFEBEE');
  esp(f++);

  titulo(f++, '📖  Para actualizar esta guía: menú → 📖 Guía de uso', '#37474F', '#FFFFFF');

  hoja.setHiddenGridlines(true);
  SpreadsheetApp.setActiveSheet(hoja);
  ss.toast('Hoja "📖 Guía de Uso" lista.', '📖 Guía de Uso', 5);
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
//  SECCIÓN 8 · INTEGRACIÓN KOBOTOOLBOX → HOJA "INTERÉS"
// ────────────────────────────────────────────────────────────────────────────
//
//  Flujo de uso:
//  1. Menú → 🌐 KoboToolbox → 🔑 Configurar token de API  (una sola vez)
//  2. Menú → 🌐 KoboToolbox → 🔄 Sync → hoja Interés
//     ↳ Solo registros con "¿Deseas inscribirte en Educación?" = Sí
//     ↳ Agrega solo registros nuevos (deduplica por DPI)
//     ↳ Las 13 columnas se llenan automáticamente
//  3. Opcional: 🔁 Sync automático (cada hora)
//
// ────────────────────────────────────────────────────────────────────────────

// ════════════════════════════════════════════════════════════════════════════
//  SECCIÓN 9 · REFERENCIAS A EDUCACIÓN & LISTA DE ESPERA
// ════════════════════════════════════════════════════════════════════════════

// ── 9.1  Configurar hoja "Referencias a Educación" ───────────────────────────
function setupHojaReferencias() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ui   = SpreadsheetApp.getUi();
  const ENCABEZADOS = [
    'Creamos ID', 'Nombre Completo', 'Nombre Preferido', 'DPI / CUI',
    'Fecha de Nacimiento', 'Edad', 'Género', 'Teléfono',
    'Zona / Colonia', 'Último Nivel Cursado',
    'Fecha de Referencia', 'Responsable de Referencia', 'Grado de Interés',
    'Acción'
  ];
  const ACCIONES_REF = ['-- Seleccionar --', 'Enviar a: Lista de Espera'];

  let hoja = ss.getSheetByName(HOJA_REFERENCIAS);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_REFERENCIAS);
  } else {
    hoja.clearContents();
    hoja.clearFormats();
    try { hoja.clearDataValidations(); } catch(e) {}
  }

  // Fila 1 — encabezados
  const rHead = hoja.getRange(1, 1, 1, ENCABEZADOS.length);
  rHead.setValues([ENCABEZADOS])
    .setBackground(COLOR_HEADER_REFERENCIAS).setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 32);

  // Anchos de columna
  [180, 220, 150, 130, 130, 60, 90, 110, 150, 200, 200].forEach(function(w, i) {
    hoja.setColumnWidth(i + 1, w);
  });

  // Validación Acción (columna 11)
  const valAccion = SpreadsheetApp.newDataValidation()
    .requireValueInList(ACCIONES_REF, true).setAllowInvalid(false).build();
  hoja.getRange(2, COL_REF.ACCION, hoja.getMaxRows() - 1, 1).setDataValidation(valAccion);

  // Validación Género (columna 7)
  const valGen = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Hombre', 'Mujer', 'Otro'], true).setAllowInvalid(true).build();
  hoja.getRange(2, COL_REF.GENERO, hoja.getMaxRows() - 1, 1).setDataValidation(valGen);

  ui.alert(
    '✅ Hoja "' + HOJA_REFERENCIAS + '" configurada.\n\n' +
    'Columnas: Creamos ID · Nombre Completo · Nombre Preferido · DPI · Fecha Nac · Edad · Género · Teléfono · Zona · Último Nivel · Fecha de Referencia · Responsable · Grado de Interés · Acción\n\n' +
    '⚠️ El Sync importa SOLO registros referidos a Educación.\n\n' +
    'Acción disponible: "Enviar a: Lista de Espera"\n\n' +
    'Usa Menú → 🌐 KoboToolbox → 🔄 Sync → Referencias para importar datos.'
  );
}

// ── 9.2  Configurar hoja "Lista de Espera" ───────────────────────────────────
function setupHojaListaEspera() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ui   = SpreadsheetApp.getUi();
  const ENCABEZADOS = [
    'Creamos ID', 'Nombre Completo', 'Nombre Preferido', 'DPI / CUI',
    'Fecha de Nacimiento', 'Edad', 'Género', 'Teléfono',
    'Zona / Colonia', 'Último Nivel Cursado', 'Acción'
  ];
  const ACCIONES_LISTA = ['-- Seleccionar --', ...GRADOS.map(function(g) { return 'Enviar a: ' + g; })];

  let hoja = ss.getSheetByName(HOJA_LISTA_ESPERA);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_LISTA_ESPERA);
  } else {
    hoja.clearContents();
    hoja.clearFormats();
    try { hoja.clearDataValidations(); } catch(e) {}
  }

  // Fila 1 — encabezados
  const rHead = hoja.getRange(1, 1, 1, ENCABEZADOS.length);
  rHead.setValues([ENCABEZADOS])
    .setBackground(COLOR_HEADER_LISTA).setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 32);

  // Anchos de columna
  [180, 220, 150, 130, 130, 60, 90, 110, 150, 200, 200].forEach(function(w, i) {
    hoja.setColumnWidth(i + 1, w);
  });

  // Validación Acción (columna 11) — enviar a cualquier grado
  const valAccion = SpreadsheetApp.newDataValidation()
    .requireValueInList(ACCIONES_LISTA, true).setAllowInvalid(false).build();
  hoja.getRange(2, COL_LISTA.ACCION, hoja.getMaxRows() - 1, 1).setDataValidation(valAccion);

  // Validación Género (columna 7)
  const valGen = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Hombre', 'Mujer', 'Otro'], true).setAllowInvalid(true).build();
  hoja.getRange(2, COL_LISTA.GENERO, hoja.getMaxRows() - 1, 1).setDataValidation(valGen);

  ui.alert(
    '✅ Hoja "' + HOJA_LISTA_ESPERA + '" configurada.\n\n' +
    'Columnas: Creamos ID · Nombre Completo · Nombre Preferido · DPI · Fecha Nac · Edad · Género · Teléfono · Zona · Último Nivel · Acción\n\n' +
    'Acción: selecciona "Enviar a: [Grado]" para transferir a la hoja de grado.'
  );
}

// ── 9.2b Transferir Hoja de Interés → Lista de Espera ───────────────────────
// Solo copia la fila — el origen queda marcado con ✅ pero NO se borra ni oculta.
function _transInteresAListaEspera(fila) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInteres = ss.getSheetByName(HOJA_INTERES);
  const datos       = hojaInteres.getRange(fila, 1, 1, 14).getValues()[0];

  const nombre = datos[COL_INTERES.NOMBRE - 1];
  if (!nombre) {
    SpreadsheetApp.getUi().alert('⚠️ La fila no tiene nombre. No se realizó la copia.');
    hojaInteres.getRange(fila, COL_INTERES.ACCION).setValue('-- Seleccionar --');
    return;
  }

  let hojaLista = ss.getSheetByName(HOJA_LISTA_ESPERA);
  if (!hojaLista) {
    setupHojaListaEspera();
    hojaLista = ss.getSheetByName(HOJA_LISTA_ESPERA);
  }

  // Construir la fila destino en estructura de Lista de Espera (11 cols)
  const filaDestino = Math.max(hojaLista.getLastRow() + 1, 2);
  const filaLista = [
    datos[COL_INTERES.CREAMOS_ID  - 1] || '',   // Creamos ID
    datos[COL_INTERES.NOMBRE      - 1] || '',   // Nombre Completo
    datos[COL_INTERES.NOMBRE_PREF - 1] || '',   // Nombre Preferido
    datos[COL_INTERES.DPI         - 1] || '',   // DPI / CUI
    datos[COL_INTERES.FECHA_NAC   - 1] || '',   // Fecha de Nacimiento
    datos[COL_INTERES.EDAD        - 1] || '',   // Edad
    datos[COL_INTERES.GENERO      - 1] || '',   // Género
    datos[COL_INTERES.TELEFONO    - 1] || '',   // Teléfono
    datos[COL_INTERES.ZONA        - 1] || '',   // Zona / Colonia
    datos[COL_INTERES.ULTIMO_ANIO - 1] || '',   // Último Nivel Cursado
    '-- Seleccionar --'                          // Acción (vacía al llegar)
  ];

  hojaLista.getRange(filaDestino, 1, 1, 11).setValues([filaLista]);
  hojaLista.setRowHeight(filaDestino, 24);

  // Aplicar validación de Acción → grado en la fila recién creada
  const valAccion = SpreadsheetApp.newDataValidation()
    .requireValueInList(['-- Seleccionar --', ...GRADOS.map(function(g) { return 'Enviar a: ' + g; })], true)
    .setAllowInvalid(false).build();
  hojaLista.getRange(filaDestino, COL_LISTA.ACCION).setDataValidation(valAccion);

  // Marcar origen como copiado — fondo verde + ✅ en Acción. La fila NO se borra ni oculta.
  hojaInteres.getRange(fila, 1, 1, 14).setBackground('#E8F5E9');
  hojaInteres.getRange(fila, COL_INTERES.ACCION)
    .setValue('✅ Lista de Espera').setDataValidation(null);

  ss.toast('"' + nombre + '" copiado → Lista de Espera', '✅ Copiado', 5);
}

// ── 9.3  Transferir Referencias → Lista de Espera ────────────────────────────
function _transReferenciasAListaEspera(fila) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRef = ss.getSheetByName(HOJA_REFERENCIAS);
  const datos   = hojaRef.getRange(fila, 1, 1, 14).getValues()[0];

  const nombre = datos[COL_REF.NOMBRE - 1];
  if (!nombre) {
    SpreadsheetApp.getUi().alert('⚠️ La fila no tiene nombre. No se realizó la transferencia.');
    hojaRef.getRange(fila, COL_REF.ACCION).setValue('-- Seleccionar --');
    return;
  }

  let hojaLista = ss.getSheetByName(HOJA_LISTA_ESPERA);
  if (!hojaLista) {
    setupHojaListaEspera();
    hojaLista = ss.getSheetByName(HOJA_LISTA_ESPERA);
  }

  const filaDestino = Math.max(hojaLista.getLastRow() + 1, 2);
  // Copiar las 10 columnas de datos (sin la Acción) y dejar Acción vacía
  hojaLista.getRange(filaDestino, 1, 1, 10).setValues([datos.slice(0, 10)]);
  hojaLista.getRange(filaDestino, COL_LISTA.ACCION).setValue('-- Seleccionar --');
  // Replicar la validación de Acción en la fila recién creada
  const ACCIONES_LISTA = ['-- Seleccionar --', ...GRADOS.map(function(g) { return 'Enviar a: ' + g; })];
  const valAccion = SpreadsheetApp.newDataValidation()
    .requireValueInList(ACCIONES_LISTA, true).setAllowInvalid(false).build();
  hojaLista.getRange(filaDestino, COL_LISTA.ACCION).setDataValidation(valAccion);

  hojaLista.setRowHeight(filaDestino, 24);

  // Marcar origen como enviado
  hojaRef.getRange(fila, 1, 1, 14).setBackground('#EDE7F6');
  hojaRef.getRange(fila, COL_REF.ACCION)
    .setValue('✅ Lista de Espera').setDataValidation(null);

  ss.toast('"' + nombre + '" → Lista de Espera', '✅ Referencia enviada', 5);
}

// ── 9.4  Transferir Lista de Espera → Hoja de Grado ─────────────────────────
function _transListaEsperaAGrado(fila, grado) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLista = ss.getSheetByName(HOJA_LISTA_ESPERA);
  const datos     = hojaLista.getRange(fila, 1, 1, 10).getValues()[0];

  const creamosId  = datos[COL_LISTA.CREAMOS_ID  - 1];
  const nombre     = datos[COL_LISTA.NOMBRE      - 1];
  const dpi        = datos[COL_LISTA.DPI         - 1];
  const edad       = datos[COL_LISTA.EDAD        - 1];
  const telefono   = datos[COL_LISTA.TELEFONO    - 1];

  if (!nombre) {
    SpreadsheetApp.getUi().alert('⚠️ La fila no tiene nombre. No se realizó la transferencia.');
    hojaLista.getRange(fila, COL_LISTA.ACCION).setValue('-- Seleccionar --');
    return;
  }

  const nombreHoja = _nombreHoja(grado);
  let   hojaGrado  = ss.getSheetByName(nombreHoja);
  if (!hojaGrado) hojaGrado = crearHojaGrado(grado, true);

  const filaDestino = Math.max(hojaGrado.getLastRow() + 1, 3);
  const id          = _siguienteId(hojaGrado, grado);

  hojaGrado.getRange(filaDestino, 1, 1, 9).setValues([[
    id, creamosId || '', nombre, dpi,
    telefono || '', edad, grado,
    'Presencial',
    'Inscritx'
  ]]);

  hojaGrado.getRange(filaDestino, 1, 1, 9)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  hojaGrado.getRange(filaDestino, COL_GRADO.NOMBRE).setHorizontalAlignment('left');
  hojaGrado.getRange(filaDestino, COL_GRADO.DPI).setHorizontalAlignment('left');
  hojaGrado.setRowHeight(filaDestino, 26);

  // Marcar origen como enviado
  hojaLista.getRange(fila, 1, 1, 11).setBackground('#E8F5E9');
  hojaLista.getRange(fila, COL_LISTA.ACCION)
    .setValue('✅ ' + grado).setDataValidation(null);

  ss.toast('"' + nombre + '" → "' + nombreHoja + '" · ID: ' + id, '✅ Transferido a grado', 5);
}

// ── 9.5  Sync KoboToolbox → "Referencias a Educación" ────────────────────────
function koboSincronizarReferencias() {
  _koboSincronizarReferenciasInterno(KOBO_URL_REFERENCIAS);
}

// ── 9.5b  Sync combinado: Interés + Referencias ───────────────────────────────
function koboSincronizarTodo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('🔄 Sincronizando Hoja de Interés...\n\nEsto puede tomar unos segundos.');
  _koboSincronizar(KOBO_URL_ACTUAL, false);
  ui.alert('🔄 Sincronizando Referencias a Educación...\n\nEsto puede tomar unos segundos.');
  _koboSincronizarReferenciasInterno(KOBO_URL_REFERENCIAS);
  ui.alert('✅ Sync completo\n\nSe sincronizaron:\n• Hoja de Interés\n• Referencias a Educación');
}

function _koboSincronizarReferenciasInterno(url) {
  const ui = SpreadsheetApp.getUi();
  try {
    const csv  = _koboFetchCsv(url);
    const rows = _koboParseCsv(csv);
    if (rows.length < 2) { ui.alert('El CSV de Referencias no tiene datos o está vacío.'); return; }

    const headers = rows[0];

    function _norm(s) {
      return String(s || '').trim().toLowerCase()
        .normalize('NFD').replace(/[̀-ͯ]/g, '')
        .replace(/\//g, ' ').replace(/\s+/g, ' ').trim();
    }
    const headersNorm = {};
    headers.forEach(function(h, i) {
      const k = _norm(h);
      if (!headersNorm.hasOwnProperty(k)) headersNorm[k] = i;
    });
    function _col(n) {
      const k = _norm(n);
      return headersNorm.hasOwnProperty(k) ? headersNorm[k] : -1;
    }
    function _colFuzzy(t) {
      const nt = _norm(t);
      const k  = Object.keys(headersNorm).find(function(h) { return h.indexOf(nt) >= 0; });
      return k !== undefined ? headersNorm[k] : -1;
    }

    // Mapeo de columnas (reutiliza lógica del sync principal)
    const idx = {};
    idx.CREAMOS_ID  = _col('Creamos ID');
    if (idx.CREAMOS_ID  < 0) idx.CREAMOS_ID  = _col('Inicio/Creamos ID');
    if (idx.CREAMOS_ID  < 0) idx.CREAMOS_ID  = _colFuzzy('creamos id');   // evita confundir con "¿Ya es participante de Creamos?"
    if (idx.CREAMOS_ID  < 0) idx.CREAMOS_ID  = _colFuzzy('creamos');

    idx.NOMBRE      = _col('Nombre Completo');
    if (idx.NOMBRE      < 0) idx.NOMBRE      = _colFuzzy('nombre completo'); // evita tomar "Nombre del responsable" primero
    if (idx.NOMBRE      < 0) idx.NOMBRE      = _col('Nombre(s)');
    if (idx.NOMBRE      < 0) idx.NOMBRE      = _col('Nombre');
    if (idx.NOMBRE      < 0) idx.NOMBRE      = _colFuzzy('nombre');

    idx.APELLIDO    = _col('Apellido(s)');
    if (idx.APELLIDO    < 0) idx.APELLIDO    = _col('Apellido');
    if (idx.APELLIDO    < 0) idx.APELLIDO    = _colFuzzy('apellido');

    idx.NOMBRE_PREF = _col('Nombre Preferido');
    if (idx.NOMBRE_PREF < 0) idx.NOMBRE_PREF = _colFuzzy('preferido');

    idx.DPI         = _col('DPI / CUI');
    if (idx.DPI         < 0) idx.DPI         = _col('DPI');
    if (idx.DPI         < 0) idx.DPI         = _col('CUI');
    if (idx.DPI         < 0) idx.DPI         = _colFuzzy('dpi');
    if (idx.DPI         < 0) idx.DPI         = _colFuzzy('cui');

    idx.FECHA_NAC   = _col('Fecha de Nacimiento');
    if (idx.FECHA_NAC   < 0) idx.FECHA_NAC   = _colFuzzy('nacimiento');

    idx.GENERO      = _col('Género');
    if (idx.GENERO      < 0) idx.GENERO      = _col('Genero');

    idx.TELEFONO    = _col('Teléfono');
    if (idx.TELEFONO    < 0) idx.TELEFONO    = _col('Número de Teléfono');
    if (idx.TELEFONO    < 0) idx.TELEFONO    = _colFuzzy('telefono');

    idx.ZONA        = _col('Zona / Colonia');
    if (idx.ZONA        < 0) idx.ZONA        = _col('Zona');
    if (idx.ZONA        < 0) idx.ZONA        = _colFuzzy('zona');

    idx.ZONA_ESPEC  = _colFuzzy('especifique zona');
    if (idx.ZONA_ESPEC  < 0) idx.ZONA_ESPEC  = _colFuzzy('especifique');

    // "Último Nivel Cursado" está bajo DETALLES EDUCACIÓN; buscar específico
    // para no confundir con "Último Nivel Académico Aprobado" de Laboral
    idx.ULTIMO_ANIO = _colFuzzy('ultimo nivel cursado');
    if (idx.ULTIMO_ANIO < 0) idx.ULTIMO_ANIO = _col('Último Nivel Cursado');
    if (idx.ULTIMO_ANIO < 0) idx.ULTIMO_ANIO = _colFuzzy('nivel de estudios');

    // Columnas nuevas: filtro programa + datos extra de Educación
    idx.PROGRAMA       = _colFuzzy('a que programa se refiere');
    if (idx.PROGRAMA   < 0) idx.PROGRAMA     = _colFuzzy('programa se refiere');
    if (idx.PROGRAMA   < 0) idx.PROGRAMA     = _colFuzzy('programa');

    idx.FECHA_REF      = _col('Fecha de Referencia');
    if (idx.FECHA_REF  < 0) idx.FECHA_REF    = _colFuzzy('fecha de referencia');

    idx.RESPONSABLE    = _colFuzzy('nombre del responsable');
    if (idx.RESPONSABLE< 0) idx.RESPONSABLE  = _colFuzzy('responsable');

    // Grado de Interés: reemplaza "Estado del estudio" (que no existe en el form)
    idx.GRADO_INTERES  = _colFuzzy('en que grado esta interesado');
    if (idx.GRADO_INTERES < 0) idx.GRADO_INTERES = _colFuzzy('grado esta interesado');
    if (idx.GRADO_INTERES < 0) idx.GRADO_INTERES = _colFuzzy('grado interesado');

    // DPIs ya en la hoja
    const ss            = SpreadsheetApp.getActiveSpreadsheet();
    let   hojaRef       = ss.getSheetByName(HOJA_REFERENCIAS);
    if (!hojaRef) { setupHojaReferencias(); hojaRef = ss.getSheetByName(HOJA_REFERENCIAS); }

    const ultimaFila    = hojaRef.getLastRow();
    const dpisExistentes = new Set();
    if (ultimaFila >= 2) {
      hojaRef.getRange(2, COL_REF.DPI, ultimaFila - 1, 1)
        .getValues().flat()
        .forEach(function(d) { if (d !== '') dpisExistentes.add(String(d).trim()); });
    }

    let importados = 0, omitidos = 0;
    const filasNuevas = [];

    rows.slice(1).forEach(function(row) {
      // ── Filtrar: solo referidos a Educación ──────────────────────────────
      if (idx.PROGRAMA >= 0) {
        const prog = _norm(String(row[idx.PROGRAMA] || ''));
        if (prog && prog.indexOf('educaci') < 0) { omitidos++; return; }
      }

      const dpi = idx.DPI >= 0 ? String(row[idx.DPI] || '').trim() : '';
      if (dpi && dpisExistentes.has(dpi)) { omitidos++; return; }

      // Nombre completo
      let nombre = idx.NOMBRE >= 0 ? String(row[idx.NOMBRE] || '').trim() : '';
      if (!nombre && idx.APELLIDO >= 0) {
        const ap = String(row[idx.APELLIDO] || '').trim();
        if (ap) nombre = nombre + (nombre ? ' ' : '') + ap;
      }
      if (!nombre) { omitidos++; return; }  // sin nombre → omitir

      const creamosId  = idx.CREAMOS_ID  >= 0 ? String(row[idx.CREAMOS_ID]  || '').trim() : '';
      const nombrePref = idx.NOMBRE_PREF >= 0 ? String(row[idx.NOMBRE_PREF] || '').trim() : '';
      const fechaNac   = idx.FECHA_NAC   >= 0 ? String(row[idx.FECHA_NAC]   || '').trim() : '';
      const genero     = _normalizarGenero(idx.GENERO >= 0 ? row[idx.GENERO] : '');
      const telefono   = idx.TELEFONO    >= 0 ? String(row[idx.TELEFONO]    || '').trim() : '';
      // Combinar Zona / Colonia + Especifique zona
      let zona = idx.ZONA >= 0 ? String(row[idx.ZONA] || '').trim() : '';
      if (idx.ZONA_ESPEC >= 0) {
        const zonaesp = String(row[idx.ZONA_ESPEC] || '').trim();
        if (zonaesp && zonaesp !== zona) zona = zona ? zona + ' - ' + zonaesp : zonaesp;
      }
      const ultimoAnio      = idx.ULTIMO_ANIO     >= 0 ? String(row[idx.ULTIMO_ANIO]     || '').trim() : '';
      const fechaRef        = idx.FECHA_REF      >= 0 ? String(row[idx.FECHA_REF]      || '').trim() : '';
      const responsable     = idx.RESPONSABLE    >= 0 ? String(row[idx.RESPONSABLE]    || '').trim() : '';
      const gradoInteres    = idx.GRADO_INTERES  >= 0 ? String(row[idx.GRADO_INTERES]  || '').trim() : '';

      // Calcular edad desde fecha de nacimiento
      let edad = '';
      if (fechaNac) {
        const parts = fechaNac.split(/[-\/]/);
        if (parts.length === 3) {
          const anioNac = parseInt(parts[0].length === 4 ? parts[0] : parts[2]);
          if (!isNaN(anioNac)) edad = new Date().getFullYear() - anioNac;
        }
      }

      if (dpi) dpisExistentes.add(dpi);
      filasNuevas.push([
        creamosId, nombre, nombrePref, dpi,
        fechaNac, edad, genero, telefono, zona, ultimoAnio,
        fechaRef, responsable, gradoInteres,
        '-- Seleccionar --'
      ]);
      importados++;
    });

    if (filasNuevas.length > 0) {
      const filaInicio = Math.max(hojaRef.getLastRow() + 1, 2);
      hojaRef.getRange(filaInicio, 1, filasNuevas.length, 14).setValues(filasNuevas);
      // Aplicar validación de Acción en las filas nuevas
      const valAccion = SpreadsheetApp.newDataValidation()
        .requireValueInList(['-- Seleccionar --', 'Enviar a: Lista de Espera'], true)
        .setAllowInvalid(false).build();
      hojaRef.getRange(filaInicio, COL_REF.ACCION, filasNuevas.length, 1).setDataValidation(valAccion);
    }

    const msg = '✅ Sync de Referencias completado\n\n' +
      '• Importados:  ' + importados + '\n' +
      '• Omitidos (ya existían o sin nombre): ' + omitidos + '\n\n' +
      'Total en hoja: ' + (Math.max(hojaRef.getLastRow() - 1, 0));
    if (ui) ui.alert(msg);
    ss.toast(importados + ' registros importados.', '✅ Referencias sync', 5);

  } catch(e) {
    if (ui) ui.alert('❌ Error en sync de Referencias:\n' + e.message);
  }
}

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
  const token   = _koboGetToken();
  const opts    = { headers: { Authorization: 'Token ' + token }, muteHttpExceptions: true };
  const esperas = [0, 3000, 6000, 12000]; // 4 intentos: 0s, 3s, 6s, 12s

  for (var i = 0; i < esperas.length; i++) {
    if (esperas[i] > 0) Utilities.sleep(esperas[i]);

    var resp = UrlFetchApp.fetch(url, opts);
    var code = resp.getResponseCode();

    if (code === 200) return resp.getContentText('UTF-8');

    if (code === 401 || code === 403) throw new Error(
      'Error ' + code + ': token inválido o sin permisos.\n' +
      'Ve a: DP Educación → 🌐 KoboToolbox → 🔑 Configurar token de API'
    );

    // 5xx o cualquier error temporal: reintentar (excepto en el último intento)
    if (i < esperas.length - 1) continue;

    throw new Error(
      'Error al descargar el CSV (HTTP ' + code + ').\n' +
      'Se intentó ' + esperas.length + ' veces. KoboToolbox puede estar caído momentáneamente.\n' +
      'Espera unos minutos y vuelve a intentarlo.'
    );
  }
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


// ── 8.4  Sincronizar datos → hoja "Interés" ───────────────────────────────────
//
//  - Descarga el CSV de KOBO_URL_ACTUAL
//  - Filtra solo registros de Educación (¿Deseas inscribirte? = Sí/Si/Yes/1)
//  - Si el DPI ya existe en Interés → omite (sin duplicados)
//  - Si no existe → agrega al final con las 13 columnas completas

// Helper: devuelve true si el valor Kobo indica "Sí" en cualquier forma
function _esValorPositivo(val) {
  if (val === 1 || val === true) return true;   // número/booleano exacto
  if (val === 0 || val === false) return false;  // número/booleano exacto
  const v = String(val || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // quita tildes
  return v === 'si' || v === 'yes' || v === '1' || v === 'true';
}

// ── Helper: considera "no seleccionado" cualquier valor vacío o explícitamente negativo ──
function _esValorNegativo(val) {
  const v = String(val || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return v === '' || v === 'no' || v === '0' || v === 'false';
}

// ── Helper: verifica si un campo multi-selección de programas incluye Educación ──
// Funciona aunque la persona haya elegido varios programas a la vez.
// Ejemplo de valor: "Inclusión Laboral, Educación Extraescolar/Alternativa, Apoyo..."
function _contieneEducacion(val) {
  const v = String(val || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // sin tildes
    .replace(/\//g, ' ').replace(/\s+/g, ' ');        // / → espacio
  return v.indexOf(KOBO_KEYWORD_EDUCACION) >= 0;
}

// ── Helper: normalizar género ─────────────────────────────────────────────────
// Acepta cualquier variante del formulario y devuelve Hombre / Mujer / Otro / ''
function _normalizarGenero(raw) {
  const v = String(raw || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // sin tildes
  if (!v) return '';
  if (v === 'h' || v === 'm' || v.indexOf('hombre') >= 0 || v.indexOf('masculino') >= 0) return 'Hombre';
  if (v === 'f' || v.indexOf('mujer') >= 0 || v.indexOf('femenino') >= 0 || v.indexOf('femenil') >= 0) return 'Mujer';
  return 'Otro';
}

// ── Botón: importar datos históricos (una sola vez) ──────────────────────────
function koboImportarHistorico() {
  const ui = SpreadsheetApp.getUi();
  const r  = ui.alert(
    '📦 Importar datos históricos',
    'Se importarán registros del formulario anterior (años pasados).\n\n' +
    '• Se agregan solo registros cuyo DPI no exista ya en "Interés"\n' +
    '• Los registros ya en la hoja NO se tocan\n' +
    '• En datos históricos NO se filtra por "¿Inscribirte en Educación?"\n' +
    '  (todos los registros del form. histórico son de educación)\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) return;
  _koboSincronizar(KOBO_URL_HISTORICO, true);
}

function koboSincronizarHojaInteres() {
  _koboSincronizar(KOBO_URL_ACTUAL, false);
}

// ── Función interna compartida de sincronización ──────────────────────────────
// url:           URL del CSV de KoboToolbox
// modoHistorico: si true, NO filtra por campo INSCRIPCION (todos son de educación)
function _koboSincronizar(url, modoHistorico) {
  const ui = SpreadsheetApp.getUi();
  try {
    const csv  = _koboFetchCsv(url);
    const rows = _koboParseCsv(csv);
    if (rows.length < 2) { ui.alert('El CSV no tiene datos o está vacío.'); return; }

    const headers = rows[0];

    // Normaliza: sin tildes, minúsculas, sin espacios, barra → espacio
    function _norm(s) {
      return String(s || '').trim().toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/\//g, ' ').replace(/\s+/g, ' ').trim();
    }

    // Mapa normalizado cabecera → índice
    const headersNorm = {};
    headers.forEach(function(h, i) {
      const k = _norm(h);
      if (!headersNorm.hasOwnProperty(k)) headersNorm[k] = i;
    });
    function _col(nombre) {
      const k = _norm(nombre);
      return headersNorm.hasOwnProperty(k) ? headersNorm[k] : -1;
    }

    // Búsqueda difusa: encuentra la primera columna cuyo nombre CONTENGA el término
    function _colFuzzy(termino) {
      const t = _norm(termino);
      const k = Object.keys(headersNorm).find(function(h) { return h.indexOf(t) >= 0; });
      return k !== undefined ? headersNorm[k] : -1;
    }

    // ── Construir índices de todos los campos mapeados ────────────────────
    const idx = {};
    Object.keys(KOBO_MAP).forEach(function(campo) {
      idx[campo] = _col(KOBO_MAP[campo]);
    });

    // ── Fallbacks: nombres de campos del formulario HISTÓRICO ─────────────
    // El formulario histórico usaba prefijos de grupo ("Inicio/", "Educación.../").
    // El formulario nuevo no los tiene. Estos fallbacks garantizan compatibilidad.

    // Datos personales (histórico: prefijo "Inicio/")
    if (idx.NOMBRE       < 0) idx.NOMBRE       = _col('Inicio/Nombre(s)');
    if (idx.NOMBRE       < 0) idx.NOMBRE       = _col('Nombre');
    if (idx.NOMBRE       < 0) idx.NOMBRE       = _colFuzzy('nombre');

    if (idx.APELLIDO     < 0) idx.APELLIDO     = _col('Inicio/Apellido(s)');
    if (idx.APELLIDO     < 0) idx.APELLIDO     = _col('Apellido');
    if (idx.APELLIDO     < 0) idx.APELLIDO     = _colFuzzy('apellido');

    if (idx.NOMBRE_PREF  < 0) idx.NOMBRE_PREF  = _col('Inicio/Nombre Preferido');
    if (idx.NOMBRE_PREF  < 0) idx.NOMBRE_PREF  = _colFuzzy('preferido');

    if (idx.CREAMOS_ID   < 0) idx.CREAMOS_ID   = _col('Inicio/Creamos ID');
    if (idx.CREAMOS_ID   < 0) idx.CREAMOS_ID   = _colFuzzy('creamos');

    if (idx.GENERO       < 0) idx.GENERO       = _col('Inicio/Género');
    if (idx.GENERO       < 0) idx.GENERO       = _col('Genero');

    if (idx.TELEFONO     < 0) idx.TELEFONO     = _col('Inicio/Número de Teléfono');
    if (idx.TELEFONO     < 0) idx.TELEFONO     = _col('Número de teléfono');
    if (idx.TELEFONO     < 0) idx.TELEFONO     = _colFuzzy('telefono');

    if (idx.ZONA         < 0) idx.ZONA         = _col('Inicio/Zona');
    if (idx.OTRA_ZONA    < 0) idx.OTRA_ZONA    = _col('Inicio/Otra zona');
    if (idx.COLONIA      < 0) idx.COLONIA      = _col('Inicio/Colonia');
    if (idx.OTRA_COLONIA < 0) idx.OTRA_COLONIA = _col('Inicio/Otra colonia');

    if (idx.ULTIMO_ANIO  < 0) idx.ULTIMO_ANIO  = _col('Inicio/¿Cuál es tu último nivel de estudios terminado?');
    if (idx.ULTIMO_ANIO  < 0) idx.ULTIMO_ANIO  = _colFuzzy('ultimo nivel');

    // DPI / CUI (solo formulario histórico)
    if (idx.DPI < 0) idx.DPI = _col('Inicio/Número de DPI');
    if (idx.DPI < 0) idx.DPI = _col('Numero de DPI');
    if (idx.DPI < 0) idx.DPI = _col('DPI');
    if (idx.DPI < 0) idx.DPI = _col('CUI');
    if (idx.DPI < 0) idx.DPI = _colFuzzy('dpi');
    if (idx.DPI < 0) idx.DPI = _colFuzzy('cui');

    // Fecha de nacimiento (solo formulario histórico)
    if (idx.FECHA_NAC < 0) idx.FECHA_NAC = _col('Inicio/Fecha de nacimiento');
    if (idx.FECHA_NAC < 0) idx.FECHA_NAC = _colFuzzy('nacimiento');

    // Sección Educación (histórico: prefijo "Educación Extraescolar/Alternativa/")
    if (idx.GRADO_KOBO < 0) idx.GRADO_KOBO = _col('Educación Extraescolar/Alternativa/¿Qué grado/etapa te toca con Creamos?');
    if (idx.GRADO_KOBO < 0) idx.GRADO_KOBO = _colFuzzy('grado');
    if (idx.GRADO_KOBO < 0) idx.GRADO_KOBO = _colFuzzy('etapa');

    if (idx.COMENTARIO < 0) idx.COMENTARIO = _col('Educación Extraescolar/Alternativa/Comentarios de papelería');
    if (idx.COMENTARIO < 0) idx.COMENTARIO = _col('Comentarios de papelería');

    // ── Fallbacks para el filtro de Educación ────────────────────────────
    // PROGRAMAS_EDUC: columna booleana del multi-select (formulario nuevo)
    // El formulario actual exporta con prefijo "Inicio/"
    if (idx.PROGRAMAS_EDUC < 0) idx.PROGRAMAS_EDUC =
      _col('Inicio/¿Qué programas te interesan?/Educación Extraescolar/Alternativa');
    if (idx.PROGRAMAS_EDUC < 0) idx.PROGRAMAS_EDUC =
      _col('¿Qué programas te interesan?/Educación Extraescolar/Alternativa');
    if (idx.PROGRAMAS_EDUC < 0) idx.PROGRAMAS_EDUC =
      _colFuzzy('programas te interesan');   // último recurso

    // INSCRIPCION: campo Sí/No dentro del grupo Educación (ambos formularios)
    if (idx.INSCRIPCION < 0) idx.INSCRIPCION = _col('Educación Extraescolar/Alternativa/¿Deseas inscribirte en el programa de Educación?');
    if (idx.INSCRIPCION < 0) idx.INSCRIPCION = _col('¿Deseas inscribirte en el programa de Educación?');
    if (idx.INSCRIPCION < 0) idx.INSCRIPCION = _colFuzzy('inscribirte en el programa de educacion');

    // ── Índices de papelería individual ───────────────────────────────────
    const idxPap = {};
    Object.keys(KOBO_MAP_PAPELERIA).forEach(function(doc) {
      idxPap[doc] = _col(KOBO_MAP_PAPELERIA[doc]);
    });

    // Mapa normalizado de grados (tolera variantes con/sin tilde)
    const gradoMapNorm = {};
    Object.keys(KOBO_GRADO_MAP).forEach(function(k) {
      gradoMapNorm[_norm(k)] = KOBO_GRADO_MAP[k];
    });

    // ── Diagnóstico: mostrar columnas no encontradas ───────────────────────
    const sinFiltro = idx.PROGRAMAS_EDUC < 0 && idx.INSCRIPCION < 0 && idx.GRADO_KOBO < 0;
    // DPI es opcional en el formulario actual (usa UUID/CreamosID para dedup)
    const camposImportantes = modoHistorico
      ? ['NOMBRE', 'DPI']
      : sinFiltro ? ['NOMBRE', 'PROGRAMAS_EDUC'] : ['NOMBRE'];
    const faltantes = camposImportantes.filter(function(k) { return idx[k] < 0; });
    if (faltantes.length > 0) {
      // Mostrar las primeras 10 columnas del CSV para diagnóstico
      const muestra = headers.slice(0, 15).join('\n  • ');
      const aviso   = faltantes.map(function(k){
        return '  • ' + k + ' → buscado como "' + KOBO_MAP[k] + '"';
      }).join('\n');
      const r = ui.alert(
        '⚠️ Campos no encontrados en el CSV',
        'No se encontraron:\n' + aviso + '\n\n' +
        'Primeras columnas del CSV:\n  • ' + muestra + '\n\n' +
        '¿Continuar de todas formas?\n' +
        '(Los campos faltantes quedarán vacíos)',
        ui.ButtonSet.YES_NO
      );
      if (r !== ui.Button.YES) return;
    }

    // ── Hoja Interés ───────────────────────────────────────────────────────
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaInteres = ss.getSheetByName(HOJA_INTERES);
    if (!hojaInteres) {
      ui.alert('❌ La hoja "Interés" no existe.\nEjecuta primero "⚙️ Configurar hoja Interés".');
      return;
    }

    // DPIs, CreamosIDs y UUIDs ya existentes (deduplicación)
    const ultimaFilaI      = hojaInteres.getLastRow();
    const dpisExistentes   = new Set();
    const creamosExistentes = new Set();
    if (ultimaFilaI >= 2) {
      const filasDatos = hojaInteres.getRange(2, 1, ultimaFilaI - 1, Math.max(COL_INTERES.DPI, COL_INTERES.CREAMOS_ID)).getValues();
      filasDatos.forEach(function(r) {
        const cid = String(r[COL_INTERES.CREAMOS_ID - 1] || '').trim();
        const dpi = String(r[COL_INTERES.DPI - 1]        || '').trim();
        if (cid) creamosExistentes.add(cid);
        if (dpi) dpisExistentes.add(dpi);
      });
    }
    // UUIDs vistos en ESTE sync (evita importar duplicados del mismo CSV)
    const uuidsSyncActual = new Set();

    let omitidosFiltro = 0, omitidosDupes = 0;
    const filasNuevas  = [];

    rows.slice(1).forEach(function(row) {

      // ── 1. Filtrar: solo personas de Educación ───────────────────────────
      // Regla única para AMBOS formularios (histórico y nuevo):
      //   "¿Deseas inscribirte en el programa de Educación?" = Sí → importar
      //   Vacío, No, o cualquier otro valor → descartar
      // No hay bypass por Creamos ID: tener ID no significa ser de Educación.
      if (idx.INSCRIPCION >= 0) {
        if (!_esValorPositivo(row[idx.INSCRIPCION])) { omitidosFiltro++; return; }
      } else if (idx.PROGRAMAS_EDUC >= 0) {
        // Sin campo de inscripción: usar el multi-select como respaldo
        if (!_esValorPositivo(row[idx.PROGRAMAS_EDUC])) { omitidosFiltro++; return; }
      } else {
        // Sin ningún campo conocido: usar presencia de grado como último recurso
        const tieneGrado = idx.GRADO_KOBO >= 0 &&
          String(row[idx.GRADO_KOBO] || '').trim() !== '';
        if (!tieneGrado) { omitidosFiltro++; return; }
      }

      const creamosId = String(idx.CREAMOS_ID >= 0 ? (row[idx.CREAMOS_ID] || '') : '').trim();

      // ── 2. Deduplicar por UUID / CreamosID / DPI ────────────────────────
      const uuid = idx.UUID >= 0 ? String(row[idx.UUID] || '').trim() : '';
      if (uuid && uuidsSyncActual.has(uuid)) { omitidosDupes++; return; }
      const dpi  = idx.DPI  >= 0 ? String(row[idx.DPI]  || '').trim() : '';
      if (dpi && dpisExistentes.has(dpi)) { omitidosDupes++; return; }
      // Para formularios sin DPI, usar Creamos ID como clave de dedup
      if (creamosId && creamosExistentes.has(creamosId)) { omitidosDupes++; return; }

      // ── 3. Construir Nombre Completo (Nombre + Apellido) ────────────────
      const nombre   = String(idx.NOMBRE  >= 0 ? (row[idx.NOMBRE]  || '') : '').trim();
      const apellido = String(idx.APELLIDO >= 0 ? (row[idx.APELLIDO] || '') : '').trim();
      const nombreCompleto = [nombre, apellido].filter(Boolean).join(' ');

      if (!nombreCompleto && !dpi && !uuid) return; // fila completamente vacía

      // ── 4. Nombre preferido ─────────────────────────────────────────────
      const nombrePref = String(idx.NOMBRE_PREF >= 0 ? (row[idx.NOMBRE_PREF] || '') : '').trim();

      // ── 5. Fecha de nacimiento / Edad ────────────────────────────────────
      // Formulario nuevo: "Edad" viene directo como número.
      // Formulario histórico: se calcula desde "Fecha de nacimiento".
      const fechaNacRaw = idx.FECHA_NAC >= 0 ? (row[idx.FECHA_NAC] || '') : '';
      const fechaNac    = String(fechaNacRaw).trim();
      const edadDirecta = idx.EDAD_DIRECTA >= 0 ? String(row[idx.EDAD_DIRECTA] || '').trim() : '';
      const edad        = edadDirecta !== '' ? edadDirecta : _calcularEdad(fechaNac);

      // ── 6. Género → normalizado; "¿Cómo te autodescribes?" como complemento ──
      const generoRaw     = idx.GENERO      >= 0 ? row[idx.GENERO]      : '';
      const autodescRaw   = idx.AUTODESCRIBE >= 0 ? row[idx.AUTODESCRIBE] : '';
      const generoNorm    = _normalizarGenero(generoRaw);
      // Si el género normalizado es "Otro", mostrar la autodescripción (si la hay)
      const genero = (generoNorm === 'Otro' && String(autodescRaw || '').trim())
        ? String(autodescRaw).trim()
        : generoNorm;

      // ── 7. Teléfono ──────────────────────────────────────────────────────
      const telefono = String(idx.TELEFONO >= 0 ? (row[idx.TELEFONO] || '') : '').trim();

      // ── 8. Zona / Colonia ────────────────────────────────────────────────
      let zona     = String(idx.ZONA     >= 0 ? (row[idx.ZONA]     || '') : '').trim();
      let colonia  = String(idx.COLONIA  >= 0 ? (row[idx.COLONIA]  || '') : '').trim();
      if (zona    === 'Otra') zona    = String(idx.OTRA_ZONA    >= 0 ? (row[idx.OTRA_ZONA]    || '') : '').trim();
      if (colonia === 'Otra') colonia = String(idx.OTRA_COLONIA >= 0 ? (row[idx.OTRA_COLONIA] || '') : '').trim();
      const zonaColonia = [zona, colonia].filter(Boolean).join(' — ');

      // ── 9. Último nivel cursado ──────────────────────────────────────────
      const ultimoAnio = String(idx.ULTIMO_ANIO >= 0 ? (row[idx.ULTIMO_ANIO] || '') : '').trim();

      // ── 10. Grado asignado por Creamos ───────────────────────────────────
      const gradoKoboRaw = String(idx.GRADO_KOBO >= 0 ? (row[idx.GRADO_KOBO] || '') : '').trim();
      const gradoKobo    = gradoMapNorm[_norm(gradoKoboRaw)] || KOBO_GRADO_MAP[gradoKoboRaw] || gradoKoboRaw;

      // ── 11. Papelería faltante (construida desde campos individuales) ─────
      const papeleriaFaltante = Object.keys(idxPap).filter(function(doc) {
        const i   = idxPap[doc];
        if (i < 0) return false;
        const val = String(row[i] || '').trim();
        // 0, '0', '', 'No' = faltante; '1', 'Sí' = entregado
        return val === '0' || val === '' || val.toLowerCase() === 'no';
      }).join(', ');

      // ── 12. Comentario de papelería ──────────────────────────────────────
      const comentario = String(idx.COMENTARIO >= 0 ? (row[idx.COMENTARIO] || '') : '').trim();

      // ── 13. Acción: siempre vacía — se selecciona manualmente ───────────
      const accion = '';

      filasNuevas.push([
        creamosId,         // 1  Creamos ID
        nombreCompleto,    // 2  Nombre Completo
        nombrePref,        // 3  Nombre Preferido
        dpi,               // 4  DPI
        fechaNac,          // 5  Fecha de Nacimiento
        edad,              // 6  Edad
        genero,            // 7  Género
        telefono,          // 8  Teléfono
        zonaColonia,       // 9  Zona / Colonia
        ultimoAnio,        // 10 Último Nivel
        gradoKobo,         // 11 Grado Asignado (Kobo)
        papeleriaFaltante, // 12 Papelería Faltante
        comentario,        // 13 Comentario
        accion             // 14 Acción
      ]);

      if (dpi)       dpisExistentes.add(dpi);
      if (uuid)      uuidsSyncActual.add(uuid);
      if (creamosId) creamosExistentes.add(creamosId);
    });

    if (!filasNuevas.length) {
      ui.alert(
        '✅ Sin novedades\n\n' +
        'Omitidos (otro programa): ' + omitidosFiltro + '\n' +
        'Ya existían (DPI):        ' + omitidosDupes
      );
      return;
    }

    // ── Escribir en hoja Interés ───────────────────────────────────────────
    const primeraFila = Math.max(hojaInteres.getLastRow() + 1, 2);
    hojaInteres.getRange(primeraFila, 1, filasNuevas.length, 14).setValues(filasNuevas);

    // Validación Acción (setAllowInvalid(true) para no bloquear escrituras por script)
    hojaInteres.getRange(primeraFila, COL_INTERES.ACCION, filasNuevas.length, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(ACCIONES, true).setAllowInvalid(true).build()
      );

    // Validación Grado Kobo
    hojaInteres.getRange(primeraFila, COL_INTERES.GRADO_KOBO, filasNuevas.length, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(GRADOS, true).setAllowInvalid(true).build()
      ).setBackground('#F3E5F5');

    // Fondo azul en papelería
    hojaInteres.getRange(primeraFila, COL_INTERES.PAPELERIA, filasNuevas.length, 1)
      .setBackground('#E3F2FD');

    // Formato de fecha
    hojaInteres.getRange(primeraFila, COL_INTERES.FECHA_NAC, filasNuevas.length, 1)
      .setNumberFormat('dd/mm/yyyy');

    ss.setActiveSheet(hojaInteres);
    const origenLabel = modoHistorico ? '📦 Histórico' : '🔄 Formulario actual';
    ss.toast(filasNuevas.length + ' registros importados desde ' + (modoHistorico ? 'histórico' : 'formulario actual') + '.', '✅ Sync completado', 7);
    ui.alert(
      '✅ ' + (modoHistorico ? 'Importación histórica completada' : 'Sincronización completada') + '\n\n' +
      'Origen: ' + origenLabel + '\n' +
      'Nuevos registros agregados:    ' + filasNuevas.length + '\n' +
      (modoHistorico ? '' : 'Omitidos (otro programa):      ' + omitidosFiltro + '\n') +
      'Ya existían (DPI duplicado):   ' + omitidosDupes + '\n\n' +
      '💡 La columna "Acción" ya viene pre-llenada con el grado asignado.\n' +
      '   Revisa y usa "🔄 Procesar acciones pendientes" para transferir.'
    );

  } catch (err) {
    ui.alert('❌ Error en sincronización\n\n' + err.message);
  }
}

// ── Helper: calcular edad desde fecha de nacimiento ───────────────────────────
function _calcularEdad(fechaStr) {
  if (!fechaStr) return '';
  // Intenta parsear formatos comunes: YYYY-MM-DD, DD/MM/YYYY, MM/DD/YYYY
  let fecha;
  if (/^\d{4}-\d{2}-\d{2}/.test(fechaStr)) {
    fecha = new Date(fechaStr);
  } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(fechaStr)) {
    const p = fechaStr.split('/');
    fecha = new Date(p[2] + '-' + p[1] + '-' + p[0]); // DD/MM/YYYY
  } else {
    fecha = new Date(fechaStr);
  }
  if (isNaN(fecha.getTime())) return '';
  const hoy  = new Date();
  let   edad = hoy.getFullYear() - fecha.getFullYear();
  const m    = hoy.getMonth() - fecha.getMonth();
  if (m < 0 || (m === 0 && hoy.getDate() < fecha.getDate())) edad--;
  return edad > 0 && edad < 120 ? edad : '';
}


// ── 8.6  Trigger automático de sincronización ─────────────────────────────────

function koboInstalarTriggerSync() {
  const ui = SpreadsheetApp.getUi();
  try { _koboGetToken(); } catch (e) { ui.alert('❌ ' + e.message); return; }

  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'koboSincronizarHojaInteres'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('koboSincronizarHojaInteres')
    .timeBased().everyMinutes(1).create();

  ui.alert(
    '✅ Sync automático activado\n\n' +
    'Cada minuto se agregarán automáticamente los registros nuevos de\n' +
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

// ────────────────────────────────────────────────────────────────────────────
//  SECCIÓN 9 · HOJA DE SEGUIMIENTO DE GRADUADOS
// ────────────────────────────────────────────────────────────────────────────
//
//  Registra a los estudiantes que completan Quinto Bachillerato.
//  Flujo automático:
//    1. En la hoja "Quinto Bachillerato [año]", cambia Estado → "Ciclo de Vida Terminado"
//    2. El trigger onEdit detecta el cambio y pregunta si deseas registrar
//       al alumno en la hoja "Seguimiento Graduados"
//    3. También puedes registrar graduados manualmente desde el menú
//
// ────────────────────────────────────────────────────────────────────────────

// ── 9.1  Configurar hoja Seguimiento ─────────────────────────────────────────

function setupHojaSeguimiento() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   hoja = ss.getSheetByName(HOJA_SEGUIMIENTO);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_SEGUIMIENTO);
  } else {
    hoja.clearFormats();
    hoja.clearConditionalFormatRules();
  }

  const anio = new Date().getFullYear();

  // Fila 1: título
  hoja.getRange(1, 1, 1, 8).merge()
    .setValue('🎓 SEGUIMIENTO DE GRADUADOS — QUINTO BACHILLERATO')
    .setBackground(COLOR_HEADER_SEGUIM)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 42);

  // Fila 2: encabezados
  hoja.getRange(2, 1, 1, 8)
    .setValues([['ID','Nombre Completo','DPI / CUI','No. Teléfono',
                 'Edad','Año Graduación','Estado Post-Grad','Observaciones']])
    .setBackground(COLOR_HEADER_SEGUIM)
    .setFontColor(COLOR_FONT_HEADER)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 28);
  hoja.setFrozenRows(2);

  // Anchos
  hoja.setColumnWidth(COL_SEGUIM.ID,          80);
  hoja.setColumnWidth(COL_SEGUIM.NOMBRE,      220);
  hoja.setColumnWidth(COL_SEGUIM.DPI,         140);
  hoja.setColumnWidth(COL_SEGUIM.TELEFONO,    130);
  hoja.setColumnWidth(COL_SEGUIM.EDAD,         60);
  hoja.setColumnWidth(COL_SEGUIM.ANIO_GRAD,   120);
  hoja.setColumnWidth(COL_SEGUIM.ESTADO_POST, 160);
  hoja.setColumnWidth(COL_SEGUIM.OBSERVACION, 240);

  const MAX = 300;

  // Validación año de graduación
  hoja.getRange(3, COL_SEGUIM.ANIO_GRAD, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(2000, 2099).setAllowInvalid(false)
      .setHelpText('Año de graduación (ej. 2026)').build()
  );

  // Validación estado post-graduación
  hoja.getRange(3, COL_SEGUIM.ESTADO_POST, MAX, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(ESTADOS_POST_GRAD, true).setAllowInvalid(false)
      .setHelpText('Estado del graduado').build()
  );

  // Formato condicional por estado
  const dr = hoja.getRange(3, 1, MAX, 8);
  hoja.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G3="Empleado"')
      .setBackground('#C8E6C9').setRanges([dr]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G3="Continúa estudiando"')
      .setBackground('#E3F2FD').setRanges([dr]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G3="Sin contacto"')
      .setBackground('#FFF9C4').setRanges([dr]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G3="Emigró"')
      .setBackground('#F3E5F5').setRanges([dr]).build()
  ]);

  SpreadsheetApp.getUi().alert(
    '✅ Hoja "' + HOJA_SEGUIMIENTO + '" configurada.\n\n' +
    'Cuando un alumno de Quinto Bachillerato cambie su\n' +
    'Estado a "Ciclo de Vida Terminado", se ofrecerá registrarlo aquí\n' +
    'automáticamente.'
  );
}


// ── 9.2  Registrar graduado (desde onEdit o manualmente) ─────────────────────

function _ofrecerRegistrarGraduado(hojaGrado, fila) {
  const ui      = SpreadsheetApp.getUi();
  const datos   = hojaGrado.getRange(fila, 1, 1, 9).getValues()[0];
  const nombre  = datos[COL_GRADO.NOMBRE - 1];

  const r = ui.alert(
    '🎓 ¿Registrar en Seguimiento?',
    '"' + nombre + '" completó el Ciclo de Vida.\n\n' +
    '¿Agregar a la hoja "' + HOJA_SEGUIMIENTO + '"?',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) return;

  _escribirGraduado(datos);
}

function registrarGraduadoManual() {
  const ui   = SpreadsheetApp.getUi();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const fila = hoja.getActiveCell().getRow();

  if (fila < 3) { ui.alert('Selecciona una fila de alumno (fila 3 o más).'); return; }

  const nombreH = hoja.getName();
  const esperada = _nombreHoja(GRADO_GRADUACION);
  if (nombreH !== esperada) {
    ui.alert(
      '⚠️ Hoja incorrecta\n\n' +
      'Navega a la hoja "' + esperada + '" y\n' +
      'selecciona la fila del graduado.'
    );
    return;
  }

  const datos  = hoja.getRange(fila, 1, 1, 9).getValues()[0];
  const nombre = datos[COL_GRADO.NOMBRE - 1];
  if (!nombre) { ui.alert('La fila seleccionada no tiene nombre.'); return; }

  const r = ui.alert(
    '🎓 Registrar graduado',
    '¿Registrar a "' + nombre + '" en la hoja "' + HOJA_SEGUIMIENTO + '"?',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) return;

  _escribirGraduado(datos);
}

function _escribirGraduado(datosGrado) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let   hoja = ss.getSheetByName(HOJA_SEGUIMIENTO);
  if (!hoja) {
    setupHojaSeguimiento();
    hoja = ss.getSheetByName(HOJA_SEGUIMIENTO);
  }

  const anio        = new Date().getFullYear();
  const ultimaFila  = Math.max(hoja.getLastRow() + 1, 3);
  const idGrado     = datosGrado[COL_GRADO.ID - 1];
  const idSeg       = 'SEG-' + String(ultimaFila - 2).padStart(3, '0');

  hoja.getRange(ultimaFila, 1, 1, 8).setValues([[
    idSeg,
    datosGrado[COL_GRADO.NOMBRE    - 1],
    datosGrado[COL_GRADO.DPI       - 1],
    datosGrado[COL_GRADO.TELEFONO  - 1],
    datosGrado[COL_GRADO.EDAD      - 1],
    anio,
    'Seguimiento activo',
    'Graduado de Quinto Bachillerato. ID origen: ' + idGrado
  ]]);

  hoja.getRange(ultimaFila, 1, 1, 8)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  hoja.getRange(ultimaFila, COL_SEGUIM.NOMBRE)
    .setHorizontalAlignment('left');
  hoja.getRange(ultimaFila, COL_SEGUIM.OBSERVACION)
    .setHorizontalAlignment('left');
  hoja.setRowHeight(ultimaFila, 26);

  ss.toast(
    datosGrado[COL_GRADO.NOMBRE - 1] + ' agregado a Seguimiento · ID: ' + idSeg,
    '🎓 Graduado registrado', 5
  );
}
