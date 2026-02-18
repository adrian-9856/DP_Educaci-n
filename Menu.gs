// ============================================================
//  MENÚ Y PUNTO DE ENTRADA
// ============================================================

/**
 * Se ejecuta automáticamente al abrir el spreadsheet.
 * Crea el menú personalizado "DP Educación".
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DP Educación')
    .addItem('⚙️ Configurar hoja Interés', 'setupHojaInteres')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('📚 Crear hoja de grado')
        .addItem('Primero Básico',        'crearHojaPrimeroBasico')
        .addItem('Segundo Básico',        'crearHojaSegundoBasico')
        .addItem('Tercero Básico',        'crearHojaTerceroBasico')
        .addItem('Cuarto Bachillerato',   'crearHojaCuartoBachillerato')
        .addItem('Quinto Bachillerato',   'crearHojaQuintoBachillerato')
        .addItem('Sexto Bachillerato',    'crearHojaSextoBachillerato')
        .addSeparator()
        .addItem('Crear TODOS los grados', 'crearTodasLasHojas')
    )
    .addSeparator()
    .addItem('🔄 Procesar acciones pendientes', 'procesarAccionesPendientes')
    .addItem('📋 Ver resumen de alumnos', 'mostrarResumen')
    .addToUi();
}

// Wrappers individuales para el sub-menú y para el botón dibujado en la hoja
function crearHojaPrimeroBasico()      { crearHojaGrado('Primero Básico'); }
function crearHojaSegundoBasico()      { crearHojaGrado('Segundo Básico'); }
function crearHojaTerceroBasico()      { crearHojaGrado('Tercero Básico'); }
function crearHojaCuartoBachillerato() { crearHojaGrado('Cuarto Bachillerato'); }
function crearHojaQuintoBachillerato() { crearHojaGrado('Quinto Bachillerato'); }
function crearHojaSextoBachillerato()  { crearHojaGrado('Sexto Bachillerato'); }

/**
 * Crea todas las hojas de grado de una sola vez.
 */
function crearTodasLasHojas() {
  GRADOS.forEach(g => crearHojaGrado(g));
  SpreadsheetApp.getUi().alert('✅ Todas las hojas de grado fueron creadas correctamente.');
}
