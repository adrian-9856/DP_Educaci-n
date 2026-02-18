// ============================================================
//  INSTALACIÓN DE TRIGGERS
// ============================================================

/**
 * Instala el trigger onEdit instalable (necesario para que onEdit
 * pueda usar SpreadsheetApp.getUi() y otras APIs avanzadas).
 *
 * Ejecutar UNA SOLA VEZ desde el menú o el editor de Apps Script.
 */
function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Eliminar triggers anteriores del mismo tipo para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Instalar nuevo trigger onEdit instalable
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Trigger instalado correctamente.\n\n' +
    'Ahora, cada vez que cambies la columna "Acción" en la hoja Interés, ' +
    'el estudiante se moverá automáticamente a su hoja de grado.'
  );
}

/**
 * Elimina todos los triggers del proyecto (útil para reiniciar).
 */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('Todos los triggers fueron eliminados.');
}
