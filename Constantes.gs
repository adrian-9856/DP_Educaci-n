// ============================================================
//  CONSTANTES GLOBALES
// ============================================================

const HOJA_INTERES = 'Interés';

// Columnas de la hoja Interés (1-based)
const COL_INTERES = {
  NOMBRE:        1,
  EDAD:          2,
  DPI:           3,
  ULTIMO_ANIO:   4,
  PAPELERIA:     5,
  COMENTARIO:    6,
  ACCION:        7
};

// Columnas de cada hoja de grado (1-based)
const COL_GRADO = {
  ID:            1,
  NOMBRE:        2,
  DPI:           3,
  TELEFONO:      4,
  EDAD:          5,
  GRADO:         6,
  MODALIDAD:     7,
  ESTADO:        8
};

// Grados disponibles
const GRADOS = [
  'Primero Básico',
  'Segundo Básico',
  'Tercero Básico',
  'Cuarto Bachillerato',
  'Quinto Bachillerato',
  'Sexto Bachillerato'
];

// Papelería (opciones multi-selección — se usan en notas de celda con checkboxes)
const PAPELERIA_OPCIONES = [
  'Partida de nacimiento',
  'DPI / CUI',
  'Constancia de notas',
  'Foto reciente',
  'Paz y salvo',
  'Formulario de inscripción',
  'Certificado médico'
];

// Modalidades
const MODALIDADES = ['Presencial', 'Semi-presencial'];

// Estados
const ESTADOS = ['Oyente', 'Deserción', 'Graduando'];

// Acciones disponibles en la hoja Interés
const ACCIONES = ['-- Seleccionar --', ...GRADOS.map(g => 'Enviar a: ' + g)];

// Último año cursado opciones
const ULTIMO_ANIO_OPCIONES = [
  'Primaria',
  'Primero Básico',
  'Segundo Básico',
  'Tercero Básico',
  'Cuarto Bachillerato',
  'Quinto Bachillerato'
];

// Color encabezados hoja Interés
const COLOR_HEADER_INTERES = '#1565C0';   // azul oscuro
const COLOR_HEADER_GRADO   = '#2E7D32';   // verde oscuro
const COLOR_FONT_HEADER    = '#FFFFFF';
