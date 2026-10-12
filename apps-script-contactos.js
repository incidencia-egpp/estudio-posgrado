/**
 * Contactos MGPP – entrega a cada encuestador/a solo las personas que le tocan.
 * Vive en el Sheets de contactos. Pestaña: la primera, o cambia HOJA abajo.
 * Columnas: las 26 de la base + una llamada "encuestador" con el código de acceso.
 */
const HOJA = 'Contactos';   // nombre de la pestaña; si no existe se usa la primera

// Nombres internos que usa la página. Los encabezados de la hoja se reconocen sin importar mayúsculas ni tildes.
const CAMPOS = [
    'Código PUCP', 'Apellidos y nombres', 'Número de DNI', 'Sexo', 'Edad', 'Especialidad',
    'Modalidad de la especialidad', 'Ciclo de ingreso', 'Ultimo ciclo matrículado', 'Ciclo de egreso',
    'CRAEST', 'Egresado o graduado', 'Fecha de sustentación', 'Correo PUCP', 'Correo externo',
    'Celular (movil)', 'Teléfono (fijo)', 'Es PNP', 'Promoción(es) PNP', 'Ciclo de matrícula en lista PNP',
    'Rango / cargo PNP', 'Grupo PNP', 'Horario PNP', 'Fecha de nacimiento', 'Estado en base histórica',
    'Archivo(s) fuente PNP'
];
const norm_ = t => String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
const CANON_ = {}; CAMPOS.forEach(c => CANON_[norm_(c)] = c);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.accion !== 'contactos') return json_({ ok: true, servicio: 'contactos-mgpp' });

  const codigo = String(p.codigo || '').trim();
  if (!codigo) return json_({ ok: false, error: 'codigo_invalido' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA) || ss.getSheets()[0];
  const datos = hoja.getDataRange().getValues();
  const enc = datos[0].map(h => norm_(h));
  const iEnc = enc.indexOf('encuestador');
  if (iEnc < 0) return json_({ ok: false, error: 'sin_columna_encuestador' });

  const mios = datos.slice(1).filter(f => String(f[iEnc]).trim() === codigo);
  if (!mios.length) return json_({ ok: false, error: 'codigo_invalido' });   // el código no figura en la columna

  const contactos = mios.map(f => {
    const o = {};
    enc.forEach((h, i) => {
      if (!h || h === 'encuestador') return;
      const v = f[i];
      o[CANON_[h] || String(datos[0][i]).trim()] = v instanceof Date ? Utilities.formatDate(v, 'America/Lima', 'yyyy-MM-dd') : (v === '' ? null : v);
    });
    return o;
  });
  return json_({ ok: true, contactos: contactos });
}

/**
 * Ejecutar UNA vez desde el editor (menú de funciones > crearColumnas > Ejecutar).
 * Crea/prepara la hoja con las 27 columnas en el orden de la base y da formato de texto
 * a los códigos, DNI y teléfonos para que no se pierdan los ceros iniciales.
 */
function crearColumnas() {
  const COLUMNAS = CAMPOS.concat(['encuestador']);
  const TEXTO = ['Código PUCP', 'Número de DNI', 'Celular (movil)', 'Teléfono (fijo)', 'encuestador',
                 'Ciclo de ingreso', 'Ultimo ciclo matrículado', 'Ciclo de egreso', 'Ciclo de matrícula en lista PNP'];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(HOJA);
  if (!hoja) {
    const primera = ss.getSheets()[0];
    if (ss.getSheets().length === 1 && primera.getLastRow() === 0) { hoja = primera; hoja.setName(HOJA); }
    else hoja = ss.insertSheet(HOJA);
  }
  if (hoja.getLastRow() > 0 && hoja.getRange(1, 1).getValue() !== '') {
    throw new Error('La hoja "' + HOJA + '" ya tiene datos en la fila 1. Bórrala o renómbrala antes de crear las columnas.');
  }
  hoja.getRange(1, 1, 1, COLUMNAS.length).setValues([COLUMNAS]).setFontWeight('bold').setBackground('#0F2150').setFontColor('#ffffff');
  TEXTO.forEach(n => hoja.getRange(2, COLUMNAS.indexOf(n) + 1, 1000, 1).setNumberFormat('@'));
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, COLUMNAS.length);
}
