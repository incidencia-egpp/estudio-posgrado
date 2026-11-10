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
    'Archivo(s) fuente PNP', 'Celda de cuota'
];
const norm_ = t => String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
const CANON_ = {}; CAMPOS.forEach(c => CANON_[norm_(c)] = c);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Metas de encuestas completas (cuota total 100). Se cumplen entre las tres encuestadoras, por celda.
const CELDAS = [
  { celda: 1, nombre: 'PNP · 2018–2022', pnp: true, periodo: '2018–2022', sexo: '', modalidad: '', meta: 17 },
  { celda: 2, nombre: 'PNP · 2023–2025', pnp: true, periodo: '2023–2025', sexo: '', modalidad: '', meta: 18 },
  { celda: 3, nombre: 'Civil presencial · 2018–2022 · Mujeres', pnp: false, periodo: '2018–2022', sexo: 'Mujeres', modalidad: 'Presencial', meta: 15 },
  { celda: 4, nombre: 'Civil presencial · 2018–2022 · Hombres', pnp: false, periodo: '2018–2022', sexo: 'Hombres', modalidad: 'Presencial', meta: 14 },
  { celda: 5, nombre: 'Civil presencial · 2023–2025 · Mujeres', pnp: false, periodo: '2023–2025', sexo: 'Mujeres', modalidad: 'Presencial', meta: 14 },
  { celda: 6, nombre: 'Civil presencial · 2023–2025 · Hombres', pnp: false, periodo: '2023–2025', sexo: 'Hombres', modalidad: 'Presencial', meta: 7 },
  { celda: 7, nombre: 'Civil a distancia · 2023–2025 · Mujeres', pnp: false, periodo: '2023–2025', sexo: 'Mujeres', modalidad: 'A distancia', meta: 9 },
  { celda: 8, nombre: 'Civil a distancia · 2023–2025 · Hombres', pnp: false, periodo: '2023–2025', sexo: 'Hombres', modalidad: 'A distancia', meta: 6 }
];
// Meta personal de referencia (solo florurbina y ximenapize llaman a policías).
const METAS_ENC = { florurbina: { pnp: 17, civil: 16 }, ximenapize: { pnp: 18, civil: 16 }, claudiatorres: { pnp: 0, civil: 33 } };

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.accion === 'avance') return avance_(String(p.codigo || '').trim());
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

// Lista de todo el equipo (código, celda y encuestadora) para calcular el avance de cuotas en la página.
function avance_(codigo) {
  if (!codigo) return json_({ ok: false, error: 'codigo_invalido' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA) || ss.getSheets()[0];
  const datos = hoja.getDataRange().getValues();
  const enc = datos[0].map(h => norm_(h));
  const iEnc = enc.indexOf('encuestador'), iCel = enc.indexOf('celda de cuota'), iCod = enc.indexOf('codigo pucp'), iNom = enc.indexOf('apellidos y nombres');
  if (iEnc < 0 || iCel < 0) return json_({ ok: false, error: 'sin_columna_celda' });
  const filas = datos.slice(1).filter(f => String(f[iCod]).trim() !== '');
  if (!filas.some(f => String(f[iEnc]).trim() === codigo)) return json_({ ok: false, error: 'codigo_invalido' });
  return json_({
    ok: true, celdas: CELDAS, metasEnc: METAS_ENC,
    personas: filas.map(f => ({ c: String(f[iCod]).trim(), n: String(f[iNom]).trim(), e: String(f[iEnc]).trim(), k: Number(f[iCel]) || 0 }))
  });
}

/**
 * Ejecutar UNA vez desde el editor (menú de funciones > crearColumnas > Ejecutar).
 * Crea/prepara la hoja con las 27 columnas en el orden de la base y da formato de texto
 * a los códigos, DNI y teléfonos para que no se pierdan los ceros iniciales.
 */
function crearColumnas() {
  const COLUMNAS = CAMPOS.concat(['encuestador']);
  const TEXTO = ['Código PUCP', 'Número de DNI', 'Celular (movil)', 'Teléfono (fijo)', 'encuestador',
                 'Ciclo de ingreso', 'Ultimo ciclo matrículado', 'Ciclo de egreso', 'Ciclo de matrícula en lista PNP', 'Celda de cuota'];
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
