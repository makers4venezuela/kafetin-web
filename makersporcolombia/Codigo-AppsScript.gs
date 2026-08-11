/**
 * Makers por Colombia — backend en Google Sheets + Drive
 *
 * PRIMER USO (una sola vez):
 *   1. Selecciona la función  setup  y pulsa Ejecutar
 *   2. Acepta los permisos que pida
 *   3. Implementar → Nueva implementación → Aplicación web
 *      Ejecutar como: Yo · Acceso: Cualquier persona
 */
var HOJA     = 'Registro';
var CARPETA  = 'Makers por Colombia - Fotos';
var CABECERA = ['Fecha','Voluntario','Correo','Telefono','Taller','Ciudad','Pais',
                'Categoria','Modelo','Archivo','Variante','Fabricadas','Destino',
                'Notas','Consentimiento','Foto'];

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(HOJA); }

  sh.getRange(1, 1, 1, CABECERA.length)
    .setValues([CABECERA])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f3864');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('M:M').setNumberFormat('@');   // Destino siempre texto
  sh.getRange('L:L').setNumberFormat('0');
  var anchos = [95,150,190,130,170,110,95,150,230,230,180,90,150,220,190,220];
  for (var i = 0; i < anchos.length; i++) sh.setColumnWidth(i + 1, anchos[i]);
  if (sh.getMaxColumns() > CABECERA.length)
    sh.deleteColumns(CABECERA.length + 1, sh.getMaxColumns() - CABECERA.length);
  try { if (!sh.getFilter()) sh.getRange(1, 1, sh.getMaxRows(), CABECERA.length).createFilter(); } catch (e) {}

  setupResumen_();
  var f = carpeta_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Listo. Carpeta de fotos: ' + f.getName(), 'Makers por Colombia', 8);
  return 'setup ok - carpeta: ' + f.getUrl();
}

var CATALOGO = [
  ['Braquiopalmar', 'Braquiopalmar termoformada — Adulto', 'Braquiopalmar-Adulto.3mf'],
  ['Braquiopalmar', 'Braquiopalmar termoformada — Infantil', 'Braquiopalmar-Infantil.3mf'],
  ['Carpo-Palmar', 'Carpo-Palmar termoformada — Mediano', 'Carpo-Palmar-Mediano.3mf'],
  ['Carpo-Palmar', 'Carpo-Palmar termoformada — Pequeño', 'Carpo-Palmar-Pequeno.3mf'],
  ['Dedo', 'Fijación de dedo', 'Dedo.3mf'],
  ['Archivo', 'Fijación de dedo (V3)', 'FIXACE PRSTU V3.stl'],
  ['Archivo', 'Fijación de dedo completo (V2)', 'FIXACE CELEHO PRSTU V2.stl'],
  ['Archivo', 'Fijación pulgar–meñique (V3)', 'FIXACE PALEC-MALIK V3.stl'],
  ['Archivo', 'FTM Pequeña — Soporte metacarpo', 'FTM PEQUEÑA - Soporte Metacarpo.stl'],
  ['Archivo', 'FTM Mediana — Soporte metacarpo', 'FTM MEDIANA - Soporte Metacarpo.stl'],
  ['Archivo', 'FTM Grande — Derecha', 'FTM GRANDE.stl'],
  ['Archivo', 'FTM Grande — Izquierda', 'FTM GRANDE IZQ(1).stl'],
  ['Archivo', 'FTM — 4.º y 5.º metacarpo', 'FTM - Mano 4toy 5to metacarpo.stl'],
  ['Archivo', 'Férula Intrínseco Plus — Derecha', 'Férula Intrinseco Plus - Derecho.stl'],
  ['Archivo', 'Férula Intrínseco Plus — Izquierda', 'Férula Intrinseco Plus - Izquierdo.stl'],
  ['Archivo', 'Férula de muñeca panal (con logo)', 'wristSplint-honeycomb-h2 logo'],
  ['Archivo', 'Antebrazo — Braquiopalmar mediana', 'Antebrazo - Modelo Braquiopalmar Mediana.stl'],
  ['Archivo', 'Antebrazo infantil', 'Ante Brazo Infantil.stl'],
  ['Archivo', 'Codo — Braquiopalmar mediana', 'Codo - Modelo Braquiopalmar Mediana.stl'],
  ['Archivo', 'Codo infantil', 'Codo Infantil.stl'],
  ['Archivo', 'Codo infantil — contextura delgada', 'Codo Infantil - Contextura Delgada.stl'],
  ['Archivo', 'Fijación de pie / dedo gordo (V2)', 'FIXACE CHODIDLA - palec V2.stl']
];
var CATEGORIAS = ['Braquiopalmar', 'Carpo-Palmar', 'Dedo', 'Archivo', 'Otro'];

function setupResumen_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Resumen') || ss.insertSheet('Resumen');
  sh.clear();

  sh.getRange('A1').setValue('Makers por Colombia — Resumen de produccion')
    .setFontSize(15).setFontWeight('bold').setFontColor('#1f3864');
  sh.getRange('A2').setValue('Todo se calcula solo desde la hoja Registro. No escribas nada aqui.')
    .setFontColor('#888888').setFontStyle('italic');

  sh.getRange(4, 1, 1, 6)
    .setValues([['Fabricadas','Registros','Talleres','Ciudades origen','Destinos','Modelos']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f3864')
    .setHorizontalAlignment('center');
  sh.getRange(5, 1, 1, 6).setValues([[
    '=SUM(Registro!L:L)',
    '=COUNTA(Registro!I:I)-1',
    '=IFERROR(SUMPRODUCT((Registro!E2:E5000<>"")/COUNTIF(Registro!E2:E5000,Registro!E2:E5000&"")),0)',
    '=IFERROR(SUMPRODUCT((Registro!F2:F5000<>"")/COUNTIF(Registro!F2:F5000,Registro!F2:F5000&"")),0)',
    '=IFERROR(SUMPRODUCT((Registro!M2:M5000<>"")/COUNTIF(Registro!M2:M5000,Registro!M2:M5000&"")),0)',
    '=IFERROR(SUMPRODUCT((Registro!I2:I5000<>"")/COUNTIF(Registro!I2:I5000,Registro!I2:I5000&"")),0)'
  ]]).setFontSize(16).setFontWeight('bold').setFontColor('#1f3864')
    .setBackground('#fff2cc').setHorizontalAlignment('center')
    .setNumberFormat('0');   // sin esto hereda el % del Resumen anterior

  sh.getRange('A7').setValue('Por categoria').setFontWeight('bold').setFontColor('#1f3864');
  sh.getRange(8, 1, 1, 2).setValues([['Categoria','Fabricadas']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f3864');
  var f = [];
  for (var i = 0; i < CATEGORIAS.length; i++) {
    var r = 9 + i;
    f.push([CATEGORIAS[i], '=SUMIF(Registro!$H:$H,$A' + r + ',Registro!$L:$L)']);
  }
  var tot = 9 + CATEGORIAS.length;
  f.push(['Total', '=SUM(B9:B' + (tot - 1) + ')']);
  sh.getRange(9, 1, f.length, 2).setValues(f);
  sh.getRange(tot, 1, 1, 2).setFontWeight('bold').setBackground('#f2f2f2');

  // Por destino: la lista crece sola, asi que se resuelve con QUERY
  sh.getRange(8, 4).setValue('Por destino')
    .setFontWeight('bold').setFontColor('#1f3864');
  sh.getRange(9, 4).setFormula(
    '=IFERROR(QUERY(Registro!L2:M, "select Col2, sum(Col1) ' +
    'where Col2 is not null group by Col2 ' +
    'order by sum(Col1) desc label Col2 \'Destino\', sum(Col1) \'Fabricadas\'"), ' +
    '"Sin destinos registrados")');
  sh.getRange(9, 4, 1, 2).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f3864');

  var ini = tot + 2;
  sh.getRange(ini, 1).setValue('Por producto')
    .setFontWeight('bold').setFontColor('#1f3864');
  sh.getRange(ini + 1, 1, 1, 4)
    .setValues([['Categoria','Modelo','Archivo .stl','Fabricadas']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f3864');
  var g = [];
  for (var j = 0; j < CATALOGO.length; j++) {
    var rr = ini + 2 + j;
    g.push([CATALOGO[j][0], CATALOGO[j][1], CATALOGO[j][2],
            '=SUMIF(Registro!$I:$I,$B' + rr + ',Registro!$L:$L)']);
  }
  sh.getRange(ini + 2, 1, g.length, 4).setValues(g);
  var fin = ini + 2 + g.length;
  sh.getRange(fin, 1, 1, 4).setValues([['', 'Total catalogo', '',
    '=SUM(D' + (ini + 2) + ':D' + (fin - 1) + ')']])
    .setFontWeight('bold').setBackground('#f2f2f2');

  sh.getRange(fin + 2, 1).setValue(
    'Los modelos escritos a mano ("Otro modelo") no salen en esta tabla pero si en la fila "Otro" de arriba. ' +
    'Para desglosar por taller o ciudad usa Datos - Tabla dinamica sobre la hoja Registro.')
    .setFontColor('#888888');

  var anchos = [220, 300, 300, 140, 110, 130, 110];
  for (var k = 0; k < anchos.length; k++) sh.setColumnWidth(k + 1, anchos[k]);
  sh.setFrozenRows(2);
}

function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(HOJA); }
  if (sh.getLastRow() === 0)
    sh.getRange(1, 1, 1, CABECERA.length).setValues([CABECERA]).setFontWeight('bold');
  return sh;
}

function carpeta_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOTOS_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var it = DriveApp.getFoldersByName(CARPETA);
  var f = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA);
  props.setProperty('FOTOS_ID', f.getId());
  return f;
}

function limpio_(s) {
  return String(s || '').replace(/[\\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').substring(0, 60).trim();
}

function guardarFoto_(dataUrl, r) {
  if (!dataUrl || dataUrl.indexOf('base64,') < 0) return '';
  try {
    var p    = dataUrl.split('base64,');
    var tipo = p[0].replace('data:', '').replace(';', '') || 'image/jpeg';
    var ext  = tipo.indexOf('png') > -1 ? '.png' : '.jpg';
    var nom  = [limpio_(r.fecha), limpio_(r.taller), limpio_(r.modelo),
                new Date().getTime()].join(' _ ') + ext;
    var file = carpeta_().createFile(
      Utilities.newBlob(Utilities.base64Decode(p[1]), tipo, nom));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return file.getUrl();
  } catch (err) { return ''; }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (err) { return json_({ ok: false, error: 'ocupado' }); }
  try {
    var sh = hoja_();
    var d  = JSON.parse(e.postData.contents);
    var f  = Array.isArray(d) ? d : [d];
    var filas = [];
    for (var i = 0; i < f.length; i++) {
      var r = f[i];
      if (!r || !r.modelo) continue;
      filas.push([r.fecha, r.voluntario, r.correo, r.telefono, r.taller, r.ciudad, r.pais,
                  r.categoria, r.modelo, r.archivo, r.variante,
                  Number(r.fabricadas) || 0, r.destino,
                  r.notas, r.consentimiento, guardarFoto_(r.foto, r)]);
    }
    if (filas.length)
      sh.getRange(sh.getLastRow() + 1, 1, filas.length, CABECERA.length).setValues(filas);
    return json_({ ok: true, n: filas.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var v = hoja_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][8]) continue;
    out.push({
      fecha:      v[i][0] instanceof Date
                    ? Utilities.formatDate(v[i][0], 'GMT-5', 'yyyy-MM-dd')
                    : String(v[i][0]),
      voluntario: v[i][1],  taller:   v[i][4],
      ciudad:     v[i][5],  pais:     v[i][6],
      categoria:  v[i][7],  modelo:   v[i][8],
      archivo:    v[i][9],  variante: v[i][10],
      fabricadas: Number(v[i][11]) || 0,
      destino:    v[i][12],
      notas:      v[i][13],
      foto:       v[i][15],
      fotoImg:    miniatura_(v[i][15])
    });
  }
  // Correo y telefono NO se exponen: el dashboard es publico.
  var cb = e && e.parameter && e.parameter.callback;
  var s  = JSON.stringify(out);
  if (cb) return ContentService.createTextOutput(cb + '(' + s + ')')
                 .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s)
           .setMimeType(ContentService.MimeType.JSON);
}

function miniatura_(url) {
  // getUrl() devuelve la pagina del visor de Drive, que no sirve dentro de un <img>.
  // Se extrae el id y se arma la URL de miniatura, que si es una imagen directa.
  if (!url) return '';
  var m = String(url).match(/[-\w]{25,}/);
  return m ? 'https://drive.google.com/thumbnail?id=' + m[0] + '&sz=w400' : '';
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
           .setMimeType(ContentService.MimeType.JSON);
}
