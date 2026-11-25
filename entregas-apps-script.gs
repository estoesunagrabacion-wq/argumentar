/**
 * Planilla de entregas de «Argumentar».
 *
 * Cómo instalarlo (una sola vez):
 * 1. Creá una planilla de Google nueva (por ejemplo, «Argumentar — entregas»).
 * 2. En la planilla: Extensiones → Apps Script. Borrá lo que haya y pegá este archivo completo. Guardá.
 * 3. Implementar → Nueva implementación → tipo «Aplicación web».
 *      Ejecutar como: Yo.   Quién tiene acceso: Cualquier persona.
 *    Autorizá los permisos que pide Google (es tu propia cuenta).
 * 4. Copiá la «URL de la aplicación web» (termina en /exec) y pasásela a Claude para que la cargue en la app.
 *
 * Cada entrega aparece como una fila nueva en la hoja «Entregas».
 */
const CLAVE = 'arzruni-argumentar'; // tiene que coincidir con ENTREGA_CLAVE en la app

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.clave !== CLAVE) return ContentService.createTextOutput('clave incorrecta');
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = libro.getSheetByName('Entregas') || libro.insertSheet('Entregas');
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(['Fecha y hora', 'Nombre', 'Curso', 'Tipo', 'Título', 'Texto entregado', 'Palabras', 'Resumen', 'Puntos', 'Código']);
      hoja.getRange(1, 1, 1, 10).setFontWeight('bold');
      hoja.setFrozenRows(1);
      hoja.setColumnWidth(6, 520);
    }
    const corta = (s, n) => String(s || '').slice(0, n);
    hoja.appendRow([
      new Date(),
      corta(d.nombre, 80),
      corta(d.curso, 20),
      corta(d.tipo, 20),
      corta(d.titulo, 120),
      corta(d.texto, 40000),
      Number(d.palabras) || 0,
      corta(d.resumen, 500),
      Number(d.puntos) || 0,
      corta(d.codigo, 10)
    ]);
    hoja.getRange(hoja.getLastRow(), 6).setWrap(true);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  }
}
