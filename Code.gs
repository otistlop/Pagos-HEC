const SHEET_ID = '1yYxVYjbs7uO38YD-3IxBJcu_kxcego_flfa5IfKgSkU';
const NOMBRE_HOJA = 'PAGOS';
const ENCABEZADOS = ['Timestamp', 'Alumno', 'Cantidad', 'Fecha', 'Comprobante', 'Recibió'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);

    let hoja = ss.getSheetByName(NOMBRE_HOJA);
    if (!hoja) {
      hoja = ss.insertSheet(NOMBRE_HOJA);
      hoja.appendRow(ENCABEZADOS);
      hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
      hoja.setFrozenRows(1);
    }

    hoja.appendRow([
      new Date(),
      data.alumno,
      data.cantidad,
      data.fecha,
      data.comprobante,
      data.recibio
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
