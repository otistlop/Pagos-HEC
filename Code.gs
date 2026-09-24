// HEC · Pagos — recibe registros de la app y los vacía en una hoja por mes.
// Cada hoja se llama "Octubre 2026", etc., y se crea sola con el primer pago de ese mes.

const ENCABEZADOS = ["Timestamp", "Alumno", "Cantidad", "Comprobante", "Recibió"];
const RECEPTORES = ["Julieta", "Paola"];
const FORMATO_MES = /^(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre) \d{4}$/;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // evita que dos pagos simultáneos se pisen

    const d = JSON.parse(e.postData.contents);

    // Validación del lado del servidor (la app valida, pero no confiamos solo en ella)
    const alumno = String(d.alumno || "").trim();
    const cantidad = Number(d.cantidad);
    const comprobante = String(d.comprobante || "").trim();
    if (!alumno) throw new Error("falta el nombre del alumno");
    if (!(cantidad > 0)) throw new Error("cantidad no válida");
    if (!FORMATO_MES.test(d.mes)) throw new Error("mes no válido");
    if (!comprobante) throw new Error("falta el comprobante");
    if (RECEPTORES.indexOf(d.recibio) === -1) throw new Error("receptor no válido");

    const hoja = obtenerHojaDelMes(d.mes);
    hoja.appendRow([new Date(), alumno, cantidad, comprobante, d.recibio]);

    return responder({ ok: true });
  } catch (err) {
    return responder({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function obtenerHojaDelMes(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(nombre);
  if (hoja) return hoja;

  hoja = ss.insertSheet(nombre);
  hoja.getRange("D:D").setNumberFormat("@");            // texto, para que "0148" no se vuelva 148
  hoja.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]).setFontWeight("bold");
  hoja.setFrozenRows(1);
  hoja.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm");
  hoja.getRange("C:C").setNumberFormat("$#,##0.00");
  hoja.setColumnWidth(1, 140);
  hoja.setColumnWidth(2, 220);
  return hoja;
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
