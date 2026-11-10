/**
 * NODIVA · Receptor de encuesta estudiantil.
 * Vincular este proyecto al Google Sheet existente.
 * Conserva la pestaña institucional antigua; escribe en "Respuestas estudiantes".
 */
const SHEET_ID = '145oJjeFw0_Gr1jZC3vKw6gMgkP1gcqTnR9I97JYYotA';
const TAB = 'Respuestas estudiantes';
const HEADERS = ['Fecha y hora','Colegio','Lo que más gusta','Satisfacción','Atención','Mejoras','Platillo sugerido'];
const ALLOWED = {
  gusto: ['El sabor de los alimentos','La variedad del menú','La cantidad de comida','La atención que recibo'],
  satisfaccion: ['Muy satisfecho','Satisfecho','Poco satisfecho','Nada satisfecho'],
  atencion: ['Excelente y amable','Buena','Regular','Necesita mejorar'],
  mejora: ['Más variedad de platillos','La sazón de algunas comidas','La cantidad de comida','Nada, me gusta como está']
};
function doGet() {
  return ContentService.createTextOutput('NODIVA: receptor disponible').setMimeType(ContentService.MimeType.TEXT);
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const p = e && e.parameter || {};
    for (const key in ALLOWED) {
      if (!ALLOWED[key].includes(p[key])) throw new Error('Respuesta no válida: ' + key);
    }
    lock.waitLock(20000);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(TAB);
    if (!sh) sh = ss.insertSheet(TAB);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
    }
    const clean = (s, max) => {
      const t = String(s || '').slice(0,max).trim();
      return /^[=+@\-\t\r]/.test(t) ? "'" + t : t;
    };
    sh.appendRow([new Date(),clean(p.institucion,160),p.gusto,p.satisfaccion,p.atencion,p.mejora,clean(p.platillo,1500)]);
    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err.message)})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
