/**
 * Receptor de cotizaciones de la App FOSKO (Google Apps Script). receptor probado.
 *
 * Qué hace: cada vez que un cliente aprieta "Pedir cotización", guarda una fila en una
 * planilla de Google (una solicitud por fila), guarda la foto de referencia y el logo en
 * una carpeta de Drive, y te manda un mail con el resumen.
 *
 * Instalación (5 minutos, con la cuenta de FOSKO):
 *  1. Crear una planilla de Google nueva y una carpeta en Drive para las imágenes.
 *  2. En la planilla: Extensiones > Apps Script. Pegar este archivo.
 *  3. Completar SHEET_ID, FOLDER_ID y MAIL_AVISO aquí abajo.
 *  4. Implementar > Nueva implementación > Aplicación web.
 *     Ejecutar como: yo. Quién tiene acceso: cualquiera.
 *  5. Copiar la URL que entrega y pegarla en ENDPOINT dentro de index.html.
 */
const MAIL_AVISO = "iquintar@fosko.cl";
const ADMIN_CLAVE = "fosko-mockup";
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function mockup_(d) {
  if (d.clave !== ADMIN_CLAVE) return out_({ok: false, error: "clave"});
  const R = recursos_();
  const m = String(d.imagen || "").match(/^data:(.+?);base64,(.*)$/);
  if (!m) return out_({ok: false, error: "imagen"});
  const f = DriveApp.getFolderById(R.fid).createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (d.folio || "mockup") + "-mockup.jpg"));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const hoja = SpreadsheetApp.openById(R.sid).getSheets()[0];
  const filas = hoja.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][3]) === String(d.folio)) { hoja.getRange(i + 1, 2).setValue("Mockup enviado"); hoja.getRange(i + 1, COLUMNAS.indexOf("Mockup IA") + 1).setValue(f.getUrl()); break; }
  }
  return out_({ok: true, id: f.getId()});
}
// La planilla y la carpeta de Drive se crean solas la primera vez que llega una solicitud.
function recursos_() {
  const p = PropertiesService.getScriptProperties();
  let sid = p.getProperty("SHEET_ID"), fid = p.getProperty("FOLDER_ID");
  if (!sid) { sid = SpreadsheetApp.create("Solicitudes App FOSKO").getId(); p.setProperty("SHEET_ID", sid); }
  if (!fid) { fid = DriveApp.createFolder("Solicitudes App FOSKO (imagenes)").getId(); p.setProperty("FOLDER_ID", fid); }
  return {sid: sid, fid: fid};
}

const CPV = ["Cliente","Producto Fosko","Producto Cliente","Tipo","Cantidad Anual","Cantidad de entregas al año","Cantidad a cotizar etiquetas","Materialidad",
  "Color","Etiqueta","En donde lleva la etiqueta?","Proceso productivo al cual se somete la etiqueta","Cantidad de colores?","Cantidad de diseños?",
  "Transparencias?","Lleva marca en relieve?","Precio Objetivo","Servicio","Comentarios del servicio","Ejecutivo que solicita"];
const COLUMNAS = ["Fecha","Estado","Duda por resolver","Folio"].concat(CPV, ["Decoración (diseño en app)","Nombre contacto","Cargo","Mail","Fono","Foto referencia","Logo","Mockup IA"]);

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.accion === "mockup") return mockup_(d);
  const R = recursos_();
  const hoja = SpreadsheetApp.openById(R.sid).getSheets()[0];
  if (hoja.getLastRow() === 0) hoja.appendRow(COLUMNAS);
  const folio = "APP-" + Utilities.formatDate(new Date(), "America/Santiago", "yyMMdd") + "-" + hoja.getLastRow();
  const carpeta = DriveApp.getFolderById(R.fid);
  const adjuntos = [];
  const guardar = (dataUrl, nombre) => {
    if (!dataUrl) return "";
    const m = dataUrl.match(/^data:(.+?);base64,(.*)$/);
    if (!m) return "";
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nombre);
    adjuntos.push(blob);
    return carpeta.createFile(blob).getUrl();
  };
  const fotoUrl = guardar(d.fotoRef, folio + "-foto.jpg");
  const logoUrl = guardar(d.logo, folio + "-logo.png");
  const mockUrl = guardar(d.mockup, folio + "-mockup.jpg");
  const c = d.contacto || {};
  hoja.appendRow([new Date(), "Nueva", "", folio].concat(CPV.map(k => d.cpv[k] || ""),
    [d.decoracion, c["Nombre"], c["Cargo"], c["Mail"], c["Fono"], fotoUrl, logoUrl, mockUrl]));
  MailApp.sendEmail(MAIL_AVISO, "Nueva solicitud App FOSKO " + folio + (d.folioRef ? " [cotiza mockup " + d.folioRef + "]" : "") + " - " + d.cpv["Cliente"],
    d.resumen + "\n\nProducto (foto): " + d.urlProducto + "\nFoto referencia: " + fotoUrl + "\nLogo: " + logoUrl + "\nMockup IA: " + mockUrl, {attachments: adjuntos});
  return ContentService.createTextOutput(JSON.stringify({ok: true, folio: folio})).setMimeType(ContentService.MimeType.JSON);
}
