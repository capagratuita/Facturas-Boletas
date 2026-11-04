// ==========================================
// GESTOR AUTOMÁTICO DE FACTURAS - VERSIÓN FINAL (CON FILTRO DE AÑO)
// ==========================================
const CONFIG = {
  MODELO_PRIMARIO: 'gemini-3.1-flash-lite',
  MODELO_FALLBACK: 'gemini-3.8-flash',
  CONFIANZA_MINIMA: 70,
  TIMEOUT_SEGUNDOS: 90
};

// 1. MENÚ
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Facturas')
      .addItem('1. Instalar Sistema', 'instalarSistema')
      .addItem('2. Configurar API Keys', 'mostrarSidebar')
      .addToUi();
}

// 2. INSTALADOR
function instalarSistema() {
  Logger.log('--- INICIANDO INSTALACIÓN ---');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  if (props.getProperty('ENTRADA_ID')) {
    SpreadsheetApp.getUi().alert('El sistema ya está instalado. Ve al paso 2 para configurar tus API Keys.');
    return;
  }

  try {
    const obtenerOCrear = (parent, nombre) => {
      const carpetas = parent.getFoldersByName(nombre);
      if (carpetas.hasNext()) return carpetas.next();
      return parent.createFolder(nombre);
    };

    const raices = DriveApp.getFoldersByName('Facturas y Boletas Perú');
    const raiz = raices.hasNext() ? raices.next() : DriveApp.createFolder('Facturas y Boletas Perú');
    
    const entrada = obtenerOCrear(raiz, '01 Entrada');
    const procesados = obtenerOCrear(raiz, '02 Procesados');
    const revisar = obtenerOCrear(raiz, '03 Revisar');
    
    props.setProperty('ENTRADA_ID', entrada.getId());
    props.setProperty('PROCESADOS_ID', procesados.getId());
    props.setProperty('REVISAR_ID', revisar.getId());
    props.setProperty('SHEET_ID', ss.getId());
    
    crearHojasExcel(ss);
    
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    ScriptApp.newTrigger('procesarFacturasNuevas').timeBased().everyMinutes(15).create();

    Logger.log('✅ INSTALACIÓN COMPLETADA.');
    SpreadsheetApp.getUi().alert('✅ ¡Carpetas creadas con éxito!\n\n⚠️ PASO IMPORTANTE:\nEl robot aún no puede leer facturas. Ve al menú "⚙️ Facturas" > "2. Configurar API Keys".');
  } catch (e) {
    Logger.log('❌ ERROR EN INSTALACIÓN: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Error al instalar: ' + e.message);
  }
}

// 3. VENTANA DE API KEYS
function mostrarSidebar() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; padding: 15px; color: #333; }
          h3 { color: #1a73e8; margin-top: 0; font-size: 18px; }
          label { font-weight: bold; font-size: 12px; display: block; margin-top: 15px; }
          input { width: 100%; padding: 10px; margin: 5px 0; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-family: monospace; }
          .help { font-size: 11px; color: #666; margin-top: 2px; margin-bottom: 10px; }
          a { color: #1a73e8; text-decoration: none; }
          button { width: 100%; padding: 12px; background-color: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 15px; font-size: 14px; }
          button:hover { background-color: #1557b0; }
          #mensaje { margin-top: 15px; font-size: 13px; text-align: center; font-weight: bold; }
        </style>
      </head>
      <body>
        <h3>Configuración de API Keys</h3>
        <p style="font-size: 12px; color: #666;">Pega tus llaves aquí para activar el robot.</p>
        
        <label>1. API Key de Gemini (IA)</label>
        <input type="password" id="geminiKey" name="k_gemini_1" autocomplete="new-password">
        <div class="help">Consíguela en <a href="https://aistudio.google.com/app/apikey" target="_blank">Google AI Studio</a></div>
        
        <label>2. Token de ApiPerú</label>
        <input type="password" id="apiperuKey" name="k_apiperu_1" autocomplete="new-password">
        <div class="help">Regístrate en <a href="https://apiperu.dev" target="_blank">ApiPerú.dev</a></div>
        
        <button onclick="guardar()">Guardar y Activar Robot</button>
        <div id="mensaje"></div>
        
        <script>
          function guardar() {
            const g = document.getElementById('geminiKey').value;
            const a = document.getElementById('apiperuKey').value;
            const msg = document.getElementById('mensaje');
            if (!g || !a) { msg.style.color = "red"; msg.innerHTML = "Llena ambos campos."; return; }
            
            msg.style.color = "#1a73e8"; 
            msg.innerHTML = "Guardando...";
            
            google.script.run
              .withSuccessHandler(function() {
                msg.style.color = "green"; 
                msg.innerHTML = "✅ ¡Guardado con éxito!";
                setTimeout(function() {
                  google.script.host.close();
                  google.script.run.withSuccessHandler(null).mostrarMensajeExitoFinal();
                }, 1500);
              })
              .withFailureHandler(function(err) {
                msg.style.color = "red"; 
                msg.innerHTML = "Error: " + err.message;
              })
              .guardarClaves(g, a);
          }
        </script>
      </body>
    </html>
  `;
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle('API Keys'));
}

function guardarClaves(g, a) {
  const props = PropertiesService.getScriptProperties();
  if (g) props.setProperty('GEMINI_API_KEY', g);
  if (a) props.setProperty('APIPERU_TOKEN', a);
  Logger.log('API Keys guardadas en el sistema.');
}

function mostrarMensajeExitoFinal() {
  SpreadsheetApp.getUi().alert('🚀 ¡Robot activado con éxito!\n\nYa puedes subir tus fotos, PDFs o XMLs a la carpeta "01 Entrada".');
}

// 4. NORMALIZADOR DE FECHAS
function normalizarFecha(fechaStr) {
  if (!fechaStr) return '';
  let str = String(fechaStr).trim().replace(/[\/.]/g, '-');
  const partes = str.split('-');
  if (partes.length !== 3) return str; 
  let dia, mes, anio;
  if (partes[0].length === 4) { anio = partes[0]; mes = partes[1].padStart(2, '0'); dia = partes[2].padStart(2, '0'); }
  else if (partes[2].length === 4) { dia = partes[0].padStart(2, '0'); mes = partes[1].padStart(2, '0'); anio = partes[2]; }
  else if (partes[2].length === 2) { dia = partes[0].padStart(2, '0'); mes = partes[1].padStart(2, '0'); anio = '20' + partes[2]; }
  else { return str; }
  return `${dia}-${mes}-${anio}`;
}

// 5. PROCESAMIENTO PRINCIPAL
function procesarFacturasNuevas() {
  Logger.log('========================================');
  Logger.log('--- INICIO DE EJECUCIÓN AUTOMÁTICA ---');
  Logger.log('========================================');
  
  const props = PropertiesService.getScriptProperties();
  const ENTRADA_ID = props.getProperty('ENTRADA_ID');
  const SHEET_ID = props.getProperty('SHEET_ID');
  if (!ENTRADA_ID || !SHEET_ID) {
    Logger.log('❌ Sistema no instalado. Abortando.');
    return;
  }
  
  const PROCESADOS_ID = props.getProperty('PROCESADOS_ID');
  const REVISAR_ID = props.getProperty('REVISAR_ID');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  crearHojasExcel(ss);
  const carpeta = DriveApp.getFolderById(ENTRADA_ID);
  const archivos = carpeta.getFiles();
  let procesados = 0;

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (yaFueProcesado(archivo.getId(), ss)) {
      Logger.log(`-> OMITIDO: ${archivo.getName()} ya fue procesado.`);
      continue;
    }
    try {
      Logger.log(`[Analizando] ${archivo.getName()}`);
      procesarUnArchivo(archivo, ss, PROCESADOS_ID, REVISAR_ID);
      procesados++;
      Logger.log(`✅ ÉXITO: ${archivo.getName()}`);
    } catch (e) {
      Logger.log(`❌ ERROR en ${archivo.getName()}: ${e.message}`);
      moverARevisar(archivo, e.message, ss, REVISAR_ID);
    }
  }
  Logger.log(`--- FIN. Procesados: ${procesados} ---`);
  Logger.log('========================================');
}

function crearHojasExcel(ss) {
  if (!ss.getSheetByName('Registro_Facturas')) {
    Logger.log('[Excel] Creando hoja "Registro_Facturas"...');
    const sheet = ss.insertSheet('Registro_Facturas');
    sheet.appendRow(['Nro_Factura', 'Nro_Boleta', 'Proveedor', 'RUC', 'Fecha_Emision', 'Base', 'IGV', 'Total', 'Moneda', 'Enlace_Archivo', 'Estado', 'ID_Archivo', 'Fecha_Proceso']);
    sheet.getRange('A1:M1').setBackground('#1a365d').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange('E:E').setNumberFormat('@'); 
    sheet.getRange('M:M').setNumberFormat('@'); 
  }
  if (!ss.getSheetByName('Observaciones')) {
    Logger.log('[Excel] Creando hoja "Observaciones"...');
    ss.insertSheet('Observaciones').appendRow(['ID_Archivo', 'Nro_Comprobante', 'Observacion', 'Fecha']);
  }
}

function yaFueProcesado(idArchivo, ss) {
  const sheet = ss.getSheetByName('Registro_Facturas');
  if (!sheet) return false;
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) { if (datos[i][11] === idArchivo) return true; } 
  return false;
}

function procesarUnArchivo(archivo, ss, PROCESADOS_ID, REVISAR_ID) {
  const idOriginal = archivo.getId();
  let nombre = archivo.getName();
  let mime = archivo.getMimeType();
  const ext = nombre.split('.').pop().toLowerCase();
  let blob, datos = null, fuenteDatos = '';
  let alertas = []; 

  Logger.log(`[Tipo] MIME: ${mime} | Ext: ${ext}`);

  if (ext === 'xml' || mime.includes('xml')) {
    Logger.log('-> Ruta: Extrayendo datos del XML...');
    datos = extraerDeXML(archivo); fuenteDatos = 'XML'; blob = archivo.getBlob();
    if (!datos) throw new Error('XML inválido');
  } 
  else if (mime.includes('image') || mime.includes('pdf') || ext === 'heic' || ext === 'heif') {
    if (ext === 'heic' || ext === 'heif') {
      Logger.log('-> Convirtiendo HEIC a JPEG...');
      blob = convertirHeicAJpeg(archivo); nombre = nombre.replace(/\.[^/.]+$/, '') + '.jpg';
    } else { 
      blob = archivo.getBlob(); 
    }
    
    if (!mime.includes('pdf')) {
      Logger.log('-> Capa 1: Intentando leer QR...');
      const qrString = leerQR(blob);
      if (qrString) { 
        datos = parsearQR(qrString); 
        if (datos) fuenteDatos = 'QR'; 
        Logger.log('-> QR leído con éxito.');
      } else {
        Logger.log('-> No se detectó QR legible.');
      }
    }
    if (!datos) {
      Logger.log('-> Capa 2: Enviando documento a Gemini...');
      datos = extraerConGemini(blob, mime); fuenteDatos = 'Gemini';
      Logger.log('-> Respuesta de Gemini: ' + JSON.stringify(datos));

      if (Array.isArray(datos) && datos.length > 0) {
        datos = datos[0];
        Logger.log('-> Nota: Gemini devolvió un Array, se extrajo el primer objeto.');
      }
      
      // VALIDACIÓN: Si la IA se equivoca diciendo que no es factura, pero extrajo datos clave, la aceptamos.
      if (datos && datos.es_factura_boleta === false && !datos.ruc && !datos.total) {
        throw new Error('No es factura ni boleta');
      }
    }
  } else { throw new Error('Formato no soportado'); }

  // Filtro de seguridad anti-errores de IA (evitar que copie las llaves de la plantilla)
  if (datos.nro_factura && String(datos.nro_factura).includes('{{')) datos.nro_factura = '';
  if (datos.nro_boleta && String(datos.nro_boleta).includes('{{')) datos.nro_boleta = '';
  if (datos.fecha_emision && String(datos.fecha_emision).includes('{{')) datos.fecha_emision = '';

  if (!datos.base || datos.base === null) {
    alertas.push('Base no visible');
    Logger.log('-> Alerta: No se encontró la Base imponible.');
  }
  if (!datos.igv || datos.igv === null) {
    alertas.push('IGV no visible');
    Logger.log('-> Alerta: No se encontró el IGV.');
  }

  datos.fecha_emision = normalizarFecha(datos.fecha_emision);

  // ---------------------------------------------------------
  // FILTRO ESTRICTO: RECHAZAR AÑOS ANTERIORES AL ACTUAL
  // ---------------------------------------------------------
  if (datos.fecha_emision && datos.fecha_emision.includes('-')) {
    const partesFecha = datos.fecha_emision.split('-');
    if (partesFecha.length === 3) {
      const anioFactura = parseInt(partesFecha[2], 10);
      const anioActual = new Date().getFullYear();
      
      if (anioFactura < anioActual) {
        Logger.log(`-> ⚠️ RECHAZADO: El comprobante es del año ${anioFactura}. Solo se aceptan del ${anioActual}.`);
        throw new Error(`Comprobante de año anterior (${anioFactura}). No válido.`);
      }
    }
  }
  // ---------------------------------------------------------

  if (datos.ruc && datos.ruc.length !== 11) alertas.push('RUC inválido');
  if (datos.ruc && !datos.proveedor) {
    Logger.log(`-> Consultando razón social para RUC ${datos.ruc}...`);
    datos.proveedor = obtenerRazonSocial(datos.ruc) || '';
    if(datos.proveedor) Logger.log(`-> Proveedor encontrado: ${datos.proveedor}`);
  }

  Logger.log('-> Archivando y convirtiendo a PDF...');
  const ruta = archivarYConvertir(archivo, blob, nombre, datos.fecha_emision, PROCESADOS_ID);
  Logger.log(`-> Guardado en: ${ruta.carpeta}`);
  
  escribirExcel(idOriginal, datos, ruta, `Procesado (${fuenteDatos})`, alertas, ss);
}

// 6. GEMINI CON SISTEMA DE REINTENTOS
function extraerConGemini(blob, mime) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if(!key) throw new Error("Falta configurar Gemini API Key.");
  const b64 = Utilities.base64Encode(blob.getBytes());
  
  const promptText = "Extrae JSON: tipo_comprobante, nro_factura, nro_boleta, proveedor, ruc, fecha_emision, base, igv, total, moneda, confianza. Si es una factura o boleta, es_factura_boleta: true. Si NO es una factura ni boleta, es_factura_boleta: false. Si falta base/igv usa null, NO INVENTES DATOS.";
  
  const payload = { contents: [{ parts: [ { text: promptText }, { inline_data: { mime_type: mime.includes('pdf')?'application/pdf':'image/jpeg', data: b64 } } ] }], generationConfig: { response_mime_type: 'application/json' } };
  
  const modelos = [CONFIG.MODELO_PRIMARIO, CONFIG.MODELO_FALLBACK];
  
  for (const mod of modelos) {
    for (let intento = 1; intento <= 3; intento++) {
      try {
        Logger.log(`[Gemini] Llamando a modelo: ${mod} (Intento ${intento}/3)...`);
        const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${mod}:generateContent?key=${key}`, { method:'post', contentType:'application/json', payload:JSON.stringify(payload), muteHttpExceptions:true, timeoutSeconds: CONFIG.TIMEOUT_SEGUNDOS });
        const code = res.getResponseCode();
        
        if (code === 200) {
          Logger.log(`[Gemini] ${mod} respondió 200 OK.`);
          return JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts[0].text);
        } 
        else if (code === 503 || code === 429) {
          Logger.log(`[Gemini] ${mod} saturado (HTTP ${code}). Esperando 10 segundos para reintentar...`);
          Utilities.sleep(10000);
        } 
        else {
          throw new Error(`HTTP ${code}`);
        }
      } catch (e) { 
        Logger.log(`[Gemini] Error con ${mod}: ${e.message}`); 
        if (intento === 3) Logger.log(`[Gemini] Se agotaron los reintentos para ${mod}.`);
      }
    }
  }
  
  throw new Error('Gemini saturado (503) tras múltiples reintentos. Reintenta más tarde.');
}

function leerQR(blob) {
  try {
    const res = UrlFetchApp.fetch('https://useqr.app/api/v1/decode', { method: 'post', payload: { file: blob }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const json = JSON.parse(res.getContentText()); 
    return (json.count > 0 && json.codes[0].text.includes('|')) ? json.codes[0].text.trim() : null;
  } catch (e) { return null; }
}

function parsearQR(str) {
  const p = str.split('|'); if (p.length < 8) return null;
  return { ruc: p[0], tipo_comprobante: p[1] === '01' ? 'factura' : 'boleta', nro_factura: p[1] === '01' ? `${p[2]}-${p[3]}` : '', nro_boleta: p[1] === '03' ? `${p[2]}-${p[3]}` : '', igv: p[4], total: p[5], fecha_emision: p[6], base: (parseFloat(p[5]) - parseFloat(p[4])).toFixed(2), proveedor: '', confianza: 100 };
}

function extraerDeXML(archivo) {
  try {
    const xml = archivo.getBlob().getDataAsString();
    const get = (tag) => { const m = xml.match(new RegExp(`<cbc:${tag}[^>]*>(.*?)</cbc:${tag}>`, 's')); return m ? m[1].replace(/<!\[CDATA\[|\]\]>/g, '').trim() : ''; };
    let ruc = ''; const sup = xml.match(/<cac:AccountingSupplierParty>([\s\S]*?)<\/cac:AccountingSupplierParty>/);
    if (sup) { const r = sup[1].match(/<cbc:ID[^>]*>(.*?)<\/cbc:ID>/); if (r) ruc = r[1].replace(/<!\[CDATA\[|\]\]>/g, '').trim(); }
    const tipo = get('InvoiceTypeCode'), num = get('ID'), fRaw = get('IssueDate');
    return { tipo_comprobante: tipo === '01' ? 'factura' : 'boleta', nro_factura: tipo === '01' ? num : '', nro_boleta: tipo === '03' ? num : '', fecha_emision: fRaw, ruc: ruc, proveedor: get('RegistrationName'), base: get('LineExtensionAmount'), igv: get('TaxAmount'), total: get('TaxInclusiveAmount'), confianza: 100 };
  } catch (e) { return null; }
}

function obtenerRazonSocial(ruc) {
  const token = PropertiesService.getScriptProperties().getProperty('APIPERU_TOKEN');
  if (!token) return null;
  const res = UrlFetchApp.fetch('https://apiperu.dev/api/ruc', { method: 'post', headers: { 'Authorization': `Bearer ${token}` }, payload: JSON.stringify({ ruc: ruc }), contentType: 'application/json', muteHttpExceptions: true });
  if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).data?.nombre_o_razon_social || null;
  return null;
}

function convertirHeicAJpeg(archivo) {
  Logger.log('[HEIC] Convirtiendo archivo usando miniatura de Drive...');
  const response = UrlFetchApp.fetch(`https://drive.google.com/thumbnail?id=${archivo.getId()}&sz=w2000`, { headers: { authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  let blobJpeg = response.getBlob(); return blobJpeg;
}

// 7. ARCHIVADO Y CONVERSIÓN A PDF USANDO GOOGLE SLIDES
function archivarYConvertir(archivo, blob, nombre, fechaStr, PROCESADOS_ID) {
  let f = new Date(); 
  if (fechaStr && /^\d{2}-\d{2}-\d{4}$/.test(fechaStr)) { const p = fechaStr.split('-'); f = new Date(`${p[2]}-${p[1]}-${p[0]}`); }
  const a = f.getFullYear().toString(), m = String(f.getMonth()+1).padStart(2,'0');
  const nMes = new Intl.DateTimeFormat('es-PE',{month:'long'}).format(f);
  const cMes = `${m}-${nMes.charAt(0).toUpperCase()+nMes.slice(1)}`;
  const raiz = DriveApp.getFolderById(PROCESADOS_ID);
  const cA = raiz.getFoldersByName(a).hasNext()?raiz.getFoldersByName(a).next():raiz.createFolder(a);
  const cM = cA.getFoldersByName(cMes).hasNext()?cA.getFoldersByName(cMes).next():cA.createFolder(cMes);
  
  const mime = archivo.getMimeType();
  const ext = nombre.split('.').pop().toLowerCase();
  let finalUrl = '';

  if (ext === 'heic' || ext === 'heif') {
    blob.setName(nombre); 
    const newFile = cM.createFile(blob);
    finalUrl = newFile.getUrl();
    archivo.setTrashed(true); 
    Logger.log('-> HEIC convertido a JPG y guardado.');
  } else {
    archivo.moveTo(cM); 
    finalUrl = archivo.getUrl();
    Logger.log('-> Archivo original movido a la carpeta del mes.');
  }

  // CONVERSIÓN A PDF CON GOOGLE SLIDES (Solo si es imagen)
  if (mime.includes('image') || ext === 'heic' || ext === 'heif') {
    try {
      Logger.log('-> Iniciando conversión a PDF con Google Slides...');
      
      const tempPres = SlidesApp.create('Temp_Conv_' + archivo.getId());
      const slide = tempPres.getSlides()[0];
      
      const img = slide.insertImage(blob);
      
      const imgWidth = img.getWidth();
      const imgHeight = img.getHeight();
      const slideWidth = tempPres.getPageWidth();
      const slideHeight = tempPres.getPageHeight();
      
      const scale = Math.min(slideWidth / imgWidth, slideHeight / imgHeight);
      img.setWidth(imgWidth * scale);
      img.setHeight(imgHeight * scale);
      
      img.setLeft((slideWidth - img.getWidth()) / 2);
      img.setTop((slideHeight - img.getHeight()) / 2);
      
      tempPres.saveAndClose();
      
      const tempFile = DriveApp.getFileById(tempPres.getId());
      const pdfBlob = tempFile.getAs(MimeType.PDF);
      pdfBlob.setName(nombre.replace(/\.[^/.]+$/, "") + ".pdf");
      
      cM.createFile(pdfBlob);
      tempFile.setTrashed(true);
      Logger.log('✅ PDF creado con éxito desde Google Slides.');
    } catch (e) {
      Logger.log('⚠️ No se pudo crear el PDF: ' + e.message);
    }
  }
  
  return { carpeta: `${a}/${cMes}`, url: finalUrl };
}

function escribirExcel(id, datos, ruta, estado, alertas, ss) {
  const sheet = ss.getSheetByName('Registro_Facturas');
  let obs = datos.observacion || ''; 
  if (alertas.length > 0) obs += ' ⚠️ ' + alertas.join('. ');
  
  const fila = sheet.getLastRow() + 1;
  const fechaProcesoStr = Utilities.formatDate(new Date(), "America/Lima", "dd-MM-yyyy HH:mm:ss");
  
  sheet.getRange(fila, 1, 1, 13).setValues([[
    datos.nro_factura || '', datos.nro_boleta || '', datos.proveedor || '', datos.ruc || '', 
    datos.fecha_emision || '', datos.base || '', datos.igv || '', datos.total || '', 
    datos.moneda || 'PEN', ruta.url, estado, id, fechaProcesoStr
  ]]);
  
  if (!datos.base) sheet.getRange(fila, 6).setBackground('#FFD966').setFontColor('#000000');
  if (!datos.igv) sheet.getRange(fila, 7).setBackground('#FFD966').setFontColor('#000000');
  
  if (obs.trim() !== '' && datos.observacion !== 'Extraído de XML') {
    const sheetObs = ss.getSheetByName('Observaciones');
    if (sheetObs) sheetObs.appendRow([id, datos.nro_factura || datos.nro_boleta, obs, fechaProcesoStr]);
  }
}

function moverARevisar(archivo, motivo, ss, REVISAR_ID) {
  try { archivo.moveTo(DriveApp.getFolderById(REVISAR_ID)); } catch(e) {}
  const sheet = ss.getSheetByName('Registro_Facturas');
  if (sheet) {
    const fila = sheet.getLastRow() + 1;
    const fechaProcesoStr = Utilities.formatDate(new Date(), "America/Lima", "dd-MM-yyyy HH:mm:ss");
    sheet.getRange(fila, 1, 1, 13).setValues([['', '', '', '', '', '', '', '', '', '', 'Revisar: ' + motivo, archivo.getId(), fechaProcesoStr]]);
    sheet.getRange(fila, 1, 1, 13).setBackground('#F4CCCC'); 
  }
}
