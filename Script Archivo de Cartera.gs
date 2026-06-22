/**
 * Sincroniza los datos del archivo CSV de Cartera (Estructura de 70 columnas).
 * Extrae los valores por posiciones fijas absolutas, limpia números y evita duplicados internos.
 */
function sincronizarCarteraSoloNuevos() {
  // --- CONFIGURACIÓN GENERAL ---
  const ID_CSV_ORIGEN = "1M5dWA8WDhfGnNBxXwmCHDB0sopZIeMhN"; 
  const ID_DESTINO = "1z5W97VUbtH6Zlzgq4hEKyoqp5P4R2WGTrG1l1f-D88E"; 
  const NOMBRE_HOJA_DESTINO = "gestion_cartera"; 

  let datosOrigen;

  // 1. LEER EL ARCHIVO CSV
  try {
    const blob = DriveApp.getFileById(ID_CSV_ORIGEN).getBlob();
    const csvTexto = blob.getDataAsString(); 
    datosOrigen = Utilities.parseCsv(csvTexto, ','); 
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error procesando el archivo CSV: " + e.message);
    return;
  }

  // 2. POSICIONES EXACTAS DEL ARCHIVO
  const idxOrigen = {
    CORTE: 2,           
    NIT: 29,            
    CENTRO_COSTO: 30,   
    FACTURA: 35,        
    EMISION: 39,        
    VALOR: 41,          
    ABONOS: 42,         
    ATRASO: 43,         
    SALDO: 50,          
    TOTAL_SALDO: 51,    
    COD_CC: 33          
  };

  // 3. CONFIGURAR HOJA DE DESTINO
  const ssDestino = SpreadsheetApp.openById(ID_DESTINO);
  const hojaDestino = ssDestino.getSheetByName(NOMBRE_HOJA_DESTINO);

  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja de destino");
    return;
  }

  const COL_DESTINO = {
    CORTE: 1, FACTURA: 4, VALOR: 6, ABONOS: 7, ATRASO: 8, SALDO: 9, TOTAL_SALDO: 10,
    CORREO: 12, GESTION: 13, ESTADO: 14, SEGUIMIENTO: 15
  };

  // 4. MAPEAR FACTURAS EXISTENTES EN DESTINO
  const datosDestino = hojaDestino.getDataRange().getValues();
  const mapaDestino = new Map(); 
  
  if (datosDestino.length > 1) {
    for (let i = 1; i < datosDestino.length; i++) {
      let numFacturaDestino = String(datosDestino[i][COL_DESTINO.FACTURA - 1]).trim();
      if (numFacturaDestino !== "") {
        mapaDestino.set(numFacturaDestino, i + 1);
      }
    }
  }

  // FUNCIÓN LIMPIADORA DE NÚMEROS
  const limpiarNumero = (texto) => {
    if (!texto) return 0;
    let limpio = String(texto).replace(/,/g, '').trim();
    let numero = parseFloat(limpio);
    return isNaN(numero) ? 0 : numero;
  };

  const registrosNuevos = [];
  let facturasActualizadas = 0;
  
  // 🛡️ ESCUDO ANTI-DUPLICADOS INTERNOS
  const facturasProcesadasEnEsteCSV = new Set();

  // 5. PROCESAR FILAS DEL CSV
  for (let i = 0; i < datosOrigen.length; i++) {
    let filaCSV = datosOrigen[i];
    
    if (filaCSV.length < 55) continue; 

    let numFacturaOrigen = String(filaCSV[idxOrigen.FACTURA]).trim();

    if (numFacturaOrigen === "" || isNaN(numFacturaOrigen)) continue; 
    
    // Si la factura ya la leímos líneas más arriba en este mismo archivo, la ignoramos
    if (facturasProcesadasEnEsteCSV.has(numFacturaOrigen)) continue;
    facturasProcesadasEnEsteCSV.add(numFacturaOrigen);

    let valCorte      = String(filaCSV[idxOrigen.CORTE]).trim();
    
    let valValor      = limpiarNumero(filaCSV[idxOrigen.VALOR]);
    let valAbonos     = limpiarNumero(filaCSV[idxOrigen.ABONOS]);
    let valAtraso     = limpiarNumero(filaCSV[idxOrigen.ATRASO]);
    let valSaldo      = limpiarNumero(filaCSV[idxOrigen.SALDO]);
    let valTotalSaldo = limpiarNumero(filaCSV[idxOrigen.TOTAL_SALDO]);
    
    let codCCRaw      = String(filaCSV[idxOrigen.COD_CC]);
    let valCodCC      = codCCRaw.replace(/Centro De Costo:/i, "").trim();

    if (mapaDestino.has(numFacturaOrigen)) {
      // ACTUALIZACIÓN
      let filaRealDestino = mapaDestino.get(numFacturaOrigen);
      
      hojaDestino.getRange(filaRealDestino, COL_DESTINO.CORTE).setValue(valCorte);
      hojaDestino.getRange(filaRealDestino, COL_DESTINO.VALOR).setValue(valValor);
      hojaDestino.getRange(filaRealDestino, COL_DESTINO.ABONOS).setValue(valAbonos);
      hojaDestino.getRange(filaRealDestino, COL_DESTINO.ATRASO).setValue(valAtraso);
      hojaDestino.getRange(filaRealDestino, COL_DESTINO.SALDO).setValue(valSaldo);
      hojaDestino.getRange(filaRealDestino, COL_DESTINO.TOTAL_SALDO).setValue(valTotalSaldo);
      
      facturasActualizadas++;
    } else {
      // INSERCIÓN
      let nuevaFila = [
        valCorte,                                             
        String(filaCSV[idxOrigen.NIT]).trim(),                
        String(filaCSV[idxOrigen.CENTRO_COSTO]).trim(),       
        numFacturaOrigen,                                     
        String(filaCSV[idxOrigen.EMISION]).trim(),            
        valValor,                                             
        valAbonos,                                            
        valAtraso,                                            
        valSaldo,                                             
        valTotalSaldo,                                        
        valCodCC,                                             
        "",                                                   
        "PENDIENTE",                                          
        "ABIERTO",                                            
        ""                                                    
      ];
      
      registrosNuevos.push(nuevaFila);
    }
  }

  // 6. INSERTAR REGISTROS NUEVOS
  if (registrosNuevos.length > 0) {
    let ultimaFilaDestino = hojaDestino.getLastRow();
    if (ultimaFilaDestino === 0) ultimaFilaDestino = 1; 
    
    hojaDestino.getRange(ultimaFilaDestino + 1, 1, registrosNuevos.length, registrosNuevos[0].length)
               .setValues(registrosNuevos);
  }

  SpreadsheetApp.flush();

  // 7. NOTIFICACIONES
  if (registrosNuevos.length > 0 || facturasActualizadas > 0) {
    SpreadsheetApp.getActive().toast(
      `✓ ${registrosNuevos.length} facturas nuevas.\n✓ ${facturasActualizadas} facturas actualizadas.`, 
      "Sincronización Completa", 
      8
    );
  } else {
    SpreadsheetApp.getActive().toast("Toda la cartera está cargada y al día. No hay cambios.", "Sin cambios", 5);
  }
}
/**
 * Se ejecuta al editar una celda para enviar correo consolidado por cliente.
 * SOLO incluye facturas con estado "Abierto".
 * Autocompleta la "Gestión" y el "Correo" en todas las filas del cliente.
 */
function alEditarGestion(e) {
  if (!e || !e.range) return;

  var hoja = e.range.getSheet();
  if (hoja.getName() !== "gestion_cartera") return;

  var filaEditada = e.range.getRow();
  var columnaEditada = e.range.getColumn();
  var valorEditado = e.value;

  if (filaEditada < 2 || !valorEditado) return;

  // =======================================================
  // ⚙️ CONFIGURACIÓN DE COLUMNAS
  // =======================================================
  var COL_GESTION = 13;   
  var COL_CORREO = 12; // 👈 AGREGADO: Definimos la columna de correo (Ej: Columna L = 12)

  // =======================================================
  // 🖼️ CONFIGURACIÓN DEL LOGO
  // =======================================================
  var ID_LOGO_DRIVE = "1Qmnilb2fKxlqu9TRZph200Zusp7XQ-Tu"; 

  if (columnaEditada === COL_GESTION && valorEditado.toLowerCase() === "envio de cartera") {
    
    // 1. Obtener todos los datos de la hoja
    var datos = hoja.getDataRange().getValues();

    var IDX_NOMBRE = 2;   // Columna 3 (Nombre)
    var IDX_FACTURA = 3;  // Columna 4 (Factura)
    var IDX_VALOR = 8;    // Columna 9 (Valor)
    var IDX_ATRASO = 7;   // Columna 8 (Atraso)
    var IDX_ABONOS = 6;   // Columna 7 (Abonos)
    var IDX_CORREO_DATOS = 11;  // Índice para buscar el correo (Columna 12 = Índice 11)
    
    var IDX_ESTADO = 13;  

    var nombreCliente = String(datos[filaEditada - 1][IDX_NOMBRE]).trim();
    // 👈 Modificado para usar IDX_CORREO_DATOS
    var correoCliente = String(datos[filaEditada - 1][IDX_CORREO_DATOS]).trim(); 

    if (!correoCliente || !correoCliente.includes("@")) {
      SpreadsheetApp.getActive().toast("No hay un correo válido para " + nombreCliente, "Error", 5);
      return;
    }

    // 2. Buscar facturas ABIERTAS y memorizar en qué filas están
    var facturasCliente = [];
    var saldoTotal = 0;
    var filasAActualizar = []; 

    for (var i = 1; i < datos.length; i++) {
      var filaActual = datos[i];
      var nombreFilaActual = String(filaActual[IDX_NOMBRE]).trim();
      var estadoFilaActual = String(filaActual[IDX_ESTADO]).trim().toLowerCase(); 

      if (nombreFilaActual === nombreCliente && estadoFilaActual === "abierto") {
        var numFactura = filaActual[IDX_FACTURA];
        var valor = parseFloat(filaActual[IDX_VALOR]) || 0; 
        var atraso = filaActual[IDX_ATRASO];
        var abonos = parseFloat(filaActual[IDX_ABONOS]) || 0;

        facturasCliente.push({
          factura: numFactura,
          valor: valor,
          atraso: atraso,
          abonos: abonos
        });

        saldoTotal += valor;
        filasAActualizar.push(i + 1); 
      }
    }

    // 🛑 VALIDACIÓN DE SEGURIDAD
    if (facturasCliente.length === 0) {
      SpreadsheetApp.getActive().toast("No hay facturas 'Abiertas' para este cliente.", "Aviso", 5);
      return; 
    }

    // 3. Obtener el logo desde Drive
    var logoBlob;
    try {
      logoBlob = DriveApp.getFileById(ID_LOGO_DRIVE).getBlob();
    } catch(error) {
      SpreadsheetApp.getActive().toast("Error: No se pudo acceder al logo en Drive.", "Aviso", 5);
      return; 
    }

    // 4. Construir las filas de la tabla
    var filasTablaHtml = "";
    facturasCliente.forEach(function(item) {
      filasTablaHtml += "<tr>" +
        "<td style='text-align: center;'>" + item.factura + "</td>" +
        "<td style='text-align: right;'>$" + item.valor.toLocaleString('es-CO') + "</td>" +
        "<td style='text-align: center;'>" + item.atraso + " días</td>" +
        "<td style='text-align: right;'>$" + item.abonos.toLocaleString('es-CO') + "</td>" +
      "</tr>";
    });

    // 5. Diseñar el cuerpo del correo en HTML
    var asunto = "Estado de cuenta consolidado - " + nombreCliente;
    var cuerpoHtml = 
      "<div style='font-family: Arial, sans-serif; color: #333;'>" +
        "<h2>Hola, " + nombreCliente + "</h2>" +
        "<p>Le escribimos para informarle sobre el estado actual de su cartera. A la fecha, su saldo total pendiente es de <b>$" + saldoTotal.toLocaleString('es-CO') + "</b>.</p>" +
        "<p>A continuación, detallamos las facturas asociadas a su cuenta que se encuentran en estado Abierto:</p>" +
        
        "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 80%;'>" +
          "<tr style='background-color: #f2f2f2;'>" +
            "<th>Factura</th>" +
            "<th>Valor Pendiente</th>" +
            "<th>Días de Atraso</th>" +
            "<th>Abonos Registrados</th>" +
          "</tr>" +
          filasTablaHtml + 
        "</table>" +
        
        "<p>Por favor, póngase en contacto con nosotros si tiene alguna duda para regularizar este saldo.</p>" +
        
        // --- INICIO DE LA FIRMA ---
        "<br><br>" +
        "<p style='margin: 0;'>Muchas gracias</p>" +
        "<br>" +
        "<img src='cid:logoFirma' style='width: 140px;' alt='Logo ASSERVI'>" +
        "<br><br>" +
        "<p style='margin: 0;'><b>Luis Albeiro Cabrera Mosquera</b></p>" +
        "<p style='margin: 2px 0 0 0; color: #333;'>Analista de facturación y cartera</p>" +
        "<p style='margin: 2px 0 0 0; color: #333;'>Tel : (606) 3267167 - (606) 3247043</p>" +
        "<p style='margin: 2px 0 0 0; color: #333;'>Urbanización Belmonte / Lote 7A / Vía Cartago, Pereira</p>" +
        "<p style='margin: 2px 0 0 0; color: #333;'>Pereira, Risaralda - Colombia</p>" +
        "<br>" +
        "<p style='margin: 15px 0 0 0; font-size: 11px; color: #555; text-align: justify; line-height: 1.4;'>" +
        "NOTA CONFIDENCIAL: La información contenida en este e-mail y en todos sus archivos anexos es confidencial de <b>Asservi SAS.</b>, sólo para uso individual del destinatario o entidad a quienes está dirigido. Si usted no es el destinatario, cualquier almacenamiento, distribución, difusión o copia de este mensaje está estrictamente prohibida y sancionada por la ley. Si por error recibe este mensaje, le ofrecemos disculpas, por favor, elimínelo inmediatamente y notifique de su error a la persona que lo envió, absteniéndose de divulgar su contenido y anexos.</p>" +
        // --- FIN DE LA FIRMA ---
        
      "</div>";

    // 6. Enviar el correo y actualizar las filas
    try {
      GmailApp.sendEmail(correoCliente, asunto, "", {
        htmlBody: cuerpoHtml,
        name: "Luis Albeiro Cabrera",
        inlineImages: {
          logoFirma: logoBlob 
        }
      });

      // Actualizamos solo las filas que cumplieron la condición
      filasAActualizar.forEach(function(numeroFila) {
        hoja.getRange(numeroFila, COL_GESTION).setValue("ENVIO DE CARTERA");
        // 👈 AGREGADO: Aquí actualizamos el correo en todas esas filas
        hoja.getRange(numeroFila, COL_CORREO).setValue(correoCliente); 
      });

      SpreadsheetApp.getActive().toast("Correo consolidado enviado a " + nombreCliente, "Éxito", 5);

    } catch (error) {
      SpreadsheetApp.getActive().toast("Error al enviar el correo: " + error.message, "Error", 5);
    }
  }
}
