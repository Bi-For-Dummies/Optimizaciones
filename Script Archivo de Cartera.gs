/**
 * Sincroniza los datos de la hoja "Cartera".
 * INSERTA las facturas nuevas y ACTUALIZA los saldos/abonos de las existentes.
 */
function sincronizarCarteraSoloNuevos() {
  // --- CONFIGURACIÓN GENERAL ---
  const ID_EXCEL_ORIGEN = "1ptbn2cBQh9jfSlt0_1In9kmUFJ9KAX70"; 
  const NOMBRE_HOJA_EXCEL = "Cartera";
  const ID_DESTINO = "1z5W97VUbtH6Zlzgq4hEKyoqp5P4R2WGTrG1l1f-D88E"; 
  const NOMBRE_HOJA_DESTINO = "gestion_cartera"; 

  // --- CONFIGURACIÓN DE ÍNDICES DE FACTURA ---
  const IDX_FACTURA_ORIGEN = 3; 
  const COL_FACTURA_DESTINO = 4; 

  // --- CONFIGURACIÓN DE ÍNDICES PARA ACTUALIZAR (Base 0: A=0, B=1, C=2...) ---
  const IDX_ABONOS = 6;       // Índice para la columna Abonos
  const IDX_ATRASO = 7;       // Índice para la columna Días de Atraso
  const IDX_VALOR = 8;        // Índice para la columna Saldo/Valor
  const IDX_TOTAL_SALDO = 9;  // Índice para la columna Total Saldo

  let datosOrigen;

  try {
    const blob = DriveApp.getFileById(ID_EXCEL_ORIGEN).getBlob();
    const tempFile = DriveApp.createFile(blob);
    const converted = Drive.Files.copy({
      mimeType: MimeType.GOOGLE_SHEETS,
      title: "TEMP_CARTERA_SYNC"
    }, tempFile.getId());

    const ssTemp = SpreadsheetApp.openById(converted.id);
    const hojaOrigen = ssTemp.getSheetByName(NOMBRE_HOJA_EXCEL);
    
    if (!hojaOrigen) throw new Error("No se encontró la pestaña '" + NOMBRE_HOJA_EXCEL + "'");
    
    datosOrigen = hojaOrigen.getDataRange().getValues();

    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    Drive.Files.remove(converted.id);

  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error procesando el archivo origen: " + e.message);
    return;
  }

  const ssDestino = SpreadsheetApp.openById(ID_DESTINO);
  const hojaDestino = ssDestino.getSheetByName(NOMBRE_HOJA_DESTINO);

  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja de destino");
    return;
  }

  // 2. Separar datos de origen y contar columnas REALES del Excel
  const encabezadosExcel = datosOrigen[0];
  const numColumnasReales = encabezadosExcel.filter(c => String(c).trim() !== "").length;
  
  // 3. Crear los encabezados de "Correo", "Gestión" y "Estado" en la Fila 1
  hojaDestino.getRange(1, numColumnasReales + 1).setValue("CORREO");
  hojaDestino.getRange(1, numColumnasReales + 2).setValue("GESTION");
  hojaDestino.getRange(1, numColumnasReales + 3).setValue("ESTADO");

  // 4. Leer toda la base de datos de destino
  const datosDestino = hojaDestino.getDataRange().getValues();
  
  // Crear un "mapa" para ubicar rápidamente en qué fila está cada factura existente
  const mapaDestino = new Map(); 
  if (datosDestino.length > 1) {
    // Empezamos en 1 para saltar el encabezado
    for (let i = 1; i < datosDestino.length; i++) {
      // COL_FACTURA_DESTINO es 4 (columna D), restamos 1 para el índice de array (3)
      let numFacturaDestino = String(datosDestino[i][COL_FACTURA_DESTINO - 1]).trim();
      if (numFacturaDestino !== "") {
        mapaDestino.set(numFacturaDestino, i); // Guardamos la factura y su índice de fila
      }
    }
  }

  const registrosNuevos = [];
  let facturasActualizadas = 0;

  // 5. Procesar las filas del Excel para saber cuáles son nuevas y cuáles actualizar
  for (let i = 1; i < datosOrigen.length; i++) {
    let filaExcel = datosOrigen[i];
    let numFacturaOrigen = String(filaExcel[IDX_FACTURA_ORIGEN]).trim();

    if (numFacturaOrigen === "") continue; // Saltar filas vacías

    if (mapaDestino.has(numFacturaOrigen)) {
      // ✅ ACTUALIZACIÓN: La factura ya existe
      let indiceEnDestino = mapaDestino.get(numFacturaOrigen);
      
      // Sobreescribimos solo las columnas de valores
      datosDestino[indiceEnDestino][IDX_ABONOS] = filaExcel[IDX_ABONOS];
      datosDestino[indiceEnDestino][IDX_ATRASO] = filaExcel[IDX_ATRASO];
      datosDestino[indiceEnDestino][IDX_VALOR] = filaExcel[IDX_VALOR];
      
      if (filaExcel.length > IDX_TOTAL_SALDO) {
        datosDestino[indiceEnDestino][IDX_TOTAL_SALDO] = filaExcel[IDX_TOTAL_SALDO];
      }
      
      facturasActualizadas++;
    } else {
      // ✅ INSERCIÓN: La factura es nueva
      let nuevaFila = filaExcel.slice(0, numColumnasReales); 
      nuevaFila.push("");          // Columna: Correo 
      nuevaFila.push("Pendiente"); // Columna: Gestión
      nuevaFila.push("Abierto");   // Columna: Estado
      
      registrosNuevos.push(nuevaFila);
    }
  }

  // 6. Guardar los datos actualizados de vuelta a la hoja destino
  // Esto sobreescribe los datos existentes con los nuevos valores de abonos/saldos de una sola vez
  hojaDestino.getRange(1, 1, datosDestino.length, datosDestino[0].length).setValues(datosDestino);

  // 7. Insertar los registros completamente nuevos al final de la tabla
  if (registrosNuevos.length > 0) {
    hojaDestino.getRange(datosDestino.length + 1, 1, registrosNuevos.length, registrosNuevos[0].length)
               .setValues(registrosNuevos);
  }

  // 8. Notificaciones de éxito
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
