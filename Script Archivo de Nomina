// ============================================================
// 🔹 CONFIGURACIÓN
// ============================================================
const ID_EXCEL_DRIVE = "1j3MNzsy7_NkP6nLbByVWC_M88l-f4mE1";

const HOJA_DESTINO = "BD";
const NOMBRE_HOJA_BASE = "Empleados";
const NOMBRE_HOJA_EXCEL_CC = "Centros de costo";


// ============================================================
// 🔹 MENÚ
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🛠️ ')
  
    .addItem('👥 Sincronizar Empleados', 'sincronizarNomina')
    .addItem('🔄 Sincronizar Novedades', 'consolidarNovedadesFinal')
    .addItem('💰 Sincronizar Horas extras', 'consolidarHorasExtras')
    .addItem('📄 Generar Archivo Plano', 'generarArchivoPlanoConFiltroFecha')
    .addToUi();
}


// ============================================================
// 🔹 NOVEDADES (NO SE TOCA)
// ============================================================
function consolidarNovedadesFinal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ss.getSheetByName("CONSOLIDADO DE NOVEDADES");
  
  const FILA_INICIO_DATOS = 7;
  const COL_CONTROL_ORIGEN = 23; // Columna W en las hojas de origen
  const TEXTO_CONTROL = "SINCRONIZADO"; 
  const ESTADO_INICIAL = "PENDIENTE"; 
  
  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja CONSOLIDADO DE NOVEDADES");
    return;
  }

  const hojas = ss.getSheets();
  let registrosParaConsolidar = [];

  hojas.forEach(hoja => {
    let nombreHoja = hoja.getName();
    
    // Filtro de hojas a ignorar
    if (nombreHoja !== "CONSOLIDADO DE NOVEDADES" && nombreHoja !== "BD" && nombreHoja !== "CONSOLIDADO HE") {
      let ultimaFila = hoja.getLastRow();
      if (ultimaFila < FILA_INICIO_DATOS) return;

      let numFilas = ultimaFila - (FILA_INICIO_DATOS - 1);
      let datos = hoja.getRange(FILA_INICIO_DATOS, 1, numFilas, hoja.getLastColumn()).getValues();
      
      for (let i = 0; i < datos.length; i++) {
        let fila = datos[i];
        let yaProcesado = fila[COL_CONTROL_ORIGEN - 1] === TEXTO_CONTROL;
        
        let tieneNovedad = false;
        for (let col = 7; col <= 16; col++) {
          if (fila[col] !== "" && fila[col] !== null) {
            tieneNovedad = true;
            break; 
          }
        }

        if (tieneNovedad && !yaProcesado) {
          // CONSTRUCCIÓN DEL REGISTRO (18 columnas en total)
          let seleccion = [
            nombreHoja, // A (0)
            fila[1],    // B (1)
            fila[2],    // C (2)
            fila[3],    // D (3)
            fila[5],    // E (4)
            fila[6],    // F (5)
            fila[7],    // G (6)
            fila[8],    // H (7)
            fila[9],    // I (8)
            fila[10],   // J (9)
            fila[11],   // K (10)
            fila[12],   // L (11)
            fila[13],   // M (12)
            fila[14],   // N (13)
            fila[15],   // O (14)
            fila[16],   // P (15)
            fila[21],   // Q (16) <- Aquí pegamos la columna V (OBSERVACIONES)
            ESTADO_INICIAL // R (17) <- El estado se mueve a la derecha
          ];
          
          registrosParaConsolidar.push(seleccion);
          
          // Marcar como procesado en la hoja de origen
          hoja.getRange(i + FILA_INICIO_DATOS, COL_CONTROL_ORIGEN).setValue(TEXTO_CONTROL);
        }
      }
    }
  });

  if (registrosParaConsolidar.length > 0) {
    let ultimaFilaDestino = hojaDestino.getLastRow();
    
    // El script detecta automáticamente que registrosParaConsolidar[0].length es 18
    hojaDestino.getRange(ultimaFilaDestino + 1, 1, registrosParaConsolidar.length, registrosParaConsolidar[0].length)
      .setValues(registrosParaConsolidar);

    SpreadsheetApp.getUi().alert("✅ Sincronizadas " + registrosParaConsolidar.length + " novedades con observaciones en Columna Q.");
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ No hay novedades nuevas.");
  }
}
// ============================================================
// 🔹 SINCRONIZAR HORAS EXTRAS
// ============================================================
function consolidarHorasExtras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ss.getSheetByName("CONSOLIDADO HE");
  
  const FILA_INICIO_DATOS = 7; 
  const COL_CONTROL_HE = 24; // Columna X
  const TEXTO_CONTROL = "HE_SINCRONIZADO"; 
  
  if (!hojaDestino) return SpreadsheetApp.getUi().alert("Error: No se encontró la hoja 'CONSOLIDADO HE'");

  const hojas = ss.getSheets();
  let registrosHE = [];
  const ignorar = ["CONSOLIDADO DE NOVEDADES", "CONSOLIDADO HE", "BD"];

  hojas.forEach(hoja => {
    let nombreHoja = hoja.getName();
    
    if (!ignorar.includes(nombreHoja)) {
      let ultimaFila = hoja.getLastRow();
      if (ultimaFila < FILA_INICIO_DATOS) return;

      let numFilas = ultimaFila - (FILA_INICIO_DATOS - 1);
      // Tomamos todas las columnas necesarias
      let datos = hoja.getRange(FILA_INICIO_DATOS, 1, numFilas, hoja.getLastColumn()).getValues();
      
      for (let i = 0; i < datos.length; i++) {
        let fila = datos[i];
        let yaProcesado = fila[COL_CONTROL_HE - 1] === TEXTO_CONTROL;
        
        // Verificamos si hay horas extras (Columnas R a U / Índices 17 a 20)
        let tieneHE = false;
        for (let col = 17; col <= 20; col++) {
          if (fila[col] !== "" && fila[col] !== 0 && fila[col] !== null) {
            tieneHE = true;
            break; 
          }
        }

        if (tieneHE && !yaProcesado) {
          let seleccion = [
            nombreHoja, // A: SUPERVISOR
            fila[1],    // B: FECHA DE CORTE
            fila[2],    // C: CEDULA
            fila[3],    // D: NOMBRE DEL TRABAJADOR
            fila[4],    // E: CODIGO CC
            fila[6],    // F: DIAS LABORADOS (Columna G de la hoja origen)
            fila[17],   // G: HORAS RECARGO NOCTURNO
            fila[18],   // H: HORAS EXTRAS DIURNAS
            fila[19],   // I: HORA EXTRA NOCTURNA
            fila[20]    // J: HORAS DOMINICALES O FESTIVAS
          ];
          registrosHE.push(seleccion);
          
          // Marcar como procesado en la columna X
          hoja.getRange(i + FILA_INICIO_DATOS, COL_CONTROL_HE).setValue(TEXTO_CONTROL);
        }
      }
    }
  });

  if (registrosHE.length > 0) {
    let ultimaFilaDestino = hojaDestino.getLastRow();
    hojaDestino.getRange(ultimaFilaDestino + 1, 1, registrosHE.length, registrosHE[0].length).setValues(registrosHE);
    SpreadsheetApp.getUi().alert("✅ Sincronizados " + registrosHE.length + " registros con Días Laborados.");
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ No hay registros nuevos.");
  }
}
// ============================================================
// 🔹 SINCRONIZAR EMPLEADOS + CENTROS DE COSTO
// ============================================================
function sincronizarNomina() {

  let datosEmp, datosCC;

  try {
    const blob = DriveApp.getFileById(ID_EXCEL_DRIVE).getBlob();
    const tempFile = DriveApp.createFile(blob);

    const converted = Drive.Files.copy({
      mimeType: MimeType.GOOGLE_SHEETS
    }, tempFile.getId());

    const ssTemp = SpreadsheetApp.openById(converted.id);

    datosEmp = ssTemp.getSheetByName(NOMBRE_HOJA_BASE).getDataRange().getValues();
    datosCC  = ssTemp.getSheetByName(NOMBRE_HOJA_EXCEL_CC).getDataRange().getValues();

    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    DriveApp.getFileById(converted.id).setTrashed(true);

  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error leyendo Excel: " + e.message);
    return;
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_DESTINO);

  if (!hoja) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja BD");
    return;
  }
  // ================= EMPLEADOS =================
  const encabezados = datosEmp[0];

  const idxCedula = encabezados.indexOf("Cedula");
  const idxNombre = encabezados.indexOf("Nombre");
  const idxCodigo = encabezados.indexOf("cod.CC");
  const idxDesc   = encabezados.indexOf("Centro.De.Costos");
  const idxArl    = encabezados.indexOf("%ARP");
  const idxEstado = encabezados.indexOf("Retirado");
  const idxContrato = encabezados.indexOf("Contrato");

  if ([idxCedula, idxNombre, idxCodigo, idxDesc, idxArl, idxEstado].includes(-1)) {
    SpreadsheetApp.getUi().alert("❌ Error: Columnas no coinciden con el Excel.");
    return;
  }

  let filas = datosEmp.slice(1).filter(f => f.some(c => c !== ""));

  // Ordenar por contrato DESC
  if (idxContrato !== -1) {
    filas.sort((a, b) => b[idxContrato] - a[idxContrato]);
  }

  // Eliminar duplicados
  const vistos = new Set();
  const filtradas = filas.filter(f => {
    const ced = String(f[idxCedula]).trim();
    if (vistos.has(ced)) return false;
    vistos.add(ced);
    return true;
  });

  // Transformar estado
  const datosFinales = filtradas.map(f => {

   let estadoBase = String(f[idxEstado]).trim().toUpperCase();

let estado;
if (estadoBase === "S") {
  estado = "RETIRADO";
} else if (estadoBase === "N") {
  estado = "ACTIVO";
} else {
  estado = "INACTIVO"; // valor por defecto por seguridad
}

    return [
      f[idxCedula],
      f[idxNombre],
      f[idxCodigo],
      f[idxDesc],
      f[idxArl],
      estado
    ];
  });

  // Limpiar sin borrar encabezados
  hoja.getRange(2, 1, hoja.getMaxRows(), 6).clearContent();

  // Insertar
  hoja.getRange(2, 1, datosFinales.length, 6).setValues(datosFinales);

// 🔻 ELIMINAR FILAS SOBRANTES SOLO EN EMPLEADOS
const totalFilasHoja = hoja.getMaxRows();
const filasUsadas = datosFinales.length + 1; // encabezado

if (totalFilasHoja > filasUsadas) {
  hoja.deleteRows(filasUsadas + 1, totalFilasHoja - filasUsadas);
}

// ================= CENTROS DE COSTO =================
  // 1. Identificar índices exactos en el archivo Excel
  const idxCod      = datosCC[0].indexOf("Código");
  const idxDescCC   = datosCC[0].indexOf("Descripción");
  const idxCedSup   = datosCC[0].indexOf("CEDULA"); // Nombre exacto proporcionado
  const idxNomSup   = datosCC[0].indexOf("NOMBRE SUPERVISOR"); // Nombre exacto proporcionado

  // Validar que se encontraron las columnas para evitar errores de ejecución
  if ([idxCod, idxDescCC, idxCedSup, idxNomSup].includes(-1)) {
    SpreadsheetApp.getActive().toast("⚠️ No se encontraron algunas columnas en Centros de Costo.", "Aviso", 5);
  }

  // 2. Obtener códigos actuales en columna I de la hoja BD
  const rangoCodigos = hoja.getRange(2, 9, hoja.getMaxRows()).getValues();
  const existentes = new Set(
    rangoCodigos.map(f => String(f[0]).trim()).filter(x => x !== "")
  );

  // 3. Filtrar los centros de costo que NO están en la BD
  const nuevos = datosCC.slice(1).filter(f => {
    const cod = String(f[idxCod]).trim();
    return cod !== "" && !existentes.has(cod);
  });

  if (nuevos.length > 0) {
    // 4. Localizar la última fila con datos en la columna I
    const colI = hoja.getRange(2, 9, hoja.getMaxRows()).getValues();
    let ultimaFilaCC = 1; 

    for (let i = 0; i < colI.length; i++) {
      if (colI[i][0] !== "") {
        ultimaFilaCC = i + 2;
      }
    }

    const filaInicio = ultimaFilaCC + 1;

    // 5. Preparar la matriz de 4 columnas: Código, Descripción, Cédula, Nombre Sup.
    const datosAgregar = nuevos.map(f => [
      f[idxCod], 
      f[idxDescCC], 
      f[idxCedSup], 
      f[idxNomSup]
    ]);

    // 6. Pegar datos en el rango I:L (9 es Columna I, el 4 es el ancho de columnas)
    hoja.getRange(filaInicio, 9, datosAgregar.length, 4)
        .setValues(datosAgregar);
  }
// =========================
// 4. ORDENAR EMPLEADOS POR DESCRIPCIÓN
// =========================

// Detectar última fila real de empleados
let ultimaFilaEmpleados = hoja.getRange("A:A").getValues().filter(String).length;

// Obtener encabezados (A-F)
let encabezadosEmp = hoja.getRange(1, 1, 1, 6).getValues()[0];

// Buscar columna "DESCRIPCIÓN"
let colOrden = encabezadosEmp.indexOf("DESCRIPCIÓN") + 1;

// Ordenar solo empleados
if (colOrden > 0 && ultimaFilaEmpleados > 1) {
  hoja.getRange(2, 1, ultimaFilaEmpleados - 1, 6)
    .sort({ column: colOrden, ascending: true });
}
  SpreadsheetApp.getActive().toast(
    `✓ ${datosFinales.length} empleados actualizados\n✓ ${nuevos.length} centros de costo agregados`,
    "Proceso completado ✅",
    6
  );
  }
// ============================================================
// 🔹 ARCHIVO PLANO
// ============================================================


function generarArchivoPlanoConFiltroFecha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. Pedir la fecha al usuario
  const respuesta = ui.prompt('Generar Archivo Plano', 'Ingresa la FECHA DE CORTE (formato DD/MM/AAAA):', ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;
  const fechaIngresada = respuesta.getResponseText().trim();

  if (!fechaIngresada) {
    ui.alert("Debes ingresar una fecha válida.");
    return;
  }

  const hojas = ss.getSheets();
  const FILA_INICIO_DATOS = 7;
  let datosParaArchivo = [];

  // Mapeo de Conceptos
  const MAPEO = {
    17: "15", // Recargo nocturno
    18: "02", // Extras diurnas
    19: "08", // Extra nocturna
    20: "03"  // Dominicales/Festivas
  };

  // 2. Procesar datos directamente de las hojas de supervisores
  hojas.forEach(hoja => {
    let nombreHoja = hoja.getName();
    if (["BD", "CONSOLIDADO HE", "CONSOLIDADO DE NOVEDADES"].indexOf(nombreHoja) === -1) {
      let ultimaFila = hoja.getLastRow();
      if (ultimaFila < FILA_INICIO_DATOS) return;
      let datos = hoja.getRange(FILA_INICIO_DATOS, 1, ultimaFila - 6, 21).getValues();

      datos.forEach(fila => {
        let fechaCelda = fila[1]; 
        if (!fechaCelda || !(fechaCelda instanceof Date)) return;
        let fechaCeldaTexto = Utilities.formatDate(fechaCelda, "GMT-5", "dd/MM/yyyy");
        if (fechaCeldaTexto !== fechaIngresada) return;

        let cedula = fila[2];
        for (let colIndex in MAPEO) {
          let cantidad = fila[colIndex];
          if (cantidad !== "" && cantidad > 0 && cantidad !== null) {
            datosParaArchivo.push([fechaCeldaTexto, cedula, "00", "'" + MAPEO[colIndex], cantidad, "00"]);
          }
        }
      });
    }
  });

  if (datosParaArchivo.length > 0) {
    // 3. CREAR EL ARCHIVO EXCEL DIRECTAMENTE EN UN NUEVO LIBRO
    const nombreExcel = "PLANO_NOMINA_" + fechaIngresada.replace(/\//g, "-");
    const nuevoSS = SpreadsheetApp.create(nombreExcel);
    const hojaNueva = nuevoSS.getSheets()[0];
    hojaNueva.setName("Archivo Plano");

    // Configurar encabezados y datos en el NUEVO libro
    const encabezados = [["Fecha", "Cedula", "Centro de Costo", "CodConcepto", "Cantidad", "Valor"]];
    hojaNueva.getRange(1, 1, 1, 6).setValues(encabezados).setFontWeight("bold");
    
    // Aplicar formatos y pegar datos en el NUEVO libro
    hojaNueva.getRange(2, 3, datosParaArchivo.length, 2).setNumberFormat("@"); 
    hojaNueva.getRange(2, 6, datosParaArchivo.length, 1).setNumberFormat("@");
    hojaNueva.getRange(2, 1, datosParaArchivo.length, 6).setValues(datosParaArchivo);
    
    // 4. GENERAR LINK DE DESCARGA
    const idTemp = nuevoSS.getId();
    const url = "https://docs.google.com/spreadsheets/d/" + idTemp + "/export?format=xlsx";
    
    const htmlOutput = HtmlService
      .createHtmlOutput('<div style="font-family:sans-serif; text-align:center; padding:15px;">' +
                        '<p>Archivo generado para: <b>'+fechaIngresada+'</b></p>' +
                        '<br><a href="' + url + '" target="_blank" onclick="google.script.host.close()" ' +
                        'style="display:inline-block;padding:12px 25px;background-color:#1d6f42;color:white;text-decoration:none;border-radius:5px;font-weight:bold;">' +
                        '📥 DESCARGAR EXCEL</a>' +
                        '<p style="font-size:10px; color:gray; margin-top:15px;">Este proceso no afecta las pestañas de tu libro.</p></div>')
      .setWidth(350).setHeight(170);
      
    ui.showModalDialog(htmlOutput, 'Descarga Lista');
    
    // Enviar el archivo temporal de Drive a papelera
    DriveApp.getFileById(idTemp).setTrashed(true);

  } else {
    ui.alert("ℹ️ No se encontraron registros para la fecha: " + fechaIngresada);
  }
}
