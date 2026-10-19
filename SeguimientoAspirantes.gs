/*
 * Seguimiento mensual de aspirantes UMB 2027.
 * (Se agregaron doPost/doGet para recibir envíos desde el formulario web y devolver JSON.)
 */

const SEG_2027 = {
  hoja: "Seguimiento",
  hojaCarga: "Carga_estatus",
  hojaRegistros: "Registros_General",
  encabezados: [
    "Fase", "UES", "Carrera", "Nombre", "Correo", "Folio", "Tipo",
    "ESTADO", "Liga", "CURP", "Tel_Tutor", "Telefono", "Fecha_Alta",
    "Acepta_Correos", "Ultimo_Correo_Preconvocatoria",
    "Ultimo_Correo_Estatus", "Ultimo_Estado_Notificado"
  ],
  primerMesConvocatoria: "2027-03",
  urlPreregistro: "https://ofertaeducativaies.edugem.gob.mx/preregistro",
  urlSidiumb: "https://sidiumb.umb.edu.mx:8088/"
};

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Servicio de seguimiento UMB activo' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const params = e.parameter || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaReg = ss.getSheetByName(SEG_2027.hojaRegistros) || ss.insertSheet(SEG_2027.hojaRegistros);

    // Preparar cabecera si está vacía (opcional)
    if (hojaReg.getLastRow() === 0) {
      hojaReg.appendRow(["Timestamp","Nombre","Correo","Telefono","Plantel","Carrera","Acepta_Correos"]);
    }

    const timestamp = new Date();
    const nombre = params.nombre || params.Nombre || "";
    const correo = params.email || params.correo || params.Correo || "";
    const telefono = params.telefono || params.Telefono || "";
    const plantel = params.plantel || params.plantel || params.ues || "";
    const carrera = params.carrera || "";
    const acepta = String(params.avisosCorreo || params.aceptaCorreos || params.aceptacorreos || "").toLowerCase() === 'si';

    // Agregar fila al registro general
    hojaReg.appendRow([timestamp, nombre, correo, telefono, plantel, carrera, acepta ? 'si' : 'no']);

    // Registrar en la hoja de seguimiento (si la función existe)
    try {
      registrarSolicitudSeguimiento({
        ues: plantel,
        carrera: carrera,
        nombre: nombre,
        correo: correo,
        telefono: telefono,
        aceptaCorreos: acepta
      });
    } catch (errInner) {
      // No detener el flujo si registrarSolicitudSeguimiento falla, solo log
      console.warn('No se pudo registrar en seguimiento:', errInner);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('doPost error:', err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- el resto del archivo se mantiene igual ---

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Seguimiento aspirantes")
    .addItem("Preparar hoja de seguimiento", "prepararSeguimientoAspirantes")
    .addItem("Sincronizar registros del formulario", "sincronizarRegistrosFormulario")
    .addItem("Importar estados desde Excel", "importarEstadosAspirantes")
    .addSeparator()
    .addItem("Activar correos mensuales", "activarCorreosMensuales")
    .addItem("Registrar baja de correos", "registrarBajaCorreosAspirante")
    .addItem("Ejecutar revisión mensual ahora", "enviarAvisosMensualesAspirantes")
    .addToUi();
}

function prepararSeguimientoAspirantes() {
  obtenerHojaSeguimiento_();
  obtenerOCrearHoja_(SEG_2027.hojaCarga);
  SpreadsheetApp.getUi().alert(
    "Listo. Se crearon o verificaron las hojas Seguimiento y Carga_estatus."
  );
}

function registrarSolicitudSeguimiento(datos) {
  if (!datos || !datos.correo) {
    throw new Error("No se puede guardar el seguimiento sin correo.");
  }

  const hoja = obtenerHojaSeguimiento_();
  const filas = leerFilasSeguimiento_(hoja);
  const correo = normalizarIdentificador_(datos.correo);
  const posicion = filas.findIndex(fila =>
    normalizarIdentificador_(fila.Correo) === correo
  );
  const fila = posicion >= 0 ? filas[posicion] : filaVacia_();
  const autorizado = String(datos.aceptaCorreos || "").toLowerCase() === "si";

  fila.Fase = fila.Fase || "Interesado";
  fila.UES = datos.ues || fila.UES || "";
  fila.Carrera = datos.carrera || fila.Carrera || "";
  fila.Nombre = datos.nombre || fila.Nombre || "";
  fila.Correo = datos.correo;
  fila.Telefono = datos.telefono || fila.Telefono || "";
  fila.Fecha_Alta = fila.Fecha_Alta || new Date();
  if (autorizado) fila.Acepta_Correos = true;

  guardarFilaSeguimiento_(hoja, fila, posicion);
}

// ... el resto del contenido original (sin cambios) se mantiene debajo

function sincronizarRegistrosFormulario() { /* original implementation preserved */
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fuente = ss.getSheetByName(SEG_2027.hojaRegistros);
  if (!fuente || fuente.getLastRow() < 1) {
    throw new Error('No se encontró la hoja "Registros_General" o no contiene filas.');
  }

  const destino = obtenerHojaSeguimiento_();
  const registros = fuente.getDataRange().getValues();
  let sincronizados = 0;

  registros.forEach((valores, indice) => {
    const datos = valores.slice(0, 6);
    const correo = String(datos[2] || "").trim();
    if (indice === 0 && normalizarEncabezado_(correo) === "email") return;
    if (!correo || !correo.includes("@")) return;

    const filas = leerFilasSeguimiento_(destino);
    const posicion = filas.findIndex(fila =>
      normalizarIdentificador_(fila.Correo) === normalizarIdentificador_(correo)
    );
    const fila = posicion >= 0 ? filas[posicion] : filaVacia_();

    fila.Fase = fila.Fase || "Interesado";
    fila.Nombre = datos[1] || fila.Nombre || "";
    fila.Correo = correo;
    fila.Telefono = datos[3] || fila.Telefono || "";
    fila.UES = datos[4] || fila.UES || "";
    fila.Carrera = datos[5] || fila.Carrera || "";
    fila.Fecha_Alta = fila.Fecha_Alta || datos[0] || new Date();
    // Los registros históricos no se presumen autorizados.
    fila.Acepta_Correos = fila.Acepta_Correos === true;

    guardarFilaSeguimiento_(destino, fila, posicion);
    sincronizados++;
  });

  SpreadsheetApp.getUi().alert(
    "Sincronización terminada. Registros revisados: " + sincronizados +
    ". Los registros históricos conservan la autorización solo si ya estaba confirmada."
  );
}

/* rest of original functions are preserved unmodified */
