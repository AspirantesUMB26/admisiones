/*
 * Seguimiento mensual de aspirantes UMB 2027.
 *
 * Agrega este archivo al proyecto de Apps Script vinculado a la hoja actual.
 * En doPost(e), después de guardar el registro en Registros_General, agrega:
 *
 * registrarSolicitudSeguimiento({
 *   ues: plantel,
 *   carrera: carrera,
 *   nombre: nombre,
 *   correo: email,
 *   telefono: telefono,
 *   aceptaCorreos: e.parameter.avisosCorreo === "si"
 * });
 *
 * registrarSolicitudSeguimiento envía una confirmación inmediata. Los avisos
 * mensuales siguen dependiendo de la autorización aceptaCorreos.
 *
 * La fecha de inicio se maneja por mes, no se presupone un día de apertura:
 * marzo de 2027 cambia los avisos previos por el seguimiento de estatus.
 * La hoja de cálculo y Apps Script deben usar la zona horaria institucional.
 *
 * Operación inicial: ejecuta "Preparar hoja de seguimiento", sincroniza los
 * registros históricos y autoriza "Activar correos mensuales". Para el Excel,
 * pega/importa la hoja con sus encabezados en la pestaña Carga_estatus.
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

function doPost(e) {
  try {
    const parametros = e && e.parameter;
    if (!parametros) {
      throw new Error("No se recibieron datos del formulario.");
    }

    const datos = {
      nombre: String(parametros.nombre || "").trim(),
      correo: String(parametros.email || "").trim(),
      telefono: String(parametros.telefono || "").trim(),
      ues: String(parametros.plantel || "").trim(),
      carrera: String(parametros.area || "").trim(),
      aceptaCorreos: String(parametros.avisosCorreo || "").toLowerCase() === "si"
    };
    if (!datos.nombre || !datos.correo || !datos.telefono || !datos.ues || !datos.carrera) {
      throw new Error("Faltan campos obligatorios del formulario.");
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(datos.correo)) {
      throw new Error("El correo electrónico no tiene un formato válido.");
    }
    if (!/^\d{10}$/.test(datos.telefono)) {
      throw new Error("El teléfono debe contener 10 dígitos.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      guardarRegistroFormulario_(datos);
    } finally {
      lock.releaseLock();
    }

    let resultadoSeguimiento;
    try {
      resultadoSeguimiento = registrarSolicitudSeguimiento(datos);
    } catch (error) {
      console.error(
        "El registro principal se guardó, pero falló el seguimiento para " +
        datos.correo + ": " + error.message
      );
      return respuestaFormulario_(
        "Success: registro guardado; no se pudo completar el seguimiento."
      );
    }

    if (resultadoSeguimiento && resultadoSeguimiento.correoEnviado === false) {
      console.warn(
        "Registro guardado, pero no se envió correo a " + datos.correo +
        ": " + resultadoSeguimiento.errorCorreo
      );
      return respuestaFormulario_(
        "Success: registro guardado; no se pudo enviar el correo de confirmación."
      );
    }
    return respuestaFormulario_("Success");
  } catch (error) {
    console.error("Error procesando el registro del formulario: " + error.message);
    return respuestaFormulario_("Error: " + error.message);
  }
}

function guardarRegistroFormulario_(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("El proyecto debe estar vinculado a la hoja de cálculo de registros.");
  }

  let hoja = ss.getSheetByName(SEG_2027.hojaRegistros);
  if (!hoja) hoja = ss.insertSheet(SEG_2027.hojaRegistros);

  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha", "Nombre", "Email", "Telefono", "Plantel", "Carrera", "Acepta_Correos"]);
  }

  const ultimaColumna = hoja.getLastColumn();
  const encabezados = ultimaColumna
    ? hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0]
    : [];
  const mapa = crearMapaEncabezados_(encabezados);
  const tieneEncabezados = ["email", "correo", "correo electronico", "e mail", "e-mail"]
    .some(nombre => Object.prototype.hasOwnProperty.call(mapa, nombre));

  if (!tieneEncabezados) {
    hoja.appendRow([
      new Date(),
      protegerValorHoja_(datos.nombre),
      protegerValorHoja_(datos.correo),
      protegerValorHoja_(datos.telefono),
      protegerValorHoja_(datos.ues),
      protegerValorHoja_(datos.carrera)
    ]);
    return;
  }

  const valores = new Array(ultimaColumna).fill("");
  const asignar = (nombres, valor) => {
    for (let i = 0; i < nombres.length; i++) {
      const indice = indiceColumnaOpcional_(mapa, nombres[i]);
      if (indice !== undefined) {
        valores[indice] = valor;
        return;
      }
    }
  };
  asignar(["fecha", "fecha alta", "marca temporal"], new Date());
  asignar(["nombre", "nombre completo"], protegerValorHoja_(datos.nombre));
  asignar(["correo", "email", "correo electronico", "e mail", "e-mail"], protegerValorHoja_(datos.correo));
  asignar(["telefono", "telefono celular", "celular"], protegerValorHoja_(datos.telefono));
  asignar(["ues", "plantel"], protegerValorHoja_(datos.ues));
  asignar(["carrera", "area"], protegerValorHoja_(datos.carrera));
  asignar(["acepta correos", "acepta correos electronicos", "avisos correo"], datos.aceptaCorreos ? "si" : "no");
  hoja.appendRow(valores);
}

function respuestaFormulario_(mensaje) {
  return ContentService.createTextOutput(mensaje)
    .setMimeType(ContentService.MimeType.TEXT);
}

function protegerValorHoja_(valor) {
  const texto = String(valor || "");
  return /^[\s]*[=+\-@]/.test(texto) ? "'" + texto : texto;
}

function prepararSeguimientoAspirantes() {
  obtenerHojaSeguimiento_();
  obtenerOCrearHoja_(SEG_2027.hojaCarga);
  SpreadsheetApp.getUi().alert(
    "Listo. Se crearon o verificaron las hojas Seguimiento y Carga_estatus."
  );
}

function registrarSolicitudSeguimiento(datos) {
  const correoOriginal = datos && String(datos.correo || "").trim();
  if (!correoOriginal || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correoOriginal)) {
    throw new Error("No se puede guardar el seguimiento sin un correo válido.");
  }

  const solicitud = {
    ues: String(datos.ues || "").trim(),
    carrera: String(datos.carrera || "").trim(),
    nombre: String(datos.nombre || "").trim(),
    correo: correoOriginal,
    telefono: String(datos.telefono || "").trim(),
    aceptaCorreos: interpretarBooleano_(datos.aceptaCorreos)
  };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const hoja = obtenerHojaSeguimiento_();
    const filas = leerFilasSeguimiento_(hoja);
    const posicion = buscarPosicionAspirante_(
      filas,
      solicitud.correo,
      solicitud.nombre,
      solicitud.telefono
    );
    const fila = posicion >= 0 ? filas[posicion] : filaVacia_();

    fila.Fase = fila.Fase || "Interesado";
    fila.UES = solicitud.ues || fila.UES || "";
    fila.Carrera = solicitud.carrera || fila.Carrera || "";
    fila.Nombre = solicitud.nombre || fila.Nombre || "";
    fila.Correo = solicitud.correo;
    fila.Telefono = solicitud.telefono || fila.Telefono || "";
    fila.Fecha_Alta = fila.Fecha_Alta || new Date();
    fila.Acepta_Correos = solicitud.aceptaCorreos;

    guardarFilaSeguimiento_(hoja, fila, posicion);

    try {
      const confirmacion = crearCorreoConfirmacion_(solicitud);
      MailApp.sendEmail({
        to: solicitud.correo,
        subject: confirmacion.asunto,
        body: confirmacion.texto,
        htmlBody: confirmacion.html
      });
      return { registrado: true, correoEnviado: true };
    } catch (error) {
      console.error(
        "Solicitud guardada, pero no se pudo enviar confirmación a " +
        solicitud.correo + ": " + error.message
      );
      return { registrado: true, correoEnviado: false, errorCorreo: error.message };
    }
  } finally {
    lock.releaseLock();
  }
}

function sincronizarRegistrosFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fuente = ss.getSheetByName(SEG_2027.hojaRegistros);
  if (!fuente || fuente.getLastRow() < 1) {
    throw new Error('No se encontró la hoja "Registros_General" o no contiene filas.');
  }

  const destino = obtenerHojaSeguimiento_();
  const registros = fuente.getDataRange().getValues();
  const mapaEncabezados = crearMapaEncabezados_(registros[0]);
  const tieneEncabezadoCorreo =
    Object.prototype.hasOwnProperty.call(mapaEncabezados, "email") ||
    Object.prototype.hasOwnProperty.call(mapaEncabezados, "correo") ||
    Object.prototype.hasOwnProperty.call(mapaEncabezados, "correo electronico") ||
    Object.prototype.hasOwnProperty.call(mapaEncabezados, "e mail") ||
    Object.prototype.hasOwnProperty.call(mapaEncabezados, "e-mail");
  let sincronizados = 0;

  registros.forEach((valores, indice) => {
    if (indice === 0 && tieneEncabezadoCorreo) return;

    let datos;
    if (tieneEncabezadoCorreo) {
      const columna = nombres => {
        for (let i = 0; i < nombres.length; i++) {
          const posicion = indiceColumnaOpcional_(mapaEncabezados, nombres[i]);
          if (posicion !== undefined) return valores[posicion];
        }
        return "";
      };
      datos = [
        columna(["fecha", "fecha alta", "marca temporal"]),
        columna(["nombre", "nombre completo"]),
        columna(["correo", "email", "correo electronico", "e mail", "e-mail"]),
        columna(["telefono", "telefono celular", "celular"]),
        columna(["ues", "plantel"]),
        columna(["carrera", "area"])
      ];
    } else {
      // Compatibilidad con la hoja histórica sin fila de encabezados.
      datos = valores.slice(0, 6);
    }

    const correo = String(datos[2] || "").trim();
    if (!correo || !correo.includes("@")) return;

    const filas = leerFilasSeguimiento_(destino);
    const posicion = buscarPosicionAspirante_(
      filas,
      correo,
      datos[1],
      datos[3]
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

function importarEstadosAspirantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const carga = ss.getSheetByName(SEG_2027.hojaCarga);
  if (!carga || carga.getLastRow() < 2) {
    throw new Error(
      'Pega o importa los datos del Excel en "Carga_estatus", con los encabezados en la primera fila.'
    );
  }

  const valores = carga.getDataRange().getValues();
  const indicesColumnas = crearMapaEncabezados_(valores[0]);
  const requeridos = ["ues", "carrera", "nombre", "correo", "folio", "tipo", "estado", "liga", "curp", "tel tutor", "telefono", "fase"];
  const faltantes = requeridos.filter(nombre => !(nombre in indicesColumnas));
  if (faltantes.length) {
    throw new Error("Faltan encabezados en Carga_estatus: " + faltantes.join(", "));
  }

  const hoja = obtenerHojaSeguimiento_();
  const rango = hoja.getLastRow() > 1
    ? hoja.getRange(2, 1, hoja.getLastRow() - 1, SEG_2027.encabezados.length)
    : null;
  const filas = rango ? rango.getValues().map(valoresFila => filaDesdeValores_(valoresFila)) : [];
  const indices = {};
  filas.forEach((fila, posicion) => agregarIndices_(indices, fila, posicion));

  let actualizados = 0;
  let nuevos = 0;
  let omitidos = 0;
  const errores = [];

  for (let numeroFila = 1; numeroFila < valores.length; numeroFila++) {
    const fuente = valores[numeroFila];
    if (fuente.every(celda => celda === "" || celda === null)) continue;

    const dato = {};
    SEG_2027.encabezados.slice(0, 12).forEach(campo => {
      const nombreEncabezado = normalizarEncabezado_(campo.replace(/_/g, " "));
      const indice = indicesColumnas[nombreEncabezado];
      dato[campo] = indice === undefined ? "" : fuente[indice];
    });

    const identidad = clavesIdentidad_(dato);
    if (!identidad.length) {
      omitidos++;
      errores.push("Fila " + (numeroFila + 1) + ": falta Correo, Folio y CURP.");
      continue;
    }

    let estadoCanonico = "";
    if (dato.ESTADO !== "" && dato.ESTADO !== null) {
      estadoCanonico = normalizarEstado_(dato.ESTADO);
      if (!estadoCanonico) {
        omitidos++;
        errores.push("Fila " + (numeroFila + 1) + ": estado no reconocido (" + dato.ESTADO + ").");
        continue;
      }
    }

    const coincidencias = Array.from(new Set(
      identidad.map(clave => indices[clave]).filter(posicionEncontrada =>
        posicionEncontrada !== undefined
      )
    ));
    if (coincidencias.length > 1) {
      omitidos++;
      errores.push(
        "Fila " + (numeroFila + 1) +
        ": Correo, Folio o CURP coinciden con aspirantes distintos; revisar manualmente."
      );
      continue;
    }

    const posicion = coincidencias.length ? coincidencias[0] : -1;
    const esNuevo = posicion < 0;
    const fila = esNuevo ? filaVacia_() : filas[posicion];
    const clavesAnteriores = esNuevo ? [] : clavesIdentidad_(fila);
    if (esNuevo) fila.Fecha_Alta = new Date();

    ["Fase", "UES", "Carrera", "Nombre", "Correo", "Folio", "Tipo", "Liga", "CURP", "Tel_Tutor", "Telefono"].forEach(campo => {
      if (dato[campo] !== "" && dato[campo] !== null) fila[campo] = dato[campo];
    });
    if (estadoCanonico) fila.ESTADO = estadoCanonico;

    const aceptaCorreos = indiceColumnaOpcional_(indicesColumnas, "acepta correos");
    if (aceptaCorreos !== undefined && fuente[aceptaCorreos] !== "") {
      fila.Acepta_Correos = interpretarBooleano_(fuente[aceptaCorreos]);
    } else if (esNuevo) {
      fila.Acepta_Correos = false;
    }

    if (esNuevo) {
      filas.push(fila);
      nuevos++;
    } else {
      actualizados++;
      clavesAnteriores.forEach(clave => {
        if (indices[clave] === posicion) delete indices[clave];
      });
    }
    agregarIndices_(indices, fila, esNuevo ? filas.length - 1 : posicion);
  }

  if (filas.length) {
    hoja.getRange(2, 1, filas.length, SEG_2027.encabezados.length)
      .setValues(filas.map(valoresDesdeFila_));
  }

  const resumen = "Importación terminada. Nuevos: " + nuevos +
    "; actualizados: " + actualizados + "; omitidos: " + omitidos + ".";
  SpreadsheetApp.getUi().alert(
    resumen + (errores.length ? "\n\nRevisar:\n" + errores.slice(0, 12).join("\n") : "")
  );
}

function activarCorreosMensuales() {
  obtenerHojaSeguimiento_();
  const handler = "enviarAvisosMensualesAspirantes";
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger(handler).timeBased().onMonthDay(1).atHour(9).create();
  SpreadsheetApp.getUi().alert(
    "Automatización activada: revisión mensual el día 1, alrededor de las 9:00 a. m. " +
    "La hora depende de la zona horaria del proyecto de Apps Script."
  );
}

function registrarBajaCorreosAspirante() {
  const correo = SpreadsheetApp.getUi().prompt(
    "Registrar baja de correos",
    "Escribe el correo de la persona que solicitó dejar de recibir recordatorios:",
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (correo.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;

  const buscado = normalizarIdentificador_(correo.getResponseText());
  if (!buscado) throw new Error("Debes escribir un correo.");

  const hoja = obtenerHojaSeguimiento_();
  const filas = leerFilasSeguimiento_(hoja);
  const posiciones = [];
  filas.forEach((fila, posicion) => {
    if (normalizarIdentificador_(fila.Correo) === buscado) {
      posiciones.push(posicion);
    }
  });
  if (!posiciones.length) {
    SpreadsheetApp.getUi().alert("No se encontró ese correo en Seguimiento.");
    return;
  }

  posiciones.forEach(posicion => {
    filas[posicion].Acepta_Correos = false;
    hoja.getRange(posicion + 2, 1, 1, SEG_2027.encabezados.length)
      .setValues([valoresDesdeFila_(filas[posicion])]);
  });
  SpreadsheetApp.getUi().alert(
    "Se registró la baja en " + posiciones.length +
    " registro(s) asociados a ese correo. No recibirán más recordatorios automáticos."
  );
}

function enviarAvisosMensualesAspirantes() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  const hoja = obtenerHojaSeguimiento_();
  const filas = leerFilasSeguimiento_(hoja);
  const mesActual = Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    "yyyy-MM"
  );
  const antesDeConvocatoria = mesActual < SEG_2027.primerMesConvocatoria;
  let enviados = 0;
  let omitidos = 0;
  const errores = [];

  filas.forEach((fila, posicion) => {
    if (fila.Acepta_Correos !== true || !fila.Correo) {
      omitidos++;
      return;
    }

    let correo;
    if (antesDeConvocatoria) {
      if (fila.Ultimo_Correo_Preconvocatoria === mesActual) {
        omitidos++;
        return;
      }
      correo = crearCorreoPreconvocatoria_(fila);
    } else {
      const estado = normalizarEstado_(fila.ESTADO);
      const estadoFinal = estado === "Confirmado" || estado === "Confirmado por acuerdo";
      if (estadoFinal) {
        if (fila.Ultimo_Estado_Notificado === estado) {
          omitidos++;
          return;
        }
      } else if (fila.Ultimo_Correo_Estatus === mesActual) {
        omitidos++;
        return;
      }
      correo = crearCorreoEstatus_(fila);
      if (!correo) {
        omitidos++;
        return;
      }
    }

    try {
      MailApp.sendEmail({
        to: String(fila.Correo),
        subject: correo.asunto,
        body: correo.texto,
        htmlBody: correo.html
      });
    } catch (error) {
      errores.push(String(fila.Correo) + ": " + error.message);
      console.error("No se pudo enviar aviso a " + fila.Correo + ": " + error.message);
      return;
    }

    enviados++;
    if (antesDeConvocatoria) {
      fila.Ultimo_Correo_Preconvocatoria = mesActual;
    } else if (correo.estadoFinal) {
      fila.Ultimo_Estado_Notificado = correo.estadoFinal;
    } else {
      fila.Ultimo_Correo_Estatus = mesActual;
    }
    try {
      hoja.getRange(posicion + 2, 1, 1, SEG_2027.encabezados.length)
        .setValues([valoresDesdeFila_(fila)]);
    } catch (error) {
      const detalle = String(fila.Correo) +
        ": correo enviado, pero no se pudo guardar la marca de envío; podría repetirse en la próxima revisión. " +
        error.message;
      errores.push(detalle);
      console.error(detalle);
    }
  });

  const resultado = "Revisión mensual: " + enviados + " correo(s) enviado(s), " +
    omitidos + " omitido(s), " + errores.length + " error(es).";
  console.log(resultado);
  if (errores.length) console.error(errores.join("\n"));
  return resultado;
  } finally {
    lock.releaseLock();
  }
}

function crearCorreoPreconvocatoria_(fila) {
  const nombre = escaparHtml_(fila.Nombre || "aspirante");
  const carrera = escaparHtml_(fila.Carrera || "la carrera de tu interés");
  const mensaje =
    "<p>Te saludamos de la Universidad Mexiquense del Bicentenario.</p>" +
    "<p>La convocatoria y el registro de aspirantes 2027 están previstos para marzo de 2027. " +
    "La fecha exacta aún está pendiente de confirmación; consulta los canales oficiales de la UMB.</p>" +
    "<p>Mientras tanto, puedes revisar la información de " + carrera +
    " y prepararte para realizar tu registro cuando se publique la convocatoria.</p>";
  return crearMensajeCorreo_(
    "Información de admisión UMB 2027",
    nombre,
    mensaje,
    "Para dejar de recibir estos recordatorios, escribe a convocatoria@umb.mx."
  );
}

function crearCorreoConfirmacion_(datos) {
  const nombre = escaparHtml_(datos.nombre || "aspirante");
  const carrera = escaparHtml_(datos.carrera || "No especificada");
  const ues = escaparHtml_(datos.ues || "No especificado");
  const contenido =
    "<p>Confirmamos que recibimos tu solicitud de información de admisión en la UMB.</p>" +
    '<table role="presentation" style="width:100%;border-collapse:collapse;background:#f7f4f1;border-left:4px solid #c09a53;margin:22px 0">' +
    '<tr><td style="padding:14px 16px;color:#333;font-size:14px;line-height:1.7">' +
    "<strong>Plantel:</strong> " + ues + "<br>" +
    "<strong>Carrera:</strong> " + carrera +
    "</td></tr></table>" +
    "<p>Para continuar con tu proceso, utiliza los siguientes accesos oficiales:</p>" +
    crearBotonesAcceso_();

  return crearMensajeCorreo_(
    "Confirmación de solicitud de información - UMB",
    nombre,
    contenido,
    "Los recordatorios mensuales solo se envían si autorizaste recibirlos en el formulario."
  );
}

function crearCorreoEstatus_(fila) {
  const estado = normalizarEstado_(fila.ESTADO);
  const nombre = escaparHtml_(fila.Nombre || "aspirante");
  let detalle;
  if (estado === "Incompleto") {
    detalle = "<p>El sistema indica que tu registro aún está incompleto. Ingresa al preregistro y termina los pasos pendientes. Revisa también SIDIUMB para completar tu proceso.</p>";
  } else if (estado === "No confirmado") {
    detalle = "<p>Tu registro está terminado, pero aparece pendiente la entrega de documentos en el plantel. Acude a tu UES para entregarlos y completar la validación.</p>";
  } else if (estado === "Confirmado") {
    detalle = "<p>Tu registro quedó confirmado. Ahora solo espera el correo con la fecha de tu examen.</p>";
  } else if (estado === "Confirmado por acuerdo") {
    detalle = "<p>Tu estatus quedó confirmado por acuerdo (pase por promedio). Ahora espera la publicación de resultados.</p>";
  } else {
    return null;
  }

  const enlaces = crearBotonesAcceso_();
  const mensaje = crearMensajeCorreo_(
    "Seguimiento de tu registro de admisión UMB 2027",
    nombre,
    detalle + ((estado === "Incompleto" || estado === "No confirmado") ? enlaces : ""),
    "Para dejar de recibir estos recordatorios, escribe a convocatoria@umb.mx."
  );
  if (estado === "Confirmado" || estado === "Confirmado por acuerdo") {
    mensaje.estadoFinal = estado;
  }
  return mensaje;
}

function crearMensajeCorreo_(asunto, nombre, contenido, pie) {
  const encabezado = "<p>Hola, " + nombre + ":</p>";
  const despedida = "<p style=\"margin:26px 0 0;color:#444;font-size:14px;line-height:1.6\">" +
    "Atentamente,<br><strong>Universidad Mexiquense del Bicentenario</strong></p>";
  const enlaceBaja = crearEnlaceSolicitudBaja_();
  const html = '<!doctype html><html lang="es"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
    '<body style="margin:0;padding:0;background:#f2f3f5;font-family:Arial,Helvetica,sans-serif;color:#333">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0">' +
    escaparHtml_(asunto) + ' - Universidad Mexiquense del Bicentenario</div>' +
    '<table role="presentation" style="width:100%;border-collapse:collapse;background:#f2f3f5">' +
    '<tr><td align="center" style="padding:30px 12px">' +
    '<table role="presentation" style="width:100%;max-width:600px;border-collapse:collapse">' +
    '<tr><td style="background:#691c32;border-bottom:5px solid #c09a53;padding:24px 28px">' +
    '<div style="font-size:12px;letter-spacing:1.5px;color:#ead9bc;text-transform:uppercase">Admisión 2027</div>' +
    '<div style="margin-top:7px;color:#fff;font-size:21px;font-weight:bold;line-height:1.3">' +
    'Universidad Mexiquense del Bicentenario</div></td></tr>' +
    '<tr><td style="background:#fff;padding:30px 28px 32px;font-size:15px;line-height:1.75">' +
    encabezado + contenido + despedida +
    '</td></tr><tr><td style="padding:22px 24px;text-align:center;color:#626262;font-size:12px;line-height:1.6">' +
    '<div>' + escaparHtml_(pie) + '</div>' +
    '<div style="margin-top:14px">' + enlaceBaja.html + '</div>' +
    '<div style="margin-top:10px">Universidad Mexiquense del Bicentenario</div>' +
    '</td></tr></table></td></tr></table></body></html>';
  const textoContenido = contenido
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<a\b[^>]*href=["']([^"']+)["'][^>]*>(.*?)<\/a>/gi, "$2: $1")
    .replace(/<[^>]*>/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+/g, " ")
    .replace(/\n\s+/g, "\n")
    .trim();
  const texto = "Hola, " + nombre.replace(/<[^>]*>/g, "") + ":\n\n" +
    textoContenido +
    "\n\nAtentamente,\nUniversidad Mexiquense del Bicentenario\n\n" +
    escaparHtml_(pie) + "\n\n" + enlaceBaja.texto;
  return { asunto: asunto, html: html, texto: texto };
}

function crearBotonesAcceso_() {
  const preregistro = escaparHtml_(SEG_2027.urlPreregistro);
  const sidiumb = escaparHtml_(SEG_2027.urlSidiumb);
  return '<table role="presentation" style="border-collapse:collapse;margin:20px 0">' +
    '<tr><td style="padding:0 10px 10px 0">' +
    '<a href="' + preregistro + '" style="display:inline-block;background:#691c32;border:1px solid #691c32;border-radius:5px;padding:12px 18px;color:#fff;text-decoration:none;font-size:14px;font-weight:bold">Abrir preregistro</a>' +
    '</td><td style="padding:0 0 10px 0">' +
    '<a href="' + sidiumb + '" style="display:inline-block;background:#fff;border:1px solid #691c32;border-radius:5px;padding:12px 18px;color:#691c32;text-decoration:none;font-size:14px;font-weight:bold">Abrir SIDIUMB</a>' +
    '</td></tr></table>';
}

function crearEnlaceSolicitudBaja_() {
  const asunto = encodeURIComponent("Solicitud de baja de correos UMB");
  const cuerpo = encodeURIComponent(
    "Solicito dejar de recibir los recordatorios mensuales de admisión UMB.\n\n" +
    "Mi correo registrado es: [escribe aquí el correo con el que te registraste]"
  );
  const href = "mailto:convocatoria@umb.mx?subject=" + asunto + "&body=" + cuerpo;
  const html = '<a href="' + escaparHtml_(href) +
    '" style="display:inline-block;background:#fff;border:1px solid #691c32;border-radius:5px;padding:10px 16px;color:#691c32;text-decoration:none;font-size:13px;font-weight:bold">Solicitar dejar de recibir correos</a>';
  return {
    html: html,
    texto: "Para solicitar la baja de los recordatorios mensuales, escribe a convocatoria@umb.mx e indica el correo registrado."
  };
}

function obtenerHojaSeguimiento_() {
  const hoja = obtenerOCrearHoja_(SEG_2027.hoja);
  const ultimaColumna = hoja.getLastColumn();
  if (!ultimaColumna || !hoja.getLastRow()) {
    hoja.getRange(1, 1, 1, SEG_2027.encabezados.length)
      .setValues([SEG_2027.encabezados]);
    hoja.setFrozenRows(1);
    return hoja;
  }

  if (ultimaColumna < SEG_2027.encabezados.length) {
    throw new Error(
      'La hoja "Seguimiento" no contiene todas las columnas requeridas.'
    );
  }
  const encabezados = hoja.getRange(1, 1, 1, SEG_2027.encabezados.length).getValues()[0]
    .map(normalizarEncabezado_);
  if (encabezados.join("|") !== SEG_2027.encabezados.map(normalizarEncabezado_).join("|")) {
    throw new Error(
      'La hoja "Seguimiento" ya existe con otros encabezados. Renómbrala o revisa su estructura antes de continuar.'
    );
  }
  return hoja;
}

function obtenerOCrearHoja_(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

function leerFilasSeguimiento_(hoja) {
  if (hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, SEG_2027.encabezados.length)
    .getValues().map(filaDesdeValores_);
}

function guardarFilaSeguimiento_(hoja, fila, posicion) {
  const valores = valoresDesdeFila_(fila);
  if (posicion >= 0) {
    hoja.getRange(posicion + 2, 1, 1, SEG_2027.encabezados.length).setValues([valores]);
  } else {
    hoja.appendRow(valores);
  }
}

function filaVacia_() {
  const fila = {};
  SEG_2027.encabezados.forEach(encabezado => fila[encabezado] = "");
  fila.Acepta_Correos = false;
  return fila;
}

function filaDesdeValores_(valores) {
  const fila = {};
  SEG_2027.encabezados.forEach((encabezado, indice) => fila[encabezado] = valores[indice]);
  fila.Acepta_Correos = interpretarBooleano_(fila.Acepta_Correos);
  return fila;
}

function valoresDesdeFila_(fila) {
  return SEG_2027.encabezados.map(encabezado => fila[encabezado] === undefined
    ? ""
    : fila[encabezado]);
}

function crearMapaEncabezados_(encabezados) {
  const mapa = {};
  encabezados.forEach((encabezado, indice) => {
    mapa[normalizarEncabezado_(encabezado)] = indice;
  });
  return mapa;
}

function indiceColumnaOpcional_(mapa, encabezado) {
  const clave = normalizarEncabezado_(encabezado);
  return mapa[clave] === undefined ? undefined : mapa[clave];
}

function normalizarEncabezado_(valor) {
  return String(valor || "").trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[_\s]+/g, " ");
}

function normalizarIdentificador_(valor) {
  return String(valor || "").trim().toLowerCase();
}

function normalizarEstado_(valor) {
  const estado = normalizarEncabezado_(valor);
  const estados = {
    "confirmado": "Confirmado",
    "confirmado por acuerdo": "Confirmado por acuerdo",
    "incompleto": "Incompleto",
    "n/a": "N/A",
    "no confirmado": "No confirmado"
  };
  return estados[estado] || "";
}

function buscarPosicionAspirante_(filas, correo, nombre, telefono) {
  const correoNormalizado = normalizarIdentificador_(correo);
  const nombreNormalizado = normalizarIdentificador_(nombre);
  const telefonoNormalizado = normalizarIdentificador_(telefono);
  const delCorreo = [];

  filas.forEach((fila, posicion) => {
    if (normalizarIdentificador_(fila.Correo) !== correoNormalizado) return;
    delCorreo.push({ fila: fila, posicion: posicion });
  });

  const exactas = delCorreo.filter(coincidencia =>
    normalizarIdentificador_(coincidencia.fila.Nombre) === nombreNormalizado &&
    normalizarIdentificador_(coincidencia.fila.Telefono) === telefonoNormalizado
  );
  if (exactas.length === 1) return exactas[0].posicion;
  if (exactas.length > 1) return -1;

  const compatibles = delCorreo.filter(coincidencia => {
    const nombreExistente = normalizarIdentificador_(coincidencia.fila.Nombre);
    const telefonoExistente = normalizarIdentificador_(coincidencia.fila.Telefono);
    return (!nombreExistente || !nombreNormalizado || nombreExistente === nombreNormalizado) &&
      (!telefonoExistente || !telefonoNormalizado || telefonoExistente === telefonoNormalizado);
  });

  return compatibles.length === 1 ? compatibles[0].posicion : -1;
}

function clavesIdentidad_(fila) {
  const claves = [];
  if (fila.Correo) {
    claves.push(
      "correo:" + normalizarIdentificador_(fila.Correo) + "|" +
      normalizarIdentificador_(fila.Nombre) + "|" +
      normalizarIdentificador_(fila.Telefono)
    );
  }
  if (fila.Folio) claves.push("folio:" + normalizarIdentificador_(fila.Folio));
  if (fila.CURP) claves.push("curp:" + normalizarIdentificador_(fila.CURP));
  return claves;
}

function agregarIndices_(indices, fila, posicion) {
  clavesIdentidad_(fila).forEach(clave => {
    if (indices[clave] === undefined) indices[clave] = posicion;
  });
}

function interpretarBooleano_(valor) {
  if (valor === true) return true;
  return ["si", "sí", "true", "1", "autorizado", "acepto"]
    .includes(normalizarEncabezado_(valor));
}

function escaparHtml_(valor) {
  return String(valor || "").replace(/[&<>"']/g, caracter => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[caracter]);
}
