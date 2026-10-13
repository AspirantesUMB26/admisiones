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

function sincronizarRegistrosFormulario() {
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
  const posicion = filas.findIndex(fila =>
    normalizarIdentificador_(fila.Correo) === buscado
  );
  if (posicion < 0) {
    SpreadsheetApp.getUi().alert("No se encontró ese correo en Seguimiento.");
    return;
  }

  filas[posicion].Acepta_Correos = false;
  hoja.getRange(posicion + 2, 1, 1, SEG_2027.encabezados.length)
    .setValues([valoresDesdeFila_(filas[posicion])]);
  SpreadsheetApp.getUi().alert("Se registró la baja. No recibirá más recordatorios automáticos.");
}

function enviarAvisosMensualesAspirantes() {
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
      if (antesDeConvocatoria) {
        fila.Ultimo_Correo_Preconvocatoria = mesActual;
      } else if (correo.estadoFinal) {
        fila.Ultimo_Estado_Notificado = correo.estadoFinal;
      } else {
        fila.Ultimo_Correo_Estatus = mesActual;
      }
      hoja.getRange(posicion + 2, 1, 1, SEG_2027.encabezados.length)
        .setValues([valoresDesdeFila_(fila)]);
      enviados++;
    } catch (error) {
      errores.push(String(fila.Correo) + ": " + error.message);
      console.error("No se pudo enviar aviso a " + fila.Correo + ": " + error.message);
    }
  });

  const resultado = "Revisión mensual: " + enviados + " correo(s) enviado(s), " +
    omitidos + " omitido(s), " + errores.length + " error(es).";
  console.log(resultado);
  if (errores.length) console.error(errores.join("\n"));
  return resultado;
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

  const enlaces =
    '<p><a href="' + SEG_2027.urlPreregistro + '">Abrir preregistro</a></p>' +
    '<p><a href="' + SEG_2027.urlSidiumb + '">Abrir SIDIUMB</a></p>';
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
  const despedida = "<p>Atentamente,<br>Universidad Mexiquense del Bicentenario</p>";
  const html = "<div style=\"font-family:Arial,sans-serif;color:#333;line-height:1.6\">" +
    "<div style=\"background:#691c32;color:#fff;padding:18px;font-weight:bold\">" +
    "Universidad Mexiquense del Bicentenario</div><div style=\"padding:20px\">" +
    encabezado + contenido + despedida + "<hr><small>" + escaparHtml_(pie) +
    "</small></div></div>";
  const texto = "Hola, " + nombre.replace(/<[^>]*>/g, "") + ":\n\n" +
    contenido.replace(/<[^>]*>/g, " ").replace(/&amp;/g, "&") +
    "\n\nAtentamente,\nUniversidad Mexiquense del Bicentenario\n\n" + pie;
  return { asunto: asunto, html: html, texto: texto };
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

  const encabezados = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0]
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

function clavesIdentidad_(fila) {
  const claves = [];
  if (fila.Correo) claves.push("correo:" + normalizarIdentificador_(fila.Correo));
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
