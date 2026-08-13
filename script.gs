// Calcula horas trabajadas entre entrada y salida
// Clasifica en diurnas, nocturnas y extras

function onFormSubmit(e) {
  const datos = e.namedValues;
  const fechaTexto = datos["Fecha "][0];
  const nombreTecnico = datos["Nombre Técnico "][0];
  const horaInicio = datos["Hora Inicio"][0] || "";
  const horaFinal = datos["Hora Final"][0] || "";
  const recibioAlimentacion = datos["Recibió Alimentación"][0];
  const actividad = datos["Actividad"][0];
  const obra = datos["Obra"][0];
  const centroCosto = datos["Centro de costo"][0];
  const equipo = datos["Equipo"][0];

  if (!nombreTecnico || !fechaTexto) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(nombreTecnico);
  if (!hoja) {
    const plantilla = ss.getSheetByName("Plantilla");
    if (!plantilla) return;
    hoja = plantilla.copyTo(ss);
    hoja.setName(nombreTecnico);
    hoja.getRange("B5").setValue(nombreTecnico);
  }

  const fecha = parseFechaFormulario(fechaTexto);
  const datosHoja = hoja.getRange("A8:O50").getValues();
  let filaActualizar = null;
  let filaLibre = null;

  for (let i = 0; i < datosHoja.length; i++) {
    const fila = datosHoja[i];
    const celdaFecha = fila[0];

    if (celdaFecha instanceof Date &&
        celdaFecha.getFullYear() === fecha.getFullYear() &&
        celdaFecha.getMonth() === fecha.getMonth() &&
        celdaFecha.getDate() === fecha.getDate() &&
        fila[3] !== "" && fila[4] === "") {
      filaActualizar = i + 8;
      break;
    }

    if (!celdaFecha && filaLibre === null) {
      filaLibre = i + 8;
    }
  }

  let filaUsar = null;

  if (filaActualizar !== null && horaFinal) {
    hoja.getRange(filaActualizar, 5).setValue(horaFinal);
    filaUsar = filaActualizar;
    if (actividad) hoja.getRange(filaUsar, 13).setValue(actividad);
    if (equipo) hoja.getRange(filaUsar, 14).setValue(equipo);
    hoja.getRange(filaUsar, 13).clearContent();
    if (recibioAlimentacion) hoja.getRange(filaUsar, 15).setValue(recibioAlimentacion);
  } else if (filaLibre !== null) {
    hoja.getRange(filaLibre, 1).setValue(fecha);
    hoja.getRange(filaLibre, 2).setValue(obra);
    hoja.getRange(filaLibre, 3).setValue(centroCosto);
    if (horaInicio) hoja.getRange(filaLibre, 4).setValue(horaInicio);
    if (horaFinal) hoja.getRange(filaLibre, 5).setValue(horaFinal);
    hoja.getRange(filaLibre, 13).setValue(actividad);
    hoja.getRange(filaLibre, 14).setValue(equipo);
    hoja.getRange(filaLibre, 15).setValue(recibioAlimentacion);
    filaUsar = filaLibre;
  }

  if (filaUsar !== null) {
    hoja.getRange(filaUsar, 15, 1, 2).clearContent();

    // Recolectar filas válidas del día
    let filasValidas = [];
    let totalHorasReales = 0;

    for (let i = 7; i < 50; i++) {
      const fila = hoja.getRange(i + 1, 1, 1, 6).getValues()[0];
      const f = fila[0];
      const entrada = fila[3];
      const salida = fila[4];

      if (f instanceof Date &&
          f.getFullYear() === fecha.getFullYear() &&
          f.getMonth() === fecha.getMonth() &&
          f.getDate() === fecha.getDate() &&
          entrada !== "" && salida !== "") {

        const hInicio = (entrada instanceof Date) ? new Date(entrada) : parseHoraFlexible(entrada);
        const hFin = (salida instanceof Date) ? new Date(salida) : parseHoraFlexible(salida);
        if (hFin < hInicio) hFin.setDate(hFin.getDate() + 1);
        const duracion = (hFin - hInicio) / (1000 * 60 * 60);

        if (duracion > 0) {
          filasValidas.push({ fila: i + 1, duracion, entradaHora: hInicio, salidaHora: hFin });
          totalHorasReales += duracion;
        }
      }
    }

    if (filasValidas.length === 0) return;

    // --- Lógica nueva por segmentos de 15 minutos ---
    const horasAjustadas = Math.min(10, totalHorasReales * 1.25);
    const ordRealTotalToAssign = horasAjustadas / 1.25;
    const ordLimitSegments = Math.round(ordRealTotalToAssign * 4);

    const ordSegs = {};
    const hedSegs = {};
    const henSegs = {};
    const rnSegs = {};
    for (let f of filasValidas) {
      ordSegs[f.fila] = 0;
      hedSegs[f.fila] = 0;
      henSegs[f.fila] = 0;
      rnSegs[f.fila] = 0;
    }

    let segments = [];
    let totalDiurnaSegsDay = 0;
    for (let f of filasValidas) {
      const start = new Date(f.entradaHora);
      const end = new Date(f.salidaHora);
      let t = new Date(start);
      if (end < t) end.setDate(end.getDate() + 1);
      while (t < end) {
        const hh = t.getHours() + t.getMinutes() / 60;
        const isDiurna = (hh >= 6 && hh < 21);
        if (isDiurna) totalDiurnaSegsDay++;
        segments.push({ time: new Date(t.getTime()), fila: f.fila, isDiurna });
        t.setMinutes(t.getMinutes() + 15);
      }
    }

    segments.sort((a, b) => a.time - b.time);

    const dayDiurnaHoras = totalDiurnaSegsDay * 0.25;
    const dayTiene8RealesDiurnas = dayDiurnaHoras >= 8 - 1e-9;

    let ordAssignedSegs = 0;
    for (let seg of segments) {
      if (ordAssignedSegs < ordLimitSegments) {
        ordAssignedSegs++;
        ordSegs[seg.fila] += 1;
        if (!seg.isDiurna) {
          rnSegs[seg.fila] += 1;
        }
      } else {
        if (seg.isDiurna) {
          hedSegs[seg.fila] += 1;
        } else {
          if (dayTiene8RealesDiurnas) henSegs[seg.fila] += 1;
          else rnSegs[seg.fila] += 1;
        }
      }
    }

    // === NUEVA LÓGICA: Descuento alimentación ===
    const valorAlimentacion = hoja.getRange(filaUsar, 14).getValue();
    if (valorAlimentacion && valorAlimentacion.toString().trim().toUpperCase() === "SI") {
      // Restar 0,50h de las horas extra diurnas o nocturnas
      if (hedSegs[filaUsar] > 2) {
        hedSegs[filaUsar] -= 2;
      } else if (henSegs[filaUsar] > 2) {
        henSegs[filaUsar] -= 2;
      }
      hoja.getRange(filaUsar, 14).setValue("SI - Descuento 0,50h");
    }

    // Escribir resultados
    for (let f of filasValidas) {
      const ordHorasReales = (ordSegs[f.fila] || 0) * 0.25;
      const ordAjustado = ordHorasReales * 1.25;
      const rnHoras = (rnSegs[f.fila] || 0) * 0.25;
      const hedHoras = (hedSegs[f.fila] || 0) * 0.25;
      const henHoras = (henSegs[f.fila] || 0) * 0.25;

     hoja.getRange(f.fila, 6).setValue(ordAjustado > 0 ? ordAjustado : "").setNumberFormat("0.00");
hoja.getRange(f.fila, 7).setValue(rnHoras > 0 ? rnHoras : "").setNumberFormat("0.00");
hoja.getRange(f.fila, 9).setValue(hedHoras > 0 ? hedHoras : "").setNumberFormat("0.00");
hoja.getRange(f.fila, 11).setValue(henHoras > 0 ? henHoras : "").setNumberFormat("0.00");

    }

    hoja.getRange("F50").setFormula("=SUM(F8:F49)").setNumberFormat("0.00");
    hoja.getRange("G50").setFormula("=SUM(G8:G49)").setNumberFormat("0.00");
    hoja.getRange("H50").setFormula("=SUM(H8:H49)").setNumberFormat("0.00");
    hoja.getRange("I50").setFormula("=SUM(I8:I49)").setNumberFormat("0.00");

  }
}


function onEdit(e) {
  const hoja = e.range.getSheet();
  const fila = e.range.getRow();
  const col = e.range.getColumn();

  // Solo aplica si se edita en la zona de registros (D o E = 4 o 5)
  if (fila >= 8 && (col === 4 || col === 5)) {
    const fecha = hoja.getRange(fila, 1).getValue(); // Col A = fecha
    if (!(fecha instanceof Date)) return;

    // Recalcular automáticamente los tiempos de ese día
    recalcularDia(hoja, fecha);
  }
}

function recalcularDia(hoja, fecha) {
  // Recoge todas las filas de ese día
  const ultimaFila = hoja.getLastRow();
  const datos = hoja.getRange(8, 1, ultimaFila - 7, 14).getValues();

  let filasValidas = [];
  let totalHorasReales = 0;

  datos.forEach((fila, i) => {
    const f = fila[0];
    const entrada = fila[3];
    const salida = fila[4];

    if (f instanceof Date &&
        f.getFullYear() === fecha.getFullYear() &&
        f.getMonth() === fecha.getMonth() &&
        f.getDate() === fecha.getDate() &&
        entrada && salida) {

      const hInicio = (entrada instanceof Date) ? new Date(entrada) : parseHoraFlexible(entrada);
      const hFin = (salida instanceof Date) ? new Date(salida) : parseHoraFlexible(salida);
      if (hFin < hInicio) hFin.setDate(hFin.getDate() + 1);

      const duracion = (hFin - hInicio) / (1000 * 60 * 60);
      if (duracion > 0) {
        filasValidas.push({ fila: i + 8, duracion, entradaHora: hInicio, salidaHora: hFin });
        totalHorasReales += duracion;
      }
    }
  });

  if (filasValidas.length === 0) return;

  // ---- Lógica de cálculo (copiada de tu onFormSubmit) ----
  const horasAjustadas = Math.min(10, totalHorasReales * 1.25);
  const ordRealTotalToAssign = horasAjustadas / 1.25;
  const ordLimitSegments = Math.round(ordRealTotalToAssign * 4);

  const ordSegs = {}, hedSegs = {}, henSegs = {}, rnSegs = {};
  for (let f of filasValidas) {
    ordSegs[f.fila] = 0;
    hedSegs[f.fila] = 0;
    henSegs[f.fila] = 0;
    rnSegs[f.fila] = 0;
  }

  let segments = [];
  let totalDiurnaSegsDay = 0;
  for (let f of filasValidas) {
    const start = new Date(f.entradaHora);
    const end = new Date(f.salidaHora);
    let t = new Date(start);
    if (end < t) end.setDate(end.getDate() + 1);
    while (t < end) {
      const hh = t.getHours() + t.getMinutes() / 60;
      const isDiurna = (hh >= 6 && hh < 21);
      if (isDiurna) totalDiurnaSegsDay++;
      segments.push({ time: new Date(t.getTime()), fila: f.fila, isDiurna });
      t.setMinutes(t.getMinutes() + 15);
    }
  }
  segments.sort((a, b) => a.time - b.time);

  const dayDiurnaHoras = totalDiurnaSegsDay * 0.25;
  const dayTiene8RealesDiurnas = dayDiurnaHoras >= 8 - 1e-9;

  let ordAssignedSegs = 0;
  for (let seg of segments) {
    if (ordAssignedSegs < ordLimitSegments) {
      ordAssignedSegs++;
      ordSegs[seg.fila]++;
      if (!seg.isDiurna) rnSegs[seg.fila]++;
    } else {
      if (seg.isDiurna) {
        hedSegs[seg.fila]++;
      } else {
        if (dayTiene8RealesDiurnas) henSegs[seg.fila]++;
        else rnSegs[seg.fila]++;
      }
    }
  }

  // Guardar resultados
  for (let f of filasValidas) {
    const ordHorasReales = (ordSegs[f.fila] || 0) * 0.25;
    const ordAjustado = ordHorasReales * 1.25;
    const rnHoras = (rnSegs[f.fila] || 0) * 0.25;
    const hedHoras = (hedSegs[f.fila] || 0) * 0.25;
    const henHoras = (henSegs[f.fila] || 0) * 0.25;

    hoja.getRange(f.fila, 6).setValue(ordAjustado || "").setNumberFormat("0.00");
    hoja.getRange(f.fila, 7).setValue(rnHoras || "").setNumberFormat("0.00");
    hoja.getRange(f.fila, 9).setValue(hedHoras || "").setNumberFormat("0.00");
    hoja.getRange(f.fila, 11).setValue(henHoras || "").setNumberFormat("0.00");
  }

  hoja.getRange("F50").setFormula("=SUM(F8:F49)").setNumberFormat("0.00");
  hoja.getRange("G50").setFormula("=SUM(G8:G49)").setNumberFormat("0.00");
  hoja.getRange("H50").setFormula("=SUM(H8:H49)").setNumberFormat("0.00");
  hoja.getRange("I50").setFormula("=SUM(I8:I49)").setNumberFormat("0.00");
}


////////////////////


function parseFechaFormulario(texto) {
  const partes = texto.split("/");
  return new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
}

function parseHoraFlexible(valor) {
  if (valor instanceof Date) return valor;
  const partes = valor.split(":");
  const ahora = new Date();
  ahora.setHours(Number(partes[0]));
  ahora.setMinutes(Number(partes[1]));
  ahora.setSeconds(0);
  ahora.setMilliseconds(0);
  return ahora;
}



function pruebaManual() {
  const e = {
    namedValues: {
      "Fecha ": ["12/08/2025"],
      "Nombre Técnico ": ["Adolfo Barbosa"],
      "Hora Inicio": ["15:00"],
      "Hora Final": ["22:00"],
      "Recibió Alimentación": ["Sí"],
      "Actividad": ["Revision"],
      "Obra": ["galias"],
      "Centro de costo": ["mp422"],
      "Equipo": ["Extractor"]
    }
  };
  onFormSubmit(e);
}

function respaldoSemanal() {
  const carpetaID = '1pn6ioGOo-FGC2wAPuAqu_nL7leMVRqt9';
  const carpeta = DriveApp.getFolderById(carpetaID);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const nombreArchivo = obtenerNombreSemanaAnterior();
  DriveApp.getFileById(spreadsheet.getId()).makeCopy(nombreArchivo, carpeta);
}

function obtenerNombreSemanaAnterior() {
  const hoy = new Date();
  const lunesAnterior = new Date(hoy);
  lunesAnterior.setDate(lunesAnterior.getDate() - 7);
  lunesAnterior.setDate(lunesAnterior.getDate() - ((lunesAnterior.getDay() + 6) % 7));

  const domingoAnterior = new Date(lunesAnterior);
  domingoAnterior.setDate(lunesAnterior.getDate() + 6);

  const opciones = { day: 'numeric', month: 'long', year: 'numeric' };
  const formato = new Intl.DateTimeFormat('es-CO', opciones);

  return `Semana - ${formato.format(lunesAnterior)} al ${formato.format(domingoAnterior)}`;
}

function limpiarPlantillas() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = spreadsheet.getSheets();
  const hojasPermitidas = ["Plantilla", "Respuestas Proyectos","TECNICOS"];

  hojas.forEach(hoja => {
    if (!hojasPermitidas.includes(hoja.getName())) {
      spreadsheet.deleteSheet(hoja);
    }
  });

  Logger.log("🧹 Hojas limpias: solo quedan 'Plantilla' , 'Respuestas' y 'TECNICOS'.");
}



function mismaFecha(f1, f2) {
  return (
    f1.getFullYear() === f2.getFullYear() &&
    f1.getMonth() === f2.getMonth() &&
    f1.getDate() === f2.getDate()
  );
}



function recordarSalidasPendientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ahora = new Date();

  const hojaTecnicos = ss.getSheetByName("TECNICOS");
  if (!hojaTecnicos) {
    Logger.log("Error: La hoja 'TECNICOS' no se encontró.");
    return;
  }
  const datosTecnicos = hojaTecnicos.getRange(2, 1, hojaTecnicos.getLastRow() - 1, 2).getValues();
  const tecnicosMap = new Map(datosTecnicos);

  const hojaLogSalidas = ss.getSheetByName("LOG SALIDAS") || ss.insertSheet("LOG SALIDAS");
  if (hojaLogSalidas.getLastRow() === 0) {
    hojaLogSalidas.appendRow(["Fecha", "Nombre Técnico", "Correo", "Hora Envío", "Obra Notificada"]);
  }

  const todasHojas = ss.getSheets();
  todasHojas.forEach(hoja => {
    const nombreHoja = hoja.getName();
    if (["TECNICOS", "PLANTILLA", "LOG", "CONFIG", "LOG ENTRADAS", "LOG SALIDAS", "Respuestas"].includes(nombreHoja.toUpperCase())) {
      return;
    }

    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 8) return;

    const datosEntradaSalida = hoja.getRange(8, 1, ultimaFila - 7, 5).getValues();

    datosEntradaSalida.forEach((fila, i) => {
      const fechaRegistro = fila[0];
      const obra = fila[1];
      const entrada = fila[3];
      const salida = fila[4];

      if (fechaRegistro instanceof Date &&
          fechaRegistro.getFullYear() === ahora.getFullYear() &&
          fechaRegistro.getMonth() === ahora.getMonth() &&
          fechaRegistro.getDate() === ahora.getDate()) {

        if (entrada && !salida) {
          const hEntrada = (entrada instanceof Date) ? entrada : parseHoraFlexible(entrada);
          const horasPasadas = (ahora - hEntrada) / (1000 * 60 * 60);

          const nombreTecnico = nombreHoja;
          const correoTecnico = tecnicosMap.get(nombreTecnico);

          if (horasPasadas >= 4 && correoTecnico && !yaSeNotificoHoy(hojaLogSalidas, nombreTecnico, ahora, obra)) {
            const horaEntradaFormato = Utilities.formatDate(hEntrada, "GMT-5", "HH:mm");

            // --- NUEVO MENSAJE ---
            const mensaje = `¡Hola ${nombreTecnico}! 👋\n\nDetectamos que has registrado una entrada pero no has registrado la salida⏰.\n\n👷‍♀️ OBRA: ${obra}\n⏱️ HORA ENTRADA: ${horaEntradaFormato}\n\n¡Por favor, Agrega la salida antes de ingresar la próxima entrada! 📝\n\n¡¡Ponte al día!! 💪`;
            // --- FIN NUEVO MENSAJE ---

            try {
              MailApp.sendEmail({
                to: correoTecnico,
                subject: `🔔 Recordatorio: Salida pendiente en obra ${obra}`,
                body: mensaje
              });

              hojaLogSalidas.appendRow([
                Utilities.formatDate(ahora, "GMT-5", "yyyy-MM-dd"),
                nombreTecnico,
                correoTecnico,
                Utilities.formatDate(new Date(), "GMT-5", "HH:mm:ss"),
                obra
              ]);
              Logger.log(`✉️ Recordatorio de salida enviado a ${nombreTecnico} (${correoTecnico}) para obra ${obra}`);

            } catch (e) {
              Logger.log(`Error al enviar correo a ${nombreTecnico} para obra ${obra}: ${e.toString()}`);
            }
          }
        }
      }
    });
  });
}

function yaSeNotificoHoy(hojaLog, nombre, fechaHoy, obraNotificada) {
  const ultimaFila = hojaLog.getLastRow();
  if (ultimaFila < 2) return false;

  const registros = hojaLog.getRange(2, 1, ultimaFila - 1, hojaLog.getLastColumn()).getValues();
  const fechaTexto = Utilities.formatDate(fechaHoy, "GMT-5", "yyyy-MM-dd");

  return registros.some(([fecha, n, , , obraLog]) =>
    fecha === fechaTexto && n === nombre && obraLog === obraNotificada
  );
}


function verificarEntradasFaltantes() {
  const ahora = new Date();
  const dia = ahora.getDay(); // 0 = Domingo, 1 = Lunes, ..., 6 = Sábado
  const hora = ahora.getHours();
  const minutos = ahora.getMinutes();

  // Definir el horario laboral para enviar recordatorios
  let dentroDelHorario = false;

  // Lunes a Viernes: de 7:00 a 16:59 (antes de las 18:00)
  if (dia >= 1 && dia <= 5) {
    dentroDelHorario = (hora >= 7 && hora < 18);
  }
  // Sábado: de 7:30 a 12:59 (antes de las 13:00)
  else if (dia === 6) {
    if (hora === 7) {
      dentroDelHorario = (minutos >= 30);
    } else {
      dentroDelHorario = (hora >= 8 && hora < 14);
    }
  }

  // Si no estamos en horario laboral, no se envían recordatorios
  if (!dentroDelHorario) {
    Logger.log("Fuera de horario laboral para verificar entradas pendientes.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTecnicos = ss.getSheetByName("TECNICOS");
  if (!hojaTecnicos) {
    Logger.log("Error: La hoja 'TECNICOS' no se encontró.");
    return;
  }
  // Suponemos que la columna 1 es el Nombre y la columna 2 es el Correo
  const datosTecnicos = hojaTecnicos.getRange(2, 1, hojaTecnicos.getLastRow() - 1, 2).getValues();

  const hojaLogEntradas = ss.getSheetByName("LOG ENTRADAS") || ss.insertSheet("LOG ENTRADAS");
  if (hojaLogEntradas.getLastRow() === 0) {
    hojaLogEntradas.appendRow(["Fecha", "Nombre", "Correo", "Hora Envío", "Hora Última Salida"]); // Añadir columna para la última salida
  }

  for (let [nombre, correo] of datosTecnicos) {
    const hojaTecnico = ss.getSheetByName(nombre);
    if (!hojaTecnico) continue; // Si no hay hoja para el técnico, saltar

    // Revisar desde los registros más recientes (fila 8 en adelante)
    const registros = hojaTecnico.getRange("A8:E" + hojaTecnico.getLastRow()).getValues().reverse();
    let ultimaSalida = null;
    let hayEntradaDespues = false;

    for (let fila of registros) {
      const fecha = fila[0]; // Columna A
      const entrada = fila[3]; // Columna D
      const salida = fila[4]; // Columna E

      // Solo consideramos registros del día actual
      if (fecha instanceof Date &&
          fecha.getFullYear() === ahora.getFullYear() &&
          fecha.getMonth() === ahora.getMonth() &&
          fecha.getDate() === ahora.getDate()) {

        if (salida instanceof Date && !ultimaSalida) {
          // Encontramos la última salida del día
          ultimaSalida = salida;
        } else if (entrada instanceof Date && ultimaSalida && entrada > ultimaSalida) {
          // Si hay una entrada después de la última salida, significa que ya reingresó
          hayEntradaDespues = true;
          break; // Salir del bucle, ya no es necesario notificar
        }
      }
    }

    // Si hay una última salida y no hay una entrada posterior a esa salida
    if (ultimaSalida && !hayEntradaDespues) {
      const minutosDesdeSalida = (ahora - ultimaSalida) / 60000;
      const UMBRAL_MINUTOS = 10; // 👉 UMBERAL DE 10 MINUTOS

      // Si han pasado más de 10 minutos desde la última salida y no se ha notificado hoy para este técnico
      if (minutosDesdeSalida >= UMBRAL_MINUTOS && !yaSeNotificoHoyEntradas(hojaLogEntradas, nombre, ahora, ultimaSalida)) {
        const horaSalidaFormato = Utilities.formatDate(ultimaSalida, "GMT-5", "HH:mm");

        // --- NUEVO MENSAJE ---
        const mensaje = `¡Hola ${nombre}! 👋\n\nHan pasado más de ${UMBRAL_MINUTOS} minutos desde tu última salida registrada a las ${horaSalidaFormato} ⏰, pero no hemos recibido un nuevo registro de entrada.\n\n¡Por favor,recuerda registrar tu entrada al comenzar tu siguiente jornada! 📝\n\n¡Gracias por tu compromiso! 💪`;
        // --- FIN NUEVO MENSAJE ---

        try {
          MailApp.sendEmail({
            to: correo,
            subject: "🔔 Registro de entrada pendiente",
            body: mensaje
          });

          // Registrar que se envió la notificación en el LOG ENTRADAS
          hojaLogEntradas.appendRow([
            Utilities.formatDate(ahora, "GMT-5", "yyyy-MM-dd"),
            nombre,
            correo,
            Utilities.formatDate(new Date(), "GMT-5", "HH:mm:ss"),
            horaSalidaFormato // Registrar la hora de la última salida
          ]);
          Logger.log(`✉️ Recordatorio de entrada enviado a ${nombre} (${correo}) por salida a las ${horaSalidaFormato}`);

        } catch (e) {
          Logger.log(`Error al enviar correo de entrada a ${nombre}: ${e.toString()}`);
        }
      }
    }
  }
}


// Función auxiliar para verificar si ya se notificó hoy para entradas (modificada)
// Se agregó 'horaUltimaSalidaNotificada' para evitar notificaciones repetidas si el técnico no registra
function yaSeNotificoHoyEntradas(hojaLog, nombre, fechaHoy, horaUltimaSalidaNotificada) {
  const ultimaFila = hojaLog.getLastRow();
  if (ultimaFila < 2) return false;

  const registros = hojaLog.getRange(2, 1, ultimaFila - 1, hojaLog.getLastColumn()).getValues();
  const fechaTexto = Utilities.formatDate(fechaHoy, "GMT-5", "yyyy-MM-dd");
  const horaSalidaTexto = Utilities.formatDate(horaUltimaSalidaNotificada, "GMT-5", "HH:mm");

  // Verificar si ya se envió una notificación para este técnico, en esta fecha Y para esta misma última salida
  return registros.some(([fechaLog, nLog, , , ultimaSalidaLog]) =>
    fechaLog === fechaTexto && nLog === nombre && ultimaSalidaLog === horaSalidaTexto
  );
}






