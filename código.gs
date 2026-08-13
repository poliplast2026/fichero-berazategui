// ============================================================
// FICHERO BERAZATEGUI — control de entrada/salida y almuerzo
// Backend en Google Apps Script + Google Sheets
// PIN + verificación facial (descriptor, no se guardan fotos)
// ============================================================

const SHEET_USUARIOS = 'Usuarios';
const SHEET_FICHAJES = 'Fichajes';
const SHEET_RESUMEN = 'Resumen diario';

const USUARIOS_HEADERS = ['ID', 'Nombre completo', 'Token dispositivo', 'PIN (protegido)', 'Rostro (descriptor)', 'Consentimiento biométrico', 'Fecha registro'];
const FICHAJES_HEADERS = ['Fecha', 'Hora', 'ID Usuario', 'Nombre', 'Tipo', 'Timestamp', 'Depósito más cercano', 'Distancia (m)', 'Precisión GPS (m)'];

// ---- Ubicación de los depósitos (para el control de "solo fichar desde ahí") ----
// Mientras MODO_PRUEBA_UBICACION esté en true, el sistema NUNCA rechaza un fichaje por
// ubicación — solo la registra en la planilla para poder revisar qué tan precisa es el
// GPS de la gente en la práctica. Cuando Felipe confirme que está todo OK, cambiar esto
// a false: ahí sí va a rechazar fichajes fuera del radio de tolerancia.
const MODO_PRUEBA_UBICACION = true;
const RADIO_TOLERANCIA_METROS = 150;

const DEPOSITOS = [
  { nombre: 'Mini Parque Vergara (Berazategui)', lat: -34.782175430731726, lon: -58.22225802328305 },
  { nombre: 'Magdalena 962 (Villa Domínico)', lat: -34.7018584, lon: -58.3372409 }
];

// Paleta de colores pastel, uno por usuario (se asigna en orden de registro).
const PALETA_COLORES = [
  '#DCEEFB', '#FBE7DC', '#E4DCFB', '#DCFBE4', '#FBDCEE',
  '#FBF6DC', '#DCFBF6', '#F6DCFB', '#E8E8E8', '#FBEDDC'
];

const COLOR_ENTRADA = '#15803d'; // verde
const COLOR_SALIDA = '#b91c1c';  // rojo

// Umbral de similitud facial (distancia euclidiana entre descriptores).
// Más bajo = más estricto (más rechazos). Más alto = más laxo (más falsos positivos).
// 0.6 es el valor por defecto recomendado por face-api.js; se puede ajustar acá si en la
// práctica rechaza gente que sí es, o deja pasar gente que no es.
const UMBRAL_ROSTRO = 0.6;

// ---- API para la página externa (la que usa la cámara) ----
// La página real (con el reconocimiento facial) va a vivir afuera de Apps Script,
// porque Apps Script bloquea el acceso a la cámara. Esta "puerta de entrada" deja que
// esa página llame a las mismas funciones de siempre (registrarUsuario, marcarFichaje, etc.)
// Se usa GET (no POST) porque el redirect interno de Apps Script convierte los POST en GET
// y se pierden los datos — con GET todo viaja en la URL y no hay ese problema.
const ACCIONES_PERMITIDAS = {
  registrarUsuario: registrarUsuario,
  getUsuarioPorToken: getUsuarioPorToken,
  getFichajesDeHoy: getFichajesDeHoy,
  marcarFichaje: marcarFichaje,
  vincularDispositivo: vincularDispositivo,
  recuperarPin: recuperarPin
};

function manejarApi_(action, paramsJson) {
  const funcion = ACCIONES_PERMITIDAS[action];
  if (!funcion) {
    return jsonOutput_({ ok: false, error: 'Acción desconocida: ' + action });
  }
  let params = [];
  if (paramsJson) {
    try {
      params = JSON.parse(paramsJson);
    } catch (err) {
      return jsonOutput_({ ok: false, error: 'Parámetros inválidos.' });
    }
  }
  try {
    const data = funcion.apply(null, params);
    return jsonOutput_({ ok: true, data: data });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---- Rutas de la web app ----
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'app';
  const action = e && e.parameter && e.parameter.action;

  if (action) {
    return manejarApi_(action, e.parameter.params);
  }

  // Endpoint de prueba: responde datos (no una página) para verificar si se puede
  // llamar a este script desde una web externa (necesario si movemos la cámara afuera).
  if (page === 'ping') {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, hora: new Date().toISOString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const tmpl = page === 'qr'
    ? HtmlService.createTemplateFromFile('QR')
    : HtmlService.createTemplateFromFile('Index');
  tmpl.appUrl = ScriptApp.getService().getUrl();
  return tmpl.evaluate()
    .setTitle('Fichero Berazategui')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---- Utilidad interna ----
function getOrCreateSheet_(name, headers, onCreate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    if (onCreate) onCreate(sheet);
  } else {
    asegurarColumnas_(sheet, headers);
  }
  return sheet;
}

// Si la lista de columnas esperadas creció (agregamos una función nueva), esto agrega
// las columnas que falten al final de la hoja ya existente, sin tocar las que ya había.
// Así no hace falta borrar y recrear pestañas cada vez que el sistema crece.
function asegurarColumnas_(sheet, headersEsperados) {
  const anchoActual = Math.max(sheet.getLastColumn(), 1);
  const headersActuales = sheet.getRange(1, 1, 1, anchoActual).getValues()[0];
  headersEsperados.forEach(function (h) {
    if (headersActuales.indexOf(h) === -1) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(h).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
      headersActuales.push(h);
    }
  });
}

// Google Sheets convierte solo los textos con forma de fecha/hora en
// objetos Fecha al guardarlos. Estas dos funciones normalizan lo que
// venga de una celda (texto u objeto Fecha) a un string comparable.
function normalizarFecha_(valor, tz) {
  if (valor instanceof Date) return Utilities.formatDate(valor, tz, 'yyyy-MM-dd');
  return String(valor);
}

function normalizarHora_(valor, tz) {
  if (valor instanceof Date) return Utilities.formatDate(valor, tz, 'HH:mm:ss');
  return String(valor);
}

// Color fijo y determinístico para cada usuario, a partir del número de su ID (U001, U002...).
function colorParaUsuario_(id) {
  const num = parseInt(String(id).replace(/[^0-9]/g, ''), 10) || 0;
  return PALETA_COLORES[num % PALETA_COLORES.length];
}

function esTipoEntrada_(tipo) {
  return tipo === 'Entrada' || tipo === 'Vuelta almuerzo';
}

// Hash del PIN (nunca se guarda el PIN en texto plano).
function hashPin_(pin) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin) + '|fichero-berazategui');
  return digest.map(function (b) {
    const v = (b + 256) % 256;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

// Distancia en metros entre dos coordenadas GPS (fórmula de Haversine).
function distanciaMetros_(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad;
  const dLon = (lon2 - lon1) * rad;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * rad) * Math.cos(lat2 * rad) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

// Busca cuál de los depósitos configurados está más cerca de una coordenada dada.
function depositoMasCercano_(lat, lon) {
  if (typeof lat !== 'number' || typeof lon !== 'number' || !DEPOSITOS.length) return null;
  let mejor = null;
  DEPOSITOS.forEach(function (d) {
    const dist = distanciaMetros_(lat, lon, d.lat, d.lon);
    if (!mejor || dist < mejor.distancia) mejor = { nombre: d.nombre, distancia: dist };
  });
  return mejor;
}

// Distancia euclidiana entre dos descriptores faciales (arrays de 128 números).
function distanciaEuclidiana_(a, b) {
  let suma = 0;
  for (let i = 0; i < a.length; i++) {
    const diff = a[i] - b[i];
    suma += diff * diff;
  }
  return Math.sqrt(suma);
}

function mensajeSaludo_(tipo, nombre) {
  const primerNombre = String(nombre).trim().split(' ')[0];
  if (tipo === 'Entrada') return '¡Hola, ' + primerNombre + '! Bienvenido.';
  if (tipo === 'Salida almuerzo') return 'Buen provecho, ' + primerNombre + '.';
  if (tipo === 'Vuelta almuerzo') return 'Que te sea leve, ' + primerNombre + '.';
  if (tipo === 'Salida') return 'Hasta la próxima, ' + primerNombre + '.';
  return 'Listo, ' + primerNombre + '.';
}

// Pinta una fila de Fichajes recién agregada: fondo = color del usuario,
// columna Tipo en verde (entrada) o rojo (salida) y en negrita.
function formatearFilaFichaje_(sheet, fila, id, tipo) {
  const bg = colorParaUsuario_(id);
  sheet.getRange(fila, 1, 1, FICHAJES_HEADERS.length).setBackground(bg);
  sheet.getRange(fila, 5)
    .setFontColor(esTipoEntrada_(tipo) ? COLOR_ENTRADA : COLOR_SALIDA)
    .setFontWeight('bold');
}

// Busca, entre todos los usuarios ya registrados, el que más se parece a un descriptor.
// Se usa para evitar registros duplicados y para "vincular este celular" a un usuario existente.
function buscarUsuarioPorRostro_(descriptorRostro) {
  if (!Array.isArray(descriptorRostro)) return null;
  const sheet = getOrCreateSheet_(SHEET_USUARIOS, USUARIOS_HEADERS);
  const data = sheet.getDataRange().getValues();
  let mejor = null;
  for (let i = 1; i < data.length; i++) {
    let descriptor = null;
    try { descriptor = JSON.parse(data[i][4]); } catch (e) { descriptor = null; }
    if (!descriptor || descriptor.length !== descriptorRostro.length) continue;
    const distancia = distanciaEuclidiana_(descriptor, descriptorRostro);
    if (distancia <= UMBRAL_ROSTRO && (!mejor || distancia < mejor.distancia)) {
      mejor = { fila: i + 1, id: String(data[i][0]), nombre: data[i][1], token: data[i][2], pinHash: data[i][3], descriptorRostro: descriptor, distancia: distancia };
    }
  }
  return mejor;
}

// ---- Registro de usuario (una vez por celular) ----
function registrarUsuario(nombreCompleto, pin, descriptorRostro, consentimiento) {
  if (!/^\d{4}$/.test(String(pin))) {
    throw new Error('El PIN tiene que ser de 4 números.');
  }
  if (!consentimiento) {
    throw new Error('Hace falta aceptar el consentimiento para procesar el rostro.');
  }
  if (!Array.isArray(descriptorRostro) || descriptorRostro.length < 64) {
    throw new Error('No se pudo capturar el rostro. Probá de nuevo con buena luz.');
  }

  const existente = buscarUsuarioPorRostro_(descriptorRostro);
  if (existente) {
    throw new Error('Ya existe un registro con este rostro: ' + existente.nombre + ' (' + existente.id + '). Si sos vos, usá "Ya me registré antes, vincular este celular" en vez de crear un usuario nuevo.');
  }

  const sheet = getOrCreateSheet_(SHEET_USUARIOS, USUARIOS_HEADERS);
  const correlativo = sheet.getLastRow(); // fila 1 = encabezado
  const id = 'U' + String(correlativo).padStart(3, '0');
  const token = Utilities.getUuid();
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([id, nombreCompleto.trim(), token, hashPin_(pin), JSON.stringify(descriptorRostro), 'Sí (' + fecha + ')', fecha]);
  sheet.getRange(sheet.getLastRow(), 1, 1, USUARIOS_HEADERS.length).setBackground(colorParaUsuario_(id));
  return { id: id, token: token, nombre: nombreCompleto.trim() };
}

function getUsuarioPorToken(token) {
  const usuario = buscarUsuarioPorToken_(token);
  if (!usuario) return null;
  return { id: usuario.id, nombre: usuario.nombre };
}

function buscarUsuarioPorToken_(token) {
  const sheet = getOrCreateSheet_(SHEET_USUARIOS, USUARIOS_HEADERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === token) {
      let descriptor = null;
      try { descriptor = JSON.parse(data[i][4]); } catch (e) { descriptor = null; }
      return { fila: i + 1, id: String(data[i][0]), nombre: data[i][1], pinHash: data[i][3], descriptorRostro: descriptor };
    }
  }
  return null;
}

// Vincula este celular (un token nuevo, sin usuario asociado todavía) a un usuario ya
// existente, identificándolo por su rostro. Útil cuando alguien cambia de teléfono.
function vincularDispositivo(descriptorRostro) {
  if (!Array.isArray(descriptorRostro) || descriptorRostro.length < 64) {
    throw new Error('No se pudo capturar el rostro. Probá de nuevo con buena luz.');
  }
  const match = buscarUsuarioPorRostro_(descriptorRostro);
  if (!match) {
    throw new Error('No encontramos ese rostro registrado. Si es tu primera vez, registrate como usuario nuevo.');
  }
  return { id: match.id, token: match.token, nombre: match.nombre };
}

// Permite a la propia persona cambiar su PIN si se lo olvidó, verificando su rostro
// (tiene que ser desde el mismo celular donde ya está registrada).
function recuperarPin(token, descriptorRostro, nuevoPin) {
  if (!/^\d{4}$/.test(String(nuevoPin))) {
    throw new Error('El PIN tiene que ser de 4 números.');
  }
  const usuario = buscarUsuarioPorToken_(token);
  if (!usuario) throw new Error('Usuario no encontrado en este celular.');
  if (!usuario.descriptorRostro) throw new Error('Este usuario no tiene un rostro registrado.');
  if (!Array.isArray(descriptorRostro) || descriptorRostro.length !== usuario.descriptorRostro.length) {
    throw new Error('No se pudo verificar el rostro. Probá de nuevo.');
  }
  const distancia = distanciaEuclidiana_(usuario.descriptorRostro, descriptorRostro);
  if (distancia > UMBRAL_ROSTRO) {
    throw new Error('El rostro no coincide con el registrado. No se puede cambiar el PIN así, pedile al administrador que te lo reinicie.');
  }
  const sheet = getOrCreateSheet_(SHEET_USUARIOS, USUARIOS_HEADERS);
  sheet.getRange(usuario.fila, 4).setValue(hashPin_(nuevoPin));
  return { ok: true };
}

// ---- Fichajes del día ----
function getFichajesDeHoy(token) {
  const usuario = getUsuarioPorToken(token);
  if (!usuario) return [];
  const tz = Session.getScriptTimeZone();
  const sheet = getOrCreateSheet_(SHEET_FICHAJES, FICHAJES_HEADERS);
  const data = sheet.getDataRange().getValues();
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const tipos = [];
  for (let i = 1; i < data.length; i++) {
    const filaFecha = normalizarFecha_(data[i][0], tz);
    if (filaFecha === hoy && String(data[i][2]) === usuario.id) {
      tipos.push(data[i][4]);
    }
  }
  return tipos;
}

function marcarFichaje(token, tipo, pin, descriptorRostro, lat, lon, precisionGps) {
  const usuario = buscarUsuarioPorToken_(token);
  if (!usuario) throw new Error('Usuario no encontrado. Registrate de nuevo.');
  if (hashPin_(pin) !== usuario.pinHash) throw new Error('PIN incorrecto.');
  if (!usuario.descriptorRostro) throw new Error('No hay un rostro registrado para este usuario. Volvé a registrarte.');
  if (!Array.isArray(descriptorRostro) || descriptorRostro.length !== usuario.descriptorRostro.length) {
    throw new Error('No se pudo verificar el rostro. Probá de nuevo.');
  }
  const distancia = distanciaEuclidiana_(usuario.descriptorRostro, descriptorRostro);
  if (distancia > UMBRAL_ROSTRO) {
    throw new Error('El rostro no coincide con el registrado.');
  }

  const cercano = depositoMasCercano_(lat, lon);
  const nombreDeposito = cercano ? cercano.nombre : 'Sin datos de ubicación';
  const distanciaDeposito = cercano ? Math.round(cercano.distancia) : '';

  // El bloqueo por ubicación solo se activa cuando MODO_PRUEBA_UBICACION pase a false.
  // Hasta entonces, un fichaje lejos del depósito se guarda igual (para poder revisarlo).
  if (!MODO_PRUEBA_UBICACION && cercano && distanciaDeposito > RADIO_TOLERANCIA_METROS) {
    throw new Error('Estás fichando muy lejos del depósito (' + distanciaDeposito + ' m de ' + nombreDeposito + '). Si creés que es un error, avisá al administrador.');
  }

  const tz = Session.getScriptTimeZone();
  const sheet = getOrCreateSheet_(SHEET_FICHAJES, FICHAJES_HEADERS, function (s) {
    s.getRange(1, 5).setNote('Verde = Entrada / Vuelta de almorzar. Rojo = Salida / Salida a almorzar. El color de fondo de cada fila identifica a la persona (ver planilla "Usuarios").');
  });
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Traba del lado del servidor: no permite repetir el mismo tipo el mismo día,
  // aunque el botón del celular se toque varias veces seguidas.
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const filaFecha = normalizarFecha_(data[i][0], tz);
    if (filaFecha === hoy && String(data[i][2]) === usuario.id && data[i][4] === tipo) {
      throw new Error('Ya marcaste "' + tipo + '" hoy.');
    }
  }

  const ahora = new Date();
  const hora = Utilities.formatDate(ahora, tz, 'HH:mm:ss');
  sheet.appendRow([hoy, hora, usuario.id, usuario.nombre, tipo, ahora, nombreDeposito, distanciaDeposito, precisionGps ? Math.round(precisionGps) : '']);
  formatearFilaFichaje_(sheet, sheet.getLastRow(), usuario.id, tipo);
  return { ok: true, fecha: hoy, hora: hora, tipo: tipo, mensaje: mensajeSaludo_(tipo, usuario.nombre) };
}

// ---- Repintar toda la planilla (para prolijizar datos ya cargados) ----
function formatearTodo() {
  const tz = Session.getScriptTimeZone();

  const uSheet = getOrCreateSheet_(SHEET_USUARIOS, USUARIOS_HEADERS);
  uSheet.getRange(1, 1, 1, USUARIOS_HEADERS.length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  const uLast = uSheet.getLastRow();
  if (uLast >= 2) {
    const uData = uSheet.getRange(2, 1, uLast - 1, USUARIOS_HEADERS.length).getValues();
    const uBg = uData.map(row => {
      const color = colorParaUsuario_(row[0]);
      return USUARIOS_HEADERS.map(() => color);
    });
    uSheet.getRange(2, 1, uData.length, USUARIOS_HEADERS.length).setBackgrounds(uBg);
  }

  const fSheet = getOrCreateSheet_(SHEET_FICHAJES, FICHAJES_HEADERS, function (s) {
    s.getRange(1, 5).setNote('Verde = Entrada / Vuelta de almorzar. Rojo = Salida / Salida a almorzar. El color de fondo de cada fila identifica a la persona (ver planilla "Usuarios").');
  });
  fSheet.getRange(1, 1, 1, FICHAJES_HEADERS.length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  fSheet.getRange(1, 5).setNote('Verde = Entrada / Vuelta de almorzar. Rojo = Salida / Salida a almorzar. El color de fondo de cada fila identifica a la persona (ver planilla "Usuarios").');
  const fLast = fSheet.getLastRow();
  if (fLast >= 2) {
    const ancho = FICHAJES_HEADERS.length;
    const fData = fSheet.getRange(2, 1, fLast - 1, ancho).getValues();
    const bgRows = [];
    const fontColors = [];
    const fontWeights = [];
    for (let i = 0; i < fData.length; i++) {
      const color = colorParaUsuario_(fData[i][2]);
      const filaColores = [];
      const filaPesos = [];
      for (let c = 0; c < ancho; c++) {
        filaColores.push(c === 4 ? (esTipoEntrada_(fData[i][4]) ? COLOR_ENTRADA : COLOR_SALIDA) : '#000000');
        filaPesos.push(c === 4 ? 'bold' : 'normal');
      }
      bgRows.push(new Array(ancho).fill(color));
      fontColors.push(filaColores);
      fontWeights.push(filaPesos);
    }
    const range = fSheet.getRange(2, 1, fData.length, ancho);
    range.setBackgrounds(bgRows);
    range.setFontColors(fontColors);
    range.setFontWeights(fontWeights);
  }

  SpreadsheetApp.getUi().alert('Listo, se repintó la planilla.');
}

// ---- Resumen diario (para pasar a Excel prolijo) ----
function generarResumenDeHoy() {
  return generarResumenDeFecha(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
}

function generarResumenDeFecha(fechaStr) {
  const tz = Session.getScriptTimeZone();
  const fSheet = getOrCreateSheet_(SHEET_FICHAJES, FICHAJES_HEADERS);
  const data = fSheet.getDataRange().getValues();
  const porUsuario = {};
  for (let i = 1; i < data.length; i++) {
    const filaFecha = normalizarFecha_(data[i][0], tz);
    if (filaFecha === fechaStr) {
      const id = String(data[i][2]);
      if (!porUsuario[id]) porUsuario[id] = { nombre: data[i][3] };
      porUsuario[id][data[i][4]] = normalizarHora_(data[i][1], tz);
    }
  }

  const rSheet = getOrCreateSheet_(SHEET_RESUMEN, ['Fecha', 'ID', 'Nombre', 'Entrada', 'Salida almuerzo', 'Vuelta almuerzo', 'Salida', 'Horas trabajadas']);
  const existing = rSheet.getDataRange().getValues();
  for (let i = existing.length - 1; i >= 1; i--) {
    if (normalizarFecha_(existing[i][0], tz) === fechaStr) rSheet.deleteRow(i + 1);
  }

  Object.keys(porUsuario).forEach(id => {
    const u = porUsuario[id];
    const horas = calcularHoras_(u['Entrada'], u['Salida almuerzo'], u['Vuelta almuerzo'], u['Salida']);
    rSheet.appendRow([fechaStr, id, u.nombre, u['Entrada'] || '', u['Salida almuerzo'] || '', u['Vuelta almuerzo'] || '', u['Salida'] || '', horas]);
    rSheet.getRange(rSheet.getLastRow(), 1, 1, 8).setBackground(colorParaUsuario_(id));
  });

  return { ok: true, filas: Object.keys(porUsuario).length };
}

function calcularHoras_(entrada, salidaAlm, vueltaAlm, salida) {
  if (!entrada || !salida) return '';
  const toMin = (h) => { const p = h.split(':').map(Number); return p[0] * 60 + p[1] + (p[2] || 0) / 60; };
  let total = toMin(salida) - toMin(entrada);
  if (salidaAlm && vueltaAlm) total -= (toMin(vueltaAlm) - toMin(salidaAlm));
  if (total < 0) return '';
  const horas = Math.floor(total / 60);
  const minutos = Math.round(total % 60);
  return horas + 'h ' + minutos + 'm';
}

// ---- Administración: reiniciar el PIN de un usuario ----
function promptReiniciarPin() {
  const ui = SpreadsheetApp.getUi();
  const respId = ui.prompt('Reiniciar PIN', 'ID del usuario (lo ves en la pestaña "Usuarios", ej: U001):', ui.ButtonSet.OK_CANCEL);
  if (respId.getSelectedButton() != ui.Button.OK) return;
  const id = respId.getResponseText().trim().toUpperCase();

  const respPin = ui.prompt('Reiniciar PIN', 'Nuevo PIN de 4 números para ' + id + ':', ui.ButtonSet.OK_CANCEL);
  if (respPin.getSelectedButton() != ui.Button.OK) return;
  const pin = respPin.getResponseText().trim();
  if (!/^\d{4}$/.test(pin)) { ui.alert('El PIN tiene que ser de 4 números.'); return; }

  const sheet = getOrCreateSheet_(SHEET_USUARIOS, USUARIOS_HEADERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.getRange(i + 1, 4).setValue(hashPin_(pin));
      ui.alert('Listo, PIN actualizado para ' + data[i][1] + '. Ojo: el rostro registrado no cambió, sigue siendo el mismo.');
      return;
    }
  }
  ui.alert('No se encontró el ID ' + id + '.');
}

// Prende un disparador automático: todas las noches a las 23:00 se genera solo el
// resumen del día, sin que haga falta entrar a tocar el menú. Se ejecuta una vez
// (tocando el botón del menú); Apps Script se encarga de repetirlo todos los días.
function activarResumenAutomatico() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'generarResumenDeHoy') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('generarResumenDeHoy').timeBased().everyDays(1).atHour(23).create();
  SpreadsheetApp.getUi().alert('Listo. Todas las noches a las 23:00 se va a generar solo el resumen del día en la pestaña "Resumen diario".');
}

// ---- Menú dentro de la planilla de Google ----
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fichero')
    .addItem('Generar resumen de hoy', 'generarResumenDeHoy')
    .addItem('Generar resumen de una fecha...', 'promptResumenFecha')
    .addItem('Activar resumen automático (todas las noches)...', 'activarResumenAutomatico')
    .addSeparator()
    .addItem('Formatear planilla (colores)', 'formatearTodo')
    .addItem('Reiniciar PIN de un usuario...', 'promptReiniciarPin')
    .addToUi();
}

function promptResumenFecha() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Generar resumen', 'Fecha (AAAA-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() == ui.Button.OK) {
    const res = generarResumenDeFecha(resp.getResponseText().trim());
    ui.alert('Listo. Se generaron ' + res.filas + ' filas en "Resumen diario".');
  }
}
