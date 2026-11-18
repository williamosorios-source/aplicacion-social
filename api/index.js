const express = require('express');
const cors = require('cors');
const { sheets } = require('@googleapis/sheets');
const { GoogleAuth } = require('google-auth-library');
const fetch = require('node-fetch');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.text());

app.use((req, res, next) => {
  if (typeof req.body === 'string' && req.body.trim()) {
    try { req.body = JSON.parse(req.body); } catch (_) {}
  }
  next();
});

const SPREADSHEET_ID = process.env.SPREADSHEET_ID || "1xY6EgKQWwHhPcmugG8kQQsPET2D5dIzAd7LZF6aU_ko";
const ONESIGNAL_APP_ID = process.env.ONESIGNAL_APP_ID || "985f9dc0-c0d9-445c-b600-75e99180a9fd";
const ONESIGNAL_REST_KEY = process.env.ONESIGNAL_REST_KEY || "";

function getSheetsClient() {
  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL;
  let privateKey = process.env.GOOGLE_PRIVATE_KEY;

  if (!clientEmail || !privateKey) {
    throw new Error("Faltan las variables de entorno GOOGLE_CLIENT_EMAIL o GOOGLE_PRIVATE_KEY en Vercel.");
  }

  privateKey = privateKey.replace(/\\n/g, '\n');

  const auth = new GoogleAuth({
    credentials: {
      client_email: clientEmail,
      private_key: privateKey,
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  return sheets({ version: 'v4', auth });
}

function seedRandom(seed) {
  var x = Math.sin(seed++) * 10000;
  return x - Math.floor(x);
}

function cellStr_(v) {
  return String(v !== null && v !== undefined ? v : '').trim();
}

function getFormattedNow() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const day = pad(now.getDate());
  const month = pad(now.getMonth() + 1);
  const year = now.getFullYear();
  const hours = pad(now.getHours());
  const mins = pad(now.getMinutes());
  const secs = pad(now.getSeconds());
  return `${day}/${month}/${year} ${hours}:${mins}:${secs}`;
}

// ======================================================
// PARTICIPANTES, ASIGNACIÓN DE TARGETS Y RETOS
// ======================================================
const sheetReads = new Map();

function cachedSheetRead(key, ttl, load) {
  const entry = sheetReads.get(key);
  if (entry && entry.expiresAt > Date.now()) return entry.promise;

  const pending = { expiresAt: Infinity };
  pending.promise = Promise.resolve().then(load).then(data => {
    pending.expiresAt = Date.now() + ttl;
    return data;
  }).catch(error => {
    if (sheetReads.get(key) === pending) sheetReads.delete(key);
    throw error;
  });
  sheetReads.set(key, pending);
  return pending.promise;
}

function isQuotaError(error) {
  return Number(error.response?.status || error.code) === 429;
}

async function computeParticipants_() {
  const sheetsClient = getSheetsClient();
  const res = await sheetsClient.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ['PARTICIPANTES!A2:K', 'PARTICIPANTES!Q2'],
  });

  const [participantsRange, startRange] = res.data.valueRanges;
  const rows = participantsRange.values || [];
  const objectivesActive = cellStr_(startRange.values?.[0]?.[0]).toUpperCase() === 'SI';
  if (rows.length === 0) return [];

  let allParticipants = [];
  const targetUpdates = [];
  let retosAsignadosCol = [];
  let changedRetos = false;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const idVal = cellStr_(row[0]);
    const nombreVal = cellStr_(row[2]);
    const cedulaVal = cellStr_(row[3]);
    const generoRaw = cellStr_(row[4]).toUpperCase();
    const codigoVal = cellStr_(row[5]);
    const targetsVal = cellStr_(row[6]);
    const confirmadosVal = parseInt(cellStr_(row[7])) || 0;
    const retosAsignadosVal = cellStr_(row[8]);
    const retosCompletadosVal = cellStr_(row[9]);

    let generoVal = "";
    if (generoRaw === 'HOMBRE') generoVal = 'HOMBRE';
    else if (generoRaw === 'MUJER') generoVal = 'MUJER';

    const present = cellStr_(row[10]).toUpperCase() === 'SI';
    let assignedArray = [];
    if (targetsVal !== '') {
      assignedArray = targetsVal.split(',').map(t => t.trim()).filter(t => t !== '');
      assignedArray = [...new Set(assignedArray)].slice(0, 4);
    }

    let retosAsignadosArray = [];
    if (retosAsignadosVal !== '') {
      retosAsignadosArray = retosAsignadosVal.split(',').map(r => parseInt(r.trim())).filter(r => !isNaN(r));
    }

    let retosCompletadosArray = [];
    if (retosCompletadosVal !== '') {
      retosCompletadosArray = retosCompletadosVal.split(',').map(r => parseInt(r.trim())).filter(r => !isNaN(r));
    }

    retosAsignadosCol.push([retosAsignadosVal]);

    if (idVal && nombreVal) {
      allParticipants.push({
        rowIndex: i + 2,
        id: idVal,
        nombre: nombreVal,
        cedula: cedulaVal,
        genero: generoVal,
        codigoSorpresa: codigoVal,
        present,
        objectivesActive,
        assignedTargets: objectivesActive && present ? assignedArray : [],
        bailesConfirmados: confirmadosVal,
        retosAsignados: retosAsignadosArray,
        retosCompletados: retosCompletadosArray
      });
    }
  }

  const hombres = allParticipants.filter(p => p.present && p.genero === 'HOMBRE');
  const mujeres = allParticipants.filter(p => p.present && p.genero === 'MUJER');

  allParticipants.forEach((user) => {
    // 1. Asignación de Objetivos de Baile
    if (objectivesActive && user.present && user.genero && user.assignedTargets.length < 4) {
      const pool = user.genero === 'MUJER' ? hombres : mujeres;
      const poolCopy = [...new Set(pool.map(p => p.id))].filter(id => id !== user.id);
      user.assignedTargets = [];

      if (poolCopy.length >= 4) {
        let userSeed = 0;
        for (let k = 0; k < user.id.length; k++) {
          userSeed += user.id.charCodeAt(k);
        }

        for (let j = poolCopy.length - 1; j > 0; j--) {
          let r = Math.floor(seedRandom(userSeed + j) * (j + 1));
          let temp = poolCopy[j];
          poolCopy[j] = poolCopy[r];
          poolCopy[r] = temp;
        }

        user.assignedTargets = poolCopy.slice(0, 4);
        targetUpdates.push({
          range: `PARTICIPANTES!G${user.rowIndex}`,
          values: [[user.assignedTargets.join(',')]]
        });
      }
    }

    // 2. Asignación de Secuencia de 4 Retos Navideños
    if (user.retosAsignados.length === 0) {
      let userSeed = 0;
      for (let k = 0; k < user.id.length; k++) {
        userSeed += user.id.charCodeAt(k);
      }

      // Pista 1: Reto aleatorio entre 3 y 10
      let poolGeneral = [3, 4, 5, 6, 7, 8, 9, 10];
      let r1Index = Math.floor(seedRandom(userSeed + 1) * poolGeneral.length);
      let r1 = poolGeneral[r1Index];
      poolGeneral.splice(r1Index, 1);

      // Pista 2: Obligatoriamente ID 1 (Anfitrión) o ID 2 (DJ)
      let r2 = (seedRandom(userSeed + 2) > 0.5) ? 1 : 2;

      // Pista 3: Reto aleatorio entre los restantes de 3 a 10
      let r3Index = Math.floor(seedRandom(userSeed + 3) * poolGeneral.length);
      let r3 = poolGeneral[r3Index];
      poolGeneral.splice(r3Index, 1);

      // Pista 4: Reto aleatorio entre los restantes de 3 a 10
      let r4Index = Math.floor(seedRandom(userSeed + 4) * poolGeneral.length);
      let r4 = poolGeneral[r4Index];

      user.retosAsignados = [r1, r2, r3, r4];
      retosAsignadosCol[user.rowIndex - 2] = [user.retosAsignados.join(',')];
      changedRetos = true;
    }
  });

  if (targetUpdates.length) {
    await sheetsClient.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'RAW', data: targetUpdates },
    });
  }

  if (changedRetos) {
    await sheetsClient.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `PARTICIPANTES!I2:I${rows.length + 1}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: retosAsignadosCol },
    });
  }

  return allParticipants;
}

async function getParticipants() {
  return cachedSheetRead('participants', 30000, computeParticipants_);
}

// ======================================================
// RETOS NAVIDEÑOS Y MISTERIO DE LA PARTITURA
// ======================================================
const NARRATIVA_PISTAS = {
  1: {
    titulo: 'Pista 1: El Ritmo de los Renos',
    historia: '❄️ El Grinch congeló la Navidad y escondió los 4 fragmentos de la partitura secreta del baile. 🎶 Saliste a la pista y recuperaste la 1ª nota musical. ¡El ritmo empieza a despertar! 📍 Sigue la primera pista en la sala para encontrar el mapa del segundo fragmento...'
  },
  2: {
    titulo: 'Pista 2: La Chispa Navideña',
    historia: '🎼 El Grinch bloqueó el sonido y dejó el 2º fragmento de la partitura bajo custodia. 👥 Conectaste con tus parejas objetivo y la música cobró fuerza. ⚡ Dirígete al DJ o al Anfitrión para superar su prueba física y recibir el código del segundo fragmento.'
  },
  3: {
    titulo: 'Pista 3: El Desierto del Grinch',
    historia: '🧩 ¡El Grinch se acorrala! El calor de la pista está derritiendo su hechizo de hielo. 🎷 Has rescatado el 3er fragmento de la partitura secreta. Solo falta una nota para completar la canción. 🕵️‍♂️ Encuentra la siguiente ubicación en la sala y resuelve el acertijo final.'
  },
  4: {
    titulo: 'Pista 4: La Estrella del Social',
    historia: '🏆 ¡Salvaste el Social Navideño! Uniste los 4 fragmentos de la partitura y la música volvió a sonar a todo volumen. 🌟 Bailaste con tus 4 objetivos y superaste cada desafío en la pista. 🎁 ¡Presenta esta pantalla al Anfitrión y reclama tu premio de Navidad!'
  }
};

function readRetosSheet_() {
  return cachedSheetRead('retos', 60000, loadRetosSheet_);
}

async function loadRetosSheet_() {
  const sheetsClient = getSheetsClient();
  const res = await sheetsClient.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ['RETOS!A2:F', 'RETOS!N2:O2'],
  });

  const [retosRange, staffRange] = res.data.valueRanges;
  const rows = retosRange.values || [];
  let retosMap = {};

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const idReto = parseInt(cellStr_(row[0]));
    if (!isNaN(idReto)) {
      retosMap[idReto] = {
        rowIndex: i + 2,
        id: idReto,
        ubicacion: cellStr_(row[1]),
        actividad: cellStr_(row[2]),
        solucion: cellStr_(row[3]),
        idsCompletados: cellStr_(row[4]),
        tipoValidacion: cellStr_(row[5]).toUpperCase()
      };
    }
  }

  const staff = staffRange.values?.[0] || [];
  const codigoAnfitrion = cellStr_(staff[0]) || "NAVIDAD2026";
  const codigoDJ = cellStr_(staff[1]) || "RITMO2026";

  return { retosMap, codigoAnfitrion, codigoDJ };
}

async function getChallengeData(userId, includeLogs = false) {
  try {
    const participants = await getParticipants();
    const uidStr = String(userId).trim();
    const user = participants.find(p => String(p.id).trim() === uidStr);

    if (!user) {
      throw new Error("Usuario no encontrado.");
    }

    const logs = await getLogs();
    const myUniqueDancesCount = user.bailesConfirmados;

    // Calcular cuántos de sus 4 Objetivos han sido confirmados mutuamente
    const assignedIds = user.assignedTargets || [];
    let targetMatchedCount = 0;

    assignedIds.forEach(targetId => {
      const iReg = logs.some(l => String(l.idEmisor).trim() === uidStr && String(l.idReceptor).trim() === String(targetId).trim());
      const targetReg = logs.some(l => String(l.idEmisor).trim() === String(targetId).trim() && String(l.idReceptor).trim() === uidStr);
      if (iReg && targetReg) targetMatchedCount++;
    });

    const { retosMap } = await readRetosSheet_();

    let tarjetas = [];

    for (let pistaNum = 1; pistaNum <= 4; pistaNum++) {
      let isUnlocked = false;
      let reqTexto = "";

      if (pistaNum === 1) {
        isUnlocked = myUniqueDancesCount >= 5;
        reqTexto = "Requiere haber registrado baile con al menos 5 personas.";
      } else if (pistaNum === 2) {
        isUnlocked = targetMatchedCount >= 2;
        reqTexto = "Requiere confirmar al menos 2 de tus 4 Objetivos.";
      } else if (pistaNum === 3) {
        isUnlocked = targetMatchedCount >= 3;
        reqTexto = "Requiere confirmar al menos 3 de tus 4 Objetivos.";
      } else if (pistaNum === 4) {
        isUnlocked = targetMatchedCount >= 4;
        reqTexto = "Requiere confirmar el baile con tus 4 parejas objetivo.";
      }

      const isCompleted = user.retosCompletados.includes(pistaNum);
      const retoId = user.retosAsignados[pistaNum - 1] || (pistaNum === 2 ? 1 : 3);
      const retoInfo = retosMap[retoId] || {
        ubicacion: "Ubicación Misteriosa",
        actividad: "Supera la prueba asignada.",
        tipoValidacion: pistaNum === 2 ? "STAFF_ANFITRION" : "RESPUESTA"
      };

      tarjetas.push({
        pistaNumero: pistaNum,
        titulo: NARRATIVA_PISTAS[pistaNum].titulo,
        historia: NARRATIVA_PISTAS[pistaNum].historia,
        unlocked: isUnlocked,
        completed: isCompleted,
        requisitoTexto: reqTexto,
        retoId: retoId,
        ubicacion: retoInfo.ubicacion,
        actividad: retoInfo.actividad,
        tipoValidacion: retoInfo.tipoValidacion
      });
    }

    const allPistasCompleted = user.retosCompletados.length >= 4;
    const allObjectivesMatched = targetMatchedCount >= 4;
    const showGrandWinner = allPistasCompleted && allObjectivesMatched;

    return {
      userId: user.id,
      nombre: user.nombre,
      codigoSorpresa: user.codigoSorpresa,
      targetMatchedCount,
      totalObjectives: assignedIds.length,
      assignedTargets: assignedIds,
      objectivesActive: user.objectivesActive,
      present: user.present,
      myUniqueDancesCount,
      tarjetas,
      showGrandWinner,
      ...(includeLogs ? { logs } : {})
    };
  } catch (e) {
    console.error("Error en getChallengeData:", e);
    throw e;
  }
}

async function validateChallenge(userId, pistaNumero, respuestaOcodigo) {
  try {
    const uidStr = String(userId).trim();
    const pistaNum = parseInt(pistaNumero);
    const inputVal = String(respuestaOcodigo || '').trim().toUpperCase();

    if (isNaN(pistaNum) || pistaNum < 1 || pistaNum > 4) {
      return { status: 'ERROR', message: 'Número de pista no válido.' };
    }

    const challengeData = await getChallengeData(uidStr);
    const tarjeta = challengeData.tarjetas.find(t => t.pistaNumero === pistaNum);

    if (!tarjeta) {
      return { status: 'ERROR', message: 'Tarjeta de pista no encontrada.' };
    }

    if (!tarjeta.unlocked) {
      return { status: 'ERROR', message: 'Esta pista aún está bloqueada. Revisa los requisitos de baile u objetivos.' };
    }

    if (tarjeta.completed) {
      return { status: 'SUCCESS', message: '¡Pista ya completada previamente!' };
    }

    const { retosMap, codigoAnfitrion, codigoDJ } = await loadRetosSheet_();
    const retoInfo = retosMap[tarjeta.retoId];

    let isValid = false;

    if (tarjeta.tipoValidacion === 'STAFF_ANFITRION') {
      isValid = (inputVal === codigoAnfitrion.toUpperCase());
    } else if (tarjeta.tipoValidacion === 'STAFF_DJ') {
      isValid = (inputVal === codigoDJ.toUpperCase());
    } else {
      // Validación por respuesta de acertijo / ubicación
      const expectedSol = retoInfo ? String(retoInfo.solucion).trim().toUpperCase() : '';
      isValid = (expectedSol !== '' && inputVal === expectedSol);
    }

    if (!isValid) {
      return { status: 'ERROR', message: 'Código o respuesta incorrecta. Inténtalo de nuevo o consulta con el Staff.' };
    }

    // Actualizar Google Sheets
    const sheetsClient = getSheetsClient();
    const participants = await computeParticipants_();
    const user = participants.find(p => String(p.id).trim() === uidStr);

    let newCompletados = user.retosCompletados.slice();
    if (!newCompletados.includes(pistaNum)) {
      newCompletados.push(pistaNum);
      newCompletados.sort((a, b) => a - b);
    }

    // 1. Guardar nueva lista de pistas completadas en Columna J de PARTICIPANTES
    await sheetsClient.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `PARTICIPANTES!J${user.rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[newCompletados.join(',')]] },
    });

    sheetReads.delete('participants');

    // 2. Registrar anidadamente el ID del usuario en Columna E de RETOS
    if (retoInfo) {
      let currentIds = retoInfo.idsCompletados ? retoInfo.idsCompletados.split(',').map(s => s.trim()).filter(s => s !== '') : [];
      if (!currentIds.includes(uidStr)) {
        currentIds.push(uidStr);
        await sheetsClient.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `RETOS!E${retoInfo.rowIndex}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[currentIds.join(', ')]] },
        });
        sheetReads.delete('retos');
      }
    }

    return { status: 'SUCCESS', message: '¡Pista completada con éxito!' };
  } catch (e) {
    console.error("Error en validateChallenge:", e);
    return { status: 'ERROR', message: e.toString() };
  }
}

// ======================================================
// REGISTROS DE BAILES Y NOTIFICACIONES
// ======================================================
async function readLogs_() {
  const sheetsClient = getSheetsClient();
  const res = await sheetsClient.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'REGISTROS!A2:D',
  });

  const rows = res.data.values || [];
  let logs = [];

  for (let i = 0; i < rows.length; i++) {
    const emisor = cellStr_(rows[i][1]);
    const receptor = cellStr_(rows[i][2]);
    const msg = cellStr_(rows[i][3]);

    if (emisor !== '' && receptor !== '') {
      logs.push({
        fecha: cellStr_(rows[i][0]),
        idEmisor: emisor,
        idReceptor: receptor,
        mensaje: msg
      });
    }
  }
  return logs;
}

async function getLogs() {
  return cachedSheetRead('logs', 10000, readLogs_);
}

function countConfirmedDances(logs, userId) {
  const incoming = new Set();
  const outgoing = new Set();
  for (const log of logs) {
    if (log.idEmisor === userId) outgoing.add(log.idReceptor);
    if (log.idReceptor === userId) incoming.add(log.idEmisor);
  }
  return [...outgoing].filter(partnerId => incoming.has(partnerId)).length;
}

async function sendPushNotification(receiverId, senderName, senderId, userMessage) {
  if (!ONESIGNAL_REST_KEY) return;

  let pushText = `${senderName} (${senderId}) solicita registrar un baile contigo. Ingresa su ID ${senderId} en la app para confirmar el baile.`;
  if (userMessage) {
    pushText += ` Te dejó este mensaje: "${userMessage}"`;
  }

  const payload = {
    "app_id": ONESIGNAL_APP_ID,
    "include_aliases": { "external_id": [String(receiverId).trim()] },
    "target_channel": "push",
    "headings": { "es": "💃 ¡Solicitud de Baile Recibida!" },
    "contents": { "es": pushText }
  };

  try {
    await fetch("https://onesignal.com/api/v1/notifications", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Basic ${ONESIGNAL_REST_KEY}`
      },
      body: JSON.stringify(payload)
    });
  } catch (e) {
    console.error("Error enviando push:", e);
  }
}

async function registerDance(idEmisor, idReceptor, mensaje) {
  const sheetsClient = getSheetsClient();
  const [logs, participants] = await Promise.all([readLogs_(), getParticipants()]);
  const emisorStr = String(idEmisor).trim();
  const receptorStr = String(idReceptor).trim();

  const previousPartnerLog = logs.find((l) => {
    return String(l.idEmisor).trim() === receptorStr &&
           String(l.idReceptor).trim() === emisorStr;
  });

  const now = getFormattedNow();
  
  await sheetsClient.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: 'REGISTROS!A:D',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[now, emisorStr, receptorStr, mensaje]]
    }
  });

  sheetReads.delete('logs');
  const matched = !!previousPartnerLog;

  if (matched) {
    const updatedLogs = [...logs, { idEmisor: emisorStr, idReceptor: receptorStr }];
    const data = participants
      .filter(p => p.id === emisorStr || p.id === receptorStr)
      .map(p => ({
        range: `PARTICIPANTES!H${p.rowIndex}`,
        values: [[countConfirmedDances(updatedLogs, p.id)]]
      }));

    if (data.length) {
      try {
        await sheetsClient.spreadsheets.values.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: { valueInputOption: 'USER_ENTERED', data }
        });
        sheetReads.delete('participants');
      } catch (error) {
        console.error('El baile quedó registrado. Falló la actualización del contador:', error.message);
      }
    }
  }

  const senderObj = participants.find(p => String(p.id).trim() === emisorStr);
  const senderName = senderObj ? senderObj.nombre : "Bailarín " + idEmisor;

  await sendPushNotification(idReceptor, senderName, idEmisor, mensaje);

  return { status: 'SUCCESS', matched: matched };
}

// ======================================================
// ENCUESTAS Y VOTACIONES
// ======================================================
async function readPollOptions_() {
  const sheetsClient = getSheetsClient();
  const resO1 = await sheetsClient.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'ENCUESTA!O1',
  });

  const statusO1 = cellStr_(resO1.data.values ? resO1.data.values[0][0] : '').toUpperCase();
  if (statusO1 !== "SI") {
    return [];
  }

  const resData = await sheetsClient.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'ENCUESTA!A2:B',
  });

  const rows = resData.data.values || [];
  let options = [];
  for (let i = 0; i < rows.length; i++) {
    const song = cellStr_(rows[i][0]);
    let imageUrl = cellStr_(rows[i][1]);
    if (imageUrl && !imageUrl.startsWith('http')) imageUrl = '';
    if (song) {
      options.push({ song: song, imageUrl: imageUrl });
    }
  }
  return options;
}

async function readVotes_() {
  const sheetsClient = getSheetsClient();
  const res = await sheetsClient.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'VOTOS!A2:C',
  });

  const rows = res.data.values || [];
  let votes = [];
  for (let i = 0; i < rows.length; i++) {
    const uId = cellStr_(rows[i][1]);
    const song = cellStr_(rows[i][2]);
    if (uId && song) votes.push({ userId: uId, song: song });
  }
  return votes;
}

async function getPollData(userId) {
  const options = await readPollOptions_();
  if (!options || options.length === 0) {
    return { options: [], counts: {}, totalVotes: 0, userVote: null, disabled: true };
  }

  const votes = await readVotes_();
  const uid = String(userId).trim();
  let userVote = null;

  let counts = {};
  options.forEach(opt => { counts[opt.song] = 0; });
  votes.forEach(v => {
    if (v.userId === uid) userVote = v.song;
    if (counts.hasOwnProperty(v.song)) counts[v.song]++;
  });

  return { options: options, counts: counts, totalVotes: votes.length, userVote: userVote, disabled: false };
}

async function submitVote(userId, song) {
  const options = await readPollOptions_();
  if (!options || options.length === 0) {
    return { status: 'ERROR', message: 'La encuesta se encuentra deshabilitada en este momento.' };
  }

  const uid = String(userId).trim();
  const chosen = String(song).trim();

  const valid = options.some(o => o.song === chosen);
  if (!valid) {
    return { status: 'ERROR', message: 'La opción elegida ya no está disponible.' };
  }

  const existingVotes = await readVotes_();
  const already = existingVotes.some(v => v.userId === uid);
  if (already) {
    return { status: 'ERROR', message: 'Ya has emitido tu voto anteriormente.' };
  }

  const now = getFormattedNow();
  const sheetsClient = getSheetsClient();

  await sheetsClient.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: 'VOTOS!A:C',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[now, uid, chosen]]
    }
  });

  return { status: 'SUCCESS' };
}

// ======================================================
// ENTRADA GENERAL DE LA API
// ======================================================
const ACTIONS = {
  getParticipants,
  getLogs,
  registerDance,
  getPollData,
  submitVote,
  getChallengeData,
  validateChallenge
};

app.all('*', async (req, res) => {
  try {
    const { action, args } = req.body || {};
    if (!action) {
      return res.json({ ok: true, data: 'API activa en Vercel/Node' });
    }
    const fn = ACTIONS[action];
    if (!fn) throw new Error('Acción no válida: ' + action);

    const data = await fn.apply(null, args || []);
    res.json({ ok: true, data: data });
  } catch (err) {
    console.error("Error en handler de API:", err);
    if (isQuotaError(err)) {
      return res.status(429).json({
        ok: false,
        error: 'El servicio de registro está ocupado. Espera un momento e inténtalo de nuevo.'
      });
    }
    res.status(200).json({ ok: false, error: String(err.message || err) });
  }
});

module.exports = app;
