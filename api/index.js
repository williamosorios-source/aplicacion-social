const express = require('express');
const cors = require('cors');
const { google } = require('@googleapis/sheets');
const fetch = require('node-fetch');

const app = express();
app.use(cors());
app.use(express.json());

const SPREADSHEET_ID = process.env.SPREADSHEET_ID || "1xY6EgKQWwHhPcmugG8kQQsPET2D5dIzAd7LZF6aU_ko";
const ONESIGNAL_APP_ID = process.env.ONESIGNAL_APP_ID || "985f9dc0-c0d9-445c-b600-75e99180a9fd";
const ONESIGNAL_REST_KEY = process.env.ONESIGNAL_REST_KEY || "";

const googleAuth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') : undefined,
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth: googleAuth });

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

async function computeParticipants_() {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'PARTICIPANTES!A2:H',
    });

    const rows = res.data.values || [];
    if (rows.length === 0) return [];

    let allParticipants = [];
    let targetsCol = [];
    let changed = false;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const idVal = cellStr_(row[0]);
      const nombreVal = cellStr_(row[2]);
      const cedulaVal = cellStr_(row[3]);
      const generoRaw = cellStr_(row[4]).toUpperCase();
      const codigoVal = cellStr_(row[5]);
      const targetsVal = cellStr_(row[6]);
      const confirmadosVal = parseInt(cellStr_(row[7])) || 0;

      let generoVal = "";
      if (generoRaw === 'HOMBRE') generoVal = 'HOMBRE';
      else if (generoRaw === 'MUJER') generoVal = 'MUJER';

      let assignedArray = [];
      if (targetsVal !== '') {
        assignedArray = targetsVal.split(',').map(t => t.trim()).filter(t => t !== '');
      }

      targetsCol.push([targetsVal]);

      if (idVal && nombreVal) {
        allParticipants.push({
          rowIndex: i + 2,
          id: idVal,
          nombre: nombreVal,
          cedula: cedulaVal,
          genero: generoVal,
          codigoSorpresa: codigoVal,
          assignedTargets: assignedArray,
          bailesConfirmados: confirmadosVal
        });
      }
    }

    const hombres = allParticipants.filter(p => p.genero === 'HOMBRE');
    const mujeres = allParticipants.filter(p => p.genero === 'MUJER');

    allParticipants.forEach((user) => {
      if (user.genero !== 'HOMBRE' && user.genero !== 'MUJER') return;

      if (user.assignedTargets.length === 0) {
        let pool = (user.genero === 'MUJER') ? hombres : mujeres;
        pool = pool.filter(p => String(p.id) !== String(user.id));

        if (pool.length > 0) {
          let poolCopy = pool.slice();
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

          user.assignedTargets = poolCopy.slice(0, 5).map(t => t.id);
          targetsCol[user.rowIndex - 2] = [user.assignedTargets.join(',')];
          changed = true;
        }
      }
    });

    if (changed) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `PARTICIPANTES!G2:G${rows.length + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: targetsCol },
      });
    }

    return allParticipants;
  } catch (e) {
    console.error("Error en computeParticipants_:", e);
    return [];
  }
}

async function getParticipants() {
  return await computeParticipants_();
}

async function readLogs_() {
  try {
    const res = await sheets.spreadsheets.values.get({
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
  } catch (e) {
    console.error("Error en readLogs_:", e);
    return [];
  }
}

async function getLogs() {
  return await readLogs_();
}

async function updateConfirmedDancesFast_(specificIds) {
  try {
    const resPart = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'PARTICIPANTES!A2:H',
    });

    const partRows = resPart.data.values || [];
    if (partRows.length === 0) return;

    const logs = await readLogs_();
    let outgoingMap = {};

    logs.forEach((l) => {
      if (!outgoingMap[l.idEmisor]) outgoingMap[l.idEmisor] = {};
      outgoingMap[l.idEmisor][l.idReceptor] = true;
    });

    let targetIdsMap = null;
    if (specificIds && Array.isArray(specificIds)) {
      targetIdsMap = {};
      specificIds.forEach(id => { targetIdsMap[String(id).trim()] = true; });
    }

    let countsColH = [];
    let needsUpdate = false;

    for (let i = 0; i < partRows.length; i++) {
      const myId = cellStr_(partRows[i][0]);
      const currentVal = partRows[i][7] ? partRows[i][7] : 0;

      if (targetIdsMap && !targetIdsMap[myId]) {
        countsColH.push([currentVal]);
        continue;
      }

      let confirmedUniqueCount = 0;

      if (myId !== '' && outgoingMap[myId]) {
        const myPartners = Object.keys(outgoingMap[myId]);

        myPartners.forEach((partnerId) => {
          if (outgoingMap[partnerId] && outgoingMap[partnerId][myId]) {
            confirmedUniqueCount++;
          }
        });
      }

      countsColH.push([confirmedUniqueCount]);
      needsUpdate = true;
    }

    if (needsUpdate) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `PARTICIPANTES!H2:H${partRows.length + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: countsColH },
      });
    }

  } catch (e) {
    console.error("Error en updateConfirmedDancesFast_:", e);
  }
}

async function sendPushNotification(receiverId, senderName, senderId, userMessage) {
  if (!ONESIGNAL_REST_KEY) {
    console.log("Falta ONESIGNAL_REST_KEY: no se envió la notificación.");
    return;
  }

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
  try {
    const logs = await readLogs_();
    const emisorStr = String(idEmisor).trim();
    const receptorStr = String(idReceptor).trim();

    const previousPartnerLog = logs.find((l) => {
      return String(l.idEmisor).trim() === receptorStr &&
             String(l.idReceptor).trim() === emisorStr;
    });

    const now = getFormattedNow();
    
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'REGISTROS!A:D',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[now, emisorStr, receptorStr, mensaje]]
      }
    });

    const matched = !!previousPartnerLog;

    if (matched) {
      await updateConfirmedDancesFast_([emisorStr, receptorStr]);
    }

    const participants = await getParticipants();
    const senderObj = participants.find(p => String(p.id).trim() === emisorStr);
    const senderName = senderObj ? senderObj.nombre : "Bailarín " + idEmisor;

    await sendPushNotification(idReceptor, senderName, idEmisor, mensaje);

    return { status: 'SUCCESS', matched: matched };
  } catch (e) {
    return { status: 'ERROR', message: e.toString() };
  }
}

async function readPollOptions_() {
  try {
    const resO1 = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'ENCUESTA!O1',
    });

    const statusO1 = cellStr_(resO1.data.values ? resO1.data.values[0][0] : '').toUpperCase();
    if (statusO1 !== "SI") {
      return [];
    }

    const resData = await sheets.spreadsheets.values.get({
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
  } catch (e) {
    console.error("Error en readPollOptions_:", e);
    return [];
  }
}

async function readVotes_() {
  try {
    const res = await sheets.spreadsheets.values.get({
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
  } catch (e) {
    console.error("Error en readVotes_:", e);
    return [];
  }
}

async function getPollData(userId) {
  try {
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
  } catch (e) {
    console.error("Error en getPollData:", e);
    return { options: [], counts: {}, totalVotes: 0, userVote: null, disabled: true };
  }
}

async function submitVote(userId, song) {
  try {
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

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'VOTOS!A:C',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[now, uid, chosen]]
      }
    });

    return { status: 'SUCCESS' };
  } catch (e) {
    return { status: 'ERROR', message: e.toString() };
  }
}

const ACTIONS = {
  getParticipants: getParticipants,
  getLogs: getLogs,
  registerDance: registerDance,
  getPollData: getPollData,
  submitVote: submitVote
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
    res.status(500).json({ ok: false, error: String(err) });
  }
});

module.exports = app;
