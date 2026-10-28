const { google } = require('googleapis');

// Configuración de Autenticación con Google Sheets
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });
const SPREADSHEET_ID = process.env.GOOGLE_SHEETS_ID;

// Helper para leer rangos
async function readSheetRange(range) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: range,
    });
    return res.data.values || [];
  } catch (err) {
    console.error(`Error al leer rango ${range}:`, err);
    return [];
  }
}

// Helper para escribir/actualizar celdas
async function updateSheetRange(range, values) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
}

// 1. Obtener lista de participantes
async function getParticipants() {
  const rows = await readSheetRange('PARTICIPANTES!A2:J');
  return rows.map((row) => ({
    id: String(row[0] || '').trim(),
    email: String(row[1] || '').trim(),
    nombre: String(row[2] || '').trim(),
    cedula: String(row[3] || '').trim(),
    genero: String(row[4] || '').trim(),
    codigoSorpresa: String(row[5] || '').trim(),
    assignedTargets: String(row[6] || '').split(',').map(s => s.trim()).filter(Boolean),
    bailesConfirmados: Number(row[7] || 0),
    retosAsignados: String(row[8] || '').trim(),
    retosCompletados: String(row[9] || '').trim()
  })).filter(p => p.id !== '');
}

// Helper para actualizar los retos completados de un usuario (Columna J)
async function updateParticipantCompletedChallenges(userId, completadosString) {
  const rows = await readSheetRange('PARTICIPANTES!A2:A');
  const rowIndex = rows.findIndex(r => String(r[0]).trim() === String(userId).trim());
  if (rowIndex !== -1) {
    const sheetRow = rowIndex + 2;
    await updateSheetRange(`PARTICIPANTES!J${sheetRow}`, [[completadosString]]);
  }
}

// Helper para actualizar los IDs completados en la pestaña RETOS (Columna E)
async function updateChallengeCompletedUsers(challengeId, idsString) {
  const rows = await readSheetRange('RETOS!A2:A');
  const rowIndex = rows.findIndex(r => String(r[0]).trim() === String(challengeId).trim());
  if (rowIndex !== -1) {
    const sheetRow = rowIndex + 2;
    await updateSheetRange(`RETOS!E${sheetRow}`, [[idsString]]);
  }
}

// Helper para asignar los 4 retos aleatorios a un usuario si no los tiene en Columna I
async function ensureUserChallengesAssigned(user) {
  if (user.retosAsignados && user.retosAsignados.length > 0) {
    return user.retosAsignados.split(',').map(s => s.trim()).filter(Boolean);
  }

  // Generar secuencia de 4 retos
  const randomPista1 = Math.floor(Math.random() * 8) + 3; // 3 a 10
  const randomPista2 = Math.random() < 0.5 ? 1 : 2; // Obligatorio ID 1 (Anfitrión) o ID 2 (DJ)

  let randomPista3 = Math.floor(Math.random() * 8) + 3;
  while (randomPista3 === randomPista1) {
    randomPista3 = Math.floor(Math.random() * 8) + 3;
  }

  let randomPista4 = Math.floor(Math.random() * 8) + 3;
  while (randomPista4 === randomPista1 || randomPista4 === randomPista3) {
    randomPista4 = Math.floor(Math.random() * 8) + 3;
  }

  const assignedArray = [randomPista1, randomPista2, randomPista3, randomPista4];
  const assignedString = assignedArray.join(',');

  try {
    const rows = await readSheetRange('PARTICIPANTES!A2:A');
    const rowIndex = rows.findIndex(r => String(r[0]).trim() === String(user.id).trim());
    if (rowIndex !== -1) {
      const sheetRow = rowIndex + 2;
      await updateSheetRange(`PARTICIPANTES!I${sheetRow}`, [[assignedString]]);
    }
  } catch (err) {
    console.error('Error al guardar retos asignados:', err);
  }

  return assignedArray.map(String);
}

// 2. Leer Logs de Bailes
async function getLogs() {
  const rows = await readSheetRange('LOGS!A2:D');
  return rows.map((row) => ({
    idEmisor: String(row[0] || '').trim(),
    idReceptor: String(row[1] || '').trim(),
    mensaje: String(row[2] || '').trim(),
    fecha: String(row[3] || '').trim()
  })).filter(l => l.idEmisor !== '');
}

// 3. Registrar Baile
async function registerDance(idEmisor, idReceptor, mensaje) {
  idEmisor = String(idEmisor).trim();
  idReceptor = String(idReceptor).trim();

  if (idEmisor === idReceptor) {
    return { status: 'ERROR', message: 'No puedes registrar un baile contigo mismo.' };
  }

  const logs = await getLogs();
  
  const alreadyRegistered = logs.some(l => l.idEmisor === idEmisor && l.idReceptor === idReceptor);
  if (alreadyRegistered) {
    return { status: 'ERROR', message: 'Ya habías registrado un baile con esta persona.' };
  }

  const now = new Date().toISOString();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: 'LOGS!A:D',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[idEmisor, idReceptor, mensaje, now]]
    }
  });

  const isMutual = logs.some(l => l.idEmisor === idReceptor && l.idReceptor === idEmisor);

  return { status: 'SUCCESS', matched: isMutual };
}

// 4. Obtener datos de "El Misterio de la Partitura Navideña"
async function getChallengeData(userId) {
  userId = String(userId).trim();
  const participants = await getParticipants();
  const user = participants.find(p => String(p.id).trim() === userId);

  if (!user) {
    return { tarjetas: [], showGrandWinner: false };
  }

  const retosAsignadosIds = await ensureUserChallengesAssigned(user);
  const logs = await getLogs();

  const myEmittedReceptors = new Set(logs.filter(l => String(l.idEmisor).trim() === userId).map(l => String(l.idReceptor).trim()));
  const myReceivedSenders = new Set(logs.filter(l => String(l.idReceptor).trim() === userId).map(l => String(l.idEmisor).trim()));
  
  const confirmedMatches = [...myEmittedReceptors].filter(id => myReceivedSenders.has(id));

  const targets = user.assignedTargets || [];
  const matchedTargetsCount = targets.filter(tId => confirmedMatches.includes(String(tId).trim())).length;

  const retosRows = await readSheetRange('RETOS!A2:F11');
  const retosMap = {};
  
  for (let i = 1; i <= 10; i++) {
    retosMap[String(i)] = {
      id: String(i),
      ubicacion: i === 1 ? 'Anfitrión del Evento' : (i === 2 ? 'Cabina del DJ' : `Punto Físico #${i}`),
      actividad: 'Presenta la prueba para obtener la validación.',
      respuesta: '',
      tipoValidacion: i === 1 ? 'STAFF_ANFITRION' : (i === 2 ? 'STAFF_DJ' : 'RESPUESTA')
    };
  }

  retosRows.forEach(r => {
    const idStr = String(r[0] || '').trim();
    if (idStr) {
      retosMap[idStr] = {
        id: idStr,
        ubicacion: String(r[1] || '').trim() || `Punto Físico #${idStr}`,
        actividad: String(r[2] || '').trim() || 'Prueba Navideña',
        respuesta: String(r[3] || '').trim(),
        tipoValidacion: String(r[5] || '').trim() || 'RESPUESTA'
      };
    }
  });

  const completadosPistas = String(user.retosCompletados || '').split(',').map(s => s.trim()).filter(Boolean);

  const historias = [
    "❄️ El Grinch congeló la Navidad y escondió los 4 fragmentos de la partitura secreta del baile. 🎶 Saliste a la pista y recuperaste la 1ª nota musical.",
    "🎼 El Grinch bloqueó el sonido y dejó el 2º fragmento de la partitura bajo custodia. 👥 Conectaste con tus parejas objetivo y la música cobró fuerza.",
    "🧩 ¡El Grinch se acorrala! El calor de la pista está derritiendo su hechizo de hielo. 🎷 Has rescatado el 3er fragmento de la partitura secreta.",
    "🏆 ¡Salvaste el Social Navideño! Uniste los 4 fragmentos de la partitura y la música volvió a sonar a todo volumen. 🌟 Superaste cada desafío en la pista."
  ];

  const titulos = [
    "Capítulo 1: El Ritmo de los Renos",
    "Capítulo 2: La Chispa Navideña",
    "Capítulo 3: El Desierto del Grinch",
    "Capítulo 4: La Estrella del Social"
  ];

  const tarjetas = [1, 2, 3, 4].map((num) => {
    let unlocked = false;
    let requisitoTexto = "";

    if (num === 1) {
      unlocked = myEmittedReceptors.size >= 5;
      requisitoTexto = "Baila con al menos 5 personas para desbloquear.";
    } else if (num === 2) {
      unlocked = matchedTargetsCount >= 2;
      requisitoTexto = "Confirma a 2 de tus 5 Parejas Objetivo.";
    } else if (num === 3) {
      unlocked = matchedTargetsCount >= 3;
      requisitoTexto = "Confirma a 3 de tus 5 Parejas Objetivo.";
    } else if (num === 4) {
      unlocked = matchedTargetsCount >= 5;
      requisitoTexto = "Confirma a tus 5 Parejas Objetivo completas.";
    }

    const completed = completadosPistas.includes(String(num));
    const challengeId = String(retosAsignadosIds[num - 1] || "1").trim();
    const retoInfo = retosMap[challengeId] || retosMap["1"];

    return {
      pistaNumero: num,
      titulo: titulos[num - 1],
      historia: historias[num - 1],
      unlocked,
      completed,
      requisitoTexto,
      ubicacion: retoInfo.ubicacion,
      actividad: retoInfo.actividad,
      tipoValidacion: retoInfo.tipoValidacion,
      retoId: challengeId
    };
  });

  const showGrandWinner = (matchedTargetsCount >= 5) && (completadosPistas.length >= 4);

  return {
    tarjetas,
    showGrandWinner,
    codigoSorpresa: user.codigoSorpresa || "¡GANADOR!"
  };
}

// 5. Validar Reto / Respuesta (Compara directamente contra Columna D de RETOS)
async function validateChallenge(userId, pistaNumero, respuestaOcodigo) {
  pistaNumero = Number(pistaNumero);
  respuestaOcodigo = String(respuestaOcodigo || '').trim();

  if (!userId || !pistaNumero || !respuestaOcodigo) {
    return { status: 'ERROR', message: 'Por favor ingresa la respuesta o el código.' };
  }

  const participants = await getParticipants();
  const user = participants.find(p => String(p.id).trim() === String(userId).trim());

  if (!user) {
    return { status: 'ERROR', message: 'Usuario no encontrado.' };
  }

  const retosAsignados = String(user.retosAsignados || '').split(',').map(s => s.trim()).filter(Boolean);
  const challengeId = retosAsignados[pistaNumero - 1];

  if (!challengeId) {
    return { status: 'ERROR', message: 'No se encontró un reto configurado para esta pista.' };
  }

  const retosData = await readSheetRange('RETOS!A2:F11');
  const challengeRow = retosData.find(r => String(r[0]).trim() === String(challengeId).trim());

  if (!challengeRow) {
    return { status: 'ERROR', message: 'Detalles del reto no encontrados.' };
  }

  // Código/Respuesta de la Columna D (Índice 3)
  const expectedAnswer = String(challengeRow[3] || '').trim();

  const isMatch = expectedAnswer.toLowerCase() === respuestaOcodigo.toLowerCase();

  if (!isMatch) {
    return { status: 'ERROR', message: 'Código o respuesta incorrecta. Revisa e inténtalo de nuevo.' };
  }

  let completados = String(user.retosCompletados || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!completados.includes(String(pistaNumero))) {
    completados.push(String(pistaNumero));
    completados.sort((a, b) => Number(a) - Number(b));
    await updateParticipantCompletedChallenges(user.id, completados.join(','));
  }

  let idsCompletados = String(challengeRow[4] || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!idsCompletados.includes(String(userId))) {
    idsCompletados.push(String(userId));
    await updateChallengeCompletedUsers(challengeId, idsCompletados.join(', '));
  }

  return {
    status: 'SUCCESS',
    message: '✨ ¡Respuesta correcta! Has rescatado este fragmento de la partitura.'
  };
}

// 6. Cargar datos de Encuesta
async function getPollData(userId) {
  const optionsRows = await readSheetRange('ENCUESTA_OPCIONES!A2:C');
  const votesRows = await readSheetRange('ENCUESTA_VOTOS!A2:C');

  const options = optionsRows.map(r => ({
    song: String(r[0] || '').trim(),
    imageUrl: String(r[1] || '').trim()
  })).filter(o => o.song !== '');

  const counts = {};
  let userVote = null;

  votesRows.forEach(r => {
    const uId = String(r[0] || '').trim();
    const song = String(r[1] || '').trim();

    if (song) {
      counts[song] = (counts[song] || 0) + 1;
    }
    if (uId === String(userId).trim()) {
      userVote = song;
    }
  });

  return {
    options,
    counts,
    userVote,
    totalVotes: votesRows.length
  };
}

// 7. Registrar Voto de Encuesta
async function submitVote(userId, selectedSong) {
  userId = String(userId).trim();
  selectedSong = String(selectedSong).trim();

  const pollData = await getPollData(userId);
  if (pollData.userVote) {
    return { status: 'ERROR', message: 'Ya has registrado tu voto anteriormente.' };
  }

  const now = new Date().toISOString();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: 'ENCUESTA_VOTOS!A:C',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[userId, selectedSong, now]]
    }
  });

  return { status: 'SUCCESS' };
}

// Handler Vercel Serverless
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const { action, args = [] } = req.body || {};

    if (action === 'getParticipants') {
      const data = await getParticipants();
      return res.status(200).json({ ok: true, data });
    }
    if (action === 'getLogs') {
      const data = await getLogs();
      return res.status(200).json({ ok: true, data });
    }
    if (action === 'registerDance') {
      const data = await registerDance(args[0], args[1], args[2]);
      return res.status(200).json({ ok: true, data });
    }
    if (action === 'getChallengeData') {
      const data = await getChallengeData(args[0]);
      return res.status(200).json({ ok: true, data });
    }
    if (action === 'validateChallenge') {
      const data = await validateChallenge(args[0], args[1], args[2]);
      return res.status(200).json({ ok: true, data });
    }
    if (action === 'getPollData') {
      const data = await getPollData(args[0]);
      return res.status(200).json({ ok: true, data });
    }
    if (action === 'submitVote') {
      const data = await submitVote(args[0], args[1]);
      return res.status(200).json({ ok: true, data });
    }

    return res.status(400).json({ ok: false, error: 'Acción no válida' });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ ok: false, error: err.message });
  }
};
