/**
 * English Learning Test — Centre Focus
 * Réception des résultats du test et écriture dans Google Sheets.
 *
 * Installation pas à pas : voir README.ar.md (même dossier).
 * Le script est lié au classeur : SpreadsheetApp.getActiveSpreadsheet()
 * renvoie donc directement la feuille de résultats.
 */

var SHEET_NAME = 'Résultats';
var TOTAL_QUESTIONS = 10;

var HEADERS = [
  'ID', 'Date', 'Nom', 'Téléphone', 'Score', 'Total', 'Pourcentage',
  'Niveau', 'Durée', 'Envoi auto', 'Langue'
];

var COL_DATE = 2;
var COL_PHONE = 4;

/* ------------------------------------------------------------------ *
 * Point d'entrée : la page envoie un POST JSON à l'URL /exec.
 * ------------------------------------------------------------------ */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    // Deux candidats terminent en même temps : la page réessaiera.
    return reply({ status: 'busy' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return reply({ status: 'error', message: 'corps de requête vide' });
    }

    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet();

    // La page garde une file d'attente et réessaie tant qu'elle n'a pas de
    // confirmation : sans ce contrôle, un même test créerait plusieurs lignes.
    if (data.id && rowExists(sheet, data.id)) {
      return reply({ status: 'duplicate', id: data.id });
    }

    sheet.appendRow(buildRow(data));
    return reply({ status: 'ok', id: data.id || '' });

  } catch (err) {
    return reply({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* Ouvrir l'URL /exec dans un navigateur doit afficher ce message :
 * c'est le moyen le plus simple de vérifier que le déploiement est actif. */
function doGet() {
  return reply({ status: 'ok', service: 'English Learning Test' });
}

/* ------------------------------------------------------------------ *
 * Feuille de calcul
 * ------------------------------------------------------------------ */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    var headers = HEADERS.slice();
    for (var i = 1; i <= TOTAL_QUESTIONS; i++) {
      headers.push('Q' + i);
    }
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#E9EFFD');
    sheet.setFrozenRows(1);

    // Sans format texte, Sheets lit « 0612345678 » comme un nombre et
    // supprime le zéro initial du numéro de téléphone.
    sheet.getRange(1, COL_PHONE, sheet.getMaxRows(), 1).setNumberFormat('@');
    sheet.getRange(1, COL_DATE, sheet.getMaxRows(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
  }

  return sheet;
}

function rowExists(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return true;
  }
  return false;
}

function buildRow(data) {
  var answers = Array.isArray(data.answers) ? data.answers : [];

  var row = [
    data.id || '',
    data.submittedAt ? new Date(data.submittedAt) : new Date(),
    data.name || '',
    data.phone || '',
    Number(data.score) || 0,
    Number(data.total) || TOTAL_QUESTIONS,
    Number(data.percent) || 0,
    data.level || '',
    formatDuration(data.durationSeconds),
    data.autoSubmitted ? 'oui' : 'non',
    data.lang || ''
  ];

  // Une colonne par question : la réponse donnée, suivie de ✓ ou ✗.
  for (var i = 0; i < TOTAL_QUESTIONS; i++) {
    var a = answers[i];
    if (!a) { row.push(''); continue; }
    var chosen = (a.chosen === null || a.chosen === undefined) ? '—' : String(a.chosen);
    row.push(chosen + (a.isCorrect ? ' ✓' : ' ✗'));
  }

  return row;
}

function formatDuration(seconds) {
  var s = Math.max(0, Math.round(Number(seconds) || 0));
  var m = Math.floor(s / 60);
  var r = s % 60;
  return (m < 10 ? '0' + m : m) + ':' + (r < 10 ? '0' + r : r);
}

function reply(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
