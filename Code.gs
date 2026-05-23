/**
 * Suivi Heure Caisson - Google Apps Script
 * Analyse les heures de travail en séance caisson à partir d'un Google Sheets.
 */

var SPREADSHEET_ID = '1pMPyXUPnmNJ7D9QekiiDSq_W5860wN6WNVIEmucWhgU';

var POSTES = ['Door Man', 'Observateur', 'Sécurité', 'Speaker', 'Suiveur'];

var AGENT_FICTIF = 'Fin brulage';

var DUREE_MIN_MINUTES = 10;
var DUREE_MAX_MINUTES = 40;

// ─────────────────────────────────────────────
// Point d'entrée de la Web App
// ─────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Suivi Heure Caisson')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────
// Inclusion de fichiers HTML partiels
// ─────────────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─────────────────────────────────────────────
// Lecture et parsing des données brutes
// ─────────────────────────────────────────────
function getRawData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('general')
            || ss.getSheetByName('Réponses au formulaire 1')
            || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) return { headers: [], rows: [] };

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return { headers: headers, rows: data };
}

/**
 * Extrait le nom de l'agent entre crochets dans un en-tête de colonne.
 * Retourne null si le format ne correspond pas.
 */
function extractAgentName(header) {
  var match = String(header).match(/\[([^\]]+)\]/);
  return match ? match[1].trim() : null;
}

/**
 * Parse une date au format français "JJ/MM/AAAA HH:MM:SS" ou objet Date Google.
 * Retourne un objet Date ou null.
 */
function parseDate(value) {
  if (!value) return null;

  // Si c'est déjà un objet Date (Google Sheets peut retourner Date directement)
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  var str = String(value).trim();
  if (!str) return null;

  // Format JJ/MM/AAAA HH:MM:SS
  var m = str.match(/^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/);
  if (m) {
    return new Date(
      parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]),
      parseInt(m[4]), parseInt(m[5]), parseInt(m[6])
    );
  }

  // Tentative générique
  var d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// ─────────────────────────────────────────────
// Calcul principal — appelé depuis le client
// ─────────────────────────────────────────────
/**
 * filtre : 'annee_courante' | '2025' | 'tout'
 * Retourne un objet :
 * {
 *   agents: [ { nom, postes: { DoorMan, Observateur, ... }, total } ],
 *   stats:  { nbSeances, dureeMoyenne, totalCumule }
 * }
 * Toutes les durées sont en minutes.
 */
function getData(filtre) {
  var raw = getRawData();
  var headers = raw.headers;
  var rows    = raw.rows;

  if (!rows || rows.length === 0) {
    return { agents: [], stats: { nbSeances: 0, dureeMoyenne: 0, totalCumule: 0 } };
  }

  // ── Identifier les colonnes agents ──
  // colAgents = [ { index, nom, isFinBrulage } ]
  var colAgents = [];
  for (var c = 1; c < headers.length; c++) {
    var nom = extractAgentName(headers[c]);
    if (nom) {
      colAgents.push({ index: c, nom: nom, isFinBrulage: (nom === AGENT_FICTIF) });
    }
  }

  // ── Construire la liste des séances ──
  // Une séance = { debut: Date, fin: Date, agents: { nom: poste } }
  var seances = [];
  var ligneDebut = null; // { date: Date, agents: { nom: poste } }

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var dateVal = parseDate(row[0]);
    if (!dateVal) continue; // ignorer les lignes sans timestamp

    // Vérifier si c'est une ligne de fin (Fin brulage coché)
    var isFinLigne = false;
    for (var i = 0; i < colAgents.length; i++) {
      if (colAgents[i].isFinBrulage) {
        var val = String(row[colAgents[i].index]).trim();
        if (val && val !== '' && val !== '0' && val !== 'false') {
          isFinLigne = true;
          break;
        }
      }
    }

    if (isFinLigne) {
      // Ligne de fin : clôturer la séance en cours
      if (ligneDebut) {
        seances.push({
          debut:  ligneDebut.date,
          fin:    dateVal,
          agents: ligneDebut.agents
        });
        ligneDebut = null;
      }
      // Si pas de ligneDebut, on ignore ce "Fin brulage" orphelin
    } else {
      // Ligne de début possible
      // Collecter les agents renseignés sur cette ligne
      var agentsLigne = {};
      var nbAgentsReels = 0;
      for (var j = 0; j < colAgents.length; j++) {
        if (colAgents[j].isFinBrulage) continue;
        var poste = String(row[colAgents[j].index]).trim();
        if (poste && poste !== '' && poste !== '0' && poste !== 'false') {
          // Vérifier que c'est un poste valide
          if (POSTES.indexOf(poste) !== -1) {
            agentsLigne[colAgents[j].nom] = poste;
            nbAgentsReels++;
          }
        }
      }

      if (nbAgentsReels > 0) {
        // Si une séance précédente n'était pas fermée, on l'abandonne
        ligneDebut = { date: dateVal, agents: agentsLigne };
      }
    }
  }
  // Note : si ligneDebut reste ouvert (pas de fin), on l'ignore

  // ── Appliquer le filtre de date ──
  var now = new Date();
  var anneeEnCours = now.getFullYear();

  var seancesFiltrees = seances.filter(function(s) {
    if (!s.debut) return false;
    var annee = s.debut.getFullYear();
    if (filtre === 'annee_courante') return annee === anneeEnCours;
    if (filtre === '2025')           return annee === 2025;
    return true; // 'tout'
  });

  // ── Calculer les durées et agréger par agent ──
  var agentsMap = {}; // { nom: { postes: { DoorMan: min, ... }, total: min } }
  var nbSeances = 0;
  var sommeSeances = 0;

  seancesFiltrees.forEach(function(s) {
    var dureeMs  = s.fin.getTime() - s.debut.getTime();
    var dureeMin = dureeMs / 60000;

    // Ignorer les séances < 10 min
    if (dureeMin < DUREE_MIN_MINUTES) return;

    // Plafonner à 40 min
    var dureeComptee = Math.min(dureeMin, DUREE_MAX_MINUTES);

    nbSeances++;
    sommeSeances += dureeComptee;

    Object.keys(s.agents).forEach(function(nom) {
      var poste = s.agents[nom];
      if (!agentsMap[nom]) {
        agentsMap[nom] = { postes: {}, total: 0 };
        POSTES.forEach(function(p) { agentsMap[nom].postes[p] = 0; });
      }
      agentsMap[nom].postes[poste] = (agentsMap[nom].postes[poste] || 0) + dureeComptee;
      agentsMap[nom].total += dureeComptee;
    });
  });

  // ── Construire le tableau résultat trié ──
  var agentsList = Object.keys(agentsMap).sort().map(function(nom) {
    return {
      nom:    nom,
      postes: agentsMap[nom].postes,
      total:  agentsMap[nom].total
    };
  });

  var totalCumule   = agentsList.reduce(function(acc, a) { return acc + a.total; }, 0);
  var dureeMoyenne  = nbSeances > 0 ? sommeSeances / nbSeances : 0;

  return {
    agents: agentsList,
    stats: {
      nbSeances:    nbSeances,
      dureeMoyenne: dureeMoyenne,
      totalCumule:  totalCumule
    }
  };
}
