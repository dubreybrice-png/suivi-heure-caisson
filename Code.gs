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
// Gestion de la visibilité des agents
// (stored in Script Properties as JSON array)
// ─────────────────────────────────────────────

function getHiddenAgents() {
  var val = PropertiesService.getScriptProperties().getProperty('hiddenAgents');
  return val ? JSON.parse(val) : [];
}

function saveHiddenAgents(list) {
  PropertiesService.getScriptProperties().setProperty('hiddenAgents', JSON.stringify(list));
}

/**
 * Retourne la liste de tous les agents réels (hors agent fictif)
 * détectés dans les en-têtes du Spreadsheet, avec leur statut masqué.
 * [ { nom, hidden } ]
 */
function getAgentsList() {
  var raw = getRawData();
  var hidden = getHiddenAgents();
  var agents = [];
  var seen = {};
  raw.headers.forEach(function(h) {
    var nom = extractAgentName(h);
    if (nom && nom !== AGENT_FICTIF && !seen[nom]) {
      seen[nom] = true;
      agents.push({ nom: nom, hidden: hidden.indexOf(nom) !== -1 });
    }
  });
  agents.sort(function(a, b) { return a.nom.localeCompare(b.nom); });
  return agents;
}

/**
 * Bascule la visibilité d'un agent.
 * Retourne la liste mise à jour.
 */
function toggleAgent(nom) {
  var hidden = getHiddenAgents();
  var idx = hidden.indexOf(nom);
  if (idx === -1) {
    hidden.push(nom);
  } else {
    hidden.splice(idx, 1);
  }
  saveHiddenAgents(hidden);
  return getAgentsList();
}

/**
 * Ajoute un nouvel agent :
 * 1. Ajoute une colonne dans le Spreadsheet (en-tête formaté).
 * 2. Ajoute une question déroulante dans le Google Form lié.
 * Retourne { ok, message, agents }.
 */
function addAgent(nom) {
  nom = nom.trim();
  if (!nom) return { ok: false, message: 'Nom vide.' };
  if (nom === AGENT_FICTIF) return { ok: false, message: 'Nom réservé.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('general')
            || ss.getSheetByName('Réponses au formulaire 1')
            || ss.getSheets()[0];

  // Vérifier que l'agent n'existe pas déjà
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var existing = extractAgentName(headers[i]);
    if (existing && existing.toLowerCase() === nom.toLowerCase()) {
      return { ok: false, message: 'Cet agent existe déjà.' };
    }
  }

  // ── 1. Ajouter la colonne dans le Spreadsheet ──
  var newCol = sheet.getLastColumn() + 1;
  var headerLabel = 'Quels sont les agents concernés [' + nom + ']';
  sheet.getRange(1, newCol).setValue(headerLabel);

  // ── 2. Ajouter la question dans le Google Form lié ──
  var formUrl = ss.getFormUrl();
  if (formUrl) {
    try {
      var form = FormApp.openByUrl(formUrl);
      var item = form.addListItem();
      item.setTitle(headerLabel);
      item.setRequired(false);
      // Options = postes de travail + une option vide
      var choices = [''].concat(POSTES).map(function(p) {
        return item.createChoice(p);
      });
      item.setChoices(choices);
    } catch(e) {
      // Le form n'est pas accessible : on continue quand même
      return { ok: true, message: 'Agent ajouté dans le Spreadsheet. Impossible d\'accéder au formulaire : ' + e.message, agents: getAgentsList() };
    }
  }

  return { ok: true, message: 'Agent "' + nom + '" ajouté avec succès.', agents: getAgentsList() };
}

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
  var hidden = getHiddenAgents();
  var agentsList = Object.keys(agentsMap)
    .filter(function(nom) { return hidden.indexOf(nom) === -1; })
    .sort().map(function(nom) {
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
