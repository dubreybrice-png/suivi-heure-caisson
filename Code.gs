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

function getAgentsOrder() {
  var val = PropertiesService.getScriptProperties().getProperty('agentsOrder');
  return val ? JSON.parse(val) : [];
}

function saveAgentsOrder(list) {
  PropertiesService.getScriptProperties().setProperty('agentsOrder', JSON.stringify(list));
}

/**
 * Retourne la liste de tous les agents réels (hors agent fictif),
 * dans l'ordre personnalisé stocké, avec leur statut masqué.
 * [ { nom, hidden } ]
 */
function getAgentsList() {
  var raw = getRawData();
  var hidden = getHiddenAgents();
  var order  = getAgentsOrder();

  // Collecter tous les agents depuis les en-têtes
  var seen = {};
  raw.headers.forEach(function(h) {
    var nom = extractAgentName(h);
    if (nom && nom !== AGENT_FICTIF) seen[nom] = true;
  });
  var allNoms = Object.keys(seen);

  // Construire la liste ordonnée :
  // d'abord ceux présents dans order (dans l'ordre), puis les nouveaux triés
  var ordered = [];
  order.forEach(function(n) { if (seen[n]) ordered.push(n); });
  allNoms.sort().forEach(function(n) {
    if (ordered.indexOf(n) === -1) ordered.push(n);
  });

  // Sauvegarder si l'ordre a évolué
  if (JSON.stringify(ordered) !== JSON.stringify(order)) saveAgentsOrder(ordered);

  return ordered.map(function(n) {
    return { nom: n, hidden: hidden.indexOf(n) !== -1 };
  });
}

/**
 * Déplace un agent vers le haut ou le bas dans l'ordre personnalisé
 * ET dans le Google Form lié.
 * direction : 'up' | 'down'
 * Retourne la liste mise à jour.
 */
function moveAgent(nom, direction) {
  var order = getAgentsOrder();
  // S'assurer que l'ordre est à jour
  var agents = getAgentsList();
  order = agents.map(function(a) { return a.nom; });

  var idx = order.indexOf(nom);
  if (idx === -1) return getAgentsList();

  var newIdx = direction === 'up' ? idx - 1 : idx + 1;
  if (newIdx < 0 || newIdx >= order.length) return getAgentsList();

  // Échanger dans l'ordre local
  var tmp = order[newIdx];
  order[newIdx] = order[idx];
  order[idx] = tmp;
  saveAgentsOrder(order);

  // Déplacer la question dans le formulaire
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var formUrl = ss.getFormUrl();
  if (formUrl) {
    try {
      var form = FormApp.openByUrl(formUrl);
      var items = form.getItems();
      var headerLabel = 'Quels sont les agents concernés [' + nom + ']';
      var targetItem = null;
      var targetFormIdx = -1;
      items.forEach(function(it, i) {
        if (it.getTitle() === headerLabel) { targetItem = it; targetFormIdx = i; }
      });
      if (targetItem !== null) {
        var newFormIdx = targetFormIdx + (direction === 'up' ? -1 : 1);
        if (newFormIdx >= 0 && newFormIdx < items.length) {
          form.moveItem(targetFormIdx, newFormIdx);
        }
      }
    } catch(e) {
      // Formulaire inaccessible : on continue
    }
  }

  return getAgentsList();
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
 * 2. Ajoute une LIGNE dans la grille du Google Form lié.
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

  // Vérifier que l'agent n'existe pas déjà dans les en-têtes
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var existing = extractAgentName(headers[i]);
    if (existing && existing.toLowerCase() === nom.toLowerCase()) {
      return { ok: false, message: 'Cet agent existe déjà.' };
    }
  }

  // ── 1. Ajouter la colonne dans le Spreadsheet ──
  // Le format d'en-tête dans le Spreadsheet est "Quels sont les agents concernés [NOM]"
  // mais dans la grille du Forms, la ligne correspond juste au nom de l'agent.
  // Google Forms export dans le Spreadsheet : "Quels sont les agents concernés [NOM]"
  var newCol = sheet.getLastColumn() + 1;
  var headerLabel = 'Quels sont les agents concernés [' + nom + ']';
  sheet.getRange(1, newCol).setValue(headerLabel);

  // ── 2. Ajouter une ligne dans la grille du Google Form ──
  var formUrl = ss.getFormUrl();
  if (formUrl) {
    try {
      var form = FormApp.openByUrl(formUrl);
      var items = form.getItems(FormApp.ItemType.GRID);
      if (items.length === 0) {
        return { ok: true, message: 'Agent ajouté dans le Spreadsheet. Aucune grille trouvée dans le formulaire.', agents: getAgentsList() };
      }
      var grid = items[0].asGridItem();
      var rows = grid.getRows();
      rows.push(nom);
      grid.setRows(rows);
    } catch(e) {
      return { ok: true, message: 'Agent ajouté dans le Spreadsheet. Erreur formulaire : ' + e.message, agents: getAgentsList() };
    }
  }

  return { ok: true, message: 'Agent "' + nom + '" ajouté avec succès.', agents: getAgentsList() };
}

/**
 * Déplace un agent vers le haut ou le bas dans la grille du formulaire
 * ET dans l'ordre personnalisé local.
 * direction : 'up' | 'down'
 */
function moveAgent(nom, direction) {
  // Mettre à jour l'ordre local
  var agents = getAgentsList();
  var order = agents.map(function(a) { return a.nom; });
  var idx = order.indexOf(nom);
  if (idx === -1) return getAgentsList();
  var newIdx = direction === 'up' ? idx - 1 : idx + 1;
  if (newIdx < 0 || newIdx >= order.length) return getAgentsList();
  var tmp = order[newIdx]; order[newIdx] = order[idx]; order[idx] = tmp;
  saveAgentsOrder(order);

  // Déplacer la ligne dans la grille du formulaire
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var formUrl = ss.getFormUrl();
  if (formUrl) {
    try {
      var form = FormApp.openByUrl(formUrl);
      var items = form.getItems(FormApp.ItemType.GRID);
      if (items.length > 0) {
        var grid = items[0].asGridItem();
        var rows = grid.getRows();
        var rowIdx = rows.indexOf(nom);
        if (rowIdx !== -1) {
          var newRowIdx = rowIdx + (direction === 'up' ? -1 : 1);
          if (newRowIdx >= 0 && newRowIdx < rows.length) {
            var t = rows[newRowIdx]; rows[newRowIdx] = rows[rowIdx]; rows[rowIdx] = t;
            grid.setRows(rows);
          }
        }
      }
    } catch(e) {
      // Formulaire inaccessible : ordre local sauvegardé quand même
    }
  }

  return getAgentsList();
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
  var seancesFiltrees = seances.filter(function(s) {
    if (!s.debut) return false;
    var annee = s.debut.getFullYear();
    if (filtre === 'annee_courante') return annee === new Date().getFullYear();
    var anneeFiltre = parseInt(filtre);
    if (!isNaN(anneeFiltre)) return annee === anneeFiltre;
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
