// ============================================================
// ECO OUEST - WEB APP API v2 (avec photos Drive)
// Déployer : Déployer > Nouveau déploiement > Web App
//   - Exécuter en tant que : Moi
//   - Accès : Tout le monde
// ============================================================

var WS_NAME      = "SUIVI COMMANDES";
var LIGNE_DEBUT  = 3;
var ANNEES_SKIP  = ["2025", "2026"];
var DRIVE_FOLDER = "1RoVY9ILj1THPrcALEQj-1H3bKOc09eN9"; // Dossier ECO OUEST

function doGet(e) {
  var result;
  try {
    result = lireCommandes();
  } catch(err) {
    result = { error: err.message };
  }
  var json = JSON.stringify(result);
  // Support JSONP pour contourner CORS
  var callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data, result;
  try {
    data = JSON.parse(e.postData.contents);
    var action = data.action;
    if      (action === "modifier")      result = modifierLigne(data);
    else if (action === "ajouter")       result = ajouterLigne(data);
    else if (action === "supprimer")     result = supprimerLigne(data);
    else if (action === "uploadPhoto")   result = uploadPhoto(data);
    else result = { error: "Action inconnue" };
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LECTURE ──────────────────────────────────────────────
function lireCommandes() {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WS_NAME);
  var lastRow = ws.getLastRow();
  var data = ws.getRange(LIGNE_DEBUT, 1, lastRow - LIGNE_DEBUT + 1, 20).getValues();
  var rows = [];
  data.forEach(function(row, i) {
    var client = String(row[0] || "").trim();
    if (!client || ANNEES_SKIP.indexOf(client) !== -1) return;
    rows.push({
      ligne:            LIGNE_DEBUT + i,
      client:           String(row[0]  || "").trim(),
      metres:           String(row[1]  || "").trim(),
      devis:            String(row[2]  || "").trim(),
      devisTTC:         String(row[3]  || "").trim(),
      fournisseur:      String(row[4]  || "").trim(),
      commande:         String(row[5]  || "").trim(),
      arc:              String(row[6]  || "").trim(),
      totalAchats:      String(row[7]  || "").trim(),
      commentaireARC:   String(row[8]  || "").trim(),
      prevLivraison:    String(row[9]  || "").trim(),
      repositionnement: String(row[10] || "").trim(),
      statutLivraison:  String(row[11] || "").trim(),
      statutPose:       String(row[12] || "").trim(),
      commentaire:      String(row[13] || "").trim(),
      alerte:           String(row[14] || "").trim(),
      pvChantier:       String(row[15] || "").trim(),
      poseur_ANATOLE:   String(row[16] || "").trim(),
      poseur_ARTIOM:    String(row[17] || "").trim(),
      poseur_ANTONIO:   String(row[18] || "").trim(),
      datePose:         String(row[19] || "").trim(),
      poseur_VINCENT:   String(row[20] || "").trim(),
      poseur_LAMRI:     String(row[21] || "").trim(),
      photo:            String(row[22] || "").trim(),
    });
  });
  return { ok: true, data: rows };
}

// ── MODIFICATION ─────────────────────────────────────────
function modifierLigne(data) {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WS_NAME);
  var ligne = data.ligne;
  var champs = data.champs;
  Object.keys(champs).forEach(function(col) {
    ws.getRange(ligne, parseInt(col)).setValue(champs[col]);
  });
  recalculerAlerte(ws, ligne);
  return { ok: true, ligne: ligne };
}

// ── AJOUT ────────────────────────────────────────────────
function ajouterLigne(data) {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WS_NAME);
  var c = data.commande;
  var targetRow = ws.getLastRow() + 1;
  ws.getRange(targetRow, 1, 1, 15).setValues([[
    c.client, c.metres||'', c.devis, c.devisTTC, c.fournisseur,
    c.commande, c.arc, '', c.commentaireARC,
    c.prevLivraison, '', c.statutLivraison, c.statutPose, c.commentaire, ""
  ]]);
  recalculerAlerte(ws, targetRow);
  return { ok: true, ligne: targetRow };
}

// ── SUPPRESSION ──────────────────────────────────────────
function supprimerLigne(data) {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WS_NAME);
  ws.deleteRow(data.ligne);
  return { ok: true };
}

// ── UPLOAD PHOTO ─────────────────────────────────────────
function uploadPhoto(data) {
  var client    = data.client    || "Inconnu";
  var poseur    = data.poseur    || "Poseur";
  var base64    = data.base64;
  var mimeType  = data.mimeType  || "image/jpeg";
  var extension = mimeType === "image/png" ? ".png" : ".jpg";

  // Trouver ou créer dossier Chantiers dans ECO OUEST
  var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER);

  var chantiersFolder;
  var cf = rootFolder.getFoldersByName("Chantiers");
  chantiersFolder = cf.hasNext() ? cf.next() : rootFolder.createFolder("Chantiers");

  // Trouver ou créer dossier du client
  var clientFolder;
  var clf = chantiersFolder.getFoldersByName(client);
  clientFolder = clf.hasNext() ? clf.next() : chantiersFolder.createFolder(client);

  // Nom du fichier : date + poseur
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  var fileName = dateStr + "_" + poseur + extension;

  // Créer le fichier
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
  var file = clientFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var url = file.getUrl();
  return { ok: true, url: url, fileName: fileName };
}

// ── ALERTE AUTO ──────────────────────────────────────────
function recalculerAlerte(ws, ligne) {
  var row = ws.getRange(ligne, 1, 1, 15).getValues()[0];
  var termines = ["Livré","Terminé","Posé","LIVRE","TERMINE","POSE"];
  var estTermine = termines.indexOf(String(row[11]||"").trim()) !== -1
                || termines.indexOf(String(row[12]||"").trim()) !== -1;
  var estAnnule  = String(row[8]||"").trim().toUpperCase() === "ANNULER";
  var alerteCell = ws.getRange(ligne, 15);
  if (estTermine) { alerteCell.setValue("OK").setBackground("#C8E6C9").setFontColor("#1B5E20"); return; }
  if (estAnnule)  { alerteCell.setValue("Annulée").setBackground("#E0E0E0").setFontColor("#757575"); return; }
  var msgs = [];
  if (!row[5] || String(row[5]).toUpperCase() === "NON") msgs.push("Commande non envoyée - URGENT");
  if (!row[6] || !String(row[6]).trim())  msgs.push("ARC non reçu");
  if (!row[8] || !String(row[8]).trim())  msgs.push("Validation ARC manquante");
  if (!row[9] || !String(row[9]).trim())  msgs.push("Date livraison manquante");
  if (msgs.length === 0) {
    alerteCell.setValue("OK").setBackground("#C8E6C9").setFontColor("#1B5E20");
  } else {
    alerteCell.setValue(msgs.join(" | ")).setBackground("#FF4444").setFontColor("#FFFFFF");
  }
  alerteCell.setFontWeight("bold").setWrap(true);
}
