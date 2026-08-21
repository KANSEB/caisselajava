/************************************************************
 *  CAISSE BAR — LA JAVA · Webhook Google Apps Script
 *  À coller dans le classeur "budget_bar_lajava" :
 *    Extensions → Apps Script → coller ce code → Enregistrer
 *
 *  Puis lancer UNE FOIS la fonction  initTabs()  (menu ▶) pour
 *  créer les onglets et autoriser le script.
 *
 *  Déploiement :
 *    Déployer → Gérer les déploiements → ✏️ → Nouvelle version → Déployer
 *    (garde la même URL /exec ; nécessaire pour que doPost utilise ce code)
 *
 *  L'app POST chaque encaissement ; le script :
 *    - journalise le ticket dans l'onglet "App_Tickets"
 *    - journalise chaque article dans "App_Ventes" (avec volume L)
 *    - recalcule l'onglet "App_Live" : litres + fûts entamés / bière
 *
 *  NB : App_Live est rempli avec des VALEURS (pas des formules) pour
 *  éviter tout souci de séparateur , / ; selon la locale du classeur.
 ************************************************************/

const TAB_TICKETS = 'App_Tickets';
const TAB_VENTES  = 'App_Ventes';
const TAB_LIVE    = 'App_Live';

// Clé exigée pour la réinitialisation (= PIN admin de l'app).
const RESET_KEY = '3216';

// Volume d'un fût par bière (L). Le nom doit matcher ce qu'envoie l'app.
const FUTS = [
  { nom: 'Britt Fresh',          volFut: 30 },
  { nom: 'Britt Blanche',        volFut: 20 },
  { nom: 'St Erwann Abbaye 7%',  volFut: 20 },
  { nom: 'St Erwann IPA 7%',     volFut: 20 },
];

// Volume d'une portion (L) selon le sous-titre du produit.
function volumePortion(sub) {
  if (!sub) return 0;
  const s = String(sub).toLowerCase();
  if (s.indexOf('demi')  !== -1) return 0.25;
  if (s.indexOf('pinte') !== -1) return 0.50;
  if (s.indexOf('25')    !== -1) return 0.25;
  if (s.indexOf('50')    !== -1) return 0.50;
  return 0;
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getOrCreateTab(name) {
  const s = ss().getSheetByName(name);
  return s ? s : ss().insertSheet(name);
}

/** À lancer une fois à la main pour préparer les onglets. */
function initTabs() {
  const t = getOrCreateTab(TAB_TICKETS);
  if (t.getLastRow() === 0) {
    t.appendRow(['Horodatage', 'Méthode', 'Montant (€)', 'Device']);
    t.getRange('1:1').setFontWeight('bold');
  }
  const v = getOrCreateTab(TAB_VENTES);
  if (v.getLastRow() === 0) {
    v.appendRow(['Horodatage', 'Méthode', 'Produit', 'Format', 'Qté', 'PU (€)', 'Montant (€)', 'Volume (L)', 'Device']);
    v.getRange('1:1').setFontWeight('bold');
  }
  getOrCreateTab(TAB_LIVE);
  rebuildLive();
}

/** Recalcule l'onglet App_Live à partir de App_Ventes (valeurs, pas formules). */
function rebuildLive() {
  const v = ss().getSheetByName(TAB_VENTES);
  // Init des agrégats par bière
  const agg = {};
  FUTS.forEach(function (f) { agg[f.nom] = { litres: 0, demis: 0, pintes: 0, volFut: f.volFut }; });

  if (v && v.getLastRow() > 1) {
    // Colonnes : C=Produit(3), D=Format(4), E=Qté(5), H=Volume(8)
    const data = v.getRange(2, 3, v.getLastRow() - 1, 6).getValues(); // C..H
    data.forEach(function (r) {
      const produit = String(r[0]).trim();       // C
      const format  = String(r[1]).toLowerCase(); // D
      const qty     = Number(r[2]) || 0;          // E
      const vol     = Number(r[5]) || 0;          // H
      if (agg[produit]) {
        agg[produit].litres += vol;
        if (format.indexOf('demi') !== -1)  agg[produit].demis  += qty;
        if (format.indexOf('pinte') !== -1) agg[produit].pintes += qty;
      }
    });
  }

  const l = getOrCreateTab(TAB_LIVE);
  l.clearContents();
  const rows = [['Bière', 'Litres vendus', 'Vol/fût (L)', 'Fûts entamés', 'Nb demis', 'Nb pintes']];
  let totL = 0, totFuts = 0;
  FUTS.forEach(function (f) {
    const a = agg[f.nom];
    const futs = a.volFut ? a.litres / a.volFut : 0;
    totL += a.litres; totFuts += futs;
    rows.push([f.nom, round2(a.litres), a.volFut, round2(futs), a.demis, a.pintes]);
  });
  rows.push(['TOTAL', round2(totL), '', round2(totFuts), '', '']);
  l.getRange(1, 1, rows.length, 6).setValues(rows);
  l.getRange('1:1').setFontWeight('bold');
  l.getRange(rows.length, 1, 1, 6).setFontWeight('bold');
}

function round2(n) { return Math.round(n * 100) / 100; }

/** Vide les ventes/tickets (garde les en-têtes) et remet App_Live à zéro. */
function resetData() {
  [TAB_TICKETS, TAB_VENTES].forEach(function (name) {
    const s = ss().getSheetByName(name);
    if (s && s.getLastRow() > 1) {
      s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
    }
  });
  rebuildLive();
}

/** Reçoit un encaissement depuis l'app. */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ts = data.ts ? new Date(data.ts) : new Date();

    const tTickets = getOrCreateTab(TAB_TICKETS);
    if (tTickets.getLastRow() === 0) initTabs();
    tTickets.appendRow([ts, data.method || '', Number(data.amount) || 0, data.deviceId || '']);

    const tVentes = getOrCreateTab(TAB_VENTES);
    (data.items || []).forEach(function (it) {
      const qty = Number(it.qty) || 0;
      const pu = Number(it.price) || 0;
      const vol = volumePortion(it.sub) * qty;
      tVentes.appendRow([ts, data.method || '', it.name || '', it.sub || '', qty, pu, pu * qty, vol, data.deviceId || '']);
    });

    rebuildLive();
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** Stats globales pour le panneau admin de l'app. */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'reset') {
    if (String(e.parameter.key) !== RESET_KEY) {
      return json({ ok: false, error: 'clé invalide' });
    }
    resetData();
    return json({ ok: true });
  }
  if (e && e.parameter && e.parameter.action === 'stats') {
    const t = ss().getSheetByName(TAB_TICKETS);
    let tickets = 0, card = 0, countTickets = 0, countCard = 0;
    if (t && t.getLastRow() > 1) {
      const rows = t.getRange(2, 2, t.getLastRow() - 1, 2).getValues(); // Méthode, Montant
      rows.forEach(function (r) {
        const m = String(r[0]).toLowerCase();
        const amt = Number(r[1]) || 0;
        if (m === 'tickets') { tickets += amt; countTickets++; }
        else if (m === 'card') { card += amt; countCard++; }
      });
    }
    return json({ tickets: tickets, card: card, countTickets: countTickets, countCard: countCard });
  }
  return ContentService.createTextOutput('Caisse Bar La Java — webhook actif');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
