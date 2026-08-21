/************************************************************
 *  CAISSE BAR — LA JAVA · Webhook Google Apps Script
 *  À coller dans le classeur "budget_bar_lajava" :
 *    Extensions → Apps Script → coller ce code → Enregistrer
 *
 *  Puis lancer UNE FOIS la fonction  initTabs()  (menu ▶) pour
 *  créer les onglets et autoriser le script.
 *
 *  Déploiement :
 *    Déployer → Nouveau déploiement → Application Web
 *      · Exécuter en tant que : Moi
 *      · Qui a accès : Tout le monde
 *    Copier l'URL /exec → la coller dans index.html (SHEETS_WEBHOOK_URL)
 *
 *  L'app POST chaque encaissement ; le script :
 *    - journalise le ticket dans l'onglet "App_Tickets"
 *    - journalise chaque article dans "App_Ventes" (avec volume L)
 *    - l'onglet "App_Live" agrège en direct : litres + fûts entamés/bière
 ************************************************************/

const TAB_TICKETS = 'App_Tickets';
const TAB_VENTES  = 'App_Ventes';
const TAB_LIVE    = 'App_Live';

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
  // Tickets
  const t = getOrCreateTab(TAB_TICKETS);
  if (t.getLastRow() === 0) {
    t.appendRow(['Horodatage', 'Méthode', 'Montant (€)', 'Device']);
    t.getRange('1:1').setFontWeight('bold');
  }
  // Ventes (une ligne par article)
  const v = getOrCreateTab(TAB_VENTES);
  if (v.getLastRow() === 0) {
    v.appendRow(['Horodatage', 'Méthode', 'Produit', 'Format', 'Qté', 'PU (€)', 'Montant (€)', 'Volume (L)', 'Device']);
    v.getRange('1:1').setFontWeight('bold');
  }
  // Live (agrégats par bière, en formules)
  const l = getOrCreateTab(TAB_LIVE);
  l.clearContents();
  l.getRange('A1:F1').setValues([['Bière', 'Litres vendus', 'Vol/fût (L)', 'Fûts entamés', 'Nb demis', 'Nb pintes']]);
  l.getRange('A1:F1').setFontWeight('bold');
  FUTS.forEach(function (f, i) {
    const r = i + 2;
    l.getRange(r, 1).setValue(f.nom);
    l.getRange(r, 2).setFormula('=SUMIF(' + TAB_VENTES + '!C:C, A' + r + ', ' + TAB_VENTES + '!H:H)');
    l.getRange(r, 3).setValue(f.volFut);
    l.getRange(r, 4).setFormula('=IF(C' + r + '=0,0,B' + r + '/C' + r + ')');
    l.getRange(r, 5).setFormula('=SUMIFS(' + TAB_VENTES + '!E:E, ' + TAB_VENTES + '!C:C, A' + r + ', ' + TAB_VENTES + '!D:D, "Demi")');
    l.getRange(r, 6).setFormula('=SUMIFS(' + TAB_VENTES + '!E:E, ' + TAB_VENTES + '!C:C, A' + r + ', ' + TAB_VENTES + '!D:D, "Pinte")');
  });
  const total = FUTS.length + 2;
  l.getRange(total, 1).setValue('TOTAL');
  l.getRange(total, 2).setFormula('=SUM(B2:B' + (total - 1) + ')');
  l.getRange(total, 4).setFormula('=SUM(D2:D' + (total - 1) + ')');
  l.getRange(total, 1, 1, 6).setFontWeight('bold');
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

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** Stats globales pour le panneau admin de l'app. */
function doGet(e) {
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
