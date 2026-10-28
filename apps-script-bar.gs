/************************************************************
 *  CAISSE BAR — Webhook Google Apps Script (générique, multi-événements)
 *  À coller dans le classeur de suivi :
 *    Extensions → Apps Script → coller ce code → Enregistrer
 *
 *  Déploiement :
 *    Déployer → Gérer les déploiements → ✏️ → Nouvelle version → Déployer
 *    (garde la même URL /exec)
 *
 *  Chaque ENCAISSEMENT est rangé dans des onglets PROPRES À L'ÉVÉNEMENT
 *  en cours (ex. "Tickets · Java Club"). Changer d'événement depuis
 *  l'admin de l'app crée automatiquement de nouveaux onglets, sans
 *  toucher à ceux des événements précédents.
 *
 *  Par événement, 3 onglets :
 *    - "Tickets · <event>"  : 1 ligne / encaissement
 *    - "Ventes · <event>"   : 1 ligne / article (avec volume L)
 *    - "Live · <event>"      : agrégat recalculé (litres + fûts entamés)
 ************************************************************/

const RESET_KEY = '3216';           // = PIN admin de l'app
const DEFAULT_EVENT = 'Java Club';  // événement par défaut au premier lancement

// Volume d'un fût par bière (L). Le nom doit matcher ce qu'envoie l'app.
// ids = identifiants des produits dans l'app (demi + pinte) : une bière masquée
// dans l'admin disparaît du Live tant qu'elle n'a pas de ventes.
const FUTS = [
  { nom: 'Britt Fresh',          volFut: 30, ids: ['fresh_d', 'fresh_p'] },
  { nom: 'Britt Blanche',        volFut: 20, ids: ['blanche_d', 'blanche_p'] },
  { nom: 'St Erwann Abbaye 7%',  volFut: 20, ids: ['abbaye_d', 'abbaye_p'] },
  { nom: 'IPA Holl',             volFut: 20, ids: ['ipa_d', 'ipa_p'] },
];

const HEAD_TICKETS = ['Horodatage', 'Méthode', 'Montant (€)', 'Device', 'Vendeur'];
const HEAD_VENTES  = ['Horodatage', 'Méthode', 'Produit', 'Format', 'Qté', 'PU (€)', 'Montant (€)', 'Volume (L)', 'Device', 'Vendeur'];

function props() { return PropertiesService.getScriptProperties(); }
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

/* ---- Événement courant (partagé entre tous les appareils) ---- */
function getEvent() { return props().getProperty('event') || DEFAULT_EVENT; }
function setEvent(name) {
  const n = String(name || '').trim() || DEFAULT_EVENT;
  props().setProperty('event', n);
  return n;
}

/* ---- Produits masqués (global, tous événements) ---- */
function getHidden() {
  const raw = props().getProperty('hidden') || '[]';
  try { const a = JSON.parse(raw); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function setHidden(ids) { props().setProperty('hidden', JSON.stringify(ids || [])); }

/* ---- Onglets par événement ---- */
function safeName(s) { return String(s).replace(/[:\\\/\?\*\[\]]/g, ' ').trim().slice(0, 60); }
function tabName(base, ev) { return (base + ' · ' + safeName(ev)).slice(0, 95); }

function getOrCreateTab(name) {
  const s = ss().getSheetByName(name);
  return s ? s : ss().insertSheet(name);
}
function setHeader(sheet, head) {
  sheet.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
}

function tabsFor(ev) {
  const t = getOrCreateTab(tabName('Tickets', ev));
  if (t.getLastRow() === 0) setHeader(t, HEAD_TICKETS);
  const v = getOrCreateTab(tabName('Ventes', ev));
  if (v.getLastRow() === 0) setHeader(v, HEAD_VENTES);
  const l = getOrCreateTab(tabName('Live', ev));
  return { t: t, v: v, l: l };
}

/** Prépare les onglets de l'événement courant + autorise le script. */
function initTabs() {
  const ev = getEvent();
  tabsFor(ev);
  rebuildLive(ev);
  nettoyerOnglets();
}

/** Supprime les onglets inutiles s'ils sont vides : "Feuille 1"/"Sheet1" par défaut
 *  et les anciens App_Tickets / App_Ventes / App_Live. Ne touche jamais un onglet
 *  qui contient des ventes. */
function nettoyerOnglets() {
  const book = ss();
  ['Feuille 1', 'Sheet1', 'App_Tickets', 'App_Ventes', 'App_Live'].forEach(function (name) {
    const sh = book.getSheetByName(name);
    if (!sh || book.getSheets().length <= 1) return;
    const vide = name === 'App_Live' ? !aDesVentesLive(sh) : sh.getLastRow() <= 1;
    if (vide) book.deleteSheet(sh);
  });
}
function aDesVentesLive(sh) {
  if (sh.getLastRow() < 2) return false;
  const litres = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  return litres.some(function (r) { return Number(r[0]) > 0; });
}

function volumePortion(sub) {
  if (!sub) return 0;
  const s = String(sub).toLowerCase();
  if (s.indexOf('demi')  !== -1) return 0.25;
  if (s.indexOf('pinte') !== -1) return 0.50;
  if (s.indexOf('25')    !== -1) return 0.25;
  if (s.indexOf('50')    !== -1) return 0.50;
  return 0;
}
function round2(n) { return Math.round(n * 100) / 100; }

/** Recalcule l'onglet "Live · <event>" (valeurs, pas formules). */
function rebuildLive(ev) {
  ev = ev || getEvent();
  const v = ss().getSheetByName(tabName('Ventes', ev));
  const agg = {};
  FUTS.forEach(function (f) { agg[f.nom] = { litres: 0, demis: 0, pintes: 0, volFut: f.volFut }; });

  if (v && v.getLastRow() > 1) {
    const data = v.getRange(2, 3, v.getLastRow() - 1, 6).getValues(); // C..H
    data.forEach(function (r) {
      const produit = String(r[0]).trim();
      const format  = String(r[1]).toLowerCase();
      const qty     = Number(r[2]) || 0;
      const vol     = Number(r[5]) || 0;
      if (agg[produit]) {
        agg[produit].litres += vol;
        if (format.indexOf('demi') !== -1)  agg[produit].demis  += qty;
        if (format.indexOf('pinte') !== -1) agg[produit].pintes += qty;
      }
    });
  }

  const l = getOrCreateTab(tabName('Live', ev));
  l.clearContents();
  const rows = [['Bière', 'Litres vendus', 'Vol/fût (L)', 'Fûts entamés', 'Nb demis', 'Nb pintes']];
  let totL = 0, totFuts = 0;
  const hidden = getHidden();
  FUTS.forEach(function (f) {
    const a = agg[f.nom];
    const masquee = f.ids.every(function (id) { return hidden.indexOf(id) !== -1; });
    if (masquee && a.litres === 0) return;
    const futs = a.volFut ? a.litres / a.volFut : 0;
    totL += a.litres; totFuts += futs;
    rows.push([f.nom, round2(a.litres), a.volFut, round2(futs), a.demis, a.pintes]);
  });
  rows.push(['TOTAL', round2(totL), '', round2(totFuts), '', '']);
  l.getRange(1, 1, rows.length, 6).setValues(rows);
  l.getRange('1:1').setFontWeight('bold');
  l.getRange(rows.length, 1, 1, 6).setFontWeight('bold');
}

/** Vide tickets/ventes de l'événement courant (garde les en-têtes) + Live à zéro. */
function resetData() {
  const ev = getEvent();
  [tabName('Tickets', ev), tabName('Ventes', ev)].forEach(function (name) {
    const s = ss().getSheetByName(name);
    if (s && s.getLastRow() > 1) {
      s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
    }
  });
  rebuildLive(ev);
}

/** Reçoit un encaissement depuis l'app → rangé dans l'événement courant. */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ts = data.ts ? new Date(data.ts) : new Date();
    const vendor = data.vendor || '';
    const ev = getEvent();
    const tabs = tabsFor(ev);

    tabs.t.appendRow([ts, data.method || '', Number(data.amount) || 0, data.deviceId || '', vendor]);

    (data.items || []).forEach(function (it) {
      const qty = Number(it.qty) || 0;
      const pu = Number(it.price) || 0;
      const vol = volumePortion(it.sub) * qty;
      tabs.v.appendRow([ts, data.method || '', it.name || '', it.sub || '', qty, pu, pu * qty, vol, data.deviceId || '', vendor]);
    });

    rebuildLive(ev);
    return json({ ok: true, event: ev });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === 'event') {
    return json({ event: getEvent() });
  }
  if (p.action === 'setevent') {
    if (String(p.key) !== RESET_KEY) return json({ ok: false, error: 'clé invalide' });
    const n = setEvent(p.name);
    tabsFor(n); rebuildLive(n);
    return json({ ok: true, event: n });
  }
  if (p.action === 'visibility') {
    return json({ hidden: getHidden() });
  }
  if (p.action === 'setvisibility') {
    if (String(p.key) !== RESET_KEY) return json({ ok: false, error: 'clé invalide' });
    const raw = String(p.hidden || '');
    const ids = raw ? raw.split(',').map(function (s) { return s.trim(); }).filter(String) : [];
    setHidden(ids);
    rebuildLive(getEvent());
    return json({ ok: true, hidden: ids });
  }
  if (p.action === 'reset') {
    if (String(p.key) !== RESET_KEY) return json({ ok: false, error: 'clé invalide' });
    resetData();
    return json({ ok: true });
  }
  if (p.action === 'stats') {
    const t = ss().getSheetByName(tabName('Tickets', getEvent()));
    let tickets = 0, cash = 0, card = 0, countTickets = 0, countCash = 0, countCard = 0;
    if (t && t.getLastRow() > 1) {
      const rows = t.getRange(2, 2, t.getLastRow() - 1, 2).getValues(); // Méthode, Montant
      rows.forEach(function (r) {
        const m = String(r[0]).toLowerCase();
        const amt = Number(r[1]) || 0;
        if (m === 'tickets') { tickets += amt; countTickets++; }
        else if (m === 'cash') { cash += amt; countCash++; }
        else if (m === 'card') { card += amt; countCard++; }
      });
    }
    return json({ tickets: tickets, cash: cash, card: card,
                  countTickets: countTickets, countCash: countCash, countCard: countCard, event: getEvent() });
  }
  return ContentService.createTextOutput('Caisse Bar — webhook actif · événement : ' + getEvent());
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
