/**
 * CAISSE MERCH — LA JAVA
 * Script Google Apps Script à coller dans la Google Sheet "Merch La Java".
 *
 * ── INSTALLATION ────────────────────────────────────────────────────────────
 * 1. Ouvre ta Google Sheet merch → menu Extensions → Apps Script
 * 2. Supprime le code existant, colle TOUT ce fichier, sauvegarde
 * 3. Lance une fois la fonction  setup()  (menu déroulant en haut → setup → ▶)
 *    Elle crée les 2 onglets avec les bons en-têtes.
 * 4. Déployer → Nouveau déploiement → type "Application Web"
 *      - Exécuter en tant que : Moi
 *      - Qui a accès          : Tout le monde
 * 5. Copie l'URL /exec et colle-la dans merch.html → SHEETS_WEBHOOK_URL
 *
 * ⚠ À chaque modification de ce script : Déployer → Gérer les déploiements →
 *   ✏ modifier → Version "Nouvelle version" → Déployer. (Sinon l'URL sert
 *   toujours l'ancien code.)
 *
 * ── COMMENT ÇA MARCHE ───────────────────────────────────────────────────────
 * Onglet INVENTAIRE = la carte ET le stock. C'est LE doc de référence.
 *   Colonnes A→F remplies à la main (après l'inventaire papier).
 *   Colonnes G→L recalculées automatiquement à chaque encaissement.
 * Onglet VENTES = le journal, une ligne par encaissement.
 */

var SH_INV   = 'Inventaire';
var SH_SALES = 'Ventes';

var INV_HEADERS = [
  'ID', 'Produit', 'Catégorie', 'Taille', 'Prix', 'Stock initial',
  'Vendus', 'Offerts', 'Tarif préf.', 'Total sorti', 'Stock restant', 'CA généré'
];
var SALES_HEADERS = [
  'Horodatage', 'Mode', 'Règlement', 'Montant encaissé', 'Prix plein',
  'Remise', 'Dons libres', 'Articles', 'Appareil'
];

/* Colonnes Inventaire (1-indexé) */
var C_ID = 1, C_NAME = 2, C_CAT = 3, C_SIZE = 4, C_PRICE = 5, C_INIT = 6,
    C_SOLD = 7, C_GIFT = 8, C_PREF = 9, C_OUT = 10, C_REST = 11, C_CA = 12;

/* ────────────────────────────────────────────────────────────────────────── */
/* SETUP                                                                      */
/* ────────────────────────────────────────────────────────────────────────── */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inv = ss.getSheetByName(SH_INV) || ss.insertSheet(SH_INV);
  if (inv.getLastRow() === 0) {
    inv.getRange(1, 1, 1, INV_HEADERS.length).setValues([INV_HEADERS]);
    // Quelques lignes d'exemple — remplace-les par ta vraie carte merch
    inv.getRange(2, 1, 4, 6).setValues([
      ['tshirt', 'T-shirt', 'textile', 'S',  20, 0],
      ['tshirt', 'T-shirt', 'textile', 'M',  20, 0],
      ['tote',   'Tote bag', 'access', '',   10, 0],
      ['sticker','Sticker',  'print',  '',    2, 0]
    ]);
  }
  inv.getRange(1, 1, 1, INV_HEADERS.length).setFontWeight('bold').setBackground('#f0f0f0');
  inv.setFrozenRows(1);

  var sales = ss.getSheetByName(SH_SALES) || ss.insertSheet(SH_SALES);
  if (sales.getLastRow() === 0) {
    sales.getRange(1, 1, 1, SALES_HEADERS.length).setValues([SALES_HEADERS]);
  }
  sales.getRange(1, 1, 1, SALES_HEADERS.length).setFontWeight('bold').setBackground('#f0f0f0');
  sales.setFrozenRows(1);

  recomputeInventory();
  SpreadsheetApp.getUi().alert('Setup terminé.\n\nRemplis maintenant l\'onglet "Inventaire" colonnes A→F avec ta carte merch et les stocks de départ.');
}

/* ────────────────────────────────────────────────────────────────────────── */
/* RECEPTION D'UNE VENTE (POST depuis merch.html)                             */
/* ────────────────────────────────────────────────────────────────────────── */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return json({ ok: false, error: 'lock' }); }

  try {
    var sale = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sales = ss.getSheetByName(SH_SALES) || ss.insertSheet(SH_SALES);
    if (sales.getLastRow() === 0) sales.getRange(1, 1, 1, SALES_HEADERS.length).setValues([SALES_HEADERS]);

    var items = sale.items || [];
    var donAmount = 0, articles = [];
    items.forEach(function (it) {
      if (it.type === 'don') { donAmount += num(it.price) * num(it.qty); articles.push('Don ' + money(num(it.price) * num(it.qty))); }
      else articles.push(num(it.qty) + '× ' + it.name + (it.size ? ' ' + it.size : ''));
    });

    sales.appendRow([
      new Date(sale.ts || new Date()),
      labelMode(sale.method),
      labelSettle(sale.settle, sale.method),
      num(sale.amount),
      num(sale.fullPrice),
      num(sale.discount),
      donAmount,
      articles.join(', '),
      sale.deviceId || ''
    ]);

    applyToInventory(sale);
    return json({ ok: true });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Met à jour l'onglet Inventaire : incrémente Vendus / Offerts / Tarif préf.
 * et le CA attribué à chaque ligne, puis recalcule Total sorti et Stock restant.
 * Une ligne inconnue est créée automatiquement (stock initial 0, à corriger à la main).
 */
function applyToInventory(sale) {
  var items = (sale.items || []).filter(function (it) { return it.type !== 'don'; });
  if (!items.length) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inv = ss.getSheetByName(SH_INV) || ss.insertSheet(SH_INV);
  if (inv.getLastRow() === 0) inv.getRange(1, 1, 1, INV_HEADERS.length).setValues([INV_HEADERS]);

  var last = inv.getLastRow();
  var data = last > 1 ? inv.getRange(2, 1, last - 1, INV_HEADERS.length).getValues() : [];

  var index = {};
  for (var i = 0; i < data.length; i++) {
    index[key(data[i][C_ID - 1], data[i][C_SIZE - 1])] = i + 2;  // n° de ligne
  }

  /* Répartition du montant encaissé au prorata du prix plein de chaque ligne */
  var basketFull = 0;
  items.forEach(function (it) { basketFull += num(it.price) * num(it.qty); });
  var cashable = num(sale.amount) - donsIn(sale);   // part hors don libre
  if (cashable < 0) cashable = 0;

  items.forEach(function (it) {
    var k = key(it.id, it.size);
    var r = index[k];

    if (!r) {   // produit inconnu de l'inventaire → on crée la ligne
      inv.appendRow([it.id, it.name, '', it.size || '', num(it.price), 0, 0, 0, 0, 0, 0, 0]);
      r = inv.getLastRow();
      index[k] = r;
    }

    var qty  = num(it.qty);
    var lineFull = num(it.price) * qty;
    var lineCA = basketFull > 0 ? cashable * (lineFull / basketFull) : 0;

    var col = sale.method === 'gift' ? C_GIFT : (sale.method === 'pref' ? C_PREF : C_SOLD);
    bump(inv, r, col, qty);
    bump(inv, r, C_CA, Math.round(lineCA * 100) / 100);
    refreshRow(inv, r);
  });
}

function bump(sheet, row, col, delta) {
  var cell = sheet.getRange(row, col);
  cell.setValue(num(cell.getValue()) + delta);
}

/** Recalcule Total sorti / Stock restant d'une ligne */
function refreshRow(inv, r) {
  var v = inv.getRange(r, 1, 1, INV_HEADERS.length).getValues()[0];
  var out = num(v[C_SOLD - 1]) + num(v[C_GIFT - 1]) + num(v[C_PREF - 1]);
  inv.getRange(r, C_OUT).setValue(out);
  inv.getRange(r, C_REST).setValue(num(v[C_INIT - 1]) - out);
}

/** Recalcule tout l'onglet Inventaire (utile après une correction manuelle) */
function recomputeInventory() {
  var inv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_INV);
  if (!inv || inv.getLastRow() < 2) return;
  for (var r = 2; r <= inv.getLastRow(); r++) refreshRow(inv, r);
}

/* ────────────────────────────────────────────────────────────────────────── */
/* LECTURE (GET depuis merch.html)                                            */
/* ────────────────────────────────────────────────────────────────────────── */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'stats';
  if (action === 'inventory') return json(getInventory());
  return json(getStats());
}

/** Carte + stock : ce que l'app charge au démarrage */
function getInventory() {
  var inv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_INV);
  if (!inv || inv.getLastRow() < 2) return { rows: [] };

  var data = inv.getRange(2, 1, inv.getLastRow() - 1, INV_HEADERS.length).getValues();
  var rows = [];
  data.forEach(function (v) {
    if (!v[C_ID - 1]) return;
    var sold = num(v[C_SOLD - 1]), gift = num(v[C_GIFT - 1]), pref = num(v[C_PREF - 1]);
    rows.push({
      id:      String(v[C_ID - 1]).trim(),
      name:    String(v[C_NAME - 1]).trim(),
      cat:     String(v[C_CAT - 1] || 'autre').trim().toLowerCase(),
      size:    String(v[C_SIZE - 1] || '').trim(),
      price:   num(v[C_PRICE - 1]),
      initial: num(v[C_INIT - 1]),
      sold:    sold,
      gifted:  gift,
      pref:    pref,
      out:     sold + gift + pref
    });
  });
  return { rows: rows };
}

/** Totaux d'encaissement, toutes caisses confondues */
function getStats() {
  var sales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SALES);
  var s = {
    cash: 0, card: 0, countCash: 0, countCard: 0,
    countPref: 0, discount: 0,
    countGift: 0, giftedUnits: 0, giftedValue: 0,
    countDon: 0, donAmount: 0
  };
  if (!sales || sales.getLastRow() < 2) return s;

  var data = sales.getRange(2, 1, sales.getLastRow() - 1, SALES_HEADERS.length).getValues();
  data.forEach(function (v) {
    var mode   = String(v[1] || '');
    var settle = String(v[2] || '');
    var amount = num(v[3]), full = num(v[4]), disc = num(v[5]), don = num(v[6]);
    var arts   = String(v[7] || '');

    if (settle.indexOf('spèce') >= 0)   { s.cash += amount; s.countCash++; }
    else if (settle.indexOf('arte') >= 0) { s.card += amount; s.countCard++; }

    if (mode.indexOf('référ') >= 0) { s.countPref++; s.discount += disc; }
    if (mode.indexOf('ffert') >= 0) { s.countGift++; s.giftedValue += full; s.giftedUnits += countUnits(arts); }
    if (don > 0) { s.countDon++; s.donAmount += don; }
  });

  ['cash', 'card', 'discount', 'giftedValue', 'donAmount'].forEach(function (k) {
    s[k] = Math.round(s[k] * 100) / 100;
  });
  return s;
}

/* ────────────────────────────────────────────────────────────────────────── */
/* Helpers                                                                    */
/* ────────────────────────────────────────────────────────────────────────── */
function key(id, size) { return String(id).trim() + '|' + String(size || '').trim(); }
function num(v) { var n = Number(v); return isNaN(n) ? 0 : n; }
function money(n) { return (Math.round(n * 100) / 100).toFixed(2) + ' €'; }

function donsIn(sale) {
  return (sale.items || []).reduce(function (t, it) {
    return it.type === 'don' ? t + num(it.price) * num(it.qty) : t;
  }, 0);
}

function countUnits(articles) {
  var total = 0;
  String(articles).split(',').forEach(function (p) {
    var m = p.match(/(\d+)×/);
    if (m) total += Number(m[1]);
  });
  return total;
}

function labelMode(m) {
  return m === 'cash' ? 'Espèces'
       : m === 'card' ? 'Carte'
       : m === 'pref' ? 'Prix préférentiel'
       : m === 'gift' ? 'Offert'
       : String(m || '');
}
function labelSettle(settle, method) {
  var s = settle || (method === 'cash' || method === 'card' ? method : '');
  return s === 'cash' ? 'Espèces' : s === 'card' ? 'Carte' : '';
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
