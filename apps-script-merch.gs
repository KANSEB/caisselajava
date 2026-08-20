/**
 * CAISSE MERCH — LA JAVA  ·  connecteur pour TON doc existant
 * ───────────────────────────────────────────────────────────────────────────
 * Doc : "Inventaire Merch" (onglets Inventaire / Rentabilité / Grille de tarif.)
 *
 * CE QUE FAIT CE SCRIPT
 *  • Lit la carte pour l'app : articles + tailles + stock (onglet Inventaire)
 *    croisés avec les prix PU Public / PU Préf (onglet Rentabilité).
 *  • À chaque encaissement envoyé par l'app :
 *      – décrémente la bonne case TAILLE dans Inventaire (le stock "vivant") ;
 *      – incrémente le compteur Normal / Préf / Don de l'article dans Rentabilité ;
 *      – journalise la vente dans un onglet "Ventes" (créé automatiquement).
 *  • Ne touche JAMAIS aux colonnes en formule (Total, Valeur, CA, Marge, ROI…).
 *
 * INSTALLATION
 *  1. Ta Sheet → Extensions → Apps Script → colle tout ce fichier → 💾
 *  2. Lance une fois  verifierStructure()  (menu ▶).
 *     ⚠ PIÈGES au 1er lancement :
 *       – Google demande d'AUTORISER le script : « Paramètres avancés » →
 *         « Accéder à … (non sécurisé) » → Autoriser. Tant que ce n'est pas
 *         fait, l'exécution semble tourner en boucle.
 *       – La fenêtre de résultat s'ouvre DANS L'ONGLET GOOGLE SHEETS (pas dans
 *         l'éditeur) : tant que tu ne cliques pas OK là-bas, l'éditeur affiche
 *         « Exécution en cours… ». Le même rapport est aussi écrit dans le
 *         « Journal d'exécution » en bas de l'éditeur.
 *     Alternative sans aucune popup :  testSansPopup()  → tout s'affiche dans
 *     le Journal d'exécution.
 *  3. Déployer → Nouveau déploiement → "Application Web"
 *        Exécuter en tant que : Moi   |   Qui a accès : Tout le monde
 *  4. Copie l'URL /exec → colle-la dans merch.html → SHEETS_WEBHOOK_URL
 *
 *  ⚠ À CHAQUE modif de ce script : Déployer → Gérer les déploiements →
 *     ✏ → Version « Nouvelle version » → Déployer (sinon l'ancienne URL sert
 *     l'ancien code).
 *
 * SI TU RENOMMES / DÉPLACES DES COLONNES : rien à recoder, le script les
 *  retrouve par leur libellé (voir la section LIBELLÉS ci-dessous).
 */

/* ─── Noms d'onglets ─────────────────────────────────────────────────────── */
var SH_INV   = 'Inventaire';
var SH_RENT  = 'Rentabilité';
var SH_SALES = 'Ventes';

/* ─── LIBELLÉS attendus (adapte ici si tu changes un intitulé) ───────────── */
var LBL_ARTICLE   = 'Article';
var LBL_MARQUE    = 'Marque';
var SIZE_LABELS   = ['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', 'TU'];  // TU = taille unique (tote)
var LBL_PRODUITE  = 'Produite';
var LBL_NORMAL    = 'Normal';
var LBL_PREF      = 'Préf';
var LBL_DON       = 'Don';
var LBL_PU_PUBLIC = 'PU Public';
var LBL_PU_PREF   = 'PU Préf';

var SALES_HEADERS = ['Horodatage', 'Mode', 'Règlement', 'Montant encaissé',
                     'Prix plein', 'Remise', 'Don libre', 'Articles', 'Vendeur', 'Appareil', 'JSON'];

/* ══════════════════════════════════════════════════════════════════════════
   OUTIL DE CONTRÔLE — lance-le une fois après avoir collé le script
   ══════════════════════════════════════════════════════════════════════════ */
function verifierStructure() {
  var msg = [];
  try {
    var inv = locateInventory();
    msg.push('✅ Inventaire : en-tête ligne ' + inv.headerRow +
             ', Article col ' + inv.cArticle +
             ', tailles trouvées : ' + Object.keys(inv.sizeCols).join(', ') +
             ', ' + inv.articles.length + ' article(s).');
  } catch (e) { msg.push('❌ Inventaire : ' + e.message); }

  try {
    var rent = locateRentabilite();
    msg.push('✅ Rentabilité : en-tête ligne ' + rent.headerRow +
             ' | Produite=' + col(rent.cProduite) + ' Normal=' + col(rent.cNormal) +
             ' Préf=' + col(rent.cPref) + ' Don=' + col(rent.cDon) +
             ' PU Public=' + col(rent.cPub) + ' PU Préf=' + col(rent.cPref2) +
             ' | ' + rent.articles.length + ' article(s).');
  } catch (e) { msg.push('❌ Rentabilité : ' + e.message); }

  var inv2;
  try { inv2 = getInventory(); msg.push('✅ Carte app : ' + inv2.rows.length + ' ligne(s) taille prêtes pour la caisse.'); }
  catch (e) { msg.push('❌ Carte app : ' + e.message); }

  var report = 'CONTRÔLE STRUCTURE\n\n' + msg.join('\n\n');

  /* Toujours dans le journal (éditeur → "Journal d'exécution"), lisible sans popup */
  Logger.log(report);

  /* La popup s'ouvre DANS L'ONGLET GOOGLE SHEETS, pas dans l'éditeur.
     Si aucune interface n'est disponible, on n'attend pas indéfiniment. */
  try { SpreadsheetApp.getUi().alert(report); } catch (e) {}

  return report;
}
function col(n) { return n ? columnLetter(n) : '—'; }

/** Variante 100% sans popup : lance-la depuis l'éditeur et lis le résultat
 *  directement dans le "Journal d'exécution" en bas de l'écran. */
function testSansPopup() {
  var inv = getInventory();
  Logger.log('Carte envoyée à l\'app (' + inv.rows.length + ' lignes) :');
  inv.rows.forEach(function (r) {
    Logger.log(' - ' + r.name + ' [' + r.size + ']  ' + r.price + '€ / préf ' + r.pricePref + '€  stock ' + r.initial);
  });
  return inv.rows.length;
}

/* ══════════════════════════════════════════════════════════════════════════
   LECTURE  (GET depuis l'app)
   ══════════════════════════════════════════════════════════════════════════ */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'stats';
  if (action === 'inventory') return json(getInventory());
  if (action === 'history')   return json(getHistory());
  if (action === 'deleteSale') return json(deleteSale(e.parameter));
  return json(getStats());
}

/** Journal des encaissements, du plus recent au plus ancien (50 max). */
function getHistory() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SALES);
  if (!sh || sh.getLastRow() < 2) return { rows: [] };
  var last = sh.getLastRow();
  var first = Math.max(2, last - 49);
  var data = sh.getRange(first, 1, last - first + 1, SALES_HEADERS.length).getValues();
  var rows = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var v = data[i];
    rows.push({
      row: first + i,
      ts: v[0] instanceof Date ? v[0].toISOString() : String(v[0]),
      mode: String(v[1] || ''), settle: String(v[2] || ''),
      amount: num(v[3]), full: num(v[4]), discount: num(v[5]), don: num(v[6]),
      articles: String(v[7] || ''), vendor: String(v[8] || ''), device: String(v[9] || ''),
      hasJson: !!String(v[10] || '')
    });
  }
  return { rows: rows, total: last - 1 };
}

/**
 * Annule un encaissement : restitue le stock (Inventaire) et les compteurs
 * (Rentabilité) grâce au JSON stocké, puis supprime la ligne du journal.
 * Sécurité : l'horodatage passé doit correspondre à celui de la ligne visée.
 */
function deleteSale(p) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (err) { return { ok: false, error: 'lock' }; }
  try {
    var row = parseInt(p.row, 10);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SALES);
    if (!sh || isNaN(row) || row < 2 || row > sh.getLastRow()) return { ok: false, error: 'ligne introuvable' };

    var v = sh.getRange(row, 1, 1, SALES_HEADERS.length).getValues()[0];
    var ts = v[0] instanceof Date ? v[0].toISOString() : String(v[0]);
    if (p.ts && ts !== p.ts) return { ok: false, error: 'ligne modifiée entre-temps — recharge l\'historique' };

    var reversed = false;
    var jsonStr = String(v[10] || '');
    if (jsonStr) {
      try {
        var sale = JSON.parse(jsonStr);
        reverseSale(sale);
        reversed = true;
      } catch (err) {}
    }
    sh.deleteRow(row);
    return { ok: true, reversed: reversed };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/** Inverse de applyToSheet : re-crédite le stock, décrémente les compteurs. */
function reverseSale(sale) {
  var items = (sale.items || []).filter(function (it) { return it.type !== 'don'; });
  if (!items.length) return;
  var inv  = locateInventory();
  var rent = locateRentabilite();
  var counterCol = sale.method === 'gift' ? rent.cDon
                 : sale.method === 'pref' ? rent.cPref
                 : rent.cNormal;
  items.forEach(function (it) {
    var nkey = norm(it.name);
    var artRow = inv.rowByName[nkey];
    if (artRow) {
      var sizeCol = resolveSizeCol(inv, it.size, categoryOf(it.name));
      if (sizeCol) {
        var cell = inv.sheet.getRange(artRow, sizeCol);
        cell.setValue(num(cell.getValue()) + num(it.qty));
      }
    }
    var rRow = rent.rowByName[nkey];
    if (rRow && counterCol) {
      var c = rent.sheet.getRange(rRow, counterCol);
      c.setValue(Math.max(0, num(c.getValue()) - num(it.qty)));
    }
  });
}

/** Carte + stock par taille + prix, prête pour la grille de l'app. */
function getInventory() {
  var inv  = locateInventory();
  var rent = locateRentabilite();
  var prices = rent.priceByName;   // nom normalisé -> { pub, pref }

  var rows = [];
  inv.articles.forEach(function (a) {   // a = { name, row, values }
    var p = prices[norm(a.name)] || { pub: 0, pref: 0 };
    var id = slug(a.name);
    var cat = categoryOf(a.name);
    inv.sizeOrder.forEach(function (s) {   // s = { label, colIndexInValues, shared }
      var raw = a.values[s.idx];
      if (raw === '' || raw === null || raw === undefined) return;   // taille non proposée
      var stock = Number(raw);
      if (isNaN(stock)) return;
      var label = s.shared ? (cat === 'tote' ? 'TU' : '3XL') : s.label;
      rows.push({
        id: id, name: a.name, cat: cat, size: label,
        price: p.pub, pricePref: p.pref,
        initial: stock, out: 0
      });
    });
  });
  return { rows: rows };
}

/** Totaux caisse depuis le journal Ventes. */
function getStats() {
  var s = { cash:0, card:0, countCash:0, countCard:0,
            countPref:0, discount:0,
            countGift:0, giftedUnits:0, giftedValue:0,
            countDon:0, donAmount:0 };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SALES);
  if (!sh || sh.getLastRow() < 2) return s;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, SALES_HEADERS.length).getValues();
  data.forEach(function (v) {
    var mode = String(v[1] || ''), settle = String(v[2] || '');
    var amount = num(v[3]), full = num(v[4]), disc = num(v[5]), don = num(v[6]), arts = String(v[7] || '');
    if (settle.indexOf('spèce') >= 0)      { s.cash += amount; s.countCash++; }
    else if (settle.indexOf('arte') >= 0)  { s.card += amount; s.countCard++; }
    if (mode.indexOf('référ') >= 0)        { s.countPref++; s.discount += disc; }
    if (mode.indexOf('ffert') >= 0)        { s.countGift++; s.giftedValue += full; s.giftedUnits += countUnits(arts); }
    if (don > 0)                            { s.countDon++; s.donAmount += don; }
  });
  ['cash','card','discount','giftedValue','donAmount'].forEach(function (k) { s[k] = round2(s[k]); });
  return s;
}

/* ══════════════════════════════════════════════════════════════════════════
   ÉCRITURE  (POST depuis l'app)
   ══════════════════════════════════════════════════════════════════════════ */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (err) { return json({ ok:false, error:'lock' }); }
  try {
    var sale = JSON.parse(e.postData.contents);
    journalize(sale);
    applyToSheet(sale);
    return json({ ok:true });
  } catch (err) {
    return json({ ok:false, error:String(err) });
  } finally {
    lock.releaseLock();
  }
}

function journalize(sale) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_SALES);
  if (!sh) {
    sh = ss.insertSheet(SH_SALES);
    sh.getRange(1, 1, 1, SALES_HEADERS.length).setValues([SALES_HEADERS])
      .setFontWeight('bold').setBackground('#f0f0f0');
    sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, 9).getValue()) !== SALES_HEADERS[8]) {
    /* migration : ancien journal sans colonnes Vendeur / JSON */
    sh.getRange(1, 1, 1, SALES_HEADERS.length).setValues([SALES_HEADERS])
      .setFontWeight('bold').setBackground('#f0f0f0');
  }
  var items = sale.items || [];
  var don = 0, arts = [];
  items.forEach(function (it) {
    if (it.type === 'don') { don += num(it.price) * num(it.qty); arts.push('Don ' + money(num(it.price) * num(it.qty))); }
    else arts.push(num(it.qty) + '× ' + it.name + (it.size ? ' ' + it.size : ''));
  });
  sh.appendRow([
    new Date(sale.ts || new Date()),
    labelMode(sale.method), labelSettle(sale.settle, sale.method),
    num(sale.amount), num(sale.fullPrice), num(sale.discount),
    don, arts.join(', '), sale.vendor || '', sale.deviceId || '',
    JSON.stringify({ method: sale.method, items: sale.items || [] })
  ]);
}

/** Décrémente le stock par taille (Inventaire) + compteurs (Rentabilité). */
function applyToSheet(sale) {
  var items = (sale.items || []).filter(function (it) { return it.type !== 'don'; });
  if (!items.length) return;

  var inv  = locateInventory();
  var rent = locateRentabilite();
  var invSheet  = inv.sheet;
  var rentSheet = rent.sheet;

  var counterCol = sale.method === 'gift' ? rent.cDon
                 : sale.method === 'pref' ? rent.cPref
                 : rent.cNormal;   // cash | card => vente "normale"

  items.forEach(function (it) {
    var nkey = norm(it.name);

    /* 1) Inventaire : baisse la case taille */
    var artRow = inv.rowByName[nkey];
    if (artRow) {
      var sizeCol = resolveSizeCol(inv, it.size, categoryOf(it.name));
      if (sizeCol) {
        var cell = invSheet.getRange(artRow, sizeCol);
        var cur = num(cell.getValue());
        cell.setValue(Math.max(0, cur - num(it.qty)));
      }
    }

    /* 2) Rentabilité : incrémente Normal / Préf / Don */
    var rRow = rent.rowByName[nkey];
    if (rRow && counterCol) {
      var c = rentSheet.getRange(rRow, counterCol);
      c.setValue(num(c.getValue()) + num(it.qty));
    }
  });
}

/** Recalcule tout au cas où (les formules Sheet le font seules ; ceci est un filet). */
function forcerRecalcul() { SpreadsheetApp.flush(); }

/* ══════════════════════════════════════════════════════════════════════════
   LOCALISATION DES ONGLETS / COLONNES  (détection par libellé)
   ══════════════════════════════════════════════════════════════════════════ */
function locateInventory() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_INV);
  if (!sheet) throw new Error('Onglet "' + SH_INV + '" introuvable');
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  var grid = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  /* en-tête = ligne contenant "Article" */
  var headerRow = -1, cArticle = -1;
  for (var r = 0; r < Math.min(grid.length, 20); r++) {
    for (var c = 0; c < grid[r].length; c++) {
      if (String(grid[r][c]).trim() === LBL_ARTICLE) { headerRow = r; cArticle = c; break; }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow < 0) throw new Error('Ligne d\'en-tête (cellule "' + LBL_ARTICLE + '") non trouvée');

  var cMarque = findInRow(grid[headerRow], LBL_MARQUE);

  /* tailles : cherchées sur la ligne d'en-tête ET la suivante (sous-titres XS/S/M…) */
  var sizeCols = {};
  [headerRow, headerRow + 1].forEach(function (rr) {
    if (rr >= grid.length) return;
    grid[rr].forEach(function (cell, ci) {
      var t = String(cell).trim();
      if (t === 'XS')  sizeCols['XS']  = ci + 1;
      else if (t === 'S')   sizeCols['S']   = ci + 1;
      else if (t === 'M')   sizeCols['M']   = ci + 1;
      else if (t === 'L')   sizeCols['L']   = ci + 1;
      else if (t === 'XL')  sizeCols['XL']  = ci + 1;
      else if (t === '2XL') sizeCols['2XL'] = ci + 1;
      else if (/3XL/.test(t) || /(^|\W)TU(\W|$)/.test(t)) { sizeCols['3XL/TU'] = ci + 1; }
    });
  });
  if (!Object.keys(sizeCols).length) throw new Error('Aucune colonne de taille (XS…3XL/TU) trouvée');

  /* ordre des tailles pour la lecture ligne par ligne */
  var order = [];
  ['XS','S','M','L','XL','2XL'].forEach(function (sz) {
    if (sizeCols[sz]) order.push({ label: sz, col: sizeCols[sz], shared: false });
  });
  if (sizeCols['3XL/TU']) order.push({ label: '3XL/TU', col: sizeCols['3XL/TU'], shared: true });

  /* données : de headerRow+2 jusqu'à une ligne Article vide / "TOTAL" */
  var firstData = headerRow + 2;   // saute la ligne des sous-titres tailles
  var articles = [], rowByName = {};
  for (var rr = firstData; rr < grid.length; rr++) {
    var name = String(grid[rr][cArticle]).trim();
    if (name === '') continue;
    if (/^total/i.test(name)) break;
    var values = order.map(function (o) { return grid[rr][o.col - 1]; });
    var sizeOrder = order.map(function (o, idx) { return { label: o.label, idx: idx, shared: o.shared }; });
    articles.push({ name: name, row: rr + 1, values: values, sizeOrder: sizeOrder });
    rowByName[norm(name)] = rr + 1;
  }
  var globalSizeOrder = order.map(function (o, idx) { return { label: o.label, idx: idx, shared: o.shared }; });

  return {
    sheet: sheet, headerRow: headerRow + 1, cArticle: cArticle + 1, cMarque: cMarque,
    sizeCols: sizeCols, sizeOrder: globalSizeOrder,
    articles: articles, rowByName: rowByName
  };
}

/** Retrouve la colonne Inventaire d'une taille donnée pour un article. */
function resolveSizeCol(inv, size, cat) {
  var s = String(size || '').trim().toUpperCase();
  if (s === 'TU' || s === '3XL') return inv.sizeCols['3XL/TU'] || null;
  return inv.sizeCols[s] || null;
}

function locateRentabilite() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_RENT);
  if (!sheet) throw new Error('Onglet "' + SH_RENT + '" introuvable');
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  var grid = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var headerRow = -1, cArticle = -1;
  for (var r = 0; r < Math.min(grid.length, 25); r++) {
    for (var c = 0; c < grid[r].length; c++) {
      if (String(grid[r][c]).trim() === LBL_ARTICLE) { headerRow = r; cArticle = c; break; }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow < 0) throw new Error('Ligne d\'en-tête non trouvée');

  var row = grid[headerRow];
  var cProduite = findInRow(row, LBL_PRODUITE);
  var cNormal   = findInRow(row, LBL_NORMAL);
  var cPref     = findExact(row, LBL_PREF);      // "Préf" seul, pas "PU Préf"
  var cDon      = findExact(row, LBL_DON);
  var cPub      = findInRow(row, LBL_PU_PUBLIC);
  var cPref2    = findInRow(row, LBL_PU_PREF);

  var priceByName = {}, rowByName = {}, articles = [];
  for (var rr = headerRow + 1; rr < grid.length; rr++) {
    var name = String(grid[rr][cArticle]).trim();
    if (name === '') continue;
    if (/^total/i.test(name)) break;
    var nkey = norm(name);
    priceByName[nkey] = {
      pub:  cPub   ? num(grid[rr][cPub - 1])   : 0,
      pref: cPref2 ? num(grid[rr][cPref2 - 1]) : 0
    };
    rowByName[nkey] = rr + 1;
    articles.push(name);
  }

  return {
    sheet: sheet, headerRow: headerRow + 1, cArticle: cArticle + 1,
    cProduite: cProduite, cNormal: cNormal, cPref: cPref, cDon: cDon,
    cPub: cPub, cPref2: cPref2,
    priceByName: priceByName, rowByName: rowByName, articles: articles
  };
}

/* ─── Helpers colonnes ───────────────────────────────────────────────────── */
function findInRow(row, label) {   // "contient" (insensible casse/espaces)
  var target = norm(label);
  for (var i = 0; i < row.length; i++) {
    if (norm(String(row[i])).indexOf(target) >= 0 && String(row[i]).trim() !== '') return i + 1;
  }
  return 0;
}
function findExact(row, label) {   // égalité stricte après trim
  for (var i = 0; i < row.length; i++) {
    if (String(row[i]).trim() === label) return i + 1;
  }
  return 0;
}
function columnLetter(n) { var s=''; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=(n-m-1)/26; } return s; }

/* ─── Helpers divers ─────────────────────────────────────────────────────── */
function categoryOf(name) {
  var n = norm(name);
  if (n.indexOf('tote') >= 0)                       return 'tote';
  if (n.indexOf('pull') >= 0 || n.indexOf('capuche') >= 0 || n.indexOf('hoodie') >= 0 || n.indexOf('sweat') >= 0) return 'pull';
  return 'tee';
}
function slug(name) {
  return norm(name).replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}
function norm(s) {
  return String(s).toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ').trim();
}
function num(v) { var n = Number(v); return isNaN(n) ? 0 : n; }
function round2(n) { return Math.round(n * 100) / 100; }
function money(n) { return round2(n).toFixed(2) + ' €'; }
function countUnits(arts) {
  var t = 0; String(arts).split(',').forEach(function (p) { var m = p.match(/(\d+)×/); if (m) t += Number(m[1]); }); return t;
}
function labelMode(m) {
  return m === 'cash' ? 'Espèces' : m === 'card' ? 'Carte'
       : m === 'pref' ? 'Prix préférentiel' : m === 'gift' ? 'Offert' : String(m || '');
}
function labelSettle(settle, method) {
  var s = settle || ((method === 'cash' || method === 'card') ? method : '');
  return s === 'cash' ? 'Espèces' : s === 'card' ? 'Carte' : '';
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
