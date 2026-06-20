/**
 * QuetFlow — Web app Google Apps Script
 * Import, normalisation et visualisation des données de quêtes paroissiales.
 *
 * Couche serveur : sert l'app web, importe les exports CSV dans un Google Sheet
 * et renvoie les agrégats au tableau de bord.
 */

const SHEET_NAME = 'Quetes';
const PROP_SS_ID = 'QUETFLOW_SS_ID';

// Colonnes canoniques stockées (l'ordre définit l'ordre des colonnes du Sheet).
const HEADERS = ['Date', 'Paroisse', 'Type', 'Moyen', 'Montant'];

// Synonymes d'en-têtes d'export -> colonne canonique (clé).
const HEADER_ALIASES = {
  date:     ['date', 'jour', 'datecelebration', 'datequete', 'horodatage'],
  paroisse: ['paroisse', 'eglise', 'lieu', 'site', 'etablissement'],
  type:     ['type', 'typequete', 'nature', 'categorie', 'celebration'],
  moyen:    ['moyen', 'moyenpaiement', 'paiement', 'reglement', 'mode'],
  montant:  ['montant', 'somme', 'total', 'valeur', 'amount', 'euros', 'eur']
};

/* ------------------------------------------------------------------ *
 *  Point d'entrée Web App
 * ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createTemplateFromFile('App')
    .evaluate()
    .setTitle('QuetFlow')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Inclusion de partiels HTML (pattern HtmlService). */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ------------------------------------------------------------------ *
 *  Stockage (Google Sheet créé à la volée)
 * ------------------------------------------------------------------ */

/** Récupère (ou crée au 1er lancement) le classeur de stockage. */
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PROP_SS_ID);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recréé ci-dessous */ }
  }
  const ss = SpreadsheetApp.create('QuetFlow — Données');
  props.setProperty(PROP_SS_ID, ss.getId());
  return ss;
}

/** Récupère (ou initialise) la feuille de données avec sa ligne d'en-tête. */
function getSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];          // réutilise la 1re feuille (nom localisé variable)
    sh.setName(SHEET_NAME);
    sh.clear();
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** URL du classeur de stockage (lien « Données » dans l'en-tête). */
function getStorageUrl() {
  return getSpreadsheet_().getUrl();
}

/* ------------------------------------------------------------------ *
 *  Import / normalisation
 * ------------------------------------------------------------------ */

/**
 * Importe un contenu CSV : détecte les colonnes, normalise et ajoute les lignes.
 * @param {string} csv  Contenu texte du fichier exporté.
 * @return {{imported:number, skipped:number, total:number}}
 */
function importCsv(csv) {
  if (!csv || !csv.trim()) throw new Error('Fichier vide.');

  const rows = Utilities.parseCsv(csv, detectDelimiter_(csv));
  if (rows.length < 2) throw new Error('Aucune donnée détectée (en-tête + lignes attendus).');

  const head = rows[0].map(normalizeKey_);
  const col = {};
  Object.keys(HEADER_ALIASES).forEach(function (canon) {
    const aliases = HEADER_ALIASES[canon].map(normalizeKey_);
    col[canon] = head.findIndex(function (h) { return aliases.indexOf(h) > -1; });
  });
  if (col.montant < 0) throw new Error('Colonne « Montant » introuvable dans le fichier.');

  const out = [];
  let skipped = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.join('').trim() === '') continue;          // ligne vide
    const montant = parseAmount_(r[col.montant]);
    if (!montant) { skipped++; continue; }            // ligne sans montant exploitable
    out.push([
      col.date     > -1 ? (parseDate_(r[col.date]) || '') : '',
      col.paroisse > -1 ? String(r[col.paroisse] || '').trim() : '',
      col.type     > -1 ? String(r[col.type] || '').trim() || 'Quête' : 'Quête',
      col.moyen    > -1 ? String(r[col.moyen] || '').trim() || 'Inconnu' : 'Inconnu',
      montant
    ]);
  }

  const sh = getSheet_();
  if (out.length) {
    sh.getRange(sh.getLastRow() + 1, 1, out.length, HEADERS.length).setValues(out);
  }
  return { imported: out.length, skipped: skipped, total: Math.max(0, sh.getLastRow() - 1) };
}

/** Détecte le séparateur le plus probable sur la 1re ligne. */
function detectDelimiter_(csv) {
  const line = (csv.split(/\r?\n/)[0]) || '';
  const semi = (line.match(/;/g) || []).length;
  const comma = (line.match(/,/g) || []).length;
  return semi > comma ? ';' : ',';
}

/** Normalise une chaîne d'en-tête : minuscules, sans accents, alphanumérique. */
function normalizeKey_(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/\p{Diacritic}/gu, '')
    .replace(/[^a-z0-9]/g, '');
}

/** Convertit un montant texte (FR ou EN) en nombre. Ex. "1 234,56 €" -> 1234.56 */
function parseAmount_(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').replace(/[^0-9,.\-]/g, '').trim();   // retire € et espaces
  if (!s) return 0;
  if (s.indexOf(',') > -1) s = s.replace(/\./g, '').replace(',', '.');  // décimale FR
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/** Parse une date (jj/mm/aaaa prioritaire, sinon natif). Renvoie Date ou null. */
function parseDate_(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    let y = m[3];
    if (y.length === 2) y = '20' + y;
    return new Date(Number(y), Number(m[2]) - 1, Number(m[1]));
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/* ------------------------------------------------------------------ *
 *  Agrégation pour le tableau de bord
 * ------------------------------------------------------------------ */

/** Calcule KPIs et répartitions pour le dashboard. */
function getDashboardData() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    return { empty: true, kpis: emptyKpis_(), byMonth: [], byMoyen: [], byParoisse: [], recent: [] };
  }

  const values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const tz = 'Europe/Paris';
  let total = 0;
  const paroisse = {}, mois = {}, moyen = {};

  values.forEach(function (row) {
    const d = row[0], par = row[1] || '—', moy = row[3] || 'Inconnu', mt = Number(row[4]) || 0;
    total += mt;
    paroisse[par] = (paroisse[par] || 0) + mt;
    moyen[moy] = (moyen[moy] || 0) + mt;
    if (d instanceof Date) {
      const key = Utilities.formatDate(d, tz, 'yyyy-MM');
      mois[key] = (mois[key] || 0) + mt;
    }
  });

  const count = values.length;
  const recent = values.slice(-12).reverse().map(function (row) {
    return {
      date: row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'dd/MM/yyyy') : '',
      paroisse: row[1], type: row[2], moyen: row[3], montant: Number(row[4]) || 0
    };
  });

  return {
    empty: false,
    kpis: { total: total, count: count, moyenne: count ? total / count : 0, paroisses: Object.keys(paroisse).length },
    byMonth: Object.keys(mois).sort().map(function (k) { return { label: k, value: mois[k] }; }),
    byMoyen: toSortedPairs_(moyen),
    byParoisse: toSortedPairs_(paroisse).slice(0, 10),
    recent: recent
  };
}

function toSortedPairs_(obj) {
  return Object.keys(obj)
    .map(function (k) { return { label: k, value: obj[k] }; })
    .sort(function (a, b) { return b.value - a.value; });
}

function emptyKpis_() { return { total: 0, count: 0, moyenne: 0, paroisses: 0 }; }

/** Réinitialise toutes les données (conserve l'en-tête). */
function resetData() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  return getDashboardData();
}
