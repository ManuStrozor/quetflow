/**
 * QuetFlow — Web app Google Apps Script
 * Import, normalisation et visualisation des données de quêtes paroissiales.
 *
 * Couche serveur : sert l'app web, importe les exports CSV (13 colonnes) dans un
 * Google Sheet et renvoie les lignes normalisées au tableau de bord, qui agrège
 * et filtre côté client (Bornes, Quêtes, ventilation matin/soir & dominical).
 */

const SHEET_NAME = 'Transactions';
const PROP_SS_ID = 'QUETFLOW_SS_ID';
const TZ = 'Europe/Paris';

// Discount appliqué au montant brut : net = ROUNDDOWN(brut * RATE ; 2) - FIXED.
const DISCOUNT_RATE = 0.9885;
const DISCOUNT_FIXED = 0.1;
// Fraction de journée séparant matin et soir (0,6 jour = 14h24).
const MORNING_THRESHOLD = 0.6;

// Colonnes canoniques stockées (l'ordre définit l'ordre des colonnes du Sheet).
const HEADERS = ['Date', 'Terminal ID', 'Terminal Name', 'Project Name',
                 'Celebration', 'Location', 'Transaction ID', 'Amount Brut', 'Amount Net',
                 'Dominicale', 'Matin/Soir'];

// Synonymes d'en-têtes d'export -> colonne canonique (clé), comparés après normalizeKey_.
// Colonnes d'export ignorées : Client Name, Currency, Type, Total Transactions, Total Amount.
const HEADER_ALIASES = {
  date:          ['transactiondate', 'date', 'datetransaction', 'horodatage'],
  terminalId:    ['terminalid'],
  terminalName:  ['terminalname'],
  projectName:   ['projectname', 'project', 'projet'],
  celebration:   ['celebration', 'quete'],
  location:      ['location', 'lieu'],
  transactionId: ['transactionid'],
  amount:        ['amount', 'montant']
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

/**
 * Récupère la feuille de données. (Re)pose la ligne d'en-tête si elle est absente
 * ou obsolète : un changement de schéma efface l'ancien contenu (migration).
 */
function getSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];          // réutilise la 1re feuille (nom localisé variable)
    sh.setName(SHEET_NAME);
    sh.clear();
  }
  const lastCol = sh.getLastColumn();
  const firstRow = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (firstRow.join('|') !== HEADERS.join('|')) {
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
 * Importe un contenu CSV (export 13 colonnes) : détecte les colonnes, normalise,
 * applique le discount et ajoute les lignes. Déduplique sur Transaction ID pour
 * qu'un ré-import du même fichier ne double pas les montants.
 * @param {string} csv  Contenu texte du fichier exporté.
 * @return {{imported:number, skipped:number, duplicates:number, total:number}}
 */
function importCsv(csv) {
  if (!csv || !csv.trim()) throw new Error('Fichier vide.');

  const rows = Utilities.parseCsv(csv, detectDelimiter_(csv));
  if (rows.length < 2) throw new Error('Aucune donnée détectée (en-tête + lignes attendus).');

  const head = rows[0].map(normalizeKey_);
  const col = {};
  Object.keys(HEADER_ALIASES).forEach(function (canon) {
    const aliases = HEADER_ALIASES[canon];
    col[canon] = head.findIndex(function (h) { return aliases.indexOf(h) > -1; });
  });
  if (col.amount < 0) throw new Error('Colonne « Amount » introuvable dans le fichier.');
  if (col.date < 0)   throw new Error('Colonne « Transaction Date » introuvable dans le fichier.');

  const sh = getSheet_();
  const seen = existingTxIds_(sh);            // doublons inter-imports + intra-fichier

  const out = [];
  let skipped = 0, duplicates = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.join('').trim() === '') continue;             // ligne vide
    const brut = parseAmount_(r[col.amount]);
    if (!brut) { skipped++; continue; }                 // ligne sans montant exploitable

    const txId = col.transactionId > -1 ? String(r[col.transactionId] || '').trim() : '';
    if (txId) {
      if (seen[txId]) { duplicates++; continue; }
      seen[txId] = true;
    }

    const date = parseDate_(r[col.date]);
    out.push([
      date || '',
      col.terminalId   > -1 ? String(r[col.terminalId]   || '').trim() : '',
      col.terminalName > -1 ? String(r[col.terminalName] || '').trim() : '',
      col.projectName  > -1 ? String(r[col.projectName]  || '').trim() : '',
      col.celebration  > -1 ? String(r[col.celebration]  || '').trim() : '',
      col.location     > -1 ? String(r[col.location]     || '').trim() : '',
      txId,
      brut,
      discount_(brut),
      date ? (isSunday_(date)  ? 'Oui'   : 'Non')  : '',
      date ? (isMorning_(date) ? 'Matin' : 'Soir') : ''
    ]);
  }

  if (out.length) {
    sh.getRange(sh.getLastRow() + 1, 1, out.length, HEADERS.length).setValues(out);
  }
  return {
    imported: out.length,
    skipped: skipped,
    duplicates: duplicates,
    total: Math.max(0, sh.getLastRow() - 1)
  };
}

/** Map des Transaction ID déjà présents dans la feuille (pour la déduplication). */
function existingTxIds_(sh) {
  const map = {};
  const last = sh.getLastRow();
  if (last < 2) return map;
  const idCol = HEADERS.indexOf('Transaction ID') + 1;
  sh.getRange(2, idCol, last - 1, 1).getValues().forEach(function (row) {
    const v = String(row[0] || '').trim();
    if (v) map[v] = true;
  });
  return map;
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

/**
 * Applique le discount au montant brut.
 * net = ROUNDDOWN(brut * 0,9885 ; 2) - 0,1   (arrondi inférieur à 2 décimales)
 * Ex. : 10 € -> ROUNDDOWN(9,885;2)=9,88 -> 9,78 €.
 */
function discount_(brut) {
  const truncated = Math.floor(brut * DISCOUNT_RATE * 100) / 100;
  return truncated - DISCOUNT_FIXED;
}

/** Parse une date au format mm/jj/aa hh:mm:ss (US, mois en premier). Renvoie Date ou null. */
function parseDate_(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    let y = m[3];
    if (y.length === 2) y = '20' + y;
    return new Date(Number(y), Number(m[1]) - 1, Number(m[2]),   // m[1]=mois, m[2]=jour
                    Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Vrai si la date tombe un dimanche (WEEKDAY = 1). */
function isSunday_(d) { return d instanceof Date && d.getDay() === 0; }

/** Vrai si la transaction est du matin : MOD(date ; 1) < 0,6 (avant 14h24). */
function isMorning_(d) {
  if (!(d instanceof Date)) return false;
  return (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400 < MORNING_THRESHOLD;
}

/* ------------------------------------------------------------------ *
 *  Lecture pour le tableau de bord (agrégation côté client)
 * ------------------------------------------------------------------ */

/**
 * Renvoie les lignes normalisées et enrichies pour le dashboard.
 * Les drapeaux dominical (dim) et matin sont pré-calculés en fuseau Europe/Paris.
 * @return {{empty:boolean, rows:Array<Object>}}
 */
function getRows() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { empty: true, rows: [] };

  const C = {};
  HEADERS.forEach(function (h, i) { C[h] = i; });

  const values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const rows = values.map(function (row) {
    const d = row[C['Date']];
    const isDate = d instanceof Date;
    return {
      ts:           isDate ? d.getTime() : null,
      mois:         isDate ? Utilities.formatDate(d, TZ, 'yyyy-MM') : '',
      dateStr:      isDate ? Utilities.formatDate(d, TZ, 'dd/MM/yyyy') : '',
      timeStr:      isDate ? Utilities.formatDate(d, TZ, 'HH:mm') : '',
      terminalId:   row[C['Terminal ID']],
      terminalName: row[C['Terminal Name']],
      projet:       row[C['Project Name']],
      celebration:  row[C['Celebration']],
      location:     row[C['Location']],
      txId:         row[C['Transaction ID']],
      brut:         Number(row[C['Amount Brut']]) || 0,
      net:          Number(row[C['Amount Net']]) || 0,
      dim:          isSunday_(d),     // WEEKDAY = 1 (dimanche)
      matin:        isMorning_(d)     // MOD(date;1) < 0,6
    };
  });

  return { empty: false, rows: rows };
}

/** Réinitialise toutes les données (conserve l'en-tête). */
function resetData() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  return { empty: true, rows: [] };
}
