/**
 * Bella's Car Search - live CARFAX refresh for Google Sheets.
 *
 * Paste this into Extensions > Apps Script in the spreadsheet, save, reload the
 * sheet, then use the "Car Search" menu that appears next to Help.
 *
 * What a refresh does:
 *   1. Reads your Search Criteria and Scoring Weights tabs and keeps them as-is.
 *   2. Reads every Status / Notes / Contacted? cell you have filled in and parks
 *      them, keyed by VIN, in a hidden _NotesStore sheet.
 *   3. Pulls current listings from CARFAX and rescores them.
 *   4. Rebuilds All Cars, Shortlist, Top 10 Compare, Model Benchmarks and
 *      Start Here, writing your notes back onto the same VINs.
 *
 * Nothing you typed is lost, and a car you made notes on that has since sold is
 * kept at the bottom of All Cars marked DELISTED rather than silently vanishing.
 */

// ============================== CONFIG ==============================
var CONFIG = {
  zip: '46514',
  radius: 75,
  yearMin: 2022,
  priceMax: 40000,
  targets: [
    { make: 'Toyota', model: '4Runner' },
    { make: 'Toyota', model: 'RAV4' },
    { make: 'Honda',  model: 'HR-V' },
    { make: 'Honda',  model: 'CR-V' },
    { make: 'Subaru', model: 'Outback' }
  ],
  mileageCaps: { Toyota: 100000, Honda: 70000, Subaru: 50000 },
  defaultWeights: { rel: 26, val: 22, mile: 18, cpo: 15, feat: 14, dist: 5 }
};

var ENDPOINT = 'https://helix.carfax.com/search/v2/vehicles';
var NOTE_COLS = ['Status', 'Notes', 'Contacted?'];
var STORE = '_NotesStore';

var FEATURES = [
  { label: 'Heated Steering',   pts: 10, match: function (o) { return has(o, 'heated steering'); } },
  { label: 'Sunroof/Moonroof',  pts: 14, match: function (o) { return has(o, 'sunroof') || has(o, 'moonroof'); } },
  { label: 'Blind Spot',        pts: 12, match: function (o) { return has(o, 'blind spot'); } },
  { label: 'Power Liftgate',    pts: 12, match: function (o) { return has(o, 'power liftgate'); } },
  { label: 'Leather',           pts: 12, match: function (o) { return hasLeather(o); } },
  { label: 'Remote Start',      pts: 10, match: function (o) { return has(o, 'remote start'); } },
  { label: 'Navigation',        pts: 8,  match: function (o) { return has(o, 'navigation'); } },
  { label: 'Wireless Charging', pts: 8,  match: function (o) { return has(o, 'wireless charg'); } },
  { label: 'Parking Sensors',   pts: 6,  match: function (o) { return has(o, 'parking sensor'); } },
  { label: 'Power Driver Seat', pts: 4,  match: function (o) { return has(o, 'power seat') || has(o, 'power driver'); } },
  { label: 'Premium Audio',     pts: 4,  match: function (o) { return hasPremiumAudio(o); } }
];
var FEAT_MAX = FEATURES.reduce(function (s, f) { return s + f.pts; }, 0);

// ============================== MENU ==============================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Car Search')
    .addItem('Refresh from CARFAX now', 'refreshFromCarfax')
    .addSeparator()
    .addItem('Turn on nightly auto-refresh', 'createNightlyTrigger')
    .addItem('Turn off auto-refresh', 'removeTriggers')
    .addToUi();
}

function createNightlyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('refreshFromCarfax').timeBased().atHour(5).everyDays(1).create();
  toast('Nightly refresh is on (around 5am).');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshFromCarfax') ScriptApp.deleteTrigger(t);
  });
  toast('Auto-refresh is off.');
}

function toast(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Car Search', 6); } catch (e) {}
}

// ============================== MAIN ==============================
function refreshFromCarfax() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  toast('Reading your settings and notes...');

  ensureCriteriaSheet(ss);
  ensureWeightsSheet(ss);
  var criteria = readCriteria(ss);
  var notes = harvestNotes(ss);

  toast('Pulling listings from CARFAX...');
  var listings;
  try {
    listings = fetchAll(criteria);
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Could not reach CARFAX.\n\n' + err.message +
      '\n\nYour sheet has not been changed. Try again in a few minutes.');
    return;
  }
  if (!listings.length) {
    SpreadsheetApp.getUi().alert('CARFAX returned no listings. Your sheet has not been changed.');
    return;
  }

  toast('Scoring ' + listings.length + ' cars...');
  var cars = scoreAll(listings);
  cars.sort(function (a, b) { return b.provisional - a.provisional; });

  var liveVins = {};
  cars.forEach(function (c, i) { liveVins[c.vin] = true; c.allRow = 5 + i; });
  var orphans = [];
  Object.keys(notes).forEach(function (vin) {
    if (!liveVins[vin] && hasNote(notes[vin])) orphans.push(notes[vin]);
  });

  toast('Rebuilding tabs...');
  buildAllCars(ss, cars, notes, orphans);
  var shortlist = cars.filter(function (c) { return meetsHardCriteria(c, criteria); });
  buildShortlist(ss, shortlist, notes);
  buildCompare(ss, shortlist.slice(0, 10));
  buildBenchmarks(ss, cars);
  buildStartHere(ss, cars, shortlist, orphans.length);
  saveNotes(ss, notes);

  orderTabs(ss);
  toast('Done. ' + cars.length + ' listings, ' + shortlist.length + ' meet every criterion.');
}

// ============================== FETCH ==============================
function fetchAll(criteria) {
  var out = [], seen = {};
  CONFIG.targets.forEach(function (t) {
    var page = 1, totalPages = 1;
    while (page <= totalPages && page <= 20) {
      var url = ENDPOINT + '?' + [
        'zip=' + encodeURIComponent(CONFIG.zip),
        'radius=' + CONFIG.radius,
        'make=' + encodeURIComponent(t.make),
        'model=' + encodeURIComponent(t.model),
        'yearMin=' + criteria.yearMin,
        'priceMax=' + criteria.priceMax,
        'vehicleCondition=USED',
        'sort=BEST',
        'page=' + page
      ].join('&');

      var res = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' }
      });
      var code = res.getResponseCode();
      if (code !== 200) {
        throw new Error('CARFAX returned HTTP ' + code + ' for ' + t.make + ' ' + t.model + '.');
      }
      var data = JSON.parse(res.getContentText());
      totalPages = data.totalPageCount || 1;
      (data.listings || []).forEach(function (l) {
        if (l.vin && !seen[l.vin]) { seen[l.vin] = true; out.push(l); }
      });
      page++;
      Utilities.sleep(300);
    }
  });
  return out;
}

// ============================== SCORING ==============================
function optionSet(l) {
  var s = [];
  ['topOptions', 'atomTopOptions', 'atomOtherOptions'].forEach(function (k) {
    (l[k] || []).forEach(function (o) { s.push(String(o).toLowerCase()); });
  });
  return s;
}
function has(opts, needle) {
  for (var i = 0; i < opts.length; i++) if (opts[i].indexOf(needle) !== -1) return true;
  return false;
}
function hasLeather(opts) {
  for (var i = 0; i < opts.length; i++) {
    var o = opts[i];
    if (o.indexOf('leather') !== -1 && o.indexOf('leatherette') === -1 && o.indexOf('steering') === -1) return true;
  }
  return false;
}
function hasPremiumAudio(opts) {
  for (var i = 0; i < opts.length; i++) {
    var o = opts[i];
    if (o.indexOf('premium') !== -1 &&
        (o.indexOf('audio') !== -1 || o.indexOf('sound') !== -1 || o.indexOf('speaker') !== -1)) return true;
  }
  return false;
}
function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

function relComponent(l) {
  var r = l.reliability;
  if (!r) return null;
  var badge = { GREAT: 100, GOOD: 72, FAIR: 50 }[r.overallReliabilityBadge];
  if (badge === undefined) badge = 60;
  var cost = { LOW: 100, AVERAGE: 65, HIGH: 30 }[r.costBadge];
  if (cost === undefined) cost = 65;
  var risk = { LOW: 100, AVERAGE: 65, HIGH: 30 }[r.riskBadge];
  if (risk === undefined) risk = 65;
  var pl = r.reliabilityPlacement;
  var place = pl === 25 ? 100 : (pl === 50 ? 72 : (pl ? 50 : 65));
  return 0.45 * badge + 0.20 * cost + 0.20 * risk + 0.15 * place;
}

/** Least-squares fit of price on (year, mileage/1000) for one model. */
function fitModel(group) {
  var n = group.length;
  if (n < 8) return null;
  var X = [], y = [];
  for (var i = 0; i < n; i++) {
    X.push([1, group[i].year, (group[i].mileage || 0) / 1000]);
    y.push(group[i].currentPrice || group[i].listPrice || 0);
  }
  var A = [], b = [];
  for (var r = 0; r < 3; r++) {
    A.push([0, 0, 0]); b.push(0);
    for (var c = 0; c < 3; c++) for (var k = 0; k < n; k++) A[r][c] += X[k][r] * X[k][c];
    for (var k2 = 0; k2 < n; k2++) b[r] += X[k2][r] * y[k2];
  }
  var M = [A[0].concat(b[0]), A[1].concat(b[1]), A[2].concat(b[2])];
  for (var col = 0; col < 3; col++) {
    var piv = col;
    for (var rr = col; rr < 3; rr++) if (Math.abs(M[rr][col]) > Math.abs(M[piv][col])) piv = rr;
    if (Math.abs(M[piv][col]) < 1e-9) return null;
    var tmp = M[col]; M[col] = M[piv]; M[piv] = tmp;
    for (var r2 = 0; r2 < 3; r2++) {
      if (r2 === col) continue;
      var f = M[r2][col] / M[col][col];
      for (var k3 = col; k3 < 4; k3++) M[r2][k3] -= f * M[col][k3];
    }
  }
  return [M[0][3] / M[0][0], M[1][3] / M[1][1], M[2][3] / M[2][2]];
}

function median(arr) {
  if (!arr.length) return null;
  var a = arr.slice().sort(function (x, y) { return x - y; });
  var m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
}

function scoreAll(listings) {
  var byModel = {}, relByModel = {};
  listings.forEach(function (l) {
    (byModel[l.model] = byModel[l.model] || []).push(l);
    var rc = relComponent(l);
    if (rc !== null) (relByModel[l.model] = relByModel[l.model] || []).push(rc);
  });
  var coefs = {}, relFallback = {};
  Object.keys(byModel).forEach(function (m) { coefs[m] = fitModel(byModel[m]); });
  Object.keys(relByModel).forEach(function (m) { relFallback[m] = median(relByModel[m]); });

  var today = new Date();
  var W = CONFIG.defaultWeights;

  return listings.map(function (l) {
    var opts = optionSet(l);
    var price = l.currentPrice || l.listPrice || 0;
    var miles = l.mileage || 0;
    var cap = CONFIG.mileageCaps[l.make] || 100000;

    var co = coefs[l.model], pred = null, delta = null;
    if (co && price) {
      pred = co[0] + co[1] * l.year + co[2] * (miles / 1000);
      if (pred > 5000) delta = (pred - price) / pred;
    }
    var badgeS = { GREAT: 100, GOOD: 78, FAIR: 55 }[l.badge];
    if (badgeS === undefined) badgeS = 62;
    var deltaS = (delta === null) ? 62 : clamp((delta + 0.15) / 0.30 * 100, 0, 100);
    var sVal = 0.55 * badgeS + 0.45 * deltaS;

    var capS = clamp((1 - miles / cap) * 100, 0, 100);
    var age = Math.max(today.getFullYear() - l.year, 0.5);
    var annS = clamp((1 - (miles / age - 6000) / 12000) * 100, 0, 100);
    var sMile = 0.7 * capS + 0.3 * annS;

    var rc = relComponent(l), relEstimated = false;
    if (rc === null) { rc = relFallback[l.model]; relEstimated = true; if (rc == null) rc = 70; }
    var nsvc = (l.serviceHistory && l.serviceHistory.number) || 0;
    var hist = (l.noAccidents ? 34 : 0) + (l.oneOwner ? 22 : 0) +
               (l.personalUse ? 14 : 0) + Math.min(30, nsvc * 3);
    var sRel = 0.6 * rc + 0.4 * hist;

    var feats = {}, pts = 0;
    FEATURES.forEach(function (f) {
      var v = f.match(opts); feats[f.label] = v; if (v) pts += f.pts;
    });
    var heated = has(opts, 'heated seat') || has(opts, 'heated front seat');
    var sFeat = FEAT_MAX ? Math.min(100, pts / FEAT_MAX * 100) : 0;

    var dist = l.distanceToDealer || 0;
    var sDist = clamp((1 - dist / 75) * 100, 0, 100);
    var sCpo = l.certified ? 100 : 0;

    var days = null;
    if (l.firstSeen) {
      var p = l.firstSeen.split('-');
      var d0 = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
      days = Math.round((today - d0) / 86400000);
    }
    var ph = l.priceHistory || [], drop = 0;
    if (ph.length > 1 && ph[ph.length - 1].listPrice && price) drop = ph[ph.length - 1].listPrice - price;
    var dealer = l.dealer || {};
    var owners = (l.ownerHistory && l.ownerHistory.history) ? l.ownerHistory.history.length
                 : (l.oneOwner ? 1 : null);
    var docFee = 0;
    (l.fees || []).forEach(function (f) { if (f.feeType === 'document_fee') docFee += f.fee || 0; });

    var r1 = round1;
    return {
      vin: l.vin, year: l.year, make: l.make, model: l.model,
      trim: ((l.trim || '') + (l.subTrim && l.subTrim !== 'Unspecified' ? ' ' + l.subTrim : '')).trim(),
      price: price, miles: miles, drivetype: l.drivetype, fuel: l.fuel,
      ext: l.exteriorColor, int: l.interiorColor,
      engine: ((l.displacement || '') + ' ' + (l.engine || '')).trim(),
      mpg: l.mpgCombined, certified: !!l.certified, noAccidents: !!l.noAccidents,
      owners: owners, personal: !!l.personalUse, svc: nsvc,
      valueBadge: l.badge || 'None',
      relBadge: (l.reliability && l.reliability.overallReliabilityBadge) || (relEstimated ? 'est.' : 'n/a'),
      repairCost: l.reliability ? l.reliability.averageCost : null,
      feats: feats, heated: heated,
      dealer: dealer.name, city: dealer.city, state: dealer.state,
      dealerRating: dealer.dealerAverageRating,
      distance: Math.round(dist * 10) / 10, days: days, drop: drop,
      monthly: l.monthlyPaymentEstimate ? l.monthlyPaymentEstimate.monthlyPayment : null,
      docFee: docFee, predPrice: pred ? Math.round(pred) : null,
      deltaPct: delta === null ? null : Math.round(delta * 1000) / 10,
      sRel: r1(sRel), sVal: r1(sVal), sMile: r1(sMile),
      sFeat: r1(sFeat), sCpo: sCpo, sDist: r1(sDist),
      url: l.vdpUrl || ('https://www.carfax.com/vehicle/' + l.vin),
      provisional: (W.rel * sRel + W.val * sVal + W.mile * sMile +
                    W.cpo * sCpo + W.feat * sFeat + W.dist * sDist) / 100
    };
  });
}
function round1(v) { return Math.round(v * 10) / 10; }

function meetsHardCriteria(c, cr) {
  var cap = cr.caps[c.make] || 100000;
  if (c.price > cr.priceMax) return false;
  if (c.year < cr.yearMin) return false;
  if (c.distance > cr.maxDistance) return false;
  if (c.miles > cap) return false;
  if (cr.requireAwd && c.drivetype === 'FWD') return false;
  if (cr.requireNoAccidents && !c.noAccidents) return false;
  if (cr.requireService && !(c.svc > 0)) return false;
  if (cr.requireHeated && !c.heated) return false;
  if (cr.requireValue && c.valueBadge !== 'GOOD' && c.valueBadge !== 'GREAT') return false;
  if (cr.requireCpo && !c.certified) return false;
  return true;
}

// ============================== SHEET HELPERS ==============================
var SHEETS = {
  start: 'Start Here', criteria: 'Search Criteria', weights: 'Scoring Weights',
  shortlist: 'Shortlist', compare: 'Top 10 Compare', benchmarks: 'Model Benchmarks',
  all: 'All Cars'
};
var NAVY = '#1f3864', ACCENT = '#2e5c8a', YELLOW = '#fff2cc', GREY = '#f2f2f2';

function getSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function resetSheet(ss, name) {
  var sh = getSheet(ss, name);
  var f = sh.getFilter(); if (f) f.remove();
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (e) {}
  sh.clear();
  sh.clearConditionalFormatRules();
  if (sh.getMaxRows() > 1) sh.setFrozenRows(0);
  if (sh.getMaxColumns() > 1) sh.setFrozenColumns(0);
  return sh;
}
function headerRow(sh, row, values, widths) {
  sh.getRange(row, 1, 1, values.length).setValues([values])
    .setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(10)
    .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(row, 34);
  if (widths) widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}
function titleRow(sh, text, subtitle) {
  sh.getRange(1, 1).setValue(text).setFontSize(16).setFontWeight('bold')
    .setFontColor(NAVY).setFontFamily('Arial');
  if (subtitle) {
    sh.getRange(2, 1).setValue(subtitle).setFontSize(10).setFontStyle('italic')
      .setFontColor('#595959').setFontFamily('Arial');
  }
}
function orderTabs(ss) {
  var order = [SHEETS.start, SHEETS.criteria, SHEETS.weights, SHEETS.shortlist,
               SHEETS.compare, SHEETS.benchmarks, SHEETS.all];
  order.forEach(function (name, i) {
    var sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  var store = ss.getSheetByName(STORE);
  if (store) store.hideSheet();
  ss.setActiveSheet(ss.getSheetByName(SHEETS.shortlist));
}

// ============================== CRITERIA / WEIGHTS ==============================
var CRIT_LABELS = [
  ['Max Price ($)', 40000, 'Listing price, before tax/title/doc fee.'],
  ['Min Year', 2022, 'Model year.'],
  ['Max Distance (mi)', 75, 'Straight-line distance from Elkhart, IN 46514.'],
  ['Require AWD / 4WD', true, 'TRUE or FALSE.'],
  ['Require No Accidents', true, 'No accident or damage reported to CARFAX.'],
  ['Require Service History', true, 'At least one service record on file.'],
  ['Require Heated Seats', true, 'From the trim equipment list.'],
  ['Require Good/Great Value', true, "CARFAX's own value badge must be GOOD or GREAT."],
  ['Require Certified Pre-Owned', false, 'Set TRUE to see only certified cars.']
];

function ensureCriteriaSheet(ss) {
  if (ss.getSheetByName(SHEETS.criteria)) return;
  var sh = getSheet(ss, SHEETS.criteria);
  titleRow(sh, 'Search Criteria',
    'Edit the yellow cells. The "Meets All" column on every tab recalculates from these.');
  headerRow(sh, 4, ['Filter', 'Value', 'Notes'], [230, 110, 460]);
  CRIT_LABELS.forEach(function (row, i) {
    var r = 5 + i;
    sh.getRange(r, 1).setValue(row[0]).setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 2).setValue(row[1]).setBackground(YELLOW).setFontColor('#0000ff')
      .setHorizontalAlignment('center').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 3).setValue(row[2]).setFontSize(9).setFontColor('#595959').setFontFamily('Arial');
  });
  sh.getRange(5, 2).setNumberFormat('$#,##0');
  sh.getRange(16, 1).setValue('Mileage Cap by Make').setFontSize(12).setFontWeight('bold').setFontColor(NAVY);
  headerRow(sh, 17, ['Make', 'Max Miles', 'Notes']);
  ['Toyota', 'Honda', 'Subaru'].forEach(function (mk, i) {
    var r = 18 + i;
    sh.getRange(r, 1).setValue(mk).setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 2).setValue(CONFIG.mileageCaps[mk]).setBackground(YELLOW)
      .setFontColor('#0000ff').setNumberFormat('#,##0').setHorizontalAlignment('center');
    sh.getRange(r, 3).setValue('Your stated cap.').setFontSize(9).setFontColor('#595959');
  });
  sh.getRange(23, 1).setValue(
    'Apple CarPlay is standard on every 2022+ car in these five model lines, so it never eliminates anyone.')
    .setFontSize(9).setFontStyle('italic').setFontColor('#595959');
}

var WEIGHT_ROWS = [
  ['Reliability & History', 26, "CARFAX model reliability, repair cost and risk (60%) blended with this car's own history - accident-free, one owner, personal use, service records (40%)."],
  ['Price / Value', 22, 'CARFAX value badge (55%) blended with how far under the fitted market price it is listed (45%).'],
  ['Mileage', 18, 'Miles against your cap for that make (70%) blended with miles-per-year against a 12,000/yr normal (30%).'],
  ['Certified Pre-Owned', 15, '100 if factory certified, 0 if not.'],
  ['Features', 14, 'Heated steering, sunroof, blind spot, power liftgate, leather, remote start, nav, wireless charging, parking sensors, power seat, premium audio.'],
  ['Distance', 5, '100 at your doorstep down to 0 at 75 miles. Tiebreaker only.']
];

function ensureWeightsSheet(ss) {
  if (ss.getSheetByName(SHEETS.weights)) return;
  var sh = getSheet(ss, SHEETS.weights);
  titleRow(sh, 'Scoring Weights',
    'Edit the yellow cells to re-rank every car instantly. They should add to 100.');
  headerRow(sh, 4, ['Component', 'Weight', 'What it measures'], [200, 90, 620]);
  WEIGHT_ROWS.forEach(function (row, i) {
    var r = 5 + i;
    sh.getRange(r, 1).setValue(row[0]).setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 2).setValue(row[1]).setBackground(YELLOW).setFontColor('#0000ff')
      .setHorizontalAlignment('center').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 3).setValue(row[2]).setFontSize(9).setFontColor('#595959')
      .setWrap(true).setFontFamily('Arial');
  });
  sh.getRange(11, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(11, 2).setFormula('=SUM(B5:B10)').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(11, 3).setFormula(
    '=IF(B11=100,"OK - weights add to 100","ADJUST - weights must add to 100, currently "&B11)')
    .setFontWeight('bold').setFontColor('#008000');
}

function readCriteria(ss) {
  var sh = ss.getSheetByName(SHEETS.criteria);
  var v = sh.getRange(5, 2, 9, 1).getValues();
  var caps = {};
  var cv = sh.getRange(18, 1, 3, 2).getValues();
  cv.forEach(function (row) { if (row[0]) caps[String(row[0])] = Number(row[1]) || 0; });
  ['Toyota', 'Honda', 'Subaru'].forEach(function (m) {
    if (!caps[m]) caps[m] = CONFIG.mileageCaps[m];
  });
  return {
    priceMax: Number(v[0][0]) || CONFIG.priceMax,
    yearMin: Number(v[1][0]) || CONFIG.yearMin,
    maxDistance: Number(v[2][0]) || CONFIG.radius,
    requireAwd: v[3][0] === true || String(v[3][0]).toUpperCase() === 'TRUE',
    requireNoAccidents: v[4][0] === true || String(v[4][0]).toUpperCase() === 'TRUE',
    requireService: v[5][0] === true || String(v[5][0]).toUpperCase() === 'TRUE',
    requireHeated: v[6][0] === true || String(v[6][0]).toUpperCase() === 'TRUE',
    requireValue: v[7][0] === true || String(v[7][0]).toUpperCase() === 'TRUE',
    requireCpo: v[8][0] === true || String(v[8][0]).toUpperCase() === 'TRUE',
    caps: caps
  };
}

// ============================== NOTES ==============================
function hasNote(n) {
  return !!(n && ((n.Status && String(n.Status).trim()) ||
                  (n.Notes && String(n.Notes).trim()) ||
                  (n['Contacted?'] && String(n['Contacted?']).trim())));
}

/** Merge notes from the hidden store and from whatever is on screen right now. */
function harvestNotes(ss) {
  var map = {};
  var store = ss.getSheetByName(STORE);
  if (store && store.getLastRow() > 1) {
    store.getRange(2, 1, store.getLastRow() - 1, 6).getValues().forEach(function (r) {
      if (!r[0]) return;
      map[r[0]] = { vin: r[0], Status: r[1], Notes: r[2], 'Contacted?': r[3], label: r[4], url: r[5] };
    });
  }
  [SHEETS.all, SHEETS.shortlist].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 5) return;
    var head = sh.getRange(4, 1, 1, sh.getLastColumn()).getValues()[0];
    var idx = {};
    head.forEach(function (h, i) { idx[String(h)] = i; });
    if (idx['VIN'] === undefined) return;
    var body = sh.getRange(5, 1, sh.getLastRow() - 4, sh.getLastColumn()).getValues();
    body.forEach(function (row) {
      var vin = row[idx['VIN']];
      if (!vin) return;
      var entry = map[vin] || { vin: vin, Status: '', Notes: '', 'Contacted?': '', label: '', url: '' };
      NOTE_COLS.forEach(function (c) {
        if (idx[c] === undefined) return;
        var val = row[idx[c]];
        if (val !== '' && val !== null && val !== undefined) entry[c] = val;
      });
      if (idx['Year'] !== undefined && row[idx['Year']]) {
        entry.label = [row[idx['Year']], row[idx['Make']], row[idx['Model']], row[idx['Trim']]].join(' ').trim();
      }
      entry.url = 'https://www.carfax.com/vehicle/' + vin;
      map[vin] = entry;
    });
  });
  return map;
}

function saveNotes(ss, map) {
  var sh = ss.getSheetByName(STORE) || ss.insertSheet(STORE);
  sh.clear();
  sh.getRange(1, 1, 1, 6).setValues([['VIN', 'Status', 'Notes', 'Contacted?', 'Car', 'URL']])
    .setFontWeight('bold');
  var rows = [];
  Object.keys(map).forEach(function (vin) {
    var n = map[vin];
    if (hasNote(n)) rows.push([vin, n.Status || '', n.Notes || '', n['Contacted?'] || '', n.label || '', n.url || '']);
  });
  if (rows.length) sh.getRange(2, 1, rows.length, 6).setValues(rows);
  sh.hideSheet();
}

// ============================== ALL CARS ==============================
function colLetter(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}
var FEAT_LABELS = FEATURES.map(function (f) { return f.label; });
var ALL_COLS = [['Rank', 55], ['Score', 65], ['Meets All', 80], ['Status', 110], ['Notes', 220],
  ['Contacted?', 90], ['Year', 55], ['Make', 70], ['Model', 80], ['Trim', 130], ['Price', 85],
  ['Mileage', 80], ['Miles/Yr', 75], ['CPO', 55], ['Drivetrain', 85], ['Fuel', 75], ['MPG', 55],
  ['Value Badge', 90], ['Est. Market', 90], ['vs Market', 80], ['Reliability', 85], ['Repair $/yr', 85],
  ['No Accidents', 90], ['Owners', 65], ['Personal Use', 90], ['Service Recs', 90],
  ['Heated Seats', 90], ['Apple CarPlay', 95]]
  .concat(FEAT_LABELS.map(function (f) { return [f, 95]; }))
  .concat([['Ext Color', 85], ['Int Color', 80], ['Engine', 90], ['Dealer', 200], ['City', 110],
    ['ST', 45], ['Miles Away', 80], ['Dealer Rating', 90], ['Days Listed', 80], ['Price Drop', 85],
    ['Est. Monthly', 90], ['Doc Fee', 70], ['Sub: Reliability', 95], ['Sub: Value', 80],
    ['Sub: Mileage', 90], ['Sub: Features', 90], ['Sub: CPO', 75], ['Sub: Distance', 90],
    ['VIN', 160], ['Listing', 180]]);
var AIDX = {};
ALL_COLS.forEach(function (c, i) { AIDX[c[0]] = i + 1; });

function yn(b) { return b ? 'Yes' : 'No'; }

function buildAllCars(ss, cars, notes, orphans) {
  var sh = resetSheet(ss, SHEETS.all);
  titleRow(sh, 'All Cars - Full Inventory',
    'Every 2022+ 4Runner / RAV4 / HR-V / CR-V / Outback under your price cap within ' +
    CONFIG.radius + ' miles of Elkhart. Refreshed ' + new Date().toLocaleString() + '.');
  headerRow(sh, 4, ALL_COLS.map(function (c) { return c[0]; }), ALL_COLS.map(function (c) { return c[1]; }));

  var CR = "'" + SHEETS.criteria + "'", SW = "'" + SHEETS.weights + "'";
  var first = 5, n = cars.length, lastRow = first + n - 1;
  var L = function (name) { return colLetter(AIDX[name]); };
  var rows = cars.map(function (c, i) {
    var R = first + i;
    var note = notes[c.vin] || {};
    var score = '=ROUND((' + L('Sub: Reliability') + R + '*' + SW + '!$B$5+' +
      L('Sub: Value') + R + '*' + SW + '!$B$6+' + L('Sub: Mileage') + R + '*' + SW + '!$B$7+' +
      L('Sub: CPO') + R + '*' + SW + '!$B$8+' + L('Sub: Features') + R + '*' + SW + '!$B$9+' +
      L('Sub: Distance') + R + '*' + SW + '!$B$10)/MAX(' + SW + '!$B$11,1),1)';
    var rank = '=RANK(' + L('Score') + R + ',$' + L('Score') + '$' + first + ':$' + L('Score') + '$' + lastRow + ')';
    var meets = '=IF(AND(' +
      L('Price') + R + '<=' + CR + '!$B$5,' +
      L('Year') + R + '>=' + CR + '!$B$6,' +
      L('Miles Away') + R + '<=' + CR + '!$B$7,' +
      L('Mileage') + R + '<=INDEX(' + CR + '!$B$18:$B$20,MATCH(' + L('Make') + R + ',' + CR + '!$A$18:$A$20,0)),' +
      'OR(NOT(' + CR + '!$B$8),' + L('Drivetrain') + R + '<>"FWD"),' +
      'OR(NOT(' + CR + '!$B$9),' + L('No Accidents') + R + '="Yes"),' +
      'OR(NOT(' + CR + '!$B$10),' + L('Service Recs') + R + '>0),' +
      'OR(NOT(' + CR + '!$B$11),' + L('Heated Seats') + R + '="Yes"),' +
      'OR(NOT(' + CR + '!$B$12),' + L('Value Badge') + R + '="GREAT",' + L('Value Badge') + R + '="GOOD"),' +
      'OR(NOT(' + CR + '!$B$13),' + L('CPO') + R + '="Yes")),"YES","no")';
    var milesYr = '=IFERROR(ROUND(' + L('Mileage') + R + '/MAX(YEAR(TODAY())-' + L('Year') + R + ',0.5),0),"")';
    var row = [rank, score, meets, note.Status || '', note.Notes || '', note['Contacted?'] || '',
      c.year, c.make, c.model, c.trim, c.price, c.miles, milesYr, yn(c.certified), c.drivetype,
      c.fuel, c.mpg, c.valueBadge, c.predPrice, c.deltaPct === null ? '' : c.deltaPct / 100,
      c.relBadge, c.repairCost, yn(c.noAccidents), c.owners, yn(c.personal), c.svc,
      yn(c.heated), 'Yes'];
    FEAT_LABELS.forEach(function (f) { row.push(yn(c.feats[f])); });
    return row.concat([c.ext, c.int, c.engine, c.dealer, c.city, c.state, c.distance,
      c.dealerRating, c.days, c.drop || '', c.monthly, c.docFee || '',
      c.sRel, c.sVal, c.sMile, c.sFeat, c.sCpo, c.sDist, c.vin,
      '=HYPERLINK("' + c.url + '","' + (c.year + ' ' + c.make + ' ' + c.model).replace(/"/g, '') + '")']);
  });

  // cars you took notes on that are no longer listed
  orphans.forEach(function (o) {
    var blank = new Array(ALL_COLS.length).fill('');
    blank[AIDX['Meets All'] - 1] = 'DELISTED';
    blank[AIDX['Status'] - 1] = o.Status || '';
    blank[AIDX['Notes'] - 1] = o.Notes || '';
    blank[AIDX['Contacted?'] - 1] = o['Contacted?'] || '';
    blank[AIDX['Trim'] - 1] = o.label || '';
    blank[AIDX['VIN'] - 1] = o.vin;
    if (o.url) blank[AIDX['Listing'] - 1] = '=HYPERLINK("' + o.url + '","last listing")';
    rows.push(blank);
  });

  sh.getRange(first, 1, rows.length, ALL_COLS.length).setValues(rows)
    .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  formatCarBlock(sh, first, rows.length, AIDX);
  sh.setFrozenRows(4);
  sh.setFrozenColumns(3);
  sh.getRange(4, 1, rows.length + 1, ALL_COLS.length).createFilter();
  addScoreGradient(sh, first, rows.length, AIDX['Score']);
  addEqualsRule(sh, first, rows.length, AIDX['Meets All'], 'YES', '#c6efce', '#006100');
  addEqualsRule(sh, first, rows.length, AIDX['CPO'], 'Yes', '#c6efce', '#006100');
  addEqualsRule(sh, first, rows.length, AIDX['Meets All'], 'DELISTED', '#ffc7ce', '#9c0006');
  highlightNoteCols(sh, first, rows.length, AIDX);
}

function formatCarBlock(sh, first, n, idx) {
  if (!n) return;
  var money = ['Price', 'Est. Market', 'Repair $/yr', 'Price Drop', 'Est. Monthly', 'Doc Fee'];
  money.forEach(function (c) { if (idx[c]) sh.getRange(first, idx[c], n, 1).setNumberFormat('$#,##0'); });
  ['Mileage', 'Miles/Yr'].forEach(function (c) {
    if (idx[c]) sh.getRange(first, idx[c], n, 1).setNumberFormat('#,##0');
  });
  if (idx['vs Market']) sh.getRange(first, idx['vs Market'], n, 1).setNumberFormat('0.0%');
  if (idx['Miles Away']) sh.getRange(first, idx['Miles Away'], n, 1).setNumberFormat('0.0');
  if (idx['Dealer Rating']) sh.getRange(first, idx['Dealer Rating'], n, 1).setNumberFormat('0.0');
  if (idx['Score']) sh.getRange(first, idx['Score'], n, 1).setNumberFormat('0.0').setFontWeight('bold');
  ['Sub: Reliability', 'Sub: Value', 'Sub: Mileage', 'Sub: Features', 'Sub: CPO', 'Sub: Distance']
    .forEach(function (c) { if (idx[c]) sh.getRange(first, idx[c], n, 1).setNumberFormat('0.0'); });
  if (idx['Year']) sh.getRange(first, idx['Year'], n, 1).setNumberFormat('0');
  ['Rank', 'Year', 'CPO', 'Drivetrain', 'MPG', 'Value Badge', 'Reliability', 'No Accidents',
   'Owners', 'Personal Use', 'Service Recs', 'Heated Seats', 'Apple CarPlay', 'Meets All',
   'ST', 'Miles Away', 'Contacted?'].concat(FEAT_LABELS).forEach(function (c) {
    if (idx[c]) sh.getRange(first, idx[c], n, 1).setHorizontalAlignment('center');
  });
}

function highlightNoteCols(sh, first, n, idx) {
  if (!n) return;
  NOTE_COLS.forEach(function (c) {
    if (idx[c]) sh.getRange(first, idx[c], n, 1).setBackground(YELLOW);
  });
}

function addScoreGradient(sh, first, n, col) {
  if (!n) return;
  var range = sh.getRange(first, col, n, 1);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#f8696b', SpreadsheetApp.InterpolationType.MIN, '')
    .setGradientMidpointWithValue('#ffeb84', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpointWithValue('#63be7b', SpreadsheetApp.InterpolationType.MAX, '')
    .setRanges([range]).build();
  var rules = sh.getConditionalFormatRules();
  rules.push(rule);
  sh.setConditionalFormatRules(rules);
}

function addEqualsRule(sh, first, n, col, text, bg, fg) {
  if (!n || !col) return;
  var range = sh.getRange(first, col, n, 1);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setBold(true)
    .setRanges([range]).build();
  var rules = sh.getConditionalFormatRules();
  rules.push(rule);
  sh.setConditionalFormatRules(rules);
}

// ============================== SHORTLIST ==============================
var SL_COLS = [['Rank', 55], ['Score', 65], ['Still Meets', 85], ['Status', 110], ['Notes', 240],
  ['Contacted?', 90], ['Year', 55], ['Make', 70], ['Model', 80], ['Trim', 130], ['Price', 85],
  ['Mileage', 80], ['Miles/Yr', 75], ['CPO', 55], ['Value Badge', 90], ['vs Market', 80],
  ['Reliability', 85], ['Owners', 65], ['Service Recs', 90], ['Sunroof/Moonroof', 95],
  ['Leather', 75], ['Power Liftgate', 95], ['Blind Spot', 80], ['Fuel', 75], ['MPG', 55],
  ['Dealer', 200], ['City', 110], ['Miles Away', 80], ['Est. Monthly', 90], ['VIN', 160],
  ['Listing', 200]];
var SIDX = {};
SL_COLS.forEach(function (c, i) { SIDX[c[0]] = i + 1; });

function allRowFor(cars, vin) {
  for (var i = 0; i < cars.length; i++) if (cars[i].vin === vin) return 5 + i;
  return null;
}

function buildShortlist(ss, shortlist, notes) {
  var sh = resetSheet(ss, SHEETS.shortlist);
  titleRow(sh, 'Shortlist - ' + shortlist.length + ' Cars That Meet Every Criterion',
    'Ranked best value first. Certified cars are highlighted. Status, Notes and Contacted? are yours to fill in - a refresh keeps them with the car.');
  headerRow(sh, 4, SL_COLS.map(function (c) { return c[0]; }), SL_COLS.map(function (c) { return c[1]; }));
  if (!shortlist.length) {
    sh.getRange(5, 1).setValue('No cars currently meet every criterion. Loosen something on the Search Criteria tab.');
    return;
  }
  var AC = "'" + SHEETS.all + "'";
  var first = 5, n = shortlist.length, lastRow = first + n - 1;
  var SC = colLetter(SIDX['Score']);
  var rows = shortlist.map(function (c, i) {
    var R = first + i;
    var note = notes[c.vin] || {};
    var ar = c.allRow;
    return [
      '=RANK(' + SC + R + ',$' + SC + '$' + first + ':$' + SC + '$' + lastRow + ')',
      '=' + AC + '!' + colLetter(AIDX['Score']) + ar,
      '=' + AC + '!' + colLetter(AIDX['Meets All']) + ar,
      note.Status || '', note.Notes || '', note['Contacted?'] || '',
      c.year, c.make, c.model, c.trim, c.price, c.miles,
      '=IFERROR(ROUND(' + colLetter(SIDX['Mileage']) + R + '/MAX(YEAR(TODAY())-' + colLetter(SIDX['Year']) + R + ',0.5),0),"")',
      yn(c.certified), c.valueBadge, c.deltaPct === null ? '' : c.deltaPct / 100,
      c.relBadge, c.owners, c.svc, yn(c.feats['Sunroof/Moonroof']), yn(c.feats['Leather']),
      yn(c.feats['Power Liftgate']), yn(c.feats['Blind Spot']), c.fuel, c.mpg,
      c.dealer, c.city, c.distance, c.monthly, c.vin,
      '=HYPERLINK("' + c.url + '","' + (c.year + ' ' + c.make + ' ' + c.model + ' ' + c.trim).replace(/"/g, '').trim() + '")'
    ];
  });
  sh.getRange(first, 1, n, SL_COLS.length).setValues(rows)
    .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  formatCarBlock(sh, first, n, SIDX);
  sh.getRange(first, SIDX['Still Meets'], n, 1).setHorizontalAlignment('center');
  sh.setFrozenRows(4);
  sh.setFrozenColumns(3);
  sh.getRange(4, 1, n + 1, SL_COLS.length).createFilter();
  addScoreGradient(sh, first, n, SIDX['Score']);
  addEqualsRule(sh, first, n, SIDX['CPO'], 'Yes', '#c6efce', '#006100');
  addEqualsRule(sh, first, n, SIDX['Still Meets'], 'no', '#ffc7ce', '#9c0006');
  highlightNoteCols(sh, first, n, SIDX);
}

// ============================== TOP 10 COMPARE ==============================
var CMP_ATTRS = ['Score', 'Year', 'Make', 'Model', 'Trim', 'Price', 'Mileage', 'Miles/Yr',
  'Est. Monthly', 'CPO', 'Value Badge', 'Est. Market', 'vs Market', 'Reliability', 'Repair $/yr',
  'No Accidents', 'Owners', 'Personal Use', 'Service Recs', 'Drivetrain', 'Fuel', 'MPG',
  'Heated Seats', 'Heated Steering', 'Sunroof/Moonroof', 'Leather', 'Power Liftgate', 'Blind Spot',
  'Remote Start', 'Navigation', 'Wireless Charging', 'Ext Color', 'Int Color', 'Dealer', 'City',
  'Miles Away', 'Dealer Rating', 'Days Listed', 'Price Drop', 'Status', 'Notes'];

function buildCompare(ss, top) {
  var sh = resetSheet(ss, SHEETS.compare);
  titleRow(sh, 'Top 10 Head-to-Head',
    'The ten highest-scoring cars from the Shortlist, side by side. Every cell is linked to All Cars.');
  if (!top.length) { sh.getRange(4, 1).setValue('Nothing to compare yet.'); return; }
  var AC = "'" + SHEETS.all + "'";
  var head = ['Attribute'].concat(top.map(function (c, i) {
    return '#' + (i + 1) + '  ' + c.year + ' ' + c.make + ' ' + c.model + ' ' + c.trim;
  }));
  headerRow(sh, 4, head, [150].concat(top.map(function () { return 150; })));
  var rows = CMP_ATTRS.map(function (attr) {
    return [attr].concat(top.map(function (c) {
      return '=' + AC + '!' + colLetter(AIDX[attr]) + c.allRow;
    }));
  });
  rows.push(['Listing'].concat(top.map(function (c) {
    return '=HYPERLINK("' + c.url + '","Open CARFAX listing")';
  })));
  sh.getRange(5, 1, rows.length, head.length).setValues(rows)
    .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
  sh.getRange(5, 1, rows.length, 1).setFontWeight('bold').setBackground(GREY)
    .setHorizontalAlignment('left');
  var f = function (attr, fmt) {
    var i = CMP_ATTRS.indexOf(attr);
    if (i >= 0) sh.getRange(5 + i, 2, 1, top.length).setNumberFormat(fmt);
  };
  ['Price', 'Est. Monthly', 'Est. Market', 'Repair $/yr', 'Price Drop'].forEach(function (a) { f(a, '$#,##0'); });
  ['Mileage', 'Miles/Yr'].forEach(function (a) { f(a, '#,##0'); });
  f('vs Market', '0.0%'); f('Score', '0.0'); f('Year', '0');
  f('Miles Away', '0.0'); f('Dealer Rating', '0.0');
  sh.setFrozenRows(4); sh.setFrozenColumns(1);
  addScoreGradientRow(sh, 5, top.length);
}

function addScoreGradientRow(sh, row, n) {
  var range = sh.getRange(row, 2, 1, n);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#f8696b', SpreadsheetApp.InterpolationType.MIN, '')
    .setGradientMidpointWithValue('#ffeb84', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpointWithValue('#63be7b', SpreadsheetApp.InterpolationType.MAX, '')
    .setRanges([range]).build();
  var rules = sh.getConditionalFormatRules(); rules.push(rule); sh.setConditionalFormatRules(rules);
}

// ============================== MODEL BENCHMARKS ==============================
function buildBenchmarks(ss, cars) {
  var sh = resetSheet(ss, SHEETS.benchmarks);
  titleRow(sh, 'Model Benchmarks',
    'What each model-year actually lists for in this market. Live formulas over the All Cars tab.');
  var cols = ['Make', 'Model', 'Year', 'Cars Listed', 'Avg Price', 'Min Price', 'Max Price',
              'Avg Mileage', 'Avg Score', 'CPO Count', 'Meets-All Count'];
  headerRow(sh, 4, cols, [90, 90, 60, 85, 95, 95, 95, 95, 80, 85, 110]);
  var seen = {}, combos = [];
  cars.forEach(function (c) {
    var k = c.make + '|' + c.model + '|' + c.year;
    if (!seen[k]) { seen[k] = true; combos.push([c.make, c.model, c.year]); }
  });
  combos.sort(function (a, b) {
    return a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]) || a[2] - b[2];
  });
  var AC = "'" + SHEETS.all + "'";
  var last = 4 + cars.length;
  var rng = function (name) {
    return AC + '!$' + colLetter(AIDX[name]) + '$5:$' + colLetter(AIDX[name]) + '$' + last;
  };
  var rows = combos.map(function (cb, i) {
    var R = 5 + i;
    var crit = rng('Make') + ',$A' + R + ',' + rng('Model') + ',$B' + R + ',' + rng('Year') + ',$C' + R;
    return [cb[0], cb[1], cb[2],
      '=COUNTIFS(' + crit + ')',
      '=IFERROR(AVERAGEIFS(' + rng('Price') + ',' + crit + '),"")',
      '=IFERROR(MINIFS(' + rng('Price') + ',' + crit + '),"")',
      '=IFERROR(MAXIFS(' + rng('Price') + ',' + crit + '),"")',
      '=IFERROR(AVERAGEIFS(' + rng('Mileage') + ',' + crit + '),"")',
      '=IFERROR(AVERAGEIFS(' + rng('Score') + ',' + crit + '),"")',
      '=COUNTIFS(' + crit + ',' + rng('CPO') + ',"Yes")',
      '=COUNTIFS(' + crit + ',' + rng('Meets All') + ',"YES")'];
  });
  sh.getRange(5, 1, rows.length, cols.length).setValues(rows)
    .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
  sh.getRange(5, 5, rows.length, 3).setNumberFormat('$#,##0');
  sh.getRange(5, 8, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(5, 9, rows.length, 1).setNumberFormat('0.0');
  sh.setFrozenRows(4);
  sh.getRange(4, 1, rows.length + 1, cols.length).createFilter();
}

// ============================== START HERE ==============================
function buildStartHere(ss, cars, shortlist, orphanCount) {
  var sh = resetSheet(ss, SHEETS.start);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 20); sh.setColumnWidth(2, 250); sh.setColumnWidth(3, 620); sh.setColumnWidth(4, 90);

  var AC = "'" + SHEETS.all + "'";
  var last = 4 + cars.length + orphanCount;
  var meetsRng = AC + '!$' + colLetter(AIDX['Meets All']) + '$5:$' + colLetter(AIDX['Meets All']) + '$' + last;
  var cpoRng = AC + '!$' + colLetter(AIDX['CPO']) + '$5:$' + colLetter(AIDX['CPO']) + '$' + last;
  var vinRng = AC + '!$' + colLetter(AIDX['VIN']) + '$5:$' + colLetter(AIDX['VIN']) + '$' + last;

  var put = function (r, c, v) { return sh.getRange(r, c).setValue(v).setFontFamily('Arial'); };
  var h2 = function (r, t) {
    sh.getRange(r, 2).setValue(t).setFontSize(12).setFontWeight('bold').setFontColor(ACCENT).setFontFamily('Arial');
    sh.getRange(r, 2, 1, 2).setBorder(null, null, true, null, null, null, ACCENT,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  };
  var note = function (r, t) {
    sh.getRange(r, 2, 1, 2).merge();
    sh.getRange(r, 2).setValue(t).setFontSize(9).setFontStyle('italic')
      .setFontColor('#595959').setWrap(true).setFontFamily('Arial');
  };

  put(2, 2, "Bella's Car Search").setFontSize(16).setFontWeight('bold').setFontColor(NAVY);
  put(3, 2, 'Live CARFAX inventory - 2022+ Toyota / Honda / Subaru SUVs near Elkhart, IN')
    .setFontSize(11).setFontStyle('italic').setFontColor('#595959');
  put(4, 2, 'Last refreshed').setFontWeight('bold');
  put(4, 3, new Date().toLocaleString());
  put(5, 2, 'Search area').setFontWeight('bold');
  put(5, 3, 'Elkhart, IN ' + CONFIG.zip + ' - ' + CONFIG.radius + ' mile radius');
  put(6, 2, 'To refresh').setFontWeight('bold');
  put(6, 3, 'Car Search menu (next to Help) > Refresh from CARFAX now');

  h2(8, 'Live Counts');
  var counts = [
    ['Cars meeting every criterion', '=COUNTIF(' + meetsRng + ',"YES")'],
    ['Of those, Certified Pre-Owned', '=COUNTIFS(' + meetsRng + ',"YES",' + cpoRng + ',"Yes")'],
    ['Total listings tracked', '=COUNTA(' + vinRng + ')']
  ];
  counts.forEach(function (row, i) {
    put(9 + i, 2, row[0]).setFontWeight('bold');
    sh.getRange(9 + i, 3).setFormula(row[1]).setFontWeight('bold').setFontFamily('Arial');
  });
  note(12, 'These recalculate from the Search Criteria tab, so you can see how many cars survive a change before going looking.');

  h2(14, 'The Tabs');
  var tabs = [
    ['Shortlist', 'The cars meeting every one of your criteria, ranked best first. Start here. Status, Notes and Contacted? are yours - they survive a refresh.'],
    ['Top 10 Compare', 'The ten best side by side for a quick head-to-head.'],
    ['All Cars', 'Every listing found, all data, filterable. Notes here are kept too.'],
    ['Search Criteria', 'Yellow cells that drive the YES/no "Meets All" column.'],
    ['Scoring Weights', 'Yellow cells that drive the Score. Change one and everything re-ranks.'],
    ['Model Benchmarks', 'What each model-year actually lists for here.']
  ];
  tabs.forEach(function (t, i) {
    put(15 + i, 2, t[0]).setFontWeight('bold').setFontColor(ACCENT);
    put(15 + i, 3, t[1]).setWrap(true);
  });

  h2(22, 'How the Score Works');
  note(23, 'Six sub-scores from 0-100, blended with the weights you set. The weighting follows your priority order: reliability, then price, then mileage, then features.');
  sh.getRange(24, 2, 1, 3).setValues([['Component', 'What it measures', 'Weight']])
    .setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold').setFontFamily('Arial');
  WEIGHT_ROWS.forEach(function (w, i) {
    var r = 25 + i;
    put(r, 2, w[0]).setFontWeight('bold');
    put(r, 3, w[2]).setWrap(true).setFontSize(10);
    sh.getRange(r, 4).setFormula("='" + SHEETS.weights + "'!B" + (5 + i) + '/100')
      .setNumberFormat('0%').setFontWeight('bold').setHorizontalAlignment('center');
  });

  h2(32, 'Before You Buy');
  var cav = [
    ['Verify CPO with the dealer', 'CARFAX flags certification from the dealer feed. Confirm it is manufacturer certified (Toyota TCUV, HondaTrue, Subaru Certified) and not an in-house dealer warranty - the warranty difference is significant.'],
    ['Features come from the trim', 'Heated seats, sunroof and the rest are read from the factory equipment list for that trim. Accurate for standard equipment, but it can miss optional packages. Confirm on the window sticker.'],
    ['"Est. Market" is our estimate', 'Fitted across the listings in this sheet, adjusting for year and mileage within each model. A sanity check on asking price, not an appraisal.'],
    ['The value badge is the tightest filter', 'Requiring GOOD or GREAT value removes more cars than any other requirement. A FAIR car is not a bad car, just not a bargain. Turning that off on the Search Criteria tab is the biggest way to widen the search.'],
    ['Prices move', 'Listings change daily and good ones go fast. "Days Listed" and "Price Drop" show which have been sitting - those have the most negotiating room.'],
    ['Get a pre-purchase inspection', 'Even on a certified car with clean history, an independent inspection is the best $150 you will spend.']
  ];
  cav.forEach(function (c, i) {
    put(33 + i, 2, c[0]).setFontWeight('bold').setWrap(true);
    put(33 + i, 3, c[1]).setWrap(true);
  });
  if (orphanCount) {
    note(40, orphanCount + ' car(s) you had notes on are no longer listed. They are kept at the bottom of All Cars marked DELISTED so your notes are not lost.');
  }
}
