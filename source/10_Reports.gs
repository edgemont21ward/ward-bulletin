/**
 * HISTORY & REPORTS
 * Every publish records who spoke and who prayed that Sunday on a
 * hidden "History" tab, and "Ward Bulletin > Reports > Speakers /
 * Prayers" reads it back as a report (Report.html).
 *
 * History is read out of the published bulletin itself, the same HTML
 * that went to GitHub Pages, not out of the week's tab: that's exactly
 * what the ward saw, and it means the first fill (from the bulletins
 * already on GitHub, see seedHistoryFromGithub_) and every publish after
 * it go through the same parser, so they can never disagree.
 *
 * The History tab has one row per person per Sunday: Sunday
 * ("yyyy-MM-dd", stored as text so Sheets can't shift it a day), Part
 * (the bulletin's label: "Opening Prayer", "Closing Prayer", "Speaker",
 * "Youth Speaker"...), and Name as printed. Republishing a Sunday
 * replaces that Sunday's rows, so nothing is counted twice. It's a
 * reference tab like Songs or Members (hidden and shown with them), and
 * Archive Past Weeks never touches it, so history outlives the week
 * tabs. Rows can be fixed by hand.
 */


/** Which report a bulletin line belongs to: 'speaker', 'prayer', or '' for neither. */
function historyKindOf_(part) {
  var p = String(part || '').toLowerCase();
  if (p.indexOf('speaker') !== -1) return 'speaker';
  if (p === 'opening prayer' || p === 'closing prayer') return 'prayer';
  return '';
}


/**
 * The Sunday a rendered bulletin is for ("yyyy-MM-dd", '' if it has no
 * readable date) and its speaker and prayer lines, in bulletin order:
 * { sunday, rows: [{ part, name }] }. Lines with no name are skipped.
 */
function participationFromHtml_(html) {
  var rows = [];
  var re = /<span class="left">([^<]*)<\/span>\s*<span class="right">([^<]*)<\/span>/g;
  var m;
  while ((m = re.exec(String(html || ''))) !== null) {
    var part = decodeHtmlEntities_(m[1]).replace(/\s+/g, ' ').trim();
    var name = decodeHtmlEntities_(m[2]).replace(/\s+/g, ' ').trim();
    if (name && historyKindOf_(part)) rows.push({ part: part, name: name });
  }
  return { sunday: bulletinDateSlugFromHtml_(html), rows: rows };
}


/**
 * Called by publishBulletinForSidebar with the HTML it just published:
 * replaces that Sunday's rows on the History tab (making the tab, and
 * filling in earlier Sundays from GitHub, the first time).
 */
function recordHistoryFromHtml_(html, token) {
  var p = participationFromHtml_(html);
  if (!p.sunday) return;
  withBulletinTabLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = historySheet_(ss, token);
    var rows = readHistory_(sheet).filter(function (r) { return r.sunday !== p.sunday; });
    p.rows.forEach(function (r) { rows.push({ sunday: p.sunday, part: r.part, name: r.name }); });
    writeHistory_(sheet, rows);
  });
}


/**
 * The History tab, creating it if there isn't one yet: hidden if the
 * other reference tabs are, and filled in from every bulletin already
 * published to GitHub (see seedHistoryFromGithub_). Creating a tab
 * switches to it, so this switches back to wherever you were.
 */
function historySheet_(ss, token) {
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME_);
  if (sheet) return sheet;

  var hide = areUtilityTabsHidden_(); // before the new tab exists, or it would count as a visible one
  var wasActive = ss.getActiveSheet();
  sheet = ss.insertSheet(HISTORY_SHEET_NAME_, ss.getSheets().length);
  sheet.getRange(1, 1, 1, 3).setValues([['Sunday', 'Part', 'Name']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('@'); // Sundays stay text, never re-dated
  if (wasActive) ss.setActiveSheet(wasActive);
  if (hide) sheet.hideSheet();

  try {
    writeHistory_(sheet, seedHistoryFromGithub_(token));
  } catch (err) {
    // The tab still works without the backfill; publishes add to it from here on.
  }
  return sheet;
}


/**
 * History rows for every bulletin already on GitHub: each file in
 * GITHUB_ARCHIVE_DIR, then the live GITHUB_FILE_PATH, which wins for
 * its Sunday (a Sunday republished after being archived has both).
 */
function seedHistoryFromGithub_(token) {
  var base = 'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/contents/';
  var opts = { method: 'get', muteHttpExceptions: true, headers: token ? githubHeaders_(token) : { Accept: 'application/vnd.github+json' } };

  var listing = UrlFetchApp.fetch(base + GITHUB_ARCHIVE_DIR + '?ref=' + GITHUB_BRANCH, opts);
  var paths = listing.getResponseCode() === 200
    ? JSON.parse(listing.getContentText())
        .filter(function (f) { return /\.html$/i.test(f.name); })
        .map(function (f) { return f.path; })
    : [];
  paths.push(GITHUB_FILE_PATH);

  var bySunday = {};
  UrlFetchApp.fetchAll(paths.map(function (path) {
    var req = { url: base + path + '?ref=' + GITHUB_BRANCH };
    for (var k in opts) req[k] = opts[k];
    return req;
  })).forEach(function (resp) {
    if (resp.getResponseCode() !== 200) return;
    var p = participationFromHtml_(decodeGithubBase64_(JSON.parse(resp.getContentText()).content));
    if (p.sunday) bySunday[p.sunday] = p.rows; // later files (the live one last) win
  });

  var rows = [];
  Object.keys(bySunday).forEach(function (sunday) {
    bySunday[sunday].forEach(function (r) { rows.push({ sunday: sunday, part: r.part, name: r.name }); });
  });
  return rows;
}


/** The History tab's rows as [{ sunday, part, name }], skipping blank ones. */
function readHistory_(sheet) {
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) { // row 1 is the header
    var v = values[i][0];
    var sunday = Object.prototype.toString.call(v) === '[object Date]'
      ? Utilities.formatDate(v, tz, 'yyyy-MM-dd')                 // someone typed a real date
      : (parseBulletinDateText_(String(v)) || String(v).trim()); // "yyyy-MM-dd", or a written-out date
    var part = String(values[i][1] == null ? '' : values[i][1]).trim();
    var name = String(values[i][2] == null ? '' : values[i][2]).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(sunday) && part && name) out.push({ sunday: sunday, part: part, name: name });
  }
  return out;
}


/**
 * Rewrites the History tab's rows: newest Sunday first, and within a
 * Sunday the order they were given in (bulletin order).
 */
function writeHistory_(sheet, rows) {
  var sorted = rows
    .map(function (r, i) { return { r: r, i: i }; })
    .sort(function (a, b) { return a.r.sunday < b.r.sunday ? 1 : a.r.sunday > b.r.sunday ? -1 : a.i - b.i; })
    .map(function (x) { return [x.r.sunday, x.r.part, x.r.name]; });

  var old = sheet.getLastRow() - 1;
  if (old > 0) sheet.getRange(2, 1, old, 3).clearContent();
  if (sorted.length) {
    var range = sheet.getRange(2, 1, sorted.length, 3);
    sheet.getRange(2, 1, sorted.length, 1).setNumberFormat('@');
    range.setValues(sorted);
  }
}


/** Menu actions for "Reports". */
function showSpeakersReport() { showReport_('speaker'); }
function showPrayersReport() { showReport_('prayer'); }


/** Opens the Speakers or Prayers report (Report.html) with its data built in. */
function showReport_(kind) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  var history = [];
  withBulletinTabLock_(function () { history = readHistory_(historySheet_(ss, token)); });

  var data = summarizeParticipation_(readMembersForReport_(ss), history, kind,
    parseBulletinDateText_(bulletinTabName_(ss)));
  data.kind = kind;

  var tmpl = HtmlService.createTemplateFromFile('Report');
  tmpl.reportJson = jsonForScript_(data);
  var html = tmpl.evaluate().setWidth(760).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, kind === 'speaker' ? 'Speakers' : 'Prayers');
}


/** The Members tab as [{ name: "First Middle Last", sortName: "Last, First Middle", age, gender }]. */
function readMembersForReport_(ss) {
  var sheet = ss.getSheetByName(MEMBERS_SHEET_NAME_);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) { // row 1 is the header
    var sortName = String(values[i][0] == null ? '' : values[i][0]).replace(/\s+/g, ' ').trim();
    if (!sortName) continue;
    var comma = sortName.indexOf(',');
    var name = comma === -1 ? sortName
      : (sortName.slice(comma + 1).trim() + ' ' + sortName.slice(0, comma).trim()).trim();
    var age = Number(values[i][2]);
    out.push({ name: name, sortName: sortName, age: isNaN(age) || values[i][2] === '' ? null : age,
               gender: String(values[i][1] == null ? '' : values[i][1]).trim().toUpperCase() });
  }
  return out;
}


/**
 * Two ways to compare a name: `full`, lowercased with titles
 * ("Bishop", "President", "Brother", "Sister", "Elder") and punctuation
 * dropped; and `firstLast`, just its first and last words. History
 * holds names as printed, often without a middle name ("Lucia Montie")
 * or with a title ("President Neil Lundberg"), where the Members tab
 * has "Montie, Lucia Grace".
 */
function nameKeys_(name) {
  var s = String(name || '').toLowerCase().replace(/[.,]/g, ' ').replace(/\s+/g, ' ').trim()
    .replace(/^(bishop|president|brother|sister|elder|bro|sis) /, '');
  var words = s.split(' ');
  return { full: s, firstLast: words.length > 1 ? words[0] + ' ' + words[words.length - 1] : s };
}


/** Whole weeks from `fromSunday` to `toSunday`, both "yyyy-MM-dd". */
function weeksBetween_(fromSunday, toSunday) {
  var ms = function (d) { var p = d.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]); };
  return Math.round((ms(toSunday) - ms(fromSunday)) / (7 * 86400000));
}


/**
 * Everything the report page shows, for `kind` ('speaker' or 'prayer'):
 *   people:       every member, with { last, lastPart, times, weeksAgo },
 *                 never-yet first, then longest ago first
 *   notOnMembers: people in the history who aren't on the Members tab
 *                 (visitors, stake leaders), most recent first
 *   weeks:        [{ sunday, parts: [{ part, name }] }], newest first
 *   historyStart: the earliest Sunday in the history ('' if it's empty)
 *   thisWeek:     this week's Sunday, which "weeks ago" counts back from
 *
 * A history name is matched to a member by the whole name first, then
 * by first and last word — but only when exactly one member has that
 * first and last word, so two Smiths are never merged by guesswork.
 */
function summarizeParticipation_(members, history, kind, thisWeek) {
  var byFull = {}, byFirstLast = {}, ambiguous = {};
  members.forEach(function (m, i) {
    var k = nameKeys_(m.name);
    byFull[k.full] = i;
    if (k.firstLast in byFirstLast && byFirstLast[k.firstLast] !== i) ambiguous[k.firstLast] = true;
    else byFirstLast[k.firstLast] = i;
  });

  var stats = members.map(function () { return { last: '', lastPart: '', times: 0 }; });
  var others = {};
  var entries = history.filter(function (h) { return historyKindOf_(h.part) === kind; });
  entries.forEach(function (h) {
    var k = nameKeys_(h.name);
    var i = k.full in byFull ? byFull[k.full]
      : (!ambiguous[k.firstLast] && k.firstLast in byFirstLast ? byFirstLast[k.firstLast] : -1);
    var s = i !== -1 ? stats[i] : (others[k.full] = others[k.full] || { name: h.name, last: '', lastPart: '', times: 0 });
    s.times++;
    if (h.sunday > s.last) { s.last = h.sunday; s.lastPart = h.part; }
  });

  var withWeeks = function (s) {
    s.weeksAgo = s.last ? weeksBetween_(s.last, thisWeek) : null;
    return s;
  };
  var people = members.map(function (m, i) {
    return withWeeks({ name: m.name, sortName: m.sortName, age: m.age, gender: m.gender,
                       last: stats[i].last, lastPart: stats[i].lastPart, times: stats[i].times });
  }).sort(function (a, b) {
    if (!a.last !== !b.last) return a.last ? 1 : -1;                // never-yet first
    if (a.last !== b.last) return a.last < b.last ? -1 : 1;         // then longest ago
    return a.sortName.toLowerCase() < b.sortName.toLowerCase() ? -1 : 1;
  });

  var notOnMembers = Object.keys(others).map(function (k) { return withWeeks(others[k]); })
    .sort(function (a, b) { return a.last < b.last ? 1 : -1; });

  var weekMap = {}, order = [];
  entries.forEach(function (h) {
    if (!weekMap[h.sunday]) { weekMap[h.sunday] = []; order.push(h.sunday); }
    weekMap[h.sunday].push({ part: h.part, name: h.name });
  });
  var weeks = order.sort().reverse().map(function (s) { return { sunday: s, parts: weekMap[s] }; });

  var start = history.reduce(function (min, h) { return !min || h.sunday < min ? h.sunday : min; }, '');
  return { people: people, notOnMembers: notOnMembers, weeks: weeks, historyStart: start, thisWeek: thisWeek };
}
