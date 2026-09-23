/**
 * SONGS TAB UPDATES
 * "Ward Bulletin > Update Songs" checks the Church's own table of
 * contents for Hymns for Home and Church and adds any hymn the Songs
 * tab doesn't have yet, so a newly released batch shows up in the song
 * dropdowns without anyone typing it in.
 *
 * Only that collection is checked. Hymns (1985) and the Children's
 * Songbook are complete; Hymns for Home and Church is still being
 * released a batch at a time.
 *
 * The source is the content API the Gospel Library website itself
 * reads from: one request for the collection's table of contents,
 * which lists titles but not numbers, then one request per hymn that
 * isn't on the tab yet, to read its number off its own page. A run
 * with nothing new makes just the one request.
 *
 * It only ever adds rows. Nothing already on the Songs tab is changed
 * or removed, and a hymn whose number is already on the tab under a
 * different title is reported rather than added — that's a hymn the
 * Church has retitled, not a new one, and which title to keep is a
 * judgment call for whoever maintains the tab.
 */

var HFHC_HYMNAL_NAME_ = 'Hymns for Home and Church'; // column A on the Songs tab
var HFHC_URI_ = '/music/hymns-for-home-and-church';
var CHURCH_CONTENT_API_ =
  'https://www.churchofjesuschrist.org/study/api/v3/language-pages/type/content?lang=eng&uri=';


/** Menu action: adds any newly released hymns to the Songs tab — see updateSongs_. */
function updateSongsFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var result;
  try {
    result = updateSongs_();
  } catch (err) {
    ui.alert('Could not update Songs', err.message || String(err), ui.ButtonSet.OK);
    return;
  }

  var lines = result.added.map(function (song) {
    return '#' + song.number + ' - ' + song.title;
  });
  if (result.skipped.length) {
    if (lines.length) lines.push('');
    lines.push('Not added:');
    lines = lines.concat(result.skipped);
  }

  if (!result.added.length && !result.skipped.length) {
    ui.alert(
      'Songs are up to date',
      'All ' + result.total + ' hymns in ' + HFHC_HYMNAL_NAME_ + ' are already on the Songs tab.',
      ui.ButtonSet.OK);
  } else {
    ui.alert(
      result.added.length
        ? 'Added ' + result.added.length + (result.added.length === 1 ? ' song' : ' songs')
        : 'No songs added',
      lines.join('\n'),
      ui.ButtonSet.OK);
  }
}


/**
 * Adds every hymn in the Church's Hymns for Home and Church table of
 * contents that the Songs tab doesn't already have, each inserted in
 * number order among that hymnal's existing rows. Then rebuilds the
 * song dropdown's helper column and re-applies the dropdowns to this
 * week's tab, so the new songs can be picked straight away.
 *
 * Titles are compared ignoring case, spacing, and curly-vs-straight
 * quotes: the Church site writes "God’s Gracious Love", the Songs tab
 * "God's Gracious Love". New titles are written with straight quotes,
 * the way the rest of the tab (and anyone typing into it) spells them.
 *
 * Returns { added: [{number, title}], skipped: [String], total: Number },
 * where total is how many hymns the Church's table of contents lists.
 * Throws if the Songs tab is missing or the Church site can't be read.
 */
function updateSongs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SONGS_SHEET_NAME_);
  if (!sheet) throw new Error('There is no "' + SONGS_SHEET_NAME_ + '" tab to update.');

  var contents = fetchHfhcContents_();

  var rows = sheet.getDataRange().getValues();
  var knownTitles = {};
  var knownNumbers = {};
  for (var i = 1; i < rows.length; i++) { // row 1 is the header
    if (String(rows[i][0]).trim() !== HFHC_HYMNAL_NAME_) continue;
    knownTitles[normalizeSongTitle_(rows[i][2])] = true;
    knownNumbers[String(rows[i][1]).trim()] = String(rows[i][2]).trim();
  }

  var candidates = contents.filter(function (hymn) {
    return !knownTitles[normalizeSongTitle_(hymn.title)];
  });

  var added = [];
  var skipped = [];
  var numbers = fetchSongNumbers_(candidates);
  candidates.forEach(function (hymn, k) {
    var number = numbers[k];
    if (number === null) {
      skipped.push(hymn.title + ' (no hymn number on its page)');
    } else if (knownNumbers[String(number)] !== undefined) {
      skipped.push('#' + number + ' - ' + hymn.title + ' (the tab has #' + number +
        ' as "' + knownNumbers[String(number)] + '")');
    } else {
      added.push({ number: number, title: hymn.title });
    }
  });

  added.sort(function (a, b) { return a.number - b.number; });
  added.forEach(function (song) { insertSongRow_(sheet, song); });

  if (added.length) {
    getSongDropdownRange_(ss); // rebuild the helper column the song dropdowns read
    try {
      var thisWeek = ss.getSheetByName(bulletinTabName_(ss));
      if (thisWeek) refreshDropdowns_(thisWeek);
    } catch (err) {
      // Non-critical — the new songs are on the tab either way, and
      // "Refresh Dropdowns" (or reopening the file) picks them up.
    }
  }

  return { added: added, skipped: skipped, total: contents.length };
}


/**
 * Inserts one Hymns for Home and Church row into the Songs tab, just
 * before the first row of that hymnal with a higher number, or just
 * after its last row if there's none higher — so the hymnal stays in
 * number order, and the other hymnals' rows are never moved. Only
 * columns A-C are written; column D is rebuilt afterwards from them.
 */
function insertSongRow_(sheet, song) {
  var rows = sheet.getDataRange().getValues();
  var lastHfhcRow = 0;
  var insertAt = 0;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== HFHC_HYMNAL_NAME_) continue;
    lastHfhcRow = i + 1; // 1-based sheet row
    if (!insertAt && Number(rows[i][1]) > song.number) insertAt = i + 1;
  }

  var row;
  if (insertAt) {
    sheet.insertRowBefore(insertAt);
    row = insertAt;
  } else if (lastHfhcRow) {
    sheet.insertRowAfter(lastHfhcRow);
    row = lastHfhcRow + 1;
  } else {
    row = sheet.getLastRow() + 1; // no rows for this hymnal yet
  }
  sheet.getRange(row, 1, 1, 3).setValues([[HFHC_HYMNAL_NAME_, song.number, song.title]]);
}


/**
 * Returns [{uri, title}] for every hymn in the Church's Hymns for Home
 * and Church table of contents, in the order it lists them. Throws if
 * the request fails or the page has no hymn links at all — that means
 * the Church site has changed shape, and saying "up to date" would hide
 * it.
 */
function fetchHfhcContents_() {
  var body = fetchChurchContentBody_(HFHC_URI_);

  var hymns = [];
  var seen = {};
  var linkRe = /<a\b[^>]*href="\/study(\/music\/hymns-for-home-and-church\/[^"?#]+)[^"]*"[^>]*>([\s\S]*?)<\/a>/g;
  var m;
  while ((m = linkRe.exec(body)) !== null) {
    var title = straightenQuotes_(decodeHtmlEntities_(m[2].replace(/<[^>]+>/g, ''))).replace(/\s+/g, ' ').trim();
    if (!title || seen[m[1]]) continue;
    seen[m[1]] = true;
    hymns.push({ uri: m[1], title: title });
  }

  if (!hymns.length) {
    throw new Error('The Church website\'s list of ' + HFHC_HYMNAL_NAME_ +
      ' had no hymns in it. Its layout may have changed; nothing was added.');
  }
  return hymns;
}


/**
 * Fetches each hymn's page in `hymns` (all in parallel) and returns its
 * number in the same order — e.g. 1001 — or null where a page has no
 * number to read.
 */
function fetchSongNumbers_(hymns) {
  if (!hymns.length) return [];

  var responses = UrlFetchApp.fetchAll(hymns.map(function (hymn) {
    return { url: CHURCH_CONTENT_API_ + encodeURIComponent(hymn.uri), muteHttpExceptions: true };
  }));

  return responses.map(function (resp) {
    if (resp.getResponseCode() !== 200) return null;
    var body;
    try {
      body = JSON.parse(resp.getContentText()).content.body;
    } catch (err) {
      return null;
    }
    var m = /<p\b[^>]*class="song-number"[^>]*>\s*(\d+)\s*<\/p>/.exec(body || '');
    return m ? Number(m[1]) : null;
  });
}


/** Returns the HTML body of `uri` from the Church's content API, or throws. */
function fetchChurchContentBody_(uri) {
  var resp = UrlFetchApp.fetch(CHURCH_CONTENT_API_ + encodeURIComponent(uri), { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('The Church website returned an error (' + resp.getResponseCode() +
      ') for ' + uri + '. Nothing was added; try again later.');
  }
  var body = JSON.parse(resp.getContentText()).content;
  body = body && body.body;
  if (!body) throw new Error('The Church website returned an empty page for ' + uri + '.');
  return body;
}


/** For comparing titles: straight quotes, single spaces, lowercase. */
function normalizeSongTitle_(title) {
  return straightenQuotes_(String(title == null ? '' : title)).replace(/\s+/g, ' ').trim().toLowerCase();
}


/** ‘ ’ -> '   and   “ ” -> " */
function straightenQuotes_(s) {
  return String(s).replace(/[‘’]/g, '\'').replace(/[“”]/g, '"');
}


/** Decodes the named and numeric HTML entities a title might contain. */
function decodeHtmlEntities_(s) {
  var named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: '\'', nbsp: ' ',
                rsquo: '’', lsquo: '‘', rdquo: '”', ldquo: '“',
                mdash: '—', ndash: '–', hellip: '…' };
  return String(s).replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, function (whole, code) {
    if (code.charAt(0) === '#') {
      var n = code.charAt(1).toLowerCase() === 'x' ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10);
      return isNaN(n) ? whole : String.fromCharCode(n);
    }
    var v = named[code.toLowerCase()];
    return v === undefined ? whole : v;
  });
}
