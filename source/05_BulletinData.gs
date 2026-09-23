/**
 * BULLETIN DATA (sheet -> structured data -> HTML text)
 * The parsing core: buildBulletinData() scans the active sheet's column A
 * for known labels and turns it into the plain data object Template.html
 * renders, including the HTML-escaping/link-detection helpers at the
 * bottom (escapeHtml_/linkify_/richTextCellToHtml_ — also called from
 * 04_Preview.gs for the "Open in a new tab" links).
 */

/**
 * Reads the active sheet's column A/B labels into a bulletin data
 * object. Plain text fields are left un-escaped here — Template.html's
 * <?= ?> tags escape on output, so escaping here too would double-escape
 * (e.g. turning an apostrophe into the literal text "&#39;"). The one
 * exception is announcementsHtml, which is deliberately built as real
 * HTML (so it can include clickable links) and is printed unescaped.
 */
function buildBulletinData(sheet) {
  var range = sheet.getDataRange();
  var data = range.getValues();
  // Read alongside the plain values so announcement links survive even
  // when they were added via Sheets' "Insert link" feature rather than
  // typed out as a bare URL — see richTextCellToHtml_ for why that
  // needs its own read.
  var richText = range.getRichTextValues();
  var tz = sheet.getParent().getSpreadsheetTimeZone();

  // Column C, when a row has anything in it, splits the number/title
  // that would otherwise be combined in column B — see resolveNumTitle_.
  var openingIdx = findLabelIndex_(data, 'Opening Hymn');
  var opening = openingIdx === -1
    ? { num: '', title: '' }
    : resolveNumTitle_(data[openingIdx][1], data[openingIdx][2], parseHymn_);

  var sacramentIdx = findLabelIndex_(data, 'Sacrament Hymn');
  var sacrament = sacramentIdx === -1
    ? { num: '', title: '' }
    : resolveNumTitle_(data[sacramentIdx][1], data[sacramentIdx][2], parseHymn_);

  var announcementsHtml = getAnnouncements_(data, richText);

  return {
    date: formatDateValue_(findFirst_(data, 'Date'), tz),
    presiding: findFirst_(data, 'Presiding'),
    conducting: findFirst_(data, 'Conducting'),
    musicDirector: findFirst_(data, 'Recognize Music Director'),
    organist: findFirst_(data, 'Recognize Organist'),

    openingHymnNum: opening.num,
    openingHymnTitle: opening.title,
    openingPrayer: findFirst_(data, 'Invocation'),

    // Both optional, heading-only lines (like "The Sacrament") — neither
    // prints a value of its own, they just switch on and off. See
    // hasStakeBusinessContent_ / hasReleasesOrSustainingNames_ for what
    // each is actually gated on.
    hasStakeBusiness: hasStakeBusinessContent_(data),
    hasWardBusiness: hasReleasesOrSustainingNames_(data),

    sacramentHymnNum: sacrament.num,
    sacramentHymnTitle: sacrament.title,

    // Any number of speakers and hymns/musical numbers, in whatever
    // order they appear between "Sacrament Hymn" and "Benediction".
    sacramentProgram: getSacramentProgram_(data),

    closingPrayer: findFirst_(data, 'Benediction'),

    announcementsHtml: announcementsHtml
  };
}


/**
 * Scans the rows between "Sacrament Hymn" and "Benediction" for any
 * number of speaker, hymn/music, and testimony rows, in whatever order
 * and count they appear — so adding or removing a speaker or musical
 * number never requires changing this script. A row's column A becomes
 * a "speaker" item if it contains the word "speaker" (matches "Youth
 * Speaker", "Speaker", "Concluding Speaker", etc.), a "testimony" item
 * if it contains "testimon" (matches "Bearing of Testimonies",
 * "Testimonies", "Testimony Meeting", etc. — for a fast Sunday, where
 * this usually stands in for any speakers at all, so column B is
 * usually left blank; rendered the same way as a music item, with an
 * optional title below, in case it's ever used for a short note
 * instead), or a "music" item if it contains "hymn" or "music"
 * (matches "Intermediate Hymn", "Hymn", "Musical Number", "Closing
 * Hymn", etc.). Anything else in that range — like the "The
 * Administration of the Sacrament" header line, which carries no value
 * — is skipped.
 */
function getSacramentProgram_(data) {
  var startIdx = -1, endIdx = -1;
  for (var i = 0; i < data.length; i++) {
    var norm = normalizeLabel_(data[i][0]);
    if (startIdx === -1 && norm === 'sacrament hymn') startIdx = i;
    if (norm === 'benediction') { endIdx = i; break; }
  }
  if (startIdx === -1 || endIdx === -1 || endIdx <= startIdx) return [];

  var items = [];
  for (var r = startIdx + 1; r < endIdx; r++) {
    var label = String(data[r][0] == null ? '' : data[r][0]).trim();
    if (!label) continue; // blank spacer row, or a value-less header cell

    var value = data[r][1];
    var titleCell = data[r][2];
    var normLabel = label.toLowerCase();

    if (normLabel.indexOf('speaker') !== -1) {
      items.push({ type: 'speaker', label: label, name: value });
    } else if (normLabel.indexOf('testimon') !== -1) {
      // Column B (and, in the two-cell format, column C) is usually
      // empty here (no one person is "assigned" to bear testimonies),
      // but resolved the same way as a music item in case it's ever
      // used for a short note instead — see the "similar to Music"
      // rendering in Template.html.
      var parsedTestimony = resolveNumTitle_(value, titleCell, parseMusicItem_);
      items.push({
        type: 'testimony', label: label,
        num: parsedTestimony.num, title: parsedTestimony.title
      });
    } else if (normLabel.indexOf('hymn') !== -1 || normLabel.indexOf('music') !== -1) {
      var parsed = resolveNumTitle_(value, titleCell, parseMusicItem_);
      items.push({
        type: 'music', label: label,
        num: parsed.num, title: parsed.title
      });
    }
    // Anything else (e.g. "The Administration of the Sacrament") is a
    // header-only line, not a speaker, testimony, or hymn row — skipped.
  }
  return items;
}


/**
 * Normalizes a label for comparison: trims, drops a trailing colon,
 * lowercases, and folds misspellings of "business" into the real word.
 *
 * That last step exists because the Stake Business header is typed by
 * hand, and has been got wrong more than once — "BUISNESS", then
 * "BUISSNESS". hasStakeBusinessContent_ finds the section by exact
 * label, so any misspelling hides the heading from the bulletin
 * without an error. /bu[is]+ness/ covers the swapped and doubled
 * letters seen so far (plus "busness"), and also matches "business"
 * itself, so a correctly spelt header passes through unchanged.
 */
function normalizeLabel_(s) {
  return String(s == null ? '' : s)
    .trim()
    .replace(/:$/, '')
    .toLowerCase()
    .replace(/bu[is]+ness/g, 'business');
}


/** Returns every column-B value whose column-A cell matches `label`, in row order. */
function findValues_(data, label) {
  var norm = normalizeLabel_(label);
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (normalizeLabel_(data[i][0]) === norm) {
      out.push(data[i][1]);
    }
  }
  return out;
}


/** Returns the first column-B value matching `label`, or '' if none found. */
function findFirst_(data, label) {
  var vals = findValues_(data, label);
  return vals.length ? vals[0] : '';
}


/**
 * Resolves a hymn/music/testimony row's number + title, supporting two
 * sheet layouts so a ward can use either one (and even mix them row to
 * row):
 *   1. Combined in one cell — column B holds "#1062 - Lord, Accept Our
 *      Humble Fast" (or, for a non-numeric group like a musical number,
 *      "Ward Choir - Lead, Kindly Light"), split on the first separator
 *      by `parseFn` (parseHymn_ or parseMusicItem_).
 *   2. Split across two cells — column B holds just the number/name
 *      ("#1062"), column C holds just the title (" Lord, Accept Our
 *      Humble Fast").
 * Column C wins whenever it has anything in it (once trimmed), so a
 * sheet that's still using format 1 — where column C is blank — is
 * completely unaffected; only a row that actually has something in
 * column C is treated as format 2.
 */
function resolveNumTitle_(colB, colC, parseFn) {
  var title = String(colC == null ? '' : colC).trim();
  if (title) {
    var num = String(colB == null ? '' : colB).trim().replace(/^#\s*/, '');
    return { num: num, title: title };
  }
  return parseFn(colB);
}


/** Splits "#97 - Lead, Kindly Light" into {num: "97", title: "Lead, Kindly Light"}. */
function parseHymn_(raw) {
  var str = String(raw == null ? '' : raw).trim();
  var m = str.match(/^#?\s*(\d+)\s*[-–—]\s*(.+)$/);
  if (m) return { num: m[1], title: m[2].trim() };
  return { num: '', title: str };
}


/**
 * Splits a sacrament-program music value on its first "-"/"–"/"—"/"|"
 * separator, the same way a hymn splits into number + title, so that
 * whatever comes BEFORE the separator displays in the right-hand slot
 * (the same spot a hymn number or a speaker's name appears) and
 * whatever comes AFTER it displays as the centered italic title line:
 *   - "#292 - O My Father" → { num: "292", title: "O My Father" } —
 *     the "#" is stripped, same as parseHymn_.
 *   - "Ward Choir - Lead, Kindly Light" (no leading number) →
 *     { num: "Ward Choir", title: "Lead, Kindly Light" } — the group
 *     name lands in the right-hand slot, the song title is what's
 *     centered below, matching how a hymn's number vs. title are laid
 *     out.
 *   - Anything else (no recognized separator) → the whole string
 *     becomes { num: "", title: <whole string> }.
 * Only the first separator is used, so a title that itself contains a
 * hyphen (e.g. "Nearer, My God, to Thee - Ward Choir - a cappella")
 * still keeps "Ward Choir - a cappella" together as the title.
 */
function parseMusicItem_(raw) {
  var str = String(raw == null ? '' : raw).trim();

  var m = str.match(/^(.*?)\s*[-–—|]\s*(.+)$/);
  if (!m) return { num: '', title: str };

  var num = m[1].trim().replace(/^#\s*/, '');
  return { num: num, title: m[2].trim() };
}


/** Formats a Date object or date-like string as "August 30, 2026". */
function formatDateValue_(raw, timeZone) {
  if (Object.prototype.toString.call(raw) === '[object Date]' && !isNaN(raw.getTime())) {
    return Utilities.formatDate(raw, timeZone, 'MMMM d, yyyy');
  }
  var s = String(raw == null ? '' : raw).trim();
  var d = new Date(s);
  if (s && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, timeZone, 'MMMM d, yyyy');
  }
  return s;
}


/** True for anything other than a blank/whitespace-only cell — used to gate optional items on and off. */
function hasText_(v) {
  return String(v == null ? '' : v).trim() !== '';
}


/**
 * Field labels this script already treats as structured data elsewhere
 * (see buildBulletinData). Listed here so the announcements scan below
 * can recognize the start of the next section even if the blank-row
 * spacing before it ever gets tightened to a single row.
 */
var KNOWN_FIELD_LABELS_ = [
  'date', 'presiding', 'conducting', 'welcome',
  'recognize music director', 'recognize organist',
  'recognize stake high councilors and other stake leaders',
  'opening hymn', 'invocation', 'stake business', 'ward business',
  'releases', 'sustaining', 'sacrament hymn',
  'youth speaker', 'speaker', 'intermediate hymn',
  'closing hymn', 'benediction'
];


/**
 * The announcement-style section headers themselves — listed as stop
 * labels too, so if "PROGRAM ANNOUNCEMENTS" is only one blank row below
 * the end of "ANNOUNCEMENTS", the header text doesn't get swallowed as
 * a bogus announcement line.
 */
var ANNOUNCEMENT_SECTION_LABELS_ = ['announcements', 'program announcements'];


/**
 * Returns the row indexes of every non-empty column-A cell after
 * `startIdx`, stopping at the first row whose label is a known field or
 * section-header label, or after two consecutive blank rows (a single
 * blank row, used for visual spacing, does NOT end the list). Indexes
 * rather than cell text so the caller can also look up that same row's
 * rich text (for hyperlinks — see richTextCellToHtml_).
 */
function collectSectionItemIndexes_(data, startIdx, stopLabels) {
  var indexes = [];
  var blankStreak = 0;
  for (var r = startIdx + 1; r < data.length; r++) {
    var cell = data[r][0];
    var norm = normalizeLabel_(cell);
    var isBlank = (norm === '');

    if (!isBlank && stopLabels.indexOf(norm) !== -1) break;

    if (isBlank) {
      blankStreak++;
      if (blankStreak >= 2) break; // two blank rows = end of section
      continue;
    }

    blankStreak = 0;
    indexes.push(r);
  }
  return indexes;
}


/** Returns the row index whose column A matches `label`, or -1 if none found. */
function findLabelIndex_(data, label) {
  var norm = normalizeLabel_(label);
  for (var i = 0; i < data.length; i++) {
    if (normalizeLabel_(data[i][0]) === norm) return i;
  }
  return -1;
}


/**
 * True if the "Stake Business" section actually has anything filled in.
 * "Stake Business" itself is a header row (column B blank, like
 * "ANNOUNCEMENTS") — the real content is whatever sub-rows follow it
 * (e.g. "Presented By:", "Stake Leaders Position:"), the same "collect
 * until a stop label or two blank rows" scan the announcements use. Any
 * one of those sub-rows having a column B value counts.
 */
function hasStakeBusinessContent_(data) {
  var stakeIdx = findLabelIndex_(data, 'Stake Business');
  if (stakeIdx === -1) return false;

  var stopLabels = KNOWN_FIELD_LABELS_.concat(ANNOUNCEMENT_SECTION_LABELS_);
  var rowIndexes = collectSectionItemIndexes_(data, stakeIdx, stopLabels);
  return rowIndexes.some(function (r) { return hasText_(data[r][1]); });
}


/**
 * True if anyone is actually listed as being released or sustained —
 * this is what "Ward Business" is gated on, since there's no "Ward
 * Business" row in the sheet at all. "Releases" and "Sustaining" sit
 * side by side as a single wide section header, with a shared block of
 * boilerplate script text below it, then a "NAMES:" / "POSITIONS:"
 * label row, then the actual entries: a released person's name in
 * column A and a sustained person's name in column C, one pair per row,
 * for as many rows as there are people (blank rows below the last entry
 * are just unused capacity). Scanning starts right after "NAMES:" so
 * the boilerplate script text above it (which always has *something* in
 * column A/C) never counts as a real entry; it stops at two consecutive
 * rows with nothing in either column, so the closing "Those who wish to
 * express..." line further down never gets reached.
 */
function hasReleasesOrSustainingNames_(data) {
  var namesIdx = findLabelIndex_(data, 'NAMES:');
  if (namesIdx === -1) return false;

  var blankStreak = 0;
  for (var r = namesIdx + 1; r < data.length; r++) {
    var releaseName = data[r][0];
    var sustainName = data[r][2];

    if (hasText_(releaseName) || hasText_(sustainName)) return true;

    blankStreak++;
    if (blankStreak >= 2) break; // two blank rows = end of the table
  }
  return false;
}

/* ---------------------------------------------------------------------
 * TYPE-AHEAD DROPDOWNS — every song row (any column-A label containing
 * "hymn" or "music" — "Opening Hymn", "Sacrament Hymn", "Intermediate
 * Hymn", "Musical Number", etc.) gets a searchable dropdown in column B
 * sourced from the Songs sheet, and every person row — any label
 * containing "speaker", plus the exact labels "Recognize Music
 * Director", "Recognize Organist", "Invocation", and "Benediction" (see
 * PERSON_FIELD_LABELS_ — those exact labels are checked before the
 * "music" substring test above, since otherwise "Recognize Music
 * Director" would match on "music" and get offered a song instead of a
 * name) — gets one sourced from the Members sheet. "Conducting" gets a
 * dropdown of just the Bishopric, "Presiding" gets one of the Bishopric
 * plus the Stake Presidency together (a Stake Presidency member
 * presides instead of the Bishop whenever one is present — see the note
 * next to Presiding in the sheet itself), and "Recognize Stake High
 * Councilors and other Stake Leaders" gets one of just the High
 * Council — all three sourced from a separate "Leadership" sheet.
 * Either way, start typing and Sheets
 * filters the list down to matches. It's built with "Show a warning"
 * validation (setAllowInvalid(true)), not "Reject input", so anything
 * NOT on the list — a guest speaker, a hymn not in the hymnal — is
 * still accepted; it just gets a small red corner marker instead of a
 * dropdown checkmark.
 *
 * Source data, and a hidden helper column:
 * - "Songs" sheet — columns A/B/C = Hymnal / Number-Page / Title (skips
 *   row 1, the header). Offered as "#182 - We'll Sing All Hail to
 *   Jesus' Name", matching how songs are already typed into column B
 *   today.
 * - "Members" sheet — column A = "Last, First Middle" (skips row 1, the
 *   header). Flipped to "First Middle Last" for the dropdown, since
 *   that's how a speaker's name actually gets printed in the bulletin.
 * - "Leadership" sheet — columns A/B = Role / Name (skips row 1, the
 *   header), e.g. "Bishop" / "Matt Smith", "Stake President" / "...",
 *   "High Council Member" / "..." (one row per person — add as many
 *   "High Council Member" rows as needed). Small enough (well under a
 *   few dozen rows even with several High Council rows) that, unlike
 *   Songs/Members, it can use "List of items" directly — see
 *   getLeadershipNames_.
 *
 * Sheets' "List of items" validation (requireValueInList) hard-caps at
 * 500 entries and errors otherwise — the Songs sheet alone is already
 * past that once the Children's Songbook is folded in with the two
 * hymnals. "List from a range" (requireValueInRange) has no such cap,
 * but it needs the display strings to actually live in real cells, not
 * just a computed array — so refreshDropdowns_ below writes them into a
 * hidden column D on each of the Songs/Members sheets (re-writing it
 * fresh every time, so it can never drift from column A/B/C) and points
 * the validation at that column. It's hidden because it's just a
 * computed mirror of columns A-C — there's nothing to edit there
 * directly (an edit would just be overwritten on the next refresh).
 *
 * Because that helper column is a snapshot taken at the moment it's
 * written (not a live formula), it's refreshed automatically every time
 * the spreadsheet is opened (see onOpenInstallable) and can also be
 * refreshed on demand from the "Ward Bulletin" menu — useful right
 * after adding a new speaker/hymn row, or after editing the Songs or
 * Members list, without having to reopen the sheet.
 * ------------------------------------------------------------------ */


/**
 * Collects the announcement lines (as ready-to-print HTML — see
 * richTextCellToHtml_) from both the "ANNOUNCEMENTS" section and, if
 * present, a separate "PROGRAM ANNOUNCEMENTS" section — with the program
 * announcements appended after the regular ones, regardless of where
 * each section actually sits in the sheet. Older sheets that only have
 * "ANNOUNCEMENTS" keep working exactly as before.
 */
function getAnnouncements_(data, richText) {
  var stopLabels = KNOWN_FIELD_LABELS_.concat(ANNOUNCEMENT_SECTION_LABELS_);
  var indexes = [];

  var annIdx = findLabelIndex_(data, 'ANNOUNCEMENTS');
  if (annIdx !== -1) {
    indexes = indexes.concat(collectSectionItemIndexes_(data, annIdx, stopLabels));
  }

  var progIdx = findLabelIndex_(data, 'PROGRAM ANNOUNCEMENTS');
  if (progIdx !== -1) {
    indexes = indexes.concat(collectSectionItemIndexes_(data, progIdx, stopLabels));
  }

  return indexes.map(function (r) {
    return richTextCellToHtml_(data[r][0], richText[r][0]);
  });
}


/** Escapes HTML-sensitive characters so sheet text can't break the markup. */
function escapeHtml_(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}


/** Turns bare URLs in already-escaped text into clickable links. */
function linkify_(escapedStr) {
  return escapedStr.replace(
    /(https?:\/\/[^\s<]+)/g,
    '<a href="$1" target="_blank" rel="noopener">$1</a>'
  );
}


/**
 * Converts one announcement cell to safe, ready-to-print HTML — the
 * piece that makes links actually publish. getValues() alone only
 * returns a cell's DISPLAY text: if someone attaches a link via Sheets'
 * own "Insert link" (Ctrl+K), turning e.g. "Sign up here" blue and
 * underlined, getValues() just returns the plain words "Sign up here"
 * with no trace of the URL — so the old escapeHtml_ + linkify_ path
 * (which only recognizes a literal "https://..." typed into the cell)
 * silently drops that link and publishes plain, unclickable text.
 *
 * getRichTextValues() (read alongside getValues() in buildBulletinData)
 * preserves that link. getRuns() splits a cell into same-formatting
 * spans, each with its own getLinkUrl(); a whole-cell link is just one
 * run covering the entire text. Any run WITH a link becomes an <a> tag
 * around its own text; any run withOUT one still goes through the old
 * escape + linkify_ path, so a bare "https://..." typed directly into
 * the text keeps working exactly as before.
 */
function richTextCellToHtml_(cellValue, richTextValue) {
  if (!richTextValue) {
    return linkify_(escapeHtml_(String(cellValue == null ? '' : cellValue).trim()));
  }

  var runs = richTextValue.getRuns();
  if (!runs || runs.length === 0) {
    return linkify_(escapeHtml_(richTextValue.getText().trim()));
  }

  var html = '';
  for (var i = 0; i < runs.length; i++) {
    var run = runs[i];
    var runText = run.getText();
    var linkUrl = run.getLinkUrl();
    html += linkUrl
      ? '<a href="' + escapeHtml_(linkUrl) + '" target="_blank" rel="noopener">' + escapeHtml_(runText) + '</a>'
      : linkify_(escapeHtml_(runText));
  }
  return html.trim();
}

