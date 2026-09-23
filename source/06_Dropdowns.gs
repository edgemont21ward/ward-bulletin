/**
 * TYPE-AHEAD DROPDOWNS (Songs/Members/Leadership -> Person & Song fields)
 * Builds and refreshes the searchable dropdowns on Person/Song fields,
 * sourced from the Songs/Members/Leadership reference tabs, plus the
 * Date field's auto-fill-if-blank behavior.
 */

var DROPDOWN_HELPER_COL_ = 4; // column D on the Songs/Members sheets


// Exact label whose column-B value is a High Council member — the one
// row that draws on Leadership's "High Council Member" rows rather
// than Bishopric/Stake Presidency.
var HIGH_COUNCIL_FIELD_LABEL_ = 'recognize stake high councilors and other stake leaders';


// Exact labels for the two rows inside the "Stake Business" section.
// Both draw on the Leadership sheet rather than Members: whoever
// presents stake business is a member of the Stake Presidency or the
// High Council, and the position named beside them is their stake
// calling, so offering the whole ward roster would be wrong on both.
var STAKE_PRESENTER_FIELD_LABEL_ = 'presented by';
var STAKE_POSITION_FIELD_LABEL_ = 'stake leaders position';


// Exact-match labels whose column-B value is a person, not a song —
// checked before the generic "hymn"/"music" substring test below, since
// "Recognize Music Director" would otherwise match on "music" and get
// offered a song instead of a name.
var PERSON_FIELD_LABELS_ = [
  'recognize music director', 'recognize organist',
  'invocation', 'benediction'
];


/**
 * Re-applies the Songs/Members dropdowns to every hymn/music and speaker
 * row on `sheet` (the active sheet, if omitted). Safe to call as often
 * as you like — it just overwrites whatever validation was already
 * there. A no-op on the Songs or Members tabs themselves, or on any
 * sheet with no matching rows.
 */
function refreshDropdowns_(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() === SONGS_SHEET_NAME_ || sheet.getName() === MEMBERS_SHEET_NAME_ ||
      sheet.getName() === LEADERSHIP_SHEET_NAME_) return;

  var ss = sheet.getParent();
  var songRange = getSongDropdownRange_(ss);
  var memberRange = getMemberDropdownRange_(ss);
  var leadership = getLeadershipNames_(ss);
  var presidingNames = leadership.bishopric.concat(leadership.stakePresidency);
  // Stake Presidency first, then the High Council — the order the
  // Leadership sheet itself lists them in.
  var stakePresenterNames = leadership.stakePresidency.concat(leadership.highCouncil);

  var songRule = songRange
    ? SpreadsheetApp.newDataValidation()
        .requireValueInRange(songRange, true)
        .setAllowInvalid(true)
        .setHelpText('Start typing to search the Songs sheet, or type your own.')
        .build()
    : null;

  var memberRule = memberRange
    ? SpreadsheetApp.newDataValidation()
        .requireValueInRange(memberRange, true)
        .setAllowInvalid(true)
        .setHelpText('Start typing to search the Members sheet, or type your own.')
        .build()
    : null;

  // Small, fixed rosters (well under the 500-item cap), so "List of
  // items" is fine here — no helper column/range needed like Songs and
  // Members above.
  var conductingRule = leadership.bishopric.length
    ? SpreadsheetApp.newDataValidation()
        .requireValueInList(leadership.bishopric, true)
        .setAllowInvalid(true)
        .setHelpText('Start typing to search the Bishopric, or type your own.')
        .build()
    : null;

  var presidingRule = presidingNames.length
    ? SpreadsheetApp.newDataValidation()
        .requireValueInList(presidingNames, true)
        .setAllowInvalid(true)
        .setHelpText('Start typing to search the Stake Presidency/Bishopric, or type your own.')
        .build()
    : null;

  var highCouncilRule = leadership.highCouncil.length
    ? SpreadsheetApp.newDataValidation()
        .requireValueInList(leadership.highCouncil, true)
        .setAllowInvalid(true)
        .setHelpText('Start typing to search the High Council, or type your own.')
        .build()
    : null;

  var stakePresenterRule = stakePresenterNames.length
    ? SpreadsheetApp.newDataValidation()
        .requireValueInList(stakePresenterNames, true)
        .setAllowInvalid(true)
        .setHelpText('Start typing to search the Stake Presidency/High Council, or type your own.')
        .build()
    : null;

  // The Role column rather than the Name column — "High Council
  // Member", "Stake President", and so on.
  var stakePositionRule = leadership.stakeRoles.length
    ? SpreadsheetApp.newDataValidation()
        .requireValueInList(leadership.stakeRoles, true)
        .setAllowInvalid(true)
        .setHelpText('Start typing to search the stake callings on the Leadership sheet, or type your own.')
        .build()
    : null;

  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    var label = normalizeLabel_(data[i][0]);
    if (!label) continue;

    var isPersonField = PERSON_FIELD_LABELS_.indexOf(label) !== -1 || label.indexOf('speaker') !== -1;

    if (conductingRule && label === 'conducting') {
      sheet.getRange(i + 1, 2).setDataValidation(conductingRule);
    } else if (presidingRule && label === 'presiding') {
      sheet.getRange(i + 1, 2).setDataValidation(presidingRule);
    } else if (highCouncilRule && label === HIGH_COUNCIL_FIELD_LABEL_) {
      sheet.getRange(i + 1, 2).setDataValidation(highCouncilRule);
    } else if (stakePresenterRule && label === STAKE_PRESENTER_FIELD_LABEL_) {
      sheet.getRange(i + 1, 2).setDataValidation(stakePresenterRule);
    } else if (stakePositionRule && label === STAKE_POSITION_FIELD_LABEL_) {
      sheet.getRange(i + 1, 2).setDataValidation(stakePositionRule);
    } else if (memberRule && isPersonField) {
      sheet.getRange(i + 1, 2).setDataValidation(memberRule);
    } else if (songRule && (label.indexOf('hymn') !== -1 || label.indexOf('music') !== -1)) {
      sheet.getRange(i + 1, 2).setDataValidation(songRule);
    }
  }
}


/** Menu action: re-applies the dropdowns to the active sheet on demand. */
function refreshDropdownsFromMenu() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    fillDefaultDate_(sheet);
  } catch (err) {
    // Non-critical — see fillDefaultDate_'s own doc comment.
  }
  refreshDropdowns_(sheet);
  SpreadsheetApp.getUi().alert(
    'Dropdowns refreshed',
    'Song and speaker dropdowns on "' + sheet.getName() + '" now match the Songs and Members sheets.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Fills the "Date:" row's column B with the closest upcoming Sunday
 * (today itself, if today already is one) — but only when that cell is
 * currently blank, so a date that's already set (including a past
 * week's tab you're revisiting) is never overwritten. Meant for a
 * freshly duplicated week's tab, which starts out with an empty Date
 * cell. A no-op if there's no "Date" row at all (e.g. on the Songs/
 * Members/Leadership tabs).
 */
function fillDefaultDate_(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var data = sheet.getDataRange().getValues();
  var dateIdx = findLabelIndex_(data, 'Date');
  if (dateIdx === -1) return;
  if (hasText_(data[dateIdx][1])) return; // already set — never overwrite

  var tz = sheet.getParent().getSpreadsheetTimeZone();
  sheet.getRange(dateIdx + 1, 2).setValue(nextSunday_(tz));
}


/**
 * The next Sunday on or after today (today itself, if today already is
 * one), as a Date representing midnight on that day. `timeZone` is used
 * only to read today's actual calendar date correctly (matching the
 * spreadsheet's own time zone, not wherever this script happens to
 * execute) — the returned Date's clock time isn't otherwise meaningful,
 * since only the calendar day is ever used from it.
 */
function nextSunday_(timeZone) {
  var parts = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd').split('-');
  var today = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  var daysUntilSunday = (7 - today.getDay()) % 7; // 0 if today is already Sunday
  today.setDate(today.getDate() + daysUntilSunday);
  return today;
}


/**
 * Writes `values` into column `col` of `sheet` (row 1 gets `header`,
 * values start at row 2), clears any leftover rows below the new list
 * (so a removed song/member doesn't linger), hides the column, and
 * returns the range the values now occupy — or null if `values` is
 * empty. Shared by getSongDropdownRange_ and getMemberDropdownRange_.
 */
function writeDropdownColumn_(sheet, col, header, values) {
  sheet.getRange(1, col).setValue(header);

  if (values.length) {
    sheet.getRange(2, col, values.length, 1).setValues(values.map(function (v) { return [v]; }));
  }

  var extra = sheet.getMaxRows() - 1 - values.length;
  if (extra > 0) {
    sheet.getRange(2 + values.length, col, extra, 1).clearContent();
  }

  try {
    sheet.hideColumns(col);
  } catch (err) {
    // Not essential — the dropdown still works with the column visible.
  }

  return values.length ? sheet.getRange(2, col, values.length, 1) : null;
}


/**
 * Builds "#182 - We'll Sing All Hail to Jesus' Name" style strings from
 * the Songs sheet, writes them into its hidden helper column, and
 * returns the range they occupy (ready for "List from a range"
 * validation), or null if the Songs sheet is missing or empty.
 */
function getSongDropdownRange_(ss) {
  var sheet = ss.getSheetByName(SONGS_SHEET_NAME_);
  if (!sheet) return null;

  var rows = sheet.getDataRange().getValues();
  var values = [];
  for (var i = 1; i < rows.length; i++) { // row 1 is the header
    var num = String(rows[i][1] == null ? '' : rows[i][1]).trim();
    var title = String(rows[i][2] == null ? '' : rows[i][2]).trim();
    if (!title) continue;
    values.push(num ? '#' + num + ' - ' + title : title);
  }
  return writeDropdownColumn_(sheet, DROPDOWN_HELPER_COL_, 'Autocomplete list (auto-generated — do not edit)', values);
}


/**
 * Flips the Members sheet's "Last, First Middle" names into "First
 * Middle Last", writes them into its hidden helper column, and returns
 * the range they occupy (ready for "List from a range" validation), or
 * null if the Members sheet is missing or empty.
 */
function getMemberDropdownRange_(ss) {
  var sheet = ss.getSheetByName(MEMBERS_SHEET_NAME_);
  if (!sheet) return null;

  var rows = sheet.getDataRange().getValues();
  var values = [];
  for (var i = 1; i < rows.length; i++) { // row 1 is the header
    var full = String(rows[i][0] == null ? '' : rows[i][0]).trim();
    if (!full) continue;

    var commaIdx = full.indexOf(',');
    if (commaIdx === -1) { values.push(full); continue; }

    var last = full.substring(0, commaIdx).trim();
    var first = full.substring(commaIdx + 1).trim();
    values.push(first ? first + ' ' + last : last);
  }
  return writeDropdownColumn_(sheet, DROPDOWN_HELPER_COL_, 'Autocomplete list (auto-generated — do not edit)', values);
}


/**
 * Reads the "Leadership" sheet — columns A/B = Role / Name, e.g. "Bishop"
 * / "Matt Smith", "Bishopric 1st Counselor" / "...", "Stake President" /
 * "...", "High Council Member" / "..." (skips row 1, the header) — and
 * sorts every row with a Name into exactly one of three buckets, by
 * checking each Role for "high council", then "bishop", then "stake"
 * (in that order, so "Stake High Council Member" lands in highCouncil,
 * not stakePresidency):
 * - bishopric: for the Conducting dropdown.
 * - stakePresidency: combined with bishopric (by the caller) for the
 *   Presiding dropdown, since a Stake Presidency member presides
 *   instead of the Bishop whenever one is present (see the "[UNLESS a
 *   member of Stake Presidency is here]" note next to Presiding in the
 *   sheet itself).
 * - highCouncil: for "Recognize Stake High Councilors and other Stake
 *   Leaders" — can hold any number of rows (a stake usually has several
 *   high councilors), unlike the other two roles.
 * Also returns stakeRoles: the Role cells themselves, for the rows that
 * landed in stakePresidency or highCouncil, de-duplicated and in sheet
 * order ("Stake President", "Stake Presidency 1st Counselor", ...,
 * "High Council Member"). That's the list behind the "Stake Leaders
 * Position:" dropdown, so it stays whatever the sheet says — fix a
 * misspelt calling there and the dropdown follows.
 * All four come back empty if the Leadership sheet doesn't exist yet.
 */
function getLeadershipNames_(ss) {
  var sheet = ss.getSheetByName(LEADERSHIP_SHEET_NAME_);
  if (!sheet) return { bishopric: [], stakePresidency: [], highCouncil: [], stakeRoles: [] };

  var rows = sheet.getDataRange().getValues();
  var bishopric = [];
  var stakePresidency = [];
  var highCouncil = [];
  var stakeRoles = [];
  for (var i = 1; i < rows.length; i++) { // row 1 is the header
    var roleText = String(rows[i][0] == null ? '' : rows[i][0]).trim();
    var role = roleText.toLowerCase();
    var name = String(rows[i][1] == null ? '' : rows[i][1]).trim();
    if (!name) continue;

    if (role.indexOf('high council') !== -1) {
      highCouncil.push(name);
      if (stakeRoles.indexOf(roleText) === -1) stakeRoles.push(roleText);
    } else if (role.indexOf('bishop') !== -1) {
      bishopric.push(name);
    } else if (role.indexOf('stake') !== -1) {
      stakePresidency.push(name);
      if (stakeRoles.indexOf(roleText) === -1) stakeRoles.push(roleText);
    }
  }
  return {
    bishopric: bishopric,
    stakePresidency: stakePresidency,
    highCouncil: highCouncil,
    stakeRoles: stakeRoles
  };
}

/* ---------------------------------------------------------------------
 * NEW BULLETIN TABS, AND HIDING/SHOWING THE REFERENCE TABS
 *
 * "Create New Bulletin" duplicates a "Template" tab — a blank week's
 * program, laid out exactly like any other week's tab (same labels in
 * column A) but with the values in column B left empty — into a new
 * tab named after the closest upcoming Sunday, the same "September 20,
 * 2026" naming already used for week tabs (see nextSunday_ above for
 * the date it picks). You need to add that Template tab yourself, the
 * same way you added Songs/Members/Leadership — it's not something this
 * script can create on its own, since it doesn't know what your
 * program's layout should look like beyond the field labels it already
 * recognizes.
 *
 * The Template/Songs/Members/Leadership tabs are all reference data,
 * not something you read during the meeting, so "Show/Hide Reference
 * Tabs" lets you tuck them out of the way (Google Sheets' own
 * hideSheet()/showSheet(), the same "right-click a tab > Hide sheet"
 * you'd do by hand) without deleting anything — the script still reads
 * them normally either way, since hiding only affects what's visible in
 * the tab bar, not what Apps Script can see.
 * ------------------------------------------------------------------ */

