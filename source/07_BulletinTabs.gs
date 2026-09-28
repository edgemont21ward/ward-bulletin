/**
 * NEW BULLETIN TABS, PROGRAM ROWS, REFERENCE TAB VISIBILITY & ARCHIVING
 * (Program rows and archiving have their own sections further down.)
 *
 * "Create New Bulletin" duplicates the Template tab (a blank week's
 * program: the same labels in column A as any week's tab, with column
 * B left empty) into a new tab named after the closest upcoming Sunday,
 * in the "September 20, 2026" format the week tabs already use (see
 * nextSunday_ in 06_Dropdowns.gs for how that Sunday is picked). The
 * Template tab has to be made by hand, like Songs/Members/Leadership:
 * the script only knows the field labels it reads, not the rest of your
 * program's layout.
 *
 * "Show/Hide Reference Tabs" tucks the Template/Songs/Members/Leadership
 * tabs out of the tab bar, or brings them back. That's Sheets' own
 * hideSheet()/showSheet(), the same as right-click > Hide sheet, so
 * nothing is deleted and the script reads hidden tabs normally.
 */

/**
 * The tab name this week's bulletin should have: the closest upcoming
 * Sunday (today itself, when today IS Sunday — see nextSunday_ in
 * 06_Dropdowns.gs), formatted the same way the existing week tabs are
 * ("September 27, 2026"). Shared by createNewBulletin_ and
 * ensureCurrentBulletinTab_ so the two can never disagree about which
 * tab name counts as "already exists".
 */
function bulletinTabName_(ss) {
  var tz = ss.getSpreadsheetTimeZone();
  return Utilities.formatDate(nextSunday_(tz), tz, 'MMMM d, yyyy');
}


/**
 * Startup helper (called from onOpenInstallable in 02_Menu.gs): if
 * there's no tab yet for the closest upcoming Sunday, creates one the
 * same way the "Create New Bulletin" menu item does. Opening the
 * spreadsheet on a Sunday with no tab for that day therefore just makes
 * one, rather than waiting for someone to remember the menu item.
 *
 * Does nothing at all when that tab already exists — in particular it
 * does NOT switch to it, so opening the spreadsheet to look at some
 * other week's tab leaves you right where you were. Only the act of
 * creating a tab switches to it (that's createNewBulletin_'s own
 * behavior, kept identical to the menu item's).
 *
 * It also deletes any leftover hidden copies of the Template tab first
 * (see removeLeftoverTemplateCopies_).
 *
 * Returns the newly created Sheet, or null if one already existed.
 * Throws only what createNewBulletinUnlocked_ throws (e.g. no "Template"
 * tab to copy from yet), or if another run holds the tab lock for 30
 * seconds — the caller swallows that, since a missing Template tab
 * shouldn't turn into an error dialog every time the file opens.
 *
 * Worth knowing: because this runs on every open, deliberately deleting
 * the current week's tab means the next open recreates it. Deleting a
 * PAST week's tab is unaffected — only the closest upcoming Sunday's is
 * ever created here.
 */
function ensureCurrentBulletinTab_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return withBulletinTabLock_(function () {
    var removed = removeLeftoverTemplateCopies_(ss);
    if (removed.length) {
      ss.toast('Removed ' + removed.length + ' leftover hidden ' +
        (removed.length === 1 ? 'copy' : 'copies') + ' of the Template tab: ' +
        removed.join(', ') + '.', 'Ward Bulletin', 8);
    }

    var tabName = bulletinTabName_(ss);
    if (ss.getSheetByName(tabName)) return null;

    var sheet = createNewBulletinUnlocked_();
    ss.toast('Created "' + tabName + '" from the Template tab.', 'Ward Bulletin', 8);
    return sheet;
  });
}


/**
 * Runs `fn` while holding the spreadsheet's document lock, so only one
 * run at a time, by anyone, can be creating a week's tab.
 *
 * Without this, two runs at once — which happens when more than one
 * editor has an open trigger, since each fires when the file opens —
 * could both see no tab for the coming Sunday and both copy Template.
 * One renames its copy; the other's rename fails because that name is
 * now taken, stranding a hidden "Copy of Template N" (hidden because
 * Template is). The error was swallowed on open, so the copies piled
 * up unnoticed. With the lock, the second run waits, then finds the
 * tab already there.
 *
 * Waits up to 30 seconds for the lock, then throws.
 */
function withBulletinTabLock_(fn) {
  // getDocumentLock() can be null outside the spreadsheet's own UI, e.g.
  // in the archive service's web app; the project-wide script lock is
  // then the nearest equivalent.
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}


/**
 * Deletes every hidden tab named exactly "Copy of Template" or "Copy
 * of Template <number>", and returns their names. Those are what
 * copying the (hidden) Template tab produces before the copy is renamed
 * and unhidden, so a hidden one with that name is a copy whose tab
 * creation never finished — see withBulletinTabLock_.
 *
 * A visible tab is never touched, whatever it's called, and neither is
 * any other name — so a copy of Template you made yourself, which
 * Sheets always creates visible, is safe.
 */
function removeLeftoverTemplateCopies_(ss) {
  var prefix = 'Copy of ' + TEMPLATE_SHEET_NAME_;
  var removed = [];
  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    var isCopyName = name === prefix || (name.indexOf(prefix + ' ') === 0 && /^\d+$/.test(name.slice(prefix.length + 1)));
    if (isCopyName && sheet.isSheetHidden()) {
      ss.deleteSheet(sheet);
      removed.push(name);
    }
  });
  return removed;
}


/** Menu action: creates (or switches to) this week's bulletin tab — see createNewBulletinUnlocked_. */
function createNewBulletinFromMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sheet = createNewBulletin_();
    ui.alert('Ready', '"' + sheet.getName() + '" is ready to fill in.', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Could not create a new bulletin', err.message || String(err), ui.ButtonSet.OK);
  }
}


/** createNewBulletinUnlocked_, holding the tab lock — see withBulletinTabLock_. */
function createNewBulletin_() {
  return withBulletinTabLock_(createNewBulletinUnlocked_);
}


/**
 * Duplicates the Template tab into a new tab named after the closest
 * upcoming Sunday, fills in its Date cell and dropdowns right away
 * (rather than waiting for the next reopen — see fillDefaultDate_ and
 * refreshDropdowns_, both in 06_Dropdowns.gs), moves it to the front
 * alongside the other week tabs, and switches to it. If a tab for that
 * date already exists (you already created this week's bulletin), just
 * switches to that one instead of making a duplicate. Throws if there's
 * no "Template" tab to copy from.
 *
 * Call it through createNewBulletin_ or ensureCurrentBulletinTab_, which
 * hold the tab lock; on its own, two runs at once can race.
 */
function createNewBulletinUnlocked_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var template = ss.getSheetByName(TEMPLATE_SHEET_NAME_);
  if (!template) {
    throw new Error(
      'No "' + TEMPLATE_SHEET_NAME_ + '" tab found. Add one first — a blank week\'s ' +
      'program, laid out like any other week\'s tab with the same labels in column A ' +
      'but blank values in column B.'
    );
  }

  var tabName = bulletinTabName_(ss);

  var existing = ss.getSheetByName(tabName);
  if (existing) {
    if (existing.isSheetHidden()) existing.showSheet();
    ss.setActiveSheet(existing);
    return existing;
  }

  var newSheet = template.copyTo(ss);
  try {
    newSheet.setName(tabName);
  } catch (err) {
    ss.deleteSheet(newSheet); // don't strand a hidden "Copy of Template N"
    throw err;
  }
  if (newSheet.isSheetHidden()) newSheet.showSheet(); // Template itself may be hidden — the copy shouldn't be

  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1); // to the front, alongside the other week tabs

  // Both non-critical: the tab exists either way, a blank Date cell is
  // easy to fill in, and the dropdowns get another chance on the next
  // open or from "Refresh Dropdowns".
  try {
    fillDefaultDate_(newSheet);
  } catch (err) {}
  try {
    refreshDropdowns_(newSheet);
  } catch (err) {}

  return newSheet;
}


/**
 * Shows or hides the Template/Songs/Members/Leadership tabs together,
 * for the two menu actions below. A tab that doesn't exist yet (e.g.
 * you haven't added Leadership) is silently skipped, not an error.
 * Returns the names it actually found and changed.
 */
function setUtilityTabsVisible_(visible) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var changed = [];
  UTILITY_SHEET_NAMES_.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    try {
      if (visible) sheet.showSheet(); else sheet.hideSheet();
      changed.push(name);
    } catch (err) {
      // e.g. Sheets refusing to hide the last remaining visible tab —
      // skip it rather than blocking the rest.
    }
  });
  return changed;
}


/**
 * True only when every Template/Songs/Members/Leadership tab that
 * exists is currently hidden — used by buildMenu_() to decide which of
 * "Show"/"Hide Reference Tabs" to offer. False (i.e. offer "Hide") when
 * any of them is visible, OR when none of them exist yet — either way
 * there's nothing left to usefully "show", so "Hide" is the safer
 * default to display (clicking it just reports nothing was found).
 */
function areUtilityTabsHidden_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var foundAny = false;
  for (var i = 0; i < UTILITY_SHEET_NAMES_.length; i++) {
    var sheet = ss.getSheetByName(UTILITY_SHEET_NAMES_[i]);
    if (!sheet) continue;
    foundAny = true;
    if (!sheet.isSheetHidden()) return false;
  }
  return foundAny;
}


/** Menu action. */
function hideUtilityTabsFromMenu() {
  var changed = setUtilityTabsVisible_(false);
  buildMenu_(); // so the menu now offers "Show" instead of "Hide"
  SpreadsheetApp.getUi().alert(changed.length
    ? ('Hid: ' + changed.join(', ') + '.')
    : 'None of the Template/Songs/Members/Leadership tabs were found.');
}


/** Menu action. */
function showUtilityTabsFromMenu() {
  var changed = setUtilityTabsVisible_(true);
  buildMenu_(); // so the menu now offers "Hide" instead of "Show"
  SpreadsheetApp.getUi().alert(changed.length
    ? ('Shown: ' + changed.join(', ') + '.')
    : 'None of the Template/Songs/Members/Leadership tabs were found.');
}



/* ---------------------------------------------------------------------
 * PROGRAM ROWS
 * "Add Program Row" (menu) and the sidebar's "+ Speaker" /
 * "+ Intermediate Hymn" / "+ Musical Number" / "+ Testimonies" buttons
 * insert a row into the sacrament program, just below whichever row is
 * selected. Google Sheets doesn't let a
 * script add to its right-click menu, so this is the nearest thing.
 *
 * Adding a row by hand means retyping the label (easy to misspell, and
 * it prints on the bulletin exactly as typed), re-merging columns B-D,
 * and reapplying the dropdown. Doing it here copies all three from the
 * rows around it.
 * ------------------------------------------------------------------ */


// The labels these add, by the name the sidebar and menu use for each.
// Any label containing "speaker", "hymn", "music" or "testimon" already
// prints as a program item (see getSacramentProgram_), so these are
// just the spellings to standardize on. "Testimonies" is what the
// September 6 fast Sunday used.
var PROGRAM_ROW_LABELS_ = {
  speaker: 'Speaker',
  hymn: 'Intermediate Hymn',
  music: 'Musical Number',
  testimony: 'Testimonies'
};


/** Menu actions for the "Add Program Row" submenu. */
function addSpeakerRowFromMenu() { addProgramRowFromMenu_('speaker'); }
function addHymnRowFromMenu() { addProgramRowFromMenu_('hymn'); }
function addMusicalNumberRowFromMenu() { addProgramRowFromMenu_('music'); }
function addTestimoniesRowFromMenu() { addProgramRowFromMenu_('testimony'); }


function addProgramRowFromMenu_(kind) {
  try {
    var added = insertProgramRow_(kind);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Added a ' + added.label + ' row at row ' + added.row + '.', 'Ward Bulletin', 4);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Could not add a row', err.message || String(err),
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/**
 * Sidebar-callable: the same insert, for the sidebar's buttons. Returns
 * {row, label} for its status line, or throws a message it can show.
 */
function addProgramRowForSidebar(kind) {
  return insertProgramRow_(kind);
}


/**
 * Inserts a `kind` row ('speaker', 'hymn', 'music' or 'testimony')
 * just below the selected row of the active tab, and moves the cursor
 * to its column B ready to type. The selected row has to be inside the sacrament
 * program: below "Sacrament Hymn" and above "Benediction" (see
 * sacramentProgramBounds_). Selecting "The Administration of the
 * Sacrament" line adds the row at the top of the program.
 *
 * The new row copies its formatting and merged cells from the nearest
 * existing program row (the selected one, if it is one), so it matches
 * its neighbours. Its dropdown is copied from another row of the same
 * kind; failing that, refreshDropdowns_ applies it. A testimonies row
 * gets no dropdown: nobody is assigned to it, so column B is usually
 * left blank, as on any fast Sunday tab.
 *
 * Returns { row: the new row's 1-based number, label }.
 */
function insertProgramRow_(kind) {
  var label = PROGRAM_ROW_LABELS_[kind];
  if (!label) throw new Error('Unknown kind of program row: ' + kind);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var bounds = sacramentProgramBounds_(data);
  if (!bounds) {
    throw new Error('"' + sheet.getName() + '" has no sacrament program to add to — it needs a ' +
      '"Sacrament Hymn" row with a "Benediction" row somewhere below it.');
  }

  var selected = sheet.getActiveRange().getLastRow(); // 1-based; the bottom row if several are selected
  var selectedIdx = selected - 1;
  if (selectedIdx <= bounds.start || selectedIdx >= bounds.end) {
    throw new Error('Select a row between "Sacrament Hymn" (row ' + (bounds.start + 1) + ') and ' +
      '"Benediction" (row ' + (bounds.end + 1) + ') first. The new row goes just below the one you select.');
  }

  var styleRow = nearestProgramItemRow_(data, bounds, selectedIdx);
  var newRow = selected + 1;
  sheet.insertRowAfter(selected);
  var width = sheet.getLastColumn();

  if (styleRow !== null) {
    var src = styleRow < newRow ? styleRow : styleRow + 1; // rows below the insert moved down one
    var srcRange = sheet.getRange(src, 1, 1, width);
    srcRange.copyFormatToRange(sheet, 1, width, newRow, newRow);
    srcRange.getMergedRanges().forEach(function (m) {
      if (m.getNumRows() === 1) sheet.getRange(newRow, m.getColumn(), 1, m.getNumColumns()).merge();
    });
  }

  sheet.getRange(newRow, 1).setValue(label);
  if (kind === 'testimony') {
    sheet.getRange(newRow, 2).clearDataValidations(); // in case copying the format brought one along
  } else if (!copyDropdownFromSimilarRow_(sheet, newRow, kind)) {
    refreshDropdowns_(sheet);
  }
  sheet.getRange(newRow, 2).activate();
  return { row: newRow, label: label };
}


/**
 * The 1-based row of the program item (a speaker, hymn/music, or
 * testimony row) nearest to `selectedIdx`, preferring the selected row
 * itself, then the rows below it, then above. That's the row whose look
 * a new row copies. Null if the program has no items yet — the new row
 * then keeps whatever formatting Sheets gave it.
 */
function nearestProgramItemRow_(data, bounds, selectedIdx) {
  var isItem = function (i) {
    var l = normalizeLabel_(data[i][0]);
    return /speaker|hymn|music|testimon/.test(l);
  };
  if (isItem(selectedIdx)) return selectedIdx + 1;
  for (var d = 1; d < bounds.end - bounds.start; d++) {
    if (selectedIdx + d < bounds.end && isItem(selectedIdx + d)) return selectedIdx + d + 1;
    if (selectedIdx - d > bounds.start && isItem(selectedIdx - d)) return selectedIdx - d + 1;
  }
  return null;
}


/**
 * Gives the new row's column B the same dropdown as another row of the
 * same kind on this tab: another speaker row for 'speaker', another
 * song row for 'hymn' or 'music' (both use the Songs list). Not used for
 * 'testimony', which has no dropdown. True if it
 * found one to copy. Much faster than refreshDropdowns_, which rebuilds
 * the Songs and Members helper columns.
 */
function copyDropdownFromSimilarRow_(sheet, newRow, kind) {
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (i === newRow - 1) continue;
    var l = normalizeLabel_(data[i][0]);
    var sameKind = kind === 'speaker'
      ? l.indexOf('speaker') !== -1
      : (l.indexOf('hymn') !== -1 || l.indexOf('music') !== -1) && PERSON_FIELD_LABELS_.indexOf(l) === -1;
    if (!sameKind) continue;
    var rule = sheet.getRange(i + 1, 2).getDataValidation();
    if (rule) {
      sheet.getRange(newRow, 2).setDataValidation(rule);
      return true;
    }
  }
  return false;
}


/* ---------------------------------------------------------------------
 * ARCHIVING PAST WEEKS
 * "Ward Bulletin > Archive Past Weeks…" moves every week tab older than
 * this week's into a brand-new spreadsheet, then removes them here, so
 * this one only holds the week being worked on (and any later ones),
 * Template, and the reference tabs.
 *
 * A "week tab" is one named exactly like a date, the way week tabs are
 * named ("September 20, 2026"); anything else is never touched. "Older
 * than this week" means before the closest upcoming Sunday, the same
 * Sunday bulletinTabName_ names: on a Sunday that's today, so that
 * day's tab stays until Monday.
 *
 * Who does the work. Any editor can use the menu item, but the archive
 * has to belong to, and sit in the Drive of, the account that owns this
 * spreadsheet — and a new spreadsheet always belongs to whoever creates
 * it. So the menu item doesn't archive anything itself: it sends the
 * confirmed tab names to the ARCHIVE SERVICE, this project's web app
 * deployment, which always runs as the account that deployed it (the
 * owner), whoever calls it. See doPost.
 *
 * The service is deliberately narrow. It serves no pages (there's no
 * doGet), so nothing can reach this project's other functions through
 * it; it answers only POSTs carrying the secret kept in Script
 * Properties, which only people who can edit this spreadsheet's script
 * can read — the same people who can use the menu anyway; and all it
 * can do is archive past week tabs.
 *
 * Safety. It confirms first, listing every tab. It works in batches of
 * about 25 seconds (each call to the service is one batch), and within
 * a batch every tab is copied and each copy checked cell for cell
 * against its original before any of them is removed; a mismatch stops
 * the run with nothing in that batch removed. Only tabs that were
 * confirmed AND are still past weeks are touched, so if the date rolls
 * over mid-run, this week's tab isn't swept up.
 * ------------------------------------------------------------------ */


var WEEK_TAB_NAME_ = /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}$/;

var ARCHIVE_SECRET_KEY_ = 'ARCHIVE_SERVICE_SECRET';   // Script Properties: what a call to the service must carry
var ARCHIVE_IN_PROGRESS_KEY_ = 'ARCHIVE_IN_PROGRESS'; // Script Properties: the archive a multi-batch run is filling
var ARCHIVE_BATCH_MS_ = 25000;                        // stop starting new tabs after this long, well inside a fetch's timeout


/** Menu action: confirms which tabs will go, has the archive service archive them, and links to the result. */
function archivePastWeeksFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ARCHIVE_SERVICE_URL) {
    ui.alert('Archive isn\'t set up yet',
      'The archive service hasn\'t been deployed, so there\'s nowhere to send the tabs. See "Archive service" in source/README.md.',
      ui.ButtonSet.OK);
    return;
  }

  var thisWeek = bulletinTabName_(ss);
  var names = pastWeekTabs_(ss).map(function (s) { return s.getName(); });
  if (!names.length) {
    ui.alert('Nothing to archive', 'There are no week tabs older than ' + thisWeek + '.', ui.ButtonSet.OK);
    return;
  }

  var owner = ownerEmail_(ss);
  var answer = ui.alert(
    'Archive ' + names.length + ' past ' + (names.length === 1 ? 'week' : 'weeks') + '?',
    'These tabs will be copied into a new spreadsheet in the Google Drive of ' +
      (owner || 'the account that owns this spreadsheet') + ', then removed from this one:\n\n' +
      names.join('\n') + '\n\n' +
      thisWeek + ', any later weeks, Template, and the reference tabs stay here.',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  var result;
  try {
    result = runArchive_(names);
  } catch (err) {
    ui.alert('Archive failed', err.message || String(err), ui.ButtonSet.OK);
    return;
  }

  // Deleting the tab you were looking at leaves Sheets to pick another one; pick this week's instead.
  var current = ss.getSheetByName(thisWeek);
  if (current) ss.setActiveSheet(current);

  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.5">' +
    '<p>Archived ' + result.tabs.length + (result.tabs.length === 1 ? ' tab' : ' tabs') + ' to ' +
    '<a href="' + escapeHtml_(result.url) + '" target="_blank">' + escapeHtml_(result.name) + ' &#8599;</a>, ' +
    'in the Google Drive of ' + escapeHtml_(result.owner) + ', which owns this spreadsheet.</p>' +
    '<p style="color:#5f6368">' + escapeHtml_(result.tabs.join(', ')) + '</p>' +
    '</div>'
  ).setWidth(460).setHeight(200);
  ui.showModalDialog(html, 'Past weeks archived');
}


/**
 * The week tabs dated before this week's Sunday (see the section note),
 * in the order they sit in the tab bar.
 */
function pastWeekTabs_(ss) {
  var thisWeek = parseBulletinDateText_(bulletinTabName_(ss)); // "yyyy-MM-dd", so strings compare as dates
  return ss.getSheets().filter(function (sheet) {
    var name = sheet.getName();
    if (!WEEK_TAB_NAME_.test(name)) return false;
    var date = parseBulletinDateText_(name);
    return date !== '' && date < thisWeek;
  });
}


/**
 * Has the archive service archive `names`, calling it once per batch
 * until it reports it's done. Returns { url, name, tabs, owner }, where
 * tabs lists every tab archived across all the batches. Throws the
 * service's own message if a batch fails; any earlier batches' tabs
 * stay archived.
 */
function runArchive_(names) {
  var archiveId = null;
  var archived = [];
  for (var batch = 0; batch < 50; batch++) {
    var r = callArchiveService_({ names: names, archiveId: archiveId });
    archiveId = r.archiveId;
    archived = archived.concat(r.archived);
    if (r.done) return { url: r.url, name: r.name, tabs: archived, owner: r.owner };
  }
  throw new Error('The archive still wasn\'t finished after 50 batches. What was archived so far is in ' +
    (archiveId ? 'https://docs.google.com/spreadsheets/d/' + archiveId : 'the archive') + '.');
}


/** One call to the archive service (see doPost); returns its result, or throws its error. */
function callArchiveService_(request) {
  request.secret = archiveServiceSecret_();
  var resp = UrlFetchApp.fetch(ARCHIVE_SERVICE_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(request),
    muteHttpExceptions: true
  });
  var body;
  try {
    body = JSON.parse(resp.getContentText());
  } catch (err) {
    throw new Error('The archive service didn\'t answer properly (HTTP ' + resp.getResponseCode() +
      '), so nothing more was archived. See "Archive service" in source/README.md.');
  }
  if (!body.ok) throw new Error(body.error);
  return body.value;
}


/**
 * The secret a call to the archive service must carry, from Script
 * Properties — made on first use. Script Properties are shared by
 * everyone who runs this project, and readable only by people who can
 * open its script, i.e. this spreadsheet's editors.
 */
function archiveServiceSecret_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty(ARCHIVE_SECRET_KEY_);
  if (secret) return secret;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    secret = props.getProperty(ARCHIVE_SECRET_KEY_); // someone else may have made it while we waited
    if (!secret) {
      secret = Utilities.getUuid() + Utilities.getUuid();
      props.setProperty(ARCHIVE_SECRET_KEY_, secret);
    }
    return secret;
  } finally {
    lock.releaseLock();
  }
}


/**
 * THE ARCHIVE SERVICE — the web app's only entry point, running as the
 * account that deployed it (this spreadsheet's owner; see the section
 * note). Takes a JSON body { secret, names, archiveId } and answers
 * { ok: true, value: <archiveBatch_'s result> } or { ok: false, error }.
 * Anything without the right secret is refused before any work.
 */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e && e.postData ? e.postData.contents : '{}');
    var secret = PropertiesService.getScriptProperties().getProperty(ARCHIVE_SECRET_KEY_);
    if (!secret || typeof req.secret !== 'string' || req.secret !== secret) throw new Error('Not authorized.');
    if (!Array.isArray(req.names) || !req.names.length) throw new Error('No tabs were named to archive.');
    out = { ok: true, value: archiveBatch_(req.names.map(String), req.archiveId ? String(req.archiveId) : null) };
  } catch (err) {
    out = { ok: false, error: err.message || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}


/**
 * One batch of an archive run, on the service side. The first batch
 * (no `archiveId`) creates the archive spreadsheet; later batches fill
 * the same one, and must name the archive this run started (kept in
 * Script Properties) so a call can't point the service at some other
 * spreadsheet.
 *
 * Copies as many of the named, still-past week tabs as fit in about 25
 * seconds, checks every copy against its original, and only then
 * removes this batch's originals. Returns { archiveId, url, name,
 * owner, archived: [names], done }.
 */
function archiveBatch_(names, archiveId) {
  return withBulletinTabLock_(function () {
    var started = Date.now();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getScriptProperties();
    var owner = ownerEmail_(ss);
    var tabs = pastWeekTabs_(ss).filter(function (s) { return names.indexOf(s.getName()) !== -1; });

    var archive;
    if (archiveId) {
      if (archiveId !== props.getProperty(ARCHIVE_IN_PROGRESS_KEY_)) {
        throw new Error('That isn\'t the archive this run started, so nothing more was archived.');
      }
      archive = SpreadsheetApp.openById(archiveId);
    } else {
      if (!tabs.length) throw new Error('None of those tabs are here to archive any more.');
      archive = createArchiveSpreadsheet_(ss, names, owner);
      props.setProperty(ARCHIVE_IN_PROGRESS_KEY_, archive.getId());
    }

    var pairs = [];
    for (var i = 0; i < tabs.length; i++) {
      if (pairs.length && Date.now() - started > ARCHIVE_BATCH_MS_) break; // always at least one per batch
      var copy = tabs[i].copyTo(archive);
      copy.setName(tabs[i].getName());
      if (copy.isSheetHidden()) copy.showSheet();
      // The dropdown rules point at Songs/Members, which the archive doesn't have.
      copy.getRange(1, 1, copy.getMaxRows(), copy.getMaxColumns()).clearDataValidations();
      pairs.push({ tab: tabs[i], copy: copy });
    }

    // A new spreadsheet starts with one blank sheet (its name depends on the locale); drop anything that isn't an archived tab.
    archive.getSheets().forEach(function (sheet) {
      if (names.indexOf(sheet.getName()) === -1 && archive.getSheets().length > 1) archive.deleteSheet(sheet);
    });

    var mismatched = pairs.filter(function (p) { return !sameSheetValues_(p.tab, p.copy); });
    if (mismatched.length) {
      throw new Error('The archive copy of ' + mismatched.map(function (p) { return p.tab.getName(); }).join(', ') +
        ' didn\'t match the original, so this batch removed nothing from this spreadsheet. The archive so far is "' +
        archive.getName() + '": ' + archive.getUrl());
    }
    pairs.forEach(function (p) { ss.deleteSheet(p.tab); });

    var done = pairs.length === tabs.length;
    if (done) props.deleteProperty(ARCHIVE_IN_PROGRESS_KEY_);
    return {
      archiveId: archive.getId(), url: archive.getUrl(), name: archive.getName(), owner: owner,
      archived: pairs.map(function (p) { return p.tab.getName(); }), done: done
    };
  });
}


/**
 * Creates the archive spreadsheet, named for the date range of `names`
 * and set to this spreadsheet's time zone — after checking it belongs
 * to this spreadsheet's owner. The service runs as whoever deployed it,
 * so a new file belonging to anyone else means it was deployed from
 * the wrong account: the still-empty file is renamed so it's obviously
 * safe to delete, and nothing is archived.
 */
function createArchiveSpreadsheet_(ss, names, owner) {
  var byDate = names.slice().sort(function (a, b) {
    return parseBulletinDateText_(a) < parseBulletinDateText_(b) ? -1 : 1;
  });
  var title = ss.getName() + ' — Archive (' + byDate[0] +
    (byDate.length > 1 ? ' to ' + byDate[byDate.length - 1] : '') + ')';

  var archive = SpreadsheetApp.create(title);
  var runner = ownerEmail_(archive); // a new spreadsheet belongs to whoever created it
  if (!owner || owner !== runner) {
    archive.rename('Unused archive — safe to delete');
    throw new Error('The archive service is running as ' + (runner || 'an unknown account') + ', not ' +
      (owner || 'the account that owns this spreadsheet') + ', so its archive wouldn\'t belong to the owner. ' +
      'Nothing was archived. It has to be deployed from the owner\'s account — see "Archive service" in ' +
      'source/README.md. (It left an empty spreadsheet called "Unused archive — safe to delete" in the Drive of ' +
      (runner || 'that account') + '.)');
  }
  archive.setSpreadsheetTimeZone(ss.getSpreadsheetTimeZone());
  return archive;
}


/** The email of the account that owns `spreadsheet`, or '' if Sheets won't say. */
function ownerEmail_(spreadsheet) {
  var owner = spreadsheet.getOwner();
  return owner ? String(owner.getEmail() || '').toLowerCase() : '';
}


/** True if two sheets hold exactly the same values in the same cells. */
function sameSheetValues_(a, b) {
  return JSON.stringify(a.getDataRange().getValues()) === JSON.stringify(b.getDataRange().getValues());
}
