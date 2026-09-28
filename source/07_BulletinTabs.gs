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
  var lock = LockService.getDocumentLock();
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
 * It's careful about the removing: nothing is deleted until every tab
 * has been copied and each copy checked cell for cell against its
 * original.
 *
 * The archive has to end up in the Drive of the account that owns this
 * spreadsheet. A new spreadsheet always belongs to, and lands in the
 * Drive of, whoever creates it — no permission changes that — so this
 * only works when run from that account, and it checks: if the new
 * archive's owner isn't this spreadsheet's owner, it stops before
 * copying or removing anything (see archivePastWeeks_). Run from the
 * owner's account, the archive lands at the top of that My Drive, which
 * is also where this spreadsheet is.
 * ------------------------------------------------------------------ */


var WEEK_TAB_NAME_ = /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}$/;


/** Menu action: confirms which tabs will go, archives them, and links to the result. */
function archivePastWeeksFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    result = archivePastWeeks_(names);
  } catch (err) {
    ui.alert('Archive failed', err.message || String(err), ui.ButtonSet.OK);
    return;
  }

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
 * Copies the past week tabs named in `names` into a new spreadsheet,
 * checks each copy against its original, and only then deletes the
 * originals. Only tabs that were confirmed AND are still past weeks
 * are touched, so if the date rolls over while the confirmation is
 * open, this week's tab isn't swept up with them.
 *
 * In the copies, the dropdown rules are removed: they point at the
 * Songs and Members tabs, which the archive doesn't have. Values,
 * formatting, merges and links are kept.
 *
 * If any copy doesn't match its original, nothing is deleted here and
 * the error names the half-made archive, so it can be checked or thrown
 * away.
 *
 * Refuses unless it's being run by this spreadsheet's owner, since the
 * new spreadsheet belongs to (and lands in the Drive of) whoever runs
 * this — see the section note. That's checked straight after creating
 * the archive, the one point where the runner's identity is known
 * without asking for another permission: the new file's owner is the
 * runner. On a mismatch the still-empty spreadsheet is renamed so it's
 * obviously safe to delete, and nothing else happens.
 *
 * Returns { url, name, tabs, owner } for the new spreadsheet.
 */
function archivePastWeeks_(names) {
  return withBulletinTabLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tabs = pastWeekTabs_(ss).filter(function (s) { return names.indexOf(s.getName()) !== -1; });
    if (!tabs.length) throw new Error('None of those tabs are here to archive any more.');

    var byDate = tabs.map(function (s) { return s.getName(); }).sort(function (a, b) {
      return parseBulletinDateText_(a) < parseBulletinDateText_(b) ? -1 : 1;
    });
    var title = ss.getName() + ' — Archive (' + byDate[0] +
      (byDate.length > 1 ? ' to ' + byDate[byDate.length - 1] : '') + ')';

    var archive = SpreadsheetApp.create(title);
    var owner = ownerEmail_(ss);
    var runner = ownerEmail_(archive); // a new spreadsheet is owned by whoever created it
    if (!owner || owner !== runner) {
      archive.rename('Unused archive — safe to delete');
      throw new Error('Archive Past Weeks has to be run from ' + (owner || 'the account that owns this spreadsheet') +
        ', which owns this spreadsheet, so the archive goes into that account\'s Google Drive. Nothing was ' +
        'archived or removed. It did leave an empty spreadsheet called "Unused archive — safe to delete" in ' +
        (runner ? 'the Drive of ' + runner : 'your Google Drive') + '; you can delete it.');
    }
    archive.setSpreadsheetTimeZone(ss.getSpreadsheetTimeZone());
    var placeholder = archive.getSheets()[0]; // every new spreadsheet starts with a "Sheet1"

    var pairs = tabs.map(function (tab) {
      var copy = tab.copyTo(archive);
      copy.setName(tab.getName());
      if (copy.isSheetHidden()) copy.showSheet();
      copy.getRange(1, 1, copy.getMaxRows(), copy.getMaxColumns()).clearDataValidations();
      return { tab: tab, copy: copy };
    });
    archive.deleteSheet(placeholder);

    var mismatched = pairs.filter(function (p) { return !sameSheetValues_(p.tab, p.copy); });
    if (mismatched.length) {
      throw new Error('The archive copy of ' + mismatched.map(function (p) { return p.tab.getName(); }).join(', ') +
        ' didn\'t match the original, so no tabs were removed from this spreadsheet. The copies made so far ' +
        'are in "' + archive.getName() + '" in your Google Drive: ' + archive.getUrl());
    }

    // Deleting the tab you're looking at leaves Sheets to pick another one; pick this week's instead.
    var current = ss.getSheetByName(bulletinTabName_(ss));
    if (current) ss.setActiveSheet(current);
    pairs.forEach(function (p) { ss.deleteSheet(p.tab); });

    return { url: archive.getUrl(), name: archive.getName(), tabs: tabs.map(function (s) { return s.getName(); }), owner: owner };
  });
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
