/**
 * NEW BULLETIN TABS & REFERENCE TAB VISIBILITY
 * "Create New Bulletin" (duplicates the Template tab for the next Sunday)
 * and "Show/Hide Reference Tabs" (tucks Template/Songs/Members/
 * Leadership out of the tab bar, or brings them back).
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
 * Returns the newly created Sheet, or null if one already existed.
 * Throws only what createNewBulletin_ throws (e.g. no "Template" tab to
 * copy from yet) — the caller swallows that, since a missing Template
 * tab shouldn't turn into an error dialog every time the file opens.
 *
 * Worth knowing: because this runs on every open, deliberately deleting
 * the current week's tab means the next open recreates it. Deleting a
 * PAST week's tab is unaffected — only the closest upcoming Sunday's is
 * ever created here.
 */
function ensureCurrentBulletinTab_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = bulletinTabName_(ss);
  if (ss.getSheetByName(tabName)) return null;

  var sheet = createNewBulletin_();
  ss.toast('Created "' + tabName + '" from the Template tab.', 'Ward Bulletin', 8);
  return sheet;
}


/** Menu action: creates (or switches to) this week's bulletin tab — see createNewBulletin_. */
function createNewBulletinFromMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sheet = createNewBulletin_();
    ui.alert('Ready', '"' + sheet.getName() + '" is ready to fill in.', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Could not create a new bulletin', err.message || String(err), ui.ButtonSet.OK);
  }
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
 */
function createNewBulletin_() {
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
  newSheet.setName(tabName);
  if (newSheet.isSheetHidden()) newSheet.showSheet(); // Template itself may be hidden — the copy shouldn't be

  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1); // to the front, alongside the other week tabs

  try {
    fillDefaultDate_(newSheet);
  } catch (err) {
    // Non-critical — see fillDefaultDate_'s own doc comment in
    // 06_Dropdowns.gs.
  }
  try {
    refreshDropdowns_(newSheet);
  } catch (err) {
    // Non-critical — see refreshDropdowns_'s own doc comment in
    // 06_Dropdowns.gs.
  }

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

