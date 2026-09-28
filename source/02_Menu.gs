/**
 * MENU & OPEN TRIGGERS
 * Builds the "Ward Bulletin" menu and gets the sidebar to auto-open.
 */

/**
 * Adds the "Ward Bulletin" menu when the spreadsheet is opened, then
 * (best-effort) makes sure the sidebar keeps auto-opening on future
 * opens — see ensureAutoOpenTrigger_() below. The menu is built first,
 * and the trigger setup is wrapped in its own try/catch, because
 * anything that throws in a simple onOpen() takes the menu down with
 * it — see buildMenu_().
 */
function onOpen() {
  buildMenu_();

  try {
    ensureAutoOpenTrigger_();
  } catch (err) {
    // A brand-new, never-yet-authorized spreadsheet can't create
    // triggers from this restricted context — showToolbar() below
    // (always fully authorized, since it only ever runs from a real
    // menu click) is the reliable fallback place for this to happen
    // instead, the first time any menu item gets used.
  }
}


/**
 * Builds (or rebuilds) the "Ward Bulletin" menu. Called from onOpen()
 * itself, and again — with nothing else in between — from
 * hideUtilityTabsFromMenu()/showUtilityTabsFromMenu() right after they
 * change the reference tabs' visibility, so the menu's Show/Hide
 * Reference Tabs item flips to the other one immediately, without
 * waiting for the spreadsheet to be reopened (calling createMenu() +
 * addToUi() again for the same menu name replaces it in place — Sheets
 * doesn't need a page reload to pick that up).
 *
 * Only one of "➕ Show Reference Tabs" / "➖ Hide Reference Tabs" is
 * ever added, chosen by areUtilityTabsHidden_() — so exactly one, never
 * both, appears in the menu at a time.
 *
 * Google Sheets' custom menu API is text-only — there's no way to
 * attach a real icon image to a menu item — so an emoji prefix on each
 * label is the closest practical stand-in for one.
 *
 * Building the menu always happens first and unconditionally, with
 * nothing that could throw ahead of it (areUtilityTabsHidden_() is
 * wrapped in its own try/catch below for exactly this reason): a plain
 * onOpen() simple trigger silently drops the ENTIRE function —
 * including the menu — if anything inside it throws before .addToUi()
 * runs, with no visible error (that's exactly how an earlier version of
 * this lost its menu entirely).
 *
 * The menu's "Publish" and "Preview" do exactly what the sidebar's
 * buttons do: Preview calls previewInModal() directly, and Publish goes
 * through publishFromMenu(), a thin wrapper around the sidebar's own
 * publishBulletinForSidebar(). Either way the same dialog opens.
 */
function buildMenu_() {
  var hidden = false;
  try {
    hidden = areUtilityTabsHidden_();
  } catch (err) {
    // Can't tell — default to offering "Hide", same as when none of
    // the reference tabs exist yet (see areUtilityTabsHidden_ in
    // 07_BulletinTabs.gs). This must never throw past this point; see
    // the doc comment above.
  }

  var ui = SpreadsheetApp.getUi();
  var reportsMenu = ui.createMenu('📊 Reports')
    .addItem('Speakers', 'showSpeakersReport')
    .addItem('Prayers', 'showPrayersReport');

  var addRowMenu = ui.createMenu('➕ Add Program Row')
    .addItem('Speaker', 'addSpeakerRowFromMenu')
    .addItem('Intermediate Hymn', 'addHymnRowFromMenu')
    .addItem('Musical Number', 'addMusicalNumberRowFromMenu')
    .addItem('Testimonies', 'addTestimoniesRowFromMenu');

  var menu = ui
    .createMenu('Ward Bulletin')
    .addItem('🧰 Show Toolbar', 'showToolbar')
    .addSeparator()
    .addItem('📤 Publish', 'publishFromMenu')
    .addItem('👁️ Preview', 'previewInModal')
    .addSeparator()
    .addItem('🆕 Create New Bulletin', 'createNewBulletinFromMenu')
    .addSubMenu(addRowMenu)
    .addItem('🔄 Refresh Dropdowns', 'refreshDropdownsFromMenu')
    .addItem('🎵 Update Songs', 'updateSongsFromMenu')
    .addItem('👥 Update Ward Members', 'updateMembersFromMenu')
    .addSeparator()
    .addSubMenu(reportsMenu)
    .addSeparator();

  if (hidden) {
    menu.addItem('➕ Show Reference Tabs', 'showUtilityTabsFromMenu');
  } else {
    menu.addItem('➖ Hide Reference Tabs', 'hideUtilityTabsFromMenu');
  }
  menu.addItem('🗄️ Archive Past Weeks…', 'archivePastWeeksFromMenu');

  menu
    .addSeparator()
    .addItem('🔑 Set GitHub Token…', 'promptForGithubToken')
    .addToUi();
}


/**
 * Creates the installable "on open" trigger that fires
 * onOpenInstallable() below, if one doesn't already exist. A plain
 * onOpen() simple trigger (above) is documented to be unable to show a
 * sidebar or dialog itself, no matter what, regardless of
 * authorization — an installable trigger, which runs at full
 * authorization, is the only way around that. Safe to call repeatedly:
 * it only ever creates the trigger once.
 */
function ensureAutoOpenTrigger_() {
  var alreadyExists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'onOpenInstallable';
  });
  if (alreadyExists) return;

  ScriptApp.newTrigger('onOpenInstallable')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();
}


/**
 * The installable "on open" trigger itself — fires on the same open
 * event as the simple onOpen() above, but (unlike it) at full
 * authorization, so it can actually show the sidebar automatically.
 *
 * It also makes sure a tab exists for the closest upcoming Sunday,
 * creating one from the Template tab if it doesn't, and deletes any
 * hidden copies of Template left over from an unfinished creation (see
 * ensureCurrentBulletinTab_ in 07_BulletinTabs.gs) — so opening this on
 * a Sunday with no tab for that day just produces one. Everything after
 * showToolbar() is wrapped in its own try/catch: none of it is allowed
 * to take the sidebar down with it if it fails.
 */
function onOpenInstallable() {
  // showToolbar() runs first and unconditionally — the sidebar showing
  // up is the whole point of this trigger, so it can never be made to
  // wait on (or be skipped because of) anything below.
  showToolbar();

  try {
    ensureCurrentBulletinTab_();
  } catch (err) {
    // Non-critical — the usual cause is simply not having added a
    // "Template" tab yet, which shouldn't produce an error dialog every
    // single time the file is opened. "Ward Bulletin > Create New
    // Bulletin" reports the real reason when it's used deliberately.
  }

  try {
    fillDefaultDate_();
  } catch (err) {
    // Non-critical, like the dropdown refresh below: a blank Date cell
    // is easy to fill in by hand.
  }

  try {
    refreshDropdowns_();
  } catch (err) {
    // Non-critical — a validation hiccup (e.g. a missing Songs/Members
    // sheet) shouldn't affect anything else this trigger does. "Refresh
    // Dropdowns" in the menu can be used to retry once fixed.
  }
}


/**
 * Menu action (and the installable trigger's target above): opens the
 * sidebar (Sidebar.html): the status card, Preview and Publish, the
 * Add to the program buttons, and the Reports buttons. The publish
 * status (getPublishStatus in 03_Publish.gs) is built into the page so
 * the status card is right the moment it opens, with no round trip.
 * Preview, Publish and the reports open in a dialog over the sheet.
 */
function showToolbar() {
  try {
    ensureAutoOpenTrigger_();
  } catch (err) {
    // Already set up, or a genuinely transient failure — either way,
    // this isn't essential to showing the sidebar right now, so don't
    // let it block that.
  }

  var tmpl = HtmlService.createTemplateFromFile('Sidebar');
  tmpl.statusJson = jsonForScript_(getPublishStatus());
  var html = tmpl.evaluate().setTitle('Ward Bulletin');
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * `value` as JSON that's safe to print raw into a page's <script> with
 * <?!= ?>: "</script>" can't end the script early, and the two line
 * separators JavaScript strings can't hold are escaped too.
 */
function jsonForScript_(value) {
  return JSON.stringify(value)
    .replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}
