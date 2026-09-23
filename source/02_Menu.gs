/**
 * MENU & OPEN TRIGGERS
 * Builds the "Ward Bulletin" menu and gets the sidebar to auto-open.
 * NOTE ON FILE ORDER: Apps Script concatenates every file in a project
 * into one script before running it. Function declarations (like all of
 * these) are hoisted, so it never matters which file a function lives
 * in or which order the files load — any file can call any other file's
 * function. Only top-level `var ... = ...` values are order-sensitive
 * (see 01_Config.gs's note), which is why this project keeps them all
 * together in one file instead of scattering them.
 */

/**
 * Adds the "Ward Bulletin" menu when the spreadsheet is opened, and
 * (best-effort) makes sure the sidebar keeps auto-opening on future
 * opens too — see ensureAutoOpenTrigger_() below. Menu creation always
 * happens first and unconditionally, with nothing that could throw
 * ahead of it: a plain onOpen() simple trigger silently drops the
 * ENTIRE function — including the menu — if anything inside it throws
 * before .addToUi() runs, with no visible error (that's exactly how an
 * earlier version of this lost its menu entirely). Auto-open setup runs
 * only after, wrapped in its own try/catch, so it can never risk that
 * again. The menu's own "Publish" and "Preview" items match the
 * sidebar's buttons in both name and behavior — Preview calls
 * previewInModal() directly, and Publish goes through publishFromMenu()
 * (a thin wrapper around the same publishBulletinForSidebar() the
 * sidebar button calls) — so either one opens the exact same modal
 * dialog.
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
 * this lost its menu entirely). The menu's own "Publish" and "Preview"
 * items match the sidebar's buttons in both name and behavior —
 * Preview calls previewInModal() directly, and Publish goes through
 * publishFromMenu() (a thin wrapper around the same
 * publishBulletinForSidebar() the sidebar button calls) — so either one
 * opens the exact same modal dialog.
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

  var menu = SpreadsheetApp.getUi()
    .createMenu('Ward Bulletin')
    .addItem('🧰 Show Toolbar', 'showToolbar')
    .addSeparator()
    .addItem('📤 Publish', 'publishFromMenu')
    .addItem('👁️ Preview', 'previewInModal')
    .addSeparator()
    .addItem('🆕 Create New Bulletin', 'createNewBulletinFromMenu')
    .addItem('🔄 Refresh Dropdowns', 'refreshDropdownsFromMenu')
    .addItem('🩹 Fix "Buisness" Typo', 'fixBusinessSpellingFromMenu')
    .addSeparator();

  if (hidden) {
    menu.addItem('➕ Show Reference Tabs', 'showUtilityTabsFromMenu');
  } else {
    menu.addItem('➖ Hide Reference Tabs', 'hideUtilityTabsFromMenu');
  }

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
 * creating one from the Template tab if it doesn't (see
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
    // Non-critical — see the try/catch below.
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
 * Menu action (and the installable trigger's target above): opens a
 * sidebar with "Publish" and "Preview" buttons, so those two actions
 * are one click away without opening the menu each time. See
 * Sidebar.html for the button markup/JS. Whether the Preview button
 * should be highlighted, and the last-published tab (if any) with a
 * link to view it (see getPublishStatus and friends in 03_Publish.gs),
 * are baked into the sidebar's JS at render time so it reflects reality
 * the instant it opens, with no server round-trip. Both buttons call a
 * server function via google.script.run that opens the actual result —
 * the preview, or the freshly-published page — in a modal dialog right
 * over the sheet (see showResultModal_ in 04_Preview.gs), so it feels
 * like part of the same window instead of jumping out to a new tab.
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
  tmpl.hasChanges = hasUnpublishedChanges();

  var lastPublished = getLastPublishedInfo_();
  tmpl.pagesUrl = lastPublished.pagesUrl;
  tmpl.lastPublishedSheet = lastPublished.sheetName;
  tmpl.lastPublishedAt = lastPublished.publishedAt;

  var html = tmpl.evaluate().setTitle('Ward Bulletin');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ---------------------------------------------------------------------
 * PUBLISH TRACKING — remembers which sheet tab was last published (and
 * when), so the sidebar can link to what's actually live and highlight
 * the Preview button whenever the active tab isn't a match for that
 * anymore: either a different tab is active than the one last
 * published, or it's the same tab but its content has changed since.
 * Content is compared as a hash (not a raw edit timestamp) so an edit
 * elsewhere in the sheet that doesn't actually change the rendered
 * bulletin — a note in an unused cell, formatting, etc. — doesn't
 * falsely flag it.
 * ------------------------------------------------------------------ */

