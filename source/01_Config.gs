/**
 * CONFIGURATION
 * Ward name, GitHub Pages destination, the deployed web app URL, and the
 * sheet/tab names every other file looks up by name (Songs, Members,
 * Leadership, Template). Loads first — SONGS_SHEET_NAME_ etc. must be
 * defined before UTILITY_SHEET_NAMES_ below reads them, and Apps Script
 * runs each file's top-level `var` statements in the order the files
 * are listed in the editor's left-hand file list (drag to reorder there
 * if that list is ever changed — see the note in 02_Menu.gs).
 */

/**
 * WARD BULLETIN GENERATOR — Google Apps Script
 * ---------------------------------------------
 * Bound to the "Sacrament Meeting 2026" spreadsheet. Reads whichever
 * sheet tab is currently active (e.g. "August 30, 2026") and renders it
 * into Template.html (a copy of ward-bulletin-template.html rewritten to
 * use Apps Script scriptlets, <?= ... ?>, instead of [Bracket] tokens).
 *
 * SETUP (one-time)
 * 1. In the spreadsheet: Extensions > Apps Script.
 * 2. Delete the default empty Code.gs. This project is split across
 *    several script files instead of one — 01_Config.gs (this file)
 *    through 07_BulletinTabs.gs. Add each one as its own file (+ > Script)
 *    named to match (Apps Script appends ".gs" itself, so name the file
 *    just "01_Config", "02_Menu", etc.) and paste in its contents. The
 *    numeric prefixes keep them in a sensible reading order in the
 *    editor's file list; see the note below on why 01_Config.gs
 *    specifically has to load before the others.
 * 3. Add a file (HTML type) named exactly "Template", and paste
 *    Template.html's content into it. Add another (HTML type) named
 *    exactly "Sidebar", and paste Sidebar.html's content into it — that's
 *    what powers the "Show Toolbar" menu item's Publish/Preview buttons.
 * 4. Deploy > New deployment > select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone with the link (or "Anyone in [org]" if you
 *      want it restricted to BYU/ward accounts)
 *    - Click Deploy, authorize the requested permissions.
 * 5. Reload the spreadsheet. A "Ward Bulletin" menu appears, with a
 *    sidebar of Publish/Preview buttons alongside it (use any menu item
 *    once to authorize this project, if prompted — that's also what
 *    lets the sidebar start auto-opening on future opens; see onOpen()
 *    in 02_Menu.gs). "Show Toolbar" reopens the sidebar any time you've
 *    closed it, and the individual menu items below it do the same thing
 *    the matching sidebar button does.
 * 6. Whenever you redeploy (Deploy > Manage deployments > edit > New
 *    version) after editing this code, the same URL keeps working.
 *
 * HOW IT FINDS DATA
 * Rather than hardcoded cell references (which would break as rows get
 * added/removed for stake business, releases, etc.), this scans column A
 * of the active sheet for known labels ("Presiding:", "Opening Hymn", ...)
 * and reads the value from column B of that same row. That matches how
 * the "Conducting Program" sheet is laid out today. If you rename a
 * label in the sheet, update the matching string in the field-label lists
 * near the top of 05_BulletinData.gs (KNOWN_FIELD_LABELS_, etc.).
 *
 * KNOWN LIMITATIONS (see the message in chat for more detail)
 * - Between "Sacrament Hymn" and "Benediction", any row whose label
 *   contains "speaker" becomes a name row, any row whose label contains
 *   "testimon" becomes a testimony-meeting row (for fast Sundays — see
 *   getSacramentProgram_), and any row whose label contains "hymn" or
 *   "music" becomes a hymn/music row — in whatever order and however
 *   many of each appear in the sheet. Anything else in that range (like
 *   the "The Administration of the Sacrament" header line) is skipped.
 * - A hymn/music/testimony row's number and title can be entered either
 *   of two ways (see resolveNumTitle_) — combined in column B, split on
 *   its first "-"/"–"/"—"/"|" (e.g. "#1062 - Lord, Accept Our Humble
 *   Fast", or "Ward Choir - Lead, Kindly Light" for a non-numeric
 *   group/performer name), or split across two cells, with column B
 *   holding just the number/name ("#1062") and column C holding just
 *   the title. Column C wins whenever a row has anything in it, so the
 *   two formats can even be mixed row to row. Either way, the number/
 *   name lands in the right-hand slot and the title is the centered
 *   line below it.
 * - getActiveSheet() reflects whichever tab an editor last had open in
 *   this spreadsheet, not a per-visitor selection — see the note in the
 *   chat reply about what that means with multiple weekly tabs.
 * - "Stake Business" and "Ward Business" (between "Invocation" and
 *   "Sacrament Hymn") are both optional, heading-only lines — like "The
 *   Sacrament" — that print no value of their own, they just switch on
 *   or off. "Stake Business" prints whenever the "Stake Business" section
 *   (its header plus whatever sub-rows follow, e.g. "Presented By:") has
 *   a value in column B on any of those sub-rows — see
 *   hasStakeBusinessContent_. "Ward Business" has no row of its own at
 *   all; it prints whenever anyone is actually listed under the
 *   "Releases"/"Sustaining" section's NAMES table — see
 *   hasReleasesOrSustainingNames_. Leave the section empty and the
 *   matching heading is skipped.
 * - Song and speaker cells (any label containing "hymn"/"music", or
 *   "speaker") get a searchable, type-ahead dropdown sourced from the
 *   "Songs" and "Members" sheet tabs — see 06_Dropdowns.gs for how it's
 *   built and refreshed. Typing something not on either list is still
 *   accepted (just flagged with a small warning triangle), so this never
 *   blocks a guest speaker or a hymn outside the hymnal.
 * - "Ward Bulletin > Create New Bulletin" duplicates a "Template" tab
 *   into a new tab for the closest upcoming Sunday — see
 *   07_BulletinTabs.gs. You need to add that Template tab yourself (a
 *   blank week's program, laid out like any other week's tab).
 *   "Show/Hide Reference Tabs" tucks the Template/Songs/Members/
 *   Leadership tabs out of the tab bar without affecting anything this
 *   script reads from them.
 *
 * PUBLISHING TO GITHUB PAGES (one-time setup)
 * 1. Create a public GitHub repo (e.g. "ward-bulletin"), and in its
 *    Settings > Pages, set Source to "Deploy from a branch" / main / root.
 *    That serves https://<your-username>.github.io/<repo-name>/.
 * 2. GitHub > Settings > Developer settings > Personal access tokens >
 *    Fine-grained tokens > Generate new token. Repository access: only
 *    this repo. Permissions: Contents = Read and write. Copy the token
 *    (shown once).
 * 3. Edit GITHUB_OWNER / GITHUB_REPO / GITHUB_BRANCH / GITHUB_FILE_PATH
 *    below to match your repo.
 * 4. Reload the spreadsheet, use "Ward Bulletin > Set GitHub Token…" and
 *    paste the token in. It's stored in this script's Properties, not in
 *    the visible code, so it stays out of anything you share or copy.
 * 5. "Ward Bulletin > Publish" (menu or sidebar) now renders whichever
 *    sheet tab is active and commits it as index.html in that repo. GitHub Pages
 *    usually picks up the change within a few seconds to a minute.
 *
 * Each publish also archives the copy it's replacing to
 * GITHUB_ARCHIVE_DIR/<that bulletin's own Sunday>.html — the date read
 * out of the file being replaced, not the day the publish happens (see
 * archiveExistingBulletin_ in 03_Publish.gs).
 *
 * A readable copy of this project's own code lives in the repo's
 * source/ folder, kept up to date by hand — nothing in the script
 * writes it. Editing a file there changes nothing on its own; the Apps
 * Script project is what actually runs. 
 * My Test!!!!
 */

var WARD_NAME = 'Edgemont 21st Ward'; // shown in the header bar and page title


// GitHub Pages destination — edit these three to match your repo.
var GITHUB_OWNER = 'edgemont21ward';


var GITHUB_REPO = 'ward-bulletin';


var GITHUB_BRANCH = 'main';


var GITHUB_FILE_PATH = 'index.html';


var GITHUB_ARCHIVE_DIR = 'archive'; // where the previous index.html goes before each publish overwrites it


// The web app's deployed URL, for the "Preview" menu item. Apps Script's
// own ScriptApp.getService().getUrl() is unreliable when called from a
// menu action (a known V8-runtime bug — it can return null or a stale
// URL), so this is hardcoded instead. Copy it from Deploy > Manage
// deployments in the Apps Script editor (it stays the same across "New
// version" updates — only creating an entirely separate deployment
// changes it).
var WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbycG3HLZ1XNAm6W1n7EQ49V4t_rE6GJ_Ig_NbzuIGE3UydFo7D5o6IXy8PMVxoqJ8XP/exec';


var SONGS_SHEET_NAME_ = 'Songs';


var MEMBERS_SHEET_NAME_ = 'Members';


var LEADERSHIP_SHEET_NAME_ = 'Leadership';


var TEMPLATE_SHEET_NAME_ = 'Template';


var UTILITY_SHEET_NAMES_ = [TEMPLATE_SHEET_NAME_, SONGS_SHEET_NAME_, MEMBERS_SHEET_NAME_, LEADERSHIP_SHEET_NAME_];


// clasp deploy test 10:54 — safe to delete

// clasp deploy test 10:54 — safe to delete
