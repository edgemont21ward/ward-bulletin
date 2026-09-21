/**
 * PUBLISHING (web app + GitHub Pages)
 * Everything involved in getting the active sheet tab's bulletin out into
 * the world: doGet() serves it at the deployed web app URL, and the
 * publishToGithub_()/archive functions commit it to GitHub Pages.
 */

/** Serves the rendered bulletin at the web app's URL. */
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = buildBulletinData(sheet);
  data.wardName = WARD_NAME;

  var tmpl = HtmlService.createTemplateFromFile('Template');
  tmpl.data = data;
  return tmpl.evaluate()
    .setTitle(WARD_NAME + ' Bulletin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/** MD5 hash of a string, base64-encoded — cheap way to compare rendered HTML for equality. */
function hashContent_(str) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}


/** The GitHub Pages URL the "Publish" button commits to — the same for every sheet tab. */
function getPagesUrl_() {
  return 'https://' + GITHUB_OWNER + '.github.io/' + GITHUB_REPO + '/';
}


/**
 * Records which sheet tab was just published: its name (for the
 * sidebar's "Last published" line and for hasUnpublishedChanges()
 * below), its rendered content's hash (so a later edit to that same
 * tab can be detected), a human-readable timestamp of when it went up,
 * and — as LAST_PUBLISHED_DATE_SLUG — a filename-safe "yyyy-MM-dd" of
 * THE SUNDAY THAT BULLETIN IS FOR, which is a different thing entirely
 * from when it was published (a bulletin for the 27th is usually
 * published some evening earlier that week).
 *
 * That slug is the standby for archiveExistingBulletin_ the *next* time
 * something is published: the copy being overwritten gets archived
 * under its own Sunday's date. It's read out of the rendered HTML
 * itself (the same date the bulletin prints), falling back to the tab
 * name, which is normally that same date anyway.
 *
 * Returns the human-readable timestamp so callers that already have
 * Document Properties open (like publishBulletinForSidebar()) don't
 * need to read it back.
 */
function recordPublished_(sheet, html) {
  var props = PropertiesService.getDocumentProperties();
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var now = new Date();
  var publishedAt = Utilities.formatDate(now, tz, 'MMM d, yyyy h:mm a');
  props.setProperty('LAST_PUBLISHED_SHEET', sheet.getName());
  props.setProperty('LAST_PUBLISHED_HASH', hashContent_(html));
  props.setProperty('LAST_PUBLISHED_AT', publishedAt);
  props.setProperty(
    'LAST_PUBLISHED_DATE_SLUG',
    bulletinDateSlugFromHtml_(html) ||
      parseBulletinDateText_(sheet.getName()) ||
      archiveDateSlug_(now, tz)
  );
  return publishedAt;
}


/**
 * Sidebar-callable: {sheetName, publishedAt, pagesUrl} for whatever was
 * last published — sheetName and publishedAt are '' if nothing's been
 * published yet (pagesUrl is always the same fixed GitHub Pages URL,
 * published or not, so the sidebar can offer it once something exists
 * to look at there).
 */
function getLastPublishedInfo_() {
  var props = PropertiesService.getDocumentProperties();
  return {
    sheetName: props.getProperty('LAST_PUBLISHED_SHEET') || '',
    publishedAt: props.getProperty('LAST_PUBLISHED_AT') || '',
    pagesUrl: getPagesUrl_()
  };
}


/**
 * Sidebar-callable: true if the active sheet tab isn't a match for
 * what's actually live on GitHub Pages — either it's a different tab
 * than the one last published, the same tab has been edited since, or
 * nothing has ever been published at all. Called once when the
 * sidebar first opens, and polled periodically while it stays open
 * (see Sidebar.html) so the highlight stays current even if you switch
 * tabs or keep editing without reopening the sidebar.
 */
function hasUnpublishedChanges() {
  var props = PropertiesService.getDocumentProperties();
  var lastSheet = props.getProperty('LAST_PUBLISHED_SHEET');
  if (!lastSheet) return true; // nothing published yet

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== lastSheet) return true; // a different tab than what's live

  var lastHash = props.getProperty('LAST_PUBLISHED_HASH') || '';
  return hashContent_(renderBulletinHtml_(sheet)) !== lastHash;
}


/**
 * Sidebar-callable: getLastPublishedInfo_() plus hasUnpublishedChanges()
 * in one round trip — what Sidebar.html's periodic poll calls so the
 * Preview button's highlight and the "last published" panel both stay
 * current no matter which surface actually did the publishing. A
 * publish can come from the sidebar's own button, the "Ward Bulletin"
 * menu, or the Preview modal's in-dialog Publish button — three
 * separate pieces of UI with no way to call back into each other
 * directly (a modal or menu action can't reach into the sidebar's own
 * JS) — so instead the sidebar just asks the server for the current
 * state every few seconds and updates itself to match, regardless of
 * who changed it.
 */
function getPublishStatus() {
  var status = getLastPublishedInfo_();
  status.hasChanges = hasUnpublishedChanges();
  return status;
}


/**
 * Does the actual publish: renders the active sheet tab, commits it to
 * GitHub Pages, and opens the freshly-published page in a modal dialog.
 * Returns a plain result object (or throws) so the sidebar's "Publish"
 * button can render a short status line too; publishFromMenu() below
 * is the menu's equivalent, adding a ui.alert() on failure since a menu
 * item has no status line of its own.
 */
function publishBulletinForSidebar() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    throw new Error('No GitHub token set yet. Use "Ward Bulletin > Set GitHub Token…" first.');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var html = renderBulletinHtml_(sheet);
  publishToGithub_(html, token);
  var publishedAt = recordPublished_(sheet, html);

  var pagesUrl = getPagesUrl_();
  showResultModal_(pagesUrl, 'Bulletin Publish — ' + sheet.getName());

  return {
    sheetName: sheet.getName(),
    pagesUrl: pagesUrl,
    publishedAt: publishedAt
  };
}


/**
 * Menu action: same publish as the sidebar's "Publish" button (renders
 * the active sheet tab, commits it to GitHub Pages, and opens the
 * result in the same modal dialog via publishBulletinForSidebar()) —
 * this wrapper just adds a plain ui.alert() on failure, since a menu
 * item has no sidebar status line to show an error in.
 */
function publishFromMenu() {
  try {
    publishBulletinForSidebar();
  } catch (err) {
    SpreadsheetApp.getUi().alert('Publish failed', err.message || String(err), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/** Menu action: prompts for and stores a GitHub token in Script Properties. */
function promptForGithubToken() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'GitHub Personal Access Token',
    'Paste a fine-grained token scoped to just the "' + GITHUB_REPO +
      '" repo, with Contents: Read and write permission.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var token = resp.getResponseText().trim();
  if (!token) return;

  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  ui.alert('Saved. The token lives in this script\'s Properties, not in the visible code.');
}

/* ---------------------------------------------------------------------
 * GITHUB PAGES PUBLISHING
 * ------------------------------------------------------------------ */


/** Renders the given sheet tab to a raw HTML string (no web app wrapper). */
function renderBulletinHtml_(sheet) {
  var data = buildBulletinData(sheet);
  data.wardName = WARD_NAME;
  var tmpl = HtmlService.createTemplateFromFile('Template');
  tmpl.data = data;
  return tmpl.evaluate().getContent();
}


/**
 * Creates or updates GITHUB_FILE_PATH in the GitHub repo via the
 * Contents API: https://docs.github.com/en/rest/repos/contents
 *
 * Before it's overwritten, whatever's currently live at GITHUB_FILE_PATH
 * is first archived (see archiveExistingBulletin_) — a no-op on the
 * very first publish, when there's nothing live yet to archive.
 */
function publishToGithub_(html, token) {
  var apiUrl = 'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO +
    '/contents/' + GITHUB_FILE_PATH;

  // An existing file must be updated with its current sha; a new file omits it.
  var sha = null;
  var getResp = UrlFetchApp.fetch(apiUrl + '?ref=' + GITHUB_BRANCH, {
    method: 'get',
    headers: githubHeaders_(token),
    muteHttpExceptions: true
  });
  if (getResp.getResponseCode() === 200) {
    var existing = JSON.parse(getResp.getContentText());
    sha = existing.sha;
    archiveExistingBulletin_(existing, token);
  }

  var payload = {
    message: 'Update bulletin (' + new Date().toISOString() + ')',
    content: Utilities.base64Encode(html, Utilities.Charset.UTF_8),
    branch: GITHUB_BRANCH
  };
  if (sha) payload.sha = sha;

  var putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: githubHeaders_(token),
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = putResp.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error('GitHub API error ' + code + ': ' + putResp.getContentText());
  }
}


/**
 * Copies `existing` — the GitHub Contents API's GET response for the
 * about-to-be-overwritten GITHUB_FILE_PATH — to
 * GITHUB_ARCHIVE_DIR/<date>.html.
 *
 * <date> is the Sunday THAT bulletin was for, read out of the file
 * being archived itself (its own printed date — see
 * bulletinDateSlugFromHtml_), NOT the date the archiving happens on and
 * not the date it was pushed. Archiving last week's bulletin while
 * publishing this week's therefore files it as
 * archive/2026-09-20.html, matching the bulletin, whatever day of the
 * week the two publishes happened to fall on.
 *
 * Two fallbacks, in order, for a file whose date can't be read that way
 * (hand-edited markup, say): LAST_PUBLISHED_DATE_SLUG, which
 * recordPublished_ stored when this script published that very file,
 * and finally today's date — so an unreadable file still gets archived
 * under something rather than being silently dropped.
 *
 * Overwrites that Sunday's archive file if one already exists —
 * republishing a correction to the same week's bulletin just replaces
 * that week's archived copy instead of erroring.
 */
function archiveExistingBulletin_(existing, token) {
  var slug = bulletinDateSlugFromHtml_(decodeGithubBase64_(existing.content)) ||
    PropertiesService.getDocumentProperties().getProperty('LAST_PUBLISHED_DATE_SLUG') ||
    archiveDateSlug_(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone());

  var archivePath = GITHUB_ARCHIVE_DIR + '/' + slug + '.html';
  var apiUrl = 'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO +
    '/contents/' + archivePath;

  // Does that day's archive file already exist (a same-day republish)?
  var archiveSha = null;
  var getResp = UrlFetchApp.fetch(apiUrl + '?ref=' + GITHUB_BRANCH, {
    method: 'get',
    headers: githubHeaders_(token),
    muteHttpExceptions: true
  });
  if (getResp.getResponseCode() === 200) {
    archiveSha = JSON.parse(getResp.getContentText()).sha;
  }

  var payload = {
    message: 'Archive bulletin for ' + slug,
    content: existing.content, // GitHub's own base64 for that file, passed straight through
    branch: GITHUB_BRANCH
  };
  if (archiveSha) payload.sha = archiveSha;

  var putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: githubHeaders_(token),
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = putResp.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error('GitHub API error archiving the previous bulletin (' + code + '): ' + putResp.getContentText());
  }
}


/** Formats `date` as a filename-safe "yyyy-MM-dd" for archiveExistingBulletin_/recordPublished_. */
function archiveDateSlug_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
}


/**
 * Digs a bulletin's own date out of its rendered HTML and returns it as
 * a filename-safe "yyyy-MM-dd", or '' if there isn't one to find. This
 * is what lets an archived copy be named after the Sunday it's for
 * rather than the day it happened to be replaced.
 *
 * First choice is the date line Template.html actually prints — the
 * .bulletin-date block, whose <span> holds exactly the text
 * buildBulletinData produced ("August 30, 2026"). Failing that (older
 * or hand-edited markup), it takes the first "Month D, YYYY" anywhere
 * in the page, which in practice is still the bulletin's own date: it
 * sits up in the header, ahead of the program and the announcements.
 */
function bulletinDateSlugFromHtml_(html) {
  if (!html) return '';

  var dateLine = String(html).match(/bulletin-date[\s\S]{0,300}?<span[^>]*>([^<]+)<\/span>/);
  return (dateLine ? parseBulletinDateText_(dateLine[1]) : '') || parseBulletinDateText_(html);
}


/**
 * Turns a written-out date like "August 30, 2026" (or "Aug 30, 2026",
 * or a tab name in that same format) into "2026-08-30". Returns '' if
 * `text` holds nothing that looks like one.
 *
 * Deliberately parsed by hand rather than through `new Date(text)`:
 * this is pure text in, text out, with no time zone anywhere in the
 * middle to shift the answer onto the day before or after — the exact
 * class of bug that makes a bulletin for the 30th get filed as the
 * 29th. A three-letter abbreviation is enough to identify the month,
 * and an ordinal suffix ("30th") is tolerated.
 */
function parseBulletinDateText_(text) {
  if (!text) return '';

  var match = String(text).match(/([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})/);
  if (!match) return '';

  var months = ['january', 'february', 'march', 'april', 'may', 'june',
                'july', 'august', 'september', 'october', 'november', 'december'];
  var name = match[1].toLowerCase();
  var month = 0;
  for (var i = 0; i < months.length; i++) {
    if (months[i].indexOf(name) === 0) { // full name or any leading abbreviation
      month = i + 1;
      break;
    }
  }
  if (!month) return '';

  var day = Number(match[2]);
  if (!day || day > 31) return '';

  return match[3] + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2);
}


/**
 * Decodes the base64 `content` GitHub's Contents API returns for a file
 * (it arrives wrapped across lines, which base64Decode won't take) into
 * a UTF-8 string. Returns '' rather than throwing if it can't be read
 * for any reason — callers treat that as "date unknown" and fall back
 * to another source, since a copy filed under an imperfect name still
 * beats a publish that dies on the way there.
 */
function decodeGithubBase64_(content) {
  if (!content) return '';
  try {
    return Utilities.newBlob(
      Utilities.base64Decode(String(content).replace(/\s/g, ''))
    ).getDataAsString();
  } catch (err) {
    return '';
  }
}


function githubHeaders_(token) {
  return {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
}

/* ---------------------------------------------------------------------
 * DATA EXTRACTION
 * ------------------------------------------------------------------ */

