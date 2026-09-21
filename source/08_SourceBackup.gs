/**
 * SOURCE BACKUP (this project's own code -> GitHub)
 * Commits every file in this Apps Script project to GITHUB_SOURCE_DIR in
 * the same repo the bulletin publishes to, so the code is versioned
 * alongside what it produces instead of living only in the editor.
 *
 * Unlike everything else here, this needs the script to be able to read
 * ITSELF, which Apps Script only allows through its own REST API — hence
 * the one-time setup (turning that API on, plus a scope in the manifest)
 * described under SOURCE BACKUP in 01_Config.gs. Nothing else in this
 * project depends on it: if the setup is missing, only this one menu item
 * fails, and it says exactly what to fix.
 */

/**
 * Menu action: commits the current contents of every file in this
 * project to GITHUB_SOURCE_DIR. Files that already match what's in the
 * repo are skipped rather than re-committed, so running it twice in a
 * row does nothing the second time and the repo's history stays honest
 * about when the code actually changed.
 */
function backupSourceToGithubFromMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = backupSourceToGithub_();
    ui.alert(
      'Source backed up',
      result.written.length
        ? (result.written.length + ' of ' + result.total + ' file(s) updated in ' +
           GITHUB_SOURCE_DIR + '/:\n\n' + result.written.join('\n'))
        : ('All ' + result.total + ' file(s) already match what\'s in ' +
           GITHUB_SOURCE_DIR + '/ — nothing needed committing.'),
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Source backup failed', err.message || String(err), ui.ButtonSet.OK);
  }
}


/**
 * Reads this project's files and writes each changed one to
 * GITHUB_SOURCE_DIR/<name>.<ext>. Returns {written: [paths], total: n}.
 *
 * Only ever adds or updates: a file you delete from the Apps Script
 * editor is left behind in the repo rather than being deleted from it,
 * since quietly removing files from someone's repo is a worse failure
 * than leaving a stale one to be cleaned up by hand.
 */
function backupSourceToGithub_() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    throw new Error('No GitHub token set yet. Use "Ward Bulletin > Set GitHub Token…" first.');
  }

  var files = fetchProjectSource_();
  var written = [];
  for (var i = 0; i < files.length; i++) {
    var path = GITHUB_SOURCE_DIR + '/' + sourceFileName_(files[i]);
    if (putGithubFileIfChanged_(path, files[i].source, token)) written.push(path);
  }

  return { written: written, total: files.length };
}


/**
 * This project's own files, via the Apps Script REST API:
 * https://developers.google.com/apps-script/api/reference/rest/v1/projects/getContent
 *
 * ScriptApp.getOAuthToken() hands over the credentials the script is
 * already running with, so there's no second login involved — but that
 * token only carries the script.projects.readonly scope if the manifest
 * asks for it, which is half of the one-time setup. The other half is
 * the API being switched on for the Google account. Both failure modes
 * come back as 403/404, so that case explains both rather than making
 * you guess which one it was.
 */
function fetchProjectSource_() {
  var url = 'https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + '/content';
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code === 403 || code === 404) {
    throw new Error(
      'Couldn\'t read this project\'s own files (HTTP ' + code + '). Two things have to ' +
      'be in place, and it\'s usually the first:\n\n' +
      '1. Open script.google.com/home/usersettings and switch ON "Google Apps Script API".\n' +
      '2. appsscript.json needs the script.projects.readonly scope — see SOURCE BACKUP in ' +
      '01_Config.gs for the exact lines. After adding it, run this once more and approve ' +
      'the new permission prompt.\n\n' +
      'GitHub and publishing are unaffected either way.\n\nDetails: ' + resp.getContentText()
    );
  }
  if (code !== 200) {
    throw new Error('Apps Script API error ' + code + ': ' + resp.getContentText());
  }

  var files = (JSON.parse(resp.getContentText()) || {}).files || [];
  if (!files.length) {
    throw new Error('The Apps Script API returned no files for this project.');
  }
  return files;
}


/**
 * The filename a project file should have in the repo. Apps Script
 * stores names without extensions and reports the kind separately, so
 * "02_Menu" + SERVER_JS becomes "02_Menu.gs" — matching what the editor
 * shows and what you'd paste back in.
 */
function sourceFileName_(file) {
  var extensions = { SERVER_JS: '.gs', HTML: '.html', JSON: '.json' };
  return file.name + (extensions[file.type] || '');
}


/**
 * Writes `text` to `path` in the repo, but only if it differs from
 * what's already there. Returns true if it committed, false if the file
 * was already identical. Same Contents API dance as publishToGithub_:
 * GET for the current sha (an update needs it; a brand-new file must
 * omit it), then PUT.
 */
function putGithubFileIfChanged_(path, text, token) {
  var apiUrl = 'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO +
    '/contents/' + path;

  var sha = null;
  var getResp = UrlFetchApp.fetch(apiUrl + '?ref=' + GITHUB_BRANCH, {
    method: 'get',
    headers: githubHeaders_(token),
    muteHttpExceptions: true
  });
  if (getResp.getResponseCode() === 200) {
    var existing = JSON.parse(getResp.getContentText());
    sha = existing.sha;
    if (decodeGithubBase64_(existing.content) === text) return false; // already identical
  }

  var payload = {
    message: 'Back up ' + path + ' from the Apps Script editor',
    content: Utilities.base64Encode(text, Utilities.Charset.UTF_8),
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
    throw new Error('GitHub API error ' + code + ' writing ' + path + ': ' + putResp.getContentText());
  }
  return true;
}
