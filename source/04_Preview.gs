/**
 * PREVIEW & RESULT MODALS
 * The modal dialogs used by both "Preview" (renders the sheet straight
 * into an iframe via srcdoc) and "Publish" (points an iframe at the
 * live GitHub Pages URL) — see showInlinePreviewModal_ and
 * showResultModal_ for why they can't share one implementation.
 */

/**
 * Shared CSS for both result-style modals below (showResultModal_ and
 * showInlinePreviewModal_): a plain action bar across the top — the
 * secondary action buttons (Refresh, an optional primary Publish
 * button, an "Open in a new tab" link) right-aligned — above an iframe
 * filling the rest of the dialog body. There's no title text or close
 * button of our own in that bar: the dialog's actual title ("Bulletin
 * Preview" / "Bulletin Publish") is just passed straight through to
 * showModalDialog and shown in Apps Script's own native title bar
 * (which also has its own × to close), so this bar only needs to hold
 * the actions. html/body are explicitly sized to 100% with zero
 * margin — without that, the dialog's default body margin and auto
 * height keep the iframe from actually filling the space
 * setWidth/setHeight reserve for it. A few rules here (button.primary)
 * only apply to the preview modal's extra Publish button — harmless
 * no-ops in the publish modal, which has no such button.
 */
var MODAL_CSS_ =
  'html, body { height:100%; margin:0; padding:0; }' +
  'body { display:flex; flex-direction:column; font-family:Arial,sans-serif; overflow:hidden; }' +
  '.bar { flex:0 0 auto; display:flex; justify-content:flex-end; align-items:center; ' +
  'gap:16px; padding:8px 12px; border-bottom:1px solid #eee; box-sizing:border-box; }' +
  '.bar a, .bar button { font-size:12px; color:#364fc7; text-decoration:none; ' +
  'background:none; border:none; cursor:pointer; padding:0; font-family:inherit; }' +
  '.bar a:hover, .bar button:hover:not(:disabled) { text-decoration:underline; }' +
  '.bar button:disabled { opacity:0.5; cursor:default; }' +
  '.bar button.primary { background:#364fc7; color:#fff; padding:4px 12px; ' +
  'border-radius:4px; font-weight:600; }' +
  '.bar button.primary:hover:not(:disabled) { text-decoration:none; opacity:0.9; }' +
  'iframe { flex:1 1 auto; width:100%; height:100%; border:0; display:block; }';


/** A link out of either modal, opening in a new tab: "Open in a new tab ↗" unless `label` says otherwise. */
function openInNewTabLink_(url, label) {
  return '<a href="' + escapeHtml_(url) + '" target="_blank">' + escapeHtml_(label || 'Open in a new tab') + ' &#8599;</a>';
}


/**
 * Opens `url` in a modal dialog: an iframe pointed straight at it, plus
 * a Refresh button and the "Open in a new tab" link (for the rare site
 * that refuses to be framed) instead of leaving the spreadsheet for a
 * separate tab. Used for the Publish result, since GitHub Pages lets
 * itself be framed. Preview can't work this way: what it shows hasn't
 * been published to any URL yet — see showInlinePreviewModal_ below.
 * Refresh reloads the iframe by changing its `src` (a real network
 * request, unlike the preview modal's server round-trip), so its
 * "Refreshing…" / "✓ Refreshed" feedback is driven by the iframe's own
 * load event rather than a google.script.run callback.
 */
function showResultModal_(url, title) {
  var safeUrl = escapeHtml_(url);
  var jsUrl = JSON.stringify(url);
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' + MODAL_CSS_ + '</style></head><body>' +
    '<div class="bar">' +
    '<button type="button" id="refreshBtn" onclick="reloadFrame()">&#8635; Refresh</button>' +
    openInNewTabLink_(url) +
    '</div>' +
    '<iframe id="resultFrame" src="' + safeUrl + '"></iframe>' +
    '<script>' +
    'var REFRESH_LABEL = "↻ Refresh";' +
    'function reloadFrame() {' +
    '  var btn = document.getElementById("refreshBtn");' +
    '  var frame = document.getElementById("resultFrame");' +
    '  btn.disabled = true;' +
    '  btn.textContent = "Refreshing…";' +
    '  function onLoaded() {' +
    '    frame.removeEventListener("load", onLoaded);' +
    '    btn.textContent = "✓ Refreshed";' +
    '    setTimeout(function () {' +
    '      btn.disabled = false;' +
    '      btn.textContent = REFRESH_LABEL;' +
    '    }, 1200);' +
    '  }' +
    '  frame.addEventListener("load", onLoaded);' +
    '  var base = ' + jsUrl + ';' +
    '  var sep = base.indexOf("?") === -1 ? "?" : "&";' +
    '  frame.src = base + sep + "_r=" + Date.now();' +
    '}' +
    '</script>' +
    '</body></html>'
  ).setWidth(780).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}


/**
 * Sidebar-callable: opens a preview of the active sheet tab in a modal.
 * The tab hasn't been published anywhere yet, so there's no URL to
 * point the iframe at: the bulletin is rendered here and fed to the
 * iframe through `srcdoc`, which also means it always shows the sheet
 * exactly as it stands. The bar's link opens the bulletin that is
 * currently published on GitHub Pages, for comparison.
 */
function previewInModal() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var html = renderBulletinHtml_(sheet);
  showInlinePreviewModal_(html, 'Bulletin Preview — ' + sheet.getName(), getPagesUrl_(), 'Published bulletin');
}


/**
 * Modal/refresh-callable: re-renders the active sheet tab's bulletin
 * HTML. Used both by previewInModal() above and by the Refresh button
 * inside the preview modal (see showInlinePreviewModal_) — clicking
 * Refresh calls this again via google.script.run rather than reloading
 * any URL, so it always reflects whatever is in the sheet right now.
 */
function renderPreviewHtml() {
  return renderBulletinHtml_(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
}


/**
 * Opens `html` directly inside a modal's iframe via `srcdoc` — no
 * network request to any URL — with a Refresh button that calls
 * renderPreviewHtml() again and swaps the iframe's content in place, a
 * Publish button that calls the same publishBulletinForSidebar() the
 * sidebar's Publish button uses (which, on success, replaces this
 * dialog with the "Bulletin Publish" modal — so there's nothing left
 * for THIS dialog's success handler to do; a failure, like no GitHub
 * token being set, leaves this dialog open and shows the error here
 * instead), and (if `linkUrl` is given) a link to it, labelled
 * `linkLabel`, that opens in a new tab. The initial content and every refresh both
 * go through the iframe's `srcdoc` *property* (not an HTML attribute),
 * so there's no HTML-escaping to worry about — JSON.stringify is enough
 * to embed it safely as a JS string literal.
 */
function showInlinePreviewModal_(html, title, linkUrl, linkLabel) {
  var jsHtml = JSON.stringify(html);
  var openInNewTab = linkUrl ? openInNewTabLink_(linkUrl, linkLabel) : '';
  var dialogHtml = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' + MODAL_CSS_ + '</style></head><body>' +
    '<div class="bar">' +
    '<button type="button" id="refreshBtn" onclick="reloadFrame()">&#8635; Refresh</button>' +
    '<button type="button" id="publishBtn" class="primary" onclick="doPublish()">📤 Publish</button>' +
    openInNewTab +
    '</div>' +
    '<iframe id="resultFrame"></iframe>' +
    '<script>' +
    'document.getElementById("resultFrame").srcdoc = ' + jsHtml + ';' +
    'var REFRESH_LABEL = "↻ Refresh";' +
    'var PUBLISH_LABEL = "📤 Publish";' +
    // Rewrites Google's raw "Authorization is required to perform that
    // action" into something actionable. This button click goes through
    // google.script.run, and Apps Script's authorization popup can never
    // appear from a google.script.run call — only a direct "Ward
    // Bulletin" menu click can trigger it — so a newly shared editor
    // who's never used one hits this exact error with no way to fix it
    // from the message alone. Any other error passes through unchanged.
    'function friendlyError_(err) {' +
    '  var msg = (err && err.message) || String(err);' +
    '  if (/authorization is required/i.test(msg)) {' +
    '    return \'This account hasn\\\'t authorized this tool yet. ' +
           'Close this and use any "Ward Bulletin" menu item once ' +
           '(Publish, Preview, or Show Toolbar) to grant access, ' +
           'then try again.\';' +
    '  }' +
    '  return msg;' +
    '}' +
    'function reloadFrame() {' +
    '  var btn = document.getElementById("refreshBtn");' +
    '  btn.disabled = true;' +
    '  btn.textContent = "Refreshing…";' +
    '  google.script.run' +
    '    .withSuccessHandler(function (newHtml) {' +
    '      document.getElementById("resultFrame").srcdoc = newHtml;' +
    '      btn.textContent = "✓ Refreshed";' +
    '      setTimeout(function () {' +
    '        btn.disabled = false;' +
    '        btn.textContent = REFRESH_LABEL;' +
    '      }, 1200);' +
    '    })' +
    '    .withFailureHandler(function (err) {' +
    '      btn.disabled = false;' +
    '      btn.textContent = REFRESH_LABEL;' +
    '      alert(friendlyError_(err));' +
    '    })' +
    '    .renderPreviewHtml();' +
    '}' +
    'function doPublish() {' +
    '  var btn = document.getElementById("publishBtn");' +
    '  btn.disabled = true;' +
    '  btn.textContent = "Publishing…";' +
    '  google.script.run' +
    '    .withSuccessHandler(function () {' +
    '      /* success already replaced this dialog with the Published modal */' +
    '    })' +
    '    .withFailureHandler(function (err) {' +
    '      btn.disabled = false;' +
    '      btn.textContent = PUBLISH_LABEL;' +
    '      alert(friendlyError_(err));' +
    '    })' +
    '    .publishBulletinForSidebar();' +
    '}' +
    '</script>' +
    '</body></html>'
  ).setWidth(780).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(dialogHtml, title);
}

