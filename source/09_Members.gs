/**
 * WARD MEMBERS UPDATES
 * "Ward Bulletin > Update Ward Members" brings the Members tab in line
 * with the ward's current member list from LCR. Membership data sits
 * behind a Church account sign-in, so the script never fetches it
 * itself: whoever runs this exports the Member List from LCR (a PDF),
 * or copies it, and gives it to the dialog in MembersDialog.html. The
 * dialog reads a PDF in the browser and turns it back into rows of
 * text; from there a PDF, a paste, and a CSV file all go through the
 * same parser and the same review.
 *
 * It's a full sync, in two steps. previewMemberUpdate reads the list
 * and returns every change it would make (members to add, members
 * whose name, gender, or age changed, and members to remove because
 * they're no longer on the list) without touching the sheet. Only once
 * that has been reviewed does applyMemberUpdate rewrite the tab. The
 * review is what stops a partial paste, only half of LCR's list, from
 * quietly deleting everyone who didn't get copied.
 *
 * Afterwards the tab holds exactly the pasted list, in name order, in
 * the three columns it always had (Name as "Last, First Middle",
 * Gender, Age). Column D, the speaker dropdown's helper list, is
 * rebuilt from them.
 *
 * previewMemberUpdate and applyMemberUpdate are called from the dialog
 * through google.script.run, which means they're public. Until
 * September 2026 this project was also deployed as a web app anyone
 * could open, running as its owner, and any public function in a web
 * app's deployed version can be called from its page's browser
 * console: a stranger could have passed in a one-line list and read
 * the whole ward back as "removals", or applied it. That deployment is
 * gone, but both functions still require the one-time token that
 * updateMembersFromMenu puts into the dialog it opens (which only an
 * editor clicking the menu ever gets), so a web app added back later
 * can't reopen the hole.
 */

// Header cells recognized for each column, compared lowercase. Anything
// else in the pasted list (birth date, phone, address...) is ignored.
var MEMBER_NAME_HEADERS_ = ['name', 'preferred name', 'member name', 'full name'];
var MEMBER_GENDER_HEADERS_ = ['gender', 'sex'];
var MEMBER_AGE_HEADERS_ = ['age'];

var MEMBERS_DIALOG_TOKEN_KEY_ = 'members-dialog:'; // + token, in the script cache
var MEMBERS_DIALOG_TOKEN_TTL_ = 6 * 60 * 60;       // seconds; the cache's maximum


/**
 * Menu action: opens the PDF/paste dialog, with a fresh token for it
 * to send back — see requireMembersDialog_.
 */
function updateMembersFromMenu() {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put(MEMBERS_DIALOG_TOKEN_KEY_ + token, '1', MEMBERS_DIALOG_TOKEN_TTL_);

  var tmpl = HtmlService.createTemplateFromFile('MembersDialog');
  tmpl.token = token;
  var html = tmpl.evaluate()
    .setWidth(620)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Ward Members');
}


/**
 * Throws unless `token` is one updateMembersFromMenu handed out in the
 * last six hours. That's what would keep previewMemberUpdate and
 * applyMemberUpdate out of reach of a public web app, if one were ever
 * deployed again — see the note at the top of this file.
 */
function requireMembersDialog_(token) {
  if (!token || !CacheService.getScriptCache().get(MEMBERS_DIALOG_TOKEN_KEY_ + token)) {
    throw new Error('This dialog has expired. Close it and open Ward Bulletin > Update Ward Members again.');
  }
}


/**
 * Called from the dialog's "Review" button, with the dialog's token
 * (see requireMembersDialog_). Works out what applying
 * `text` would change and returns it for display, changing nothing:
 *   { adds: [String], updates: [String], removes: [String],
 *     unchanged: Number, ignored: [String], before: Number,
 *     after: Number, warning: String, fingerprint: String }
 * Throws (shown in the dialog) if `text` can't be read as a member list.
 */
function previewMemberUpdate(token, text) {
  requireMembersDialog_(token);
  var plan = planMemberUpdate_(text);
  return {
    adds: plan.diff.adds.map(function (m) { return describeMember_(m); }),
    updates: plan.diff.updates.map(function (u) { return u.to.name + ': ' + u.changes.join(', '); }),
    removes: plan.diff.removes.map(function (m) { return describeMember_(m); }),
    unchanged: plan.diff.unchanged,
    ignored: plan.ignored,
    before: plan.current.length,
    after: plan.members.length,
    warning: plan.warning,
    fingerprint: plan.fingerprint
  };
}


/**
 * Called from the dialog's "Apply" button, with the dialog's token, the
 * same `text` that was reviewed, and the fingerprint previewMemberUpdate
 * returned for it.
 * Rewrites the Members tab to match, rebuilds the speaker dropdown's
 * helper column, and re-applies this week's dropdowns.
 *
 * Refuses if the fingerprint no longer matches — someone edited the
 * Members tab between the review and the click, so what was reviewed
 * isn't what would be applied.
 */
function applyMemberUpdate(token, text, fingerprint) {
  requireMembersDialog_(token);
  var plan = planMemberUpdate_(text);
  if (plan.fingerprint !== fingerprint) {
    throw new Error('The Members tab changed after you reviewed it, so nothing was applied. ' +
      'Go back and review again.');
  }

  writeMembersTab_(plan.sheet, plan.members);
  CacheService.getScriptCache().remove(MEMBERS_DIALOG_TOKEN_KEY_ + token); // one apply per dialog

  var ss = plan.sheet.getParent();
  getMemberDropdownRange_(ss); // rebuild the helper column the speaker dropdowns read
  try {
    var thisWeek = ss.getSheetByName(bulletinTabName_(ss));
    if (thisWeek) refreshDropdowns_(thisWeek);
  } catch (err) {
    // Non-critical — the tab is updated either way, and "Refresh
    // Dropdowns" (or reopening the file) picks the new names up.
  }

  return {
    added: plan.diff.adds.length,
    updated: plan.diff.updates.length,
    removed: plan.diff.removes.length,
    before: plan.current.length,
    after: plan.members.length
  };
}


/**
 * The shared first half of preview and apply: parses `text`, reads the
 * Members tab, diffs the two, and fingerprints the result. Everything
 * here is read-only.
 */
function planMemberUpdate_(text) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET_NAME_);
  if (!sheet) throw new Error('There is no "' + MEMBERS_SHEET_NAME_ + '" tab to update.');

  var parsed = parseMemberList_(text);
  var members = sortMembers_(parsed.members);
  var current = readMembersTab_(sheet);
  var diff = diffMembers_(current, members);

  var warning = '';
  if (current.length && diff.removes.length > Math.max(10, current.length * 0.2)) {
    warning = 'This removes ' + diff.removes.length + ' of the ' + current.length +
      ' people on the Members tab. If only part of the list was copied out of LCR, ' +
      'go back and copy all of it.';
  }

  return {
    sheet: sheet,
    members: members,
    current: current,
    diff: diff,
    ignored: parsed.ignored,
    warning: warning,
    fingerprint: memberFingerprint_(current, members)
  };
}


/** Returns the Members tab's rows as [{name, gender, age}], skipping blank names. */
function readMembersTab_(sheet) {
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) { // row 1 is the header
    var name = normalizeMemberName_(rows[i][0]);
    if (!name) continue;
    out.push({
      name: name,
      gender: String(rows[i][1] == null ? '' : rows[i][1]).trim(),
      age: parseMemberAge_(rows[i][2])
    });
  }
  return out;
}


/**
 * Writes `members` into columns A-C from row 2 down and clears A-C on
 * any rows left over below them. Column D is left for
 * getMemberDropdownRange_ to rebuild.
 */
function writeMembersTab_(sheet, members) {
  var lastRow = sheet.getLastRow();
  if (members.length) {
    sheet.getRange(2, 1, members.length, 3).setValues(members.map(function (m) {
      return [m.name, m.gender, m.age];
    }));
  }
  var leftover = lastRow - 1 - members.length;
  if (leftover > 0) {
    sheet.getRange(2 + members.length, 1, leftover, 3).clearContent();
  }
}


/**
 * Parses a member list in any of the shapes it can arrive in:
 * - tab-separated: what the dialog makes of a PDF export, and what
 *   copying a table off a web page or out of a spreadsheet gives;
 * - comma-separated: a CSV file, with quoted "Last, First" names;
 * - space-separated lines: text copied straight out of a PDF viewer,
 *   where the column boundaries are gone (see parseSpacedMemberLines_).
 *
 * Finds the header row by its "Name" cell, within the first 20 rows,
 * so a report title or date line above it is skipped, then reads the
 * Name, Gender and Age columns wherever they are. If there's no header
 * at all, but the first row already looks like "Last, First" / M or F /
 * a number, the first three columns are taken as Name, Gender, Age.
 *
 * A row whose name has no comma isn't a member (a page footer, a count,
 * a section heading), so it goes into `ignored` for the review to show
 * instead of being added. A repeated header row — printed reports
 * repeat it on every page — is skipped silently.
 *
 * Returns { members: [{name, gender, age}], ignored: [String] }.
 */
function parseMemberList_(text) {
  text = String(text == null ? '' : text).replace(/^\uFEFF/, '');
  if (!text.trim()) throw new Error('Choose the member list PDF, or paste the list, first.');

  var delim = text.indexOf('\t') !== -1 ? '\t' : ',';
  var rows = parseDelimited_(text, delim);
  if (delim === ',' && !hasMemberHeaderOrShape_(rows)) {
    // No tabs, and it doesn't read as CSV either: text copied out of a
    // PDF viewer, one person per line with only spaces between columns.
    var spaced = parseSpacedMemberLines_(text);
    if (spaced.members.length) return spaced;
  }

  var headerIdx = -1;
  var cols = null;
  for (var i = 0; i < rows.length && i < 20; i++) {
    var found = findMemberColumns_(rows[i]);
    if (found.name !== -1) {
      headerIdx = i;
      cols = found;
      break;
    }
  }

  if (cols) {
    var missing = [];
    if (cols.gender === -1) missing.push('Gender');
    if (cols.age === -1) missing.push('Age');
    if (missing.length) {
      throw new Error('The list has a Name column but no ' + missing.join(' or ') +
        ' column. Include ' + (missing.length > 1 ? 'them' : 'it') + ' when copying out of LCR.');
    }
  } else {
    var firstRow = rows.filter(function (r) { return r.join('').trim() !== ''; })[0];
    if (!firstRow || !looksLikeMemberRow_(firstRow)) {
      throw new Error('Couldn\'t find the column headings (Name, Gender, Age) in the list. ' +
        'Copy the whole table out of LCR, including its heading row.');
    }
    cols = { name: 0, gender: 1, age: 2 };
  }

  var members = [];
  var ignored = [];
  for (var r = headerIdx + 1; r < rows.length; r++) {
    var row = rows[r];
    if (row.join('').trim() === '') continue;
    if (findMemberColumns_(row).name !== -1) continue; // repeated header

    var name = normalizeMemberName_(row[cols.name]);
    if (name.indexOf(',') === -1) {
      ignored.push(row.join(' ').replace(/\s+/g, ' ').trim());
      continue;
    }
    members.push({
      name: name,
      gender: normalizeMemberGender_(row[cols.gender]),
      age: parseMemberAge_(row[cols.age])
    });
  }

  if (!members.length) {
    throw new Error('No members found in the list — every row was either blank or had no ' +
      '"Last, First" name in its Name column.');
  }
  return { members: members, ignored: ignored };
}


/** True if `rows` has a Name header in its first 20 rows, or starts with a Name/Gender/Age-shaped row. */
function hasMemberHeaderOrShape_(rows) {
  for (var i = 0; i < rows.length && i < 20; i++) {
    if (findMemberColumns_(rows[i]).name !== -1) return true;
  }
  var firstRow = rows.filter(function (r) { return r.join('').trim() !== ''; })[0];
  return !!firstRow && looksLikeMemberRow_(firstRow);
}


/**
 * Reads text copied out of a PDF viewer, where a table row comes out as
 * one line with only spaces between its columns:
 *   Wilcox, Lenora Ruth   F   2   14 Mar 2024   ...
 * Each line is read as a "Last, First Middle" name, then M or F (or
 * Male/Female), then a whole-number age; anything after that is
 * ignored. The name is taken as short as it can be while still being
 * followed by a gender and an age, so a middle initial that happens to
 * be M or F ("Smith, John M   M   41") stays part of the name.
 *
 * This relies on Name, Gender, Age coming first and in that order, as
 * LCR's Member List prints them. Lines that don't fit — the report's
 * title, page numbers — go into `ignored` for the review to show; the
 * column-heading line is dropped silently.
 */
function parseSpacedMemberLines_(text) {
  var lineRe = /^\s*([^,]+,\s*\S.*?)\s+(M|F|Male|Female)\s+(\d{1,3})(?=\s|$)/i;
  var headingRe = /^\s*(preferred |member |full )?name\s+(gender|sex)\s+age\b/i;
  var members = [];
  var ignored = [];
  String(text).split(/\r\n|\r|\n/).forEach(function (line) {
    if (!line.trim() || headingRe.test(line)) return;
    var m = lineRe.exec(line);
    if (!m) {
      ignored.push(line.replace(/\s+/g, ' ').trim());
      return;
    }
    members.push({
      name: normalizeMemberName_(m[1]),
      gender: normalizeMemberGender_(m[2]),
      age: Number(m[3])
    });
  });
  return { members: members, ignored: ignored };
}


/** {name, gender, age}: the index of each recognized header cell in `row`, or -1. */
function findMemberColumns_(row) {
  var cols = { name: -1, gender: -1, age: -1 };
  for (var c = 0; c < row.length; c++) {
    var h = String(row[c]).trim().toLowerCase();
    if (cols.name === -1 && MEMBER_NAME_HEADERS_.indexOf(h) !== -1) cols.name = c;
    else if (cols.gender === -1 && MEMBER_GENDER_HEADERS_.indexOf(h) !== -1) cols.gender = c;
    else if (cols.age === -1 && MEMBER_AGE_HEADERS_.indexOf(h) !== -1) cols.age = c;
  }
  return cols;
}


/** True for a headerless first row shaped like "Last, First" / M or F / a number. */
function looksLikeMemberRow_(row) {
  return row.length >= 3 &&
    String(row[0]).indexOf(',') !== -1 &&
    /^[MF]$/.test(normalizeMemberGender_(row[1])) &&
    /^\s*\d{1,3}\s*$/.test(String(row[2]));
}


/**
 * Compares the Members tab (`current`) with the incoming list, matching
 * people by name ignoring case and spacing. Two people with the same
 * name are matched up in order, so a second "Smith, John" is an add
 * (or a removal), not an update to the first one. Returns:
 *   adds:      [member]           on the list, not on the tab
 *   removes:   [member]           on the tab, not on the list
 *   updates:   [{from, to, changes: [String]}]   name spelling, gender, or age differs
 *   unchanged: Number
 */
function diffMembers_(current, incoming) {
  var pool = {};
  current.forEach(function (m) {
    var key = memberKey_(m.name);
    (pool[key] = pool[key] || []).push(m);
  });

  var adds = [];
  var updates = [];
  var unchanged = 0;
  incoming.forEach(function (m) {
    var matches = pool[memberKey_(m.name)];
    var was = matches && matches.length ? matches.shift() : null;
    if (!was) {
      adds.push(m);
      return;
    }
    var changes = [];
    if (was.name !== m.name) changes.push('name "' + was.name + '" → "' + m.name + '"');
    if (was.gender !== m.gender) changes.push('gender ' + (was.gender || '(blank)') + ' → ' + (m.gender || '(blank)'));
    if (String(was.age) !== String(m.age)) changes.push('age ' + (was.age === '' ? '(blank)' : was.age) + ' → ' + (m.age === '' ? '(blank)' : m.age));
    if (changes.length) updates.push({ from: was, to: m, changes: changes });
    else unchanged++;
  });

  var removes = [];
  Object.keys(pool).forEach(function (key) {
    removes = removes.concat(pool[key]);
  });
  return { adds: adds, removes: sortMembers_(removes), updates: updates, unchanged: unchanged };
}


/** A copy of `members` sorted by name, ignoring case — the order the Members tab has always used. */
function sortMembers_(members) {
  return members.slice().sort(function (a, b) {
    var x = memberKey_(a.name), y = memberKey_(b.name);
    return x < y ? -1 : (x > y ? 1 : 0);
  });
}


/** "Walker, Lee (M, 89)" — for the review lists. */
function describeMember_(m) {
  var details = [m.gender, m.age].filter(function (v) { return v !== ''; }).join(', ');
  return details ? m.name + ' (' + details + ')' : m.name;
}


/**
 * A short hash of the tab as it stands and the list as it would be
 * written. Same inputs, same fingerprint; any edit to either in between
 * changes it — see applyMemberUpdate.
 */
function memberFingerprint_(current, members) {
  var s = JSON.stringify([current, members]);
  var h = 5381;
  for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36) + '.' + s.length.toString(36);
}


/** Trimmed, with runs of spaces collapsed to one. */
function normalizeMemberName_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
}


/** How two names are matched: normalized and lowercased. */
function memberKey_(name) {
  return normalizeMemberName_(name).toLowerCase();
}


/** "Male"/"m" -> "M", "Female"/"f" -> "F"; anything else is kept as written. */
function normalizeMemberGender_(v) {
  var g = String(v == null ? '' : v).trim();
  if (/^m(ale)?$/i.test(g)) return 'M';
  if (/^f(emale)?$/i.test(g)) return 'F';
  return g;
}


/** A whole-number age as a Number, or '' when the cell isn't one. */
function parseMemberAge_(v) {
  var m = /^\s*(\d{1,3})\s*$/.exec(String(v == null ? '' : v));
  return m ? Number(m[1]) : '';
}


/**
 * Splits delimited text into rows of cells, honouring double-quoted
 * cells (which may contain the delimiter, line breaks, and "" for a
 * literal quote) — enough for CSV exports and for text copied out of a
 * spreadsheet or web table. A quote only opens a quoted cell at the
 * very start of that cell; anywhere else it's an ordinary character,
 * so a nickname like Smith, John "Jack" survives unquoted.
 */
function parseDelimited_(text, delim) {
  var rows = [];
  var row = [];
  var cell = '';
  var inQuotes = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    if (inQuotes) {
      if (ch === '"') {
        if (text.charAt(i + 1) === '"') {
          cell += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cell += ch;
      }
    } else if (ch === '"' && cell === '') {
      inQuotes = true;
    } else if (ch === delim) {
      row.push(cell);
      cell = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text.charAt(i + 1) === '\n') i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += ch;
    }
  }
  if (cell !== '' || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}
