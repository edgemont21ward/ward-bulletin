# Ward Bulletin Generator — source

The Google Apps Script behind the Edgemont 21st Ward sacrament meeting
bulletin. This folder *is* the script: committing here pushes it into the
Apps Script project bound to the "Sacrament Meeting 2026" spreadsheet
(Extensions > Apps Script). `../index.html` is whatever it last published,
and previous weeks land in `../archive/`, each named for the Sunday that
bulletin was for.

## Files

| File | What's in it |
| --- | --- |
| `01_Config.gs` | Ward name, GitHub destinations, sheet/tab names. Full setup instructions are in the comment at the top. |
| `02_Menu.gs` | The "Ward Bulletin" menu, the on-open triggers, and the sidebar. |
| `03_Publish.gs` | Publishing to GitHub Pages, and archiving the previous bulletin. |
| `04_Preview.gs` | The Preview and Publish result dialogs. |
| `05_BulletinData.gs` | Reads the sheet and turns it into the data the template renders. |
| `06_Dropdowns.gs` | Type-ahead dropdowns for songs, speakers and leadership, plus the Date auto-fill. |
| `07_BulletinTabs.gs` | Creating each week's tab, adding speaker, song and testimony rows to its program, showing/hiding the reference tabs, and archiving past weeks to a new spreadsheet. |
| `08_Songs.gs` | "Update Songs": adds newly released Hymns for Home and Church hymns to the Songs tab, from the Church's website. |
| `09_Members.gs` | "Update Ward Members": syncs the Members tab to LCR's Member List (its PDF export, or the list pasted in), after a review of every change. |
| `10_Reports.gs` | The History tab each publish writes (who spoke and prayed), and the Reports > Speakers / Prayers dialogs that read it. |
| `Template.html` | The bulletin itself — layout and styling. |
| `Sidebar.html` | The Publish/Preview sidebar. |
| `Report.html` | The Speakers / Prayers report: "Who's due" and "Week by week". |
| `MembersDialog.html` | The Update Ward Members dialog: choose the PDF (read in the browser with pdf.js) or paste, review, apply. |
| `appsscript.json` | The project manifest — timezone, V8 runtime, and error logging. Apps Script hides this by default; see "Editing in the browser" below. |

This README is the only file here that isn't part of the script —
`../.claspignore` keeps it out of the push.

## Load order

Apps Script joins every file in the project into one script before running
it, in the order the files are listed in the editor's sidebar — hence the
number prefixes, which `../.clasp.json` repeats in `filePushOrder` so a
push reproduces that order. Functions are hoisted, so any file can call any
other file's functions regardless of order. The one thing that *is*
order-sensitive is a top-level `var` that reads another top-level `var`, so
all of those live together in `01_Config.gs`, which loads first.

## Updating

Edit the file here, commit, and push to `main`. The **Push to Apps Script**
workflow (`../.github/workflows/push-to-apps-script.yml`) runs `clasp push
--force`, which replaces the editor's contents with this folder's. Reload
the spreadsheet afterwards to pick up menu or sidebar changes.

A push doesn't touch Script Properties. The GitHub token that
`03_Publish.gs` publishes with lives there, set by hand, and survives
every push.

## Archive service

"Archive Past Weeks…" can be run by any editor, but the archive it makes
has to belong to the account that owns the spreadsheet, and Google always
makes whoever creates a file its owner. So the archiving is done by a web
app deployment of this project that runs as the account that deployed it:
the owner. The menu item sends it the tab names; it does the work.

It's locked down: it serves no pages (there's no `doGet`), and it refuses
any request that doesn't carry the secret kept in the script's Script
Properties (`ARCHIVE_SERVICE_SECRET`, made on first use), which only the
spreadsheet's editors can read.

- **It must be deployed from the owner's account.** If it runs as anyone
  else, it refuses to archive, because the archive wouldn't belong to the
  owner. The deployment settings (run as the deployer, reachable without
  signing in) come from the `webapp` block in `appsscript.json`.
- **Its URL is `ARCHIVE_SERVICE_URL` in `01_Config.gs`.** The Push to Apps
  Script workflow reads the deployment ID from that line and points the
  deployment at a new version after every push, since a deployment runs a
  frozen copy of the code.
- **Each of those refreshes uses up a version.** Google caps a project at
  around 200. If that's ever reached, delete old versions in the Apps
  Script editor (Project history).
- **To rotate the secret,** delete `ARCHIVE_SERVICE_SECRET` in Project
  Settings > Script Properties; the next archive run makes a new one.

### Editing in the browser

You can still edit in the Apps Script editor — for a quick experiment, or
to read the execution log — but the next push to `main` overwrites it
without warning, because of `--force`. Copy anything worth keeping back
into this folder and commit it.

`appsscript.json` is hidden in the editor until you turn it on: Project
Settings > "Show 'appsscript.json' manifest file in editor". The workflow
fails early if that file is missing here, since clasp can't push without
it.

## Automatic deploys

The workflow authenticates as a Google account using clasp credentials
stored in a repository secret. To set it up, or to repair it after the
credentials stop working:

1. **Enable the Apps Script API** for the Google account that owns the
   script, at <https://script.google.com/home/usersettings>. Pushes fail
   with a "User has not enabled the Apps Script API" error until you do.
2. **Log in locally**, with the same clasp major version the workflow
   installs:

   ```
   npm install -g @google/clasp@3
   clasp login
   ```

   That opens a browser, and writes the resulting credentials to
   `~/.clasprc.json`.
3. **Copy that file into a secret.** In the repo: Settings > Secrets and
   variables > Actions > New repository secret, named exactly
   `CLASPRC_JSON`, with the entire contents of `~/.clasprc.json` as the
   value — the whole JSON object, not just the token inside it. The
   workflow writes it back out to `~/.clasprc.json` on the runner and
   fails with a pointer to this section if the secret is empty or missing.
4. **Check `../.clasp.json`** points at the right project. `scriptId` is
   the ID in the Apps Script editor's URL (Project Settings > IDs also
   shows it), and `rootDir` is `source`.

The secret holds a refresh token, so it keeps working indefinitely without
being rotated — but it does stop working if the Google password changes,
the account's access is revoked, or the token goes ~6 months unused.
When that happens the push step fails on an auth error; redo steps 2 and 3
to replace the secret.

Treat the secret as a password to the whole Google account: anyone who can
read it can push arbitrary code into the script, which runs as that
account. That's also why the workflow triggers only on pushes to `main`
and never on `pull_request` — a public repo accepts PRs from anyone, and a
workflow running on those would hand a fork's code the credential.

### Running it by hand

The workflow only fires for changes under `source/`, `.clasp.json`,
`.claspignore`, or the workflow file itself. To push without changing any
of those — after repairing the secret, say — use the Actions tab > Push to
Apps Script > Run workflow, or:

```
gh workflow run push-to-apps-script.yml
```

A successful run logs the file count, e.g. `Pushed 15 files`.
