# Ward Bulletin Generator — source

The Google Apps Script behind the Edgemont 21st Ward sacrament meeting
bulletin. This folder is the readable copy of the code; the script itself
lives in the Apps Script project bound to the "Sacrament Meeting 2026"
spreadsheet (Extensions > Apps Script), and `../index.html` is whatever it
last published. Previous weeks land in `../archive/`, each named for the
Sunday that bulletin was for.

## Files

| File | What's in it |
| --- | --- |
| `01_Config.gs` | Ward name, GitHub destinations, web app URL, sheet/tab names. Full setup instructions are in the comment at the top. |
| `02_Menu.gs` | The "Ward Bulletin" menu, the on-open triggers, and the sidebar. |
| `03_Publish.gs` | Serving the web app, publishing to GitHub Pages, and archiving the previous bulletin. |
| `04_Preview.gs` | The Preview and Publish result dialogs. |
| `05_BulletinData.gs` | Reads the sheet and turns it into the data the template renders. |
| `06_Dropdowns.gs` | Type-ahead dropdowns for songs, speakers and leadership, plus the Date auto-fill. |
| `07_BulletinTabs.gs` | Creating each week's tab, and showing/hiding the reference tabs. |
| `Template.html` | The bulletin itself — layout and styling. |
| `Sidebar.html` | The Publish/Preview sidebar. |

## Load order

Apps Script joins every file in the project into one script before running
it, in the order the files are listed in the editor's sidebar — hence the
number prefixes. Functions are hoisted, so any file can call any other
file's functions regardless of order. The one thing that *is*
order-sensitive is a top-level `var` that reads another top-level `var`, so
all of those live together in `01_Config.gs`, which loads first.

## Updating

This folder is maintained by hand — nothing in the script writes to it,
and editing a file here changes nothing on its own. The Apps Script
project is what actually runs, so a change goes into the matching file in
the editor (the `.gs` extension is added by Apps Script, so the file there
is named `01_Config`, not `01_Config.gs`), gets saved, and the spreadsheet
reloaded. Commit the same change here to keep the two in step.
