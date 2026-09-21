# Ward Bulletin — Edgemont 21st Ward

The weekly sacrament meeting bulletin, published to GitHub Pages.

| Path | What it is |
| --- | --- |
| `index.html` | The current week's bulletin — what GitHub Pages serves. Overwritten by each publish. |
| `archive/` | Previous weeks, one file per bulletin, named for the Sunday that bulletin was for (`2026-09-20.html`). Written automatically, just before `index.html` is overwritten. |
| `source/` | The Google Apps Script that generates all of it. See `source/README.md`. |

Nothing here is edited by hand except `source/`. The bulletin is produced
from a Google Sheet ("Sacrament Meeting 2026") by the Apps Script in
`source/`, and lands here when you use **Ward Bulletin > Publish** from the
sheet's own menu. The script commits `index.html` directly through the
GitHub API using a token stored in its Script Properties.

## Archive naming

An archived file is named for the Sunday its bulletin was written for, not
for the day it got archived — the date is read out of the bulletin being
replaced. Publishing next week's bulletin on a Wednesday therefore files
the outgoing one as its own Sunday, not as that Wednesday.
