# HD Geo email filing add-in

An Outlook add-in that files project emails to the right HD Geo project site and searches filed emails. It's HD Geo's own replacement for paid tools such as Ideagen Mail Manager.

## What it does
- **File:** suggests the project for the open email. It checks the HD number in the subject, then the email thread, then the client contact, then the client's company. You confirm the project, pick the folder and file. Meeting invites are never filed.
- **Search:** type a plain sentence, e.g. "Sarah from Clientco about retaining wall levels last month". Fixed rules (no AI) pick out the sender, client, HD number, time period and keywords, show them as chips, and search filed emails across all projects. For open-ended questions, use Microsoft 365 Copilot. All staff have licences.

## Files
| File | Purpose |
|---|---|
| `docs/taskpane.html`, `taskpane.css`, `taskpane.js` | The panel |
| `docs/matching.js` | Project-matching and search rules (no Outlook/SharePoint calls) |
| `docs/sample-data.js` | Made-up projects and emails for the preview, to be replaced by real data |
| `manifest.xml` | Tells Outlook where the add-in lives and adds the **File to project** button |

## Preview (no setup needed)
Run a local web server on `docs` and open `taskpane.html` in a browser. Choose sample emails from the yellow bar.

## Status
- [x] Panel, matching rules and search, previewed with sample data
- [x] Filing works without Andrew (27/11/2026): the add-in sends the open email (`getAsFileAsync`, Mailbox 1.14) to the Power Automate flow "Email filing - File email" (HDGeoAutomations). The flow finds the site in the Project Register and saves `yyyyMMddHHmmss - subject.eml` to `/Shared Documents/Emails`, or to `<HD-n> - …/01 Emails` for dash jobs. The flow link is pasted once in the panel's **Settings** and stored in the user's Outlook settings, never in this repo.
- [ ] HDGeoAutomations on the 382 private HD teams (Andrew, at rollout)
- [ ] Real project list in the panel (names, clients) instead of HD numbers only
- [ ] **Andrew (later, for rollout):** Entra app registration, if we move from the flow link to signed-in Microsoft Graph calls
- [ ] "Project Email Register" list for thread matching and search
- [ ] Client contact emails from Scoro for sender matching
- [ ] "File this?" prompt on send
- [x] Copilot test (25/11/2026): Copilot reads .eml, .msg and PDF filed to SharePoint, so no PDF copy is needed (see `copilot-test/TEST-STEPS.md`)
- [x] Hosted on GitHub Pages (`https://yl-hdgeo.github.io/hdgeo-email-filing-addin/`)
- [ ] Test in Yiling's Outlook, pilot, deploy
