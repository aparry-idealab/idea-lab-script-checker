# IDEA Lab Script Checker

A **Word (Office.js) task pane add-in** for the IDEA Lab media team. It runs a
**last-hurdle, read-only quick check** of video scripts received from module
leaders/learning designers, before the scripts go into production.

> ⚠️ **This is a quick checker, not a substitute for editorial or subject
> review.** It is aimed at media team members who are not subject or
> editorial specialists. It flags *possible* issues for a human to glance
> over — it never edits the document automatically.

## What it does

- Reads the currently open Word document (paragraph text, Word list items,
  and paragraph style names) — **read-only**, no automatic edits.
- Runs a set of deterministic regex/heuristic JavaScript rules (no AI calls,
  no network calls, no data leaves the machine).
- Displays a categorised, tallied list of flagged issues in the task pane:
  for every flag you get the **category**, the **paragraph location** and a
  **short quote**, a **plain-language explanation**, and the **specific rule
  reference**.
- **Click an issue to highlight it in the document.** Clicking (or pressing
  Enter/Space on) a flagged item selects the matching text in the open Word
  document so you can quickly see it in context. This is a **selection
  only** — it uses Office.js's read-only `Range.select()` API and never
  edits the document. If the exact wording can no longer be found (e.g. the
  script was edited after the last check), the whole paragraph is selected
  instead and the status line tells you so — just click **Refresh** and try
  again.
- An optional "fix" control per flagged item is a possible future stretch
  feature and is **not implemented** in this version — the add-in never
  writes to the document.

## Rule categories

1. **Obvious errors** — doubled words, repeated/stray punctuation (`??`,
   `..` that isn't a genuine `...` ellipsis, a stray space before a comma),
   unmatched quotation marks/brackets, and a small, clearly-labelled list of
   "possible typos" (not a full spellcheck).
2. **Read-aloud flow** — sentences over ~30 words (flagged with the word
   count), overall pacing vs. the ~125–150 words-per-minute / ~400–750 word
   guidance for a 3–5 minute video (informational), common AI-sounding
   filler phrases ("it's important to note", "delve into", "in conclusion",
   etc.), and an informational Flesch-Kincaid-style readability estimate.
3. **Bullet points in scripts** — any bulleted/numbered list in the script
   body, detected via Word's own list-paragraph metadata *and* via
   markdown-style bullet characters (`-`, `*`, `•`) or numbered patterns at
   the start of a line, since scripts should be read-aloud prose.
4. **House style & editorial compliance** — hyphenation (e.g.
   "cost-effective", "part-time"), en/em dash usage, forward-slash spacing,
   Latin abbreviations (e.g./i.e./etc.), number style (spell out 1–10 in
   prose, never start a sentence with a digit), percentage-form consistency,
   date formatting, British vs. American spelling, practice/practise,
   "vs.", ampersands, the Oxford comma (low-confidence note), title
   capitalisation, "Prof"/"Dr."/"Mr."/"Mrs."/"Syndicate group", and
   bullet/list formatting rules (stem colon, capitalisation, full-stop
   consistency) for any lists intentionally kept.
5. **Session/activity numbering** — every "session N" / "activity N.N"
   reference is surfaced in its own category so the media team can see at a
   glance what might break if the module is reordered later. This covers
   both numeral forms ("session 5") **and** spelled-out number words
   ("Session one", "session five") — numeral forms and incorrectly-cased
   word forms are flagged as style warnings (should be lowercase spelled-out
   words, e.g. "session five"), while correctly-formatted word references
   are still surfaced as a lower-severity informational note, since any
   specific session/activity reference is worth a second glance before
   production in case the module is reordered later.
6. **Inclusivity & diversity language** — a lookup table of terms to avoid
   (with suggested alternatives), plus a low-confidence heuristic for
   generic gendered pronouns following a non-specific subject (e.g. "a
   student... he...").

All rules are implemented as pure, framework-agnostic JavaScript functions in
[`src/rules/ruleEngine.js`](./src/rules/ruleEngine.js), independent of
Office.js, so they can be unit tested without Word running and reused
elsewhere.

## Project structure

```
idea-lab-script-checker/
├── manifest.xml                 # Office Add-in manifest (sideloading)
├── server.js                    # Minimal local HTTPS static server (dev only)
├── package.json
├── src/
│   ├── rules/
│   │   └── ruleEngine.js        # Pure rule-checking logic (no Office.js dependency)
│   ├── canvas/
│   │   └── htmlToParagraphs.js  # HTML -> paragraph converter for Canvas rich-text content
│   └── taskpane/
│       ├── taskpane.html
│       ├── taskpane.css
│       └── taskpane.js          # Office.js glue: reads doc, calls ruleEngine, renders UI
├── chrome-extension/            # Browser extension: run the same checks on Canvas LMS pages
│   ├── manifest.json
│   ├── background.js
│   ├── lib/ruleEngine.js        # auto-generated copy, see "Chrome extension" section below
│   ├── content/
│   │   ├── domExtractor.js
│   │   ├── panel.js
│   │   └── panel.css
│   └── icons/
├── scripts/
│   └── copy-rule-engine.js      # keeps chrome-extension/lib/ruleEngine.js in sync
├── tests/
│   ├── ruleEngine.test.js       # Vitest unit tests for the rule engine
│   └── htmlToParagraphs.test.js # Vitest unit tests for the Canvas HTML converter
└── assets/
    └── icon-32.png / icon-64.png / icon-80.png
```

## Running the tests

```bash
npm install
npm test
```

This runs the Vitest suite (`tests/ruleEngine.test.js` and
`tests/htmlToParagraphs.test.js`) covering realistic example scripts (clean
scripts, and scripts deliberately containing each class of flagged issue)
for every rule category above.

## Chrome extension: checking Canvas LMS course content

The same rule engine also ships as a small, dependency-free **Chrome
extension** so the media team can run the identical checks directly on
Canvas course content (Pages, Assignments, Discussion Topics, the
Syllabus) — no API token, no server, no network calls at all. It's the
lowest-friction way to get this tool working against Canvas: it just reads
whatever rich-text content is already rendered on the Canvas page you have
open in the browser.

**How it works:**

- Click the toolbar icon only when you want to run a check (`activeTab`
  permission — nothing runs automatically or in the background, and Chrome
  doesn't even show a host-permission warning on install).
- It finds the page's main rich-text container (Syllabus, Assignment
  description, Discussion message, or a Page's body), walks it into the
  same paragraph shape the rule engine expects, and runs the unchanged
  `ScriptCheckerRules.checkDocument()` function from `src/rules/ruleEngine.js`.
- Results render in a slide-in side panel: a tally per category, then a
  collapsible, categorised list of flagged issues — same model as the Word
  task pane.
- **Click a flagged issue to jump to it**: the page scrolls to and briefly
  highlights (a yellow outline, not an edit) the exact paragraph/list item
  on the Canvas page.
- It is **read-only** — it only reads `textContent` and toggles a temporary
  CSS class for highlighting. It never edits the page or saves anything
  back to Canvas.

**Known limitation:** it checks whatever is currently *rendered* on the
page (a Page/Assignment/Discussion/Syllabus "view", or the Syllabus). It
does not yet reach inside the Rich Content Editor while you're actively
editing (TinyMCE's edit iframe), and it checks one page at a time rather
than bulk-scanning an entire course — view/publish the page first, then run
the checker, as a last-hurdle pass before sharing with students.

**Installing it locally (unpacked, for now — not published to the Chrome
Web Store):**

1. `cd` into the repo and run `npm run build:extension` (this copies the
   latest `src/rules/ruleEngine.js` into `chrome-extension/lib/`; re-run it
   any time the rule engine changes).
2. In Chrome, go to `chrome://extensions`, turn on **Developer mode**
   (top-right toggle).
3. Click **Load unpacked** and select the `chrome-extension/` folder in
   this repo.
4. Open a Canvas Page, Assignment, Discussion Topic, or Syllabus page, then
   click the IDEA Lab Script Checker icon in the toolbar to run a check.
   Click it again to close the panel.

By default the extension works on any page (it only activates when you
click the icon), so there's nothing Canvas-domain-specific to configure —
if your institution's Canvas uses a custom domain rather than
`*.instructure.com`, it will still work, since `activeTab` grants access to
whatever tab is currently active regardless of domain.

## Sharing this with colleagues (recommended, no setup required for them)

The add-in is hosted on **GitHub Pages** at: 

**https://aparry-idealab.github.io/idea-lab-script-checker/**

This is kept in sync with the `main` branch — every push rebuilds it
automatically within a minute or two, so updates reach everyone
immediately with no reinstall needed.

To let a colleague use it, just send them **`manifest-share.xml`** (from
the repo root) — it points at the GitHub Pages URL above rather than
`localhost`, so they do **not** need Node.js, this repository, or a local
dev server/certificate at all. They:

1. Open Word (desktop or [Word Online](https://office.com)).
2. **Home tab → Add-ins → Upload My Add-in** (on Word Online) or
   **Home tab → Add-ins → More Add-ins → Developer Add-ins → Upload My
   Add-in** (on Word desktop/Mac — this UI varies a bit by Word version and
   platform; see "Sideloading in Word on a Mac" below for the full set of
   menu paths we found work).
3. Select the `manifest-share.xml` file you sent them.
4. The **"Check Script"** button appears — click it to open the task pane.

For wider/automatic rollout across the whole team (no manual upload step
per person at all), ask your Microsoft 365 administrator to deploy
`manifest-share.xml` via **Microsoft 365 admin center → Settings →
Integrated apps** (a.k.a. Centralized Deployment) — this requires tenant
admin access but then every targeted user gets the add-in automatically.

> `manifest.xml` (no suffix) is the **local development** manifest — it
> points at `https://localhost:3000` and is only useful on a machine that
> also has this repo checked out and `npm start` running (see below).
> `manifest-share.xml` is the one to actually hand out.

## Sideloading in Word on a Mac (for local development)

Office Add-ins must be served over **HTTPS** with a certificate Word trusts
— `file://` URLs are not supported for the task pane. The included
`server.js` uses `office-addin-dev-certs` to generate and trust a local
development certificate, so you don't need any external hosting. (This is
only needed if you're developing/testing changes locally — see "Sharing
this with colleagues" above for the no-setup option everyone else should
use.)

1. Install dependencies and trust the local dev certificate (one-time):

   ```bash
   npm install
   npx office-addin-dev-certs install
   ```

   macOS will prompt you to add the certificate to your keychain — accept
   this so Word trusts `https://localhost:3000`.

2. Start the local server:

   ```bash
   npm start
   ```

   This serves the project at `https://localhost:3000` (leave it running).

3. Sideload the manifest using the **shared folder** method (the most
   reliable approach on Mac):

   - Create a trusted catalog folder, e.g.:
     ```bash
     mkdir -p ~/Library/Containers/com.Microsoft.Word/Data/Documents/wef
     cp manifest.xml ~/Library/Containers/com.Microsoft.Word/Data/Documents/wef/
     ```
   - Open **Word** → **Preferences** → **Ribbon & Toolbar** → not here;
     instead go to the **Insert** tab → **Add-ins** → **My Add-ins** →
     the gear/"..." menu → **Upload My Add-in**, or, if your Word build
     already watches the shared folder automatically, the add-in should
     simply appear under **My Add-ins → Shared Folder**.
   - If prompted, trust the add-in.

4. Open any `.docx`, go to the **Insert** tab (or wherever the add-in
   placed its button — "Script Checker" group, "Check Script" button), and
   click it to open the task pane. It auto-runs on load; use **Refresh** to
   re-run after editing the script.

> If the shared-folder method isn't picking up the manifest, the
> alternative is Word's **Upload My Add-in** dialog (Insert → Add-ins → My
> Add-ins → "..." → Upload My Add-in) and point it directly at
> `manifest.xml` — this works the same way and avoids any folder-watching
> quirks on some Word builds.

### Troubleshooting sideloading (found while testing on Word for Mac)

- **The ribbon "Add-ins" button opens the Store, not your sideloaded
  add-in.** Look for a small dropdown arrow next to the Add-ins icon, or a
  **"Developer Add-ins"** tab inside the Add-ins dialog (separate from
  Store / Admin Managed / My Add-ins) — that's where shared-folder
  manifests show up.
- **Don't confuse this with the "Templates and Add-ins" dialog** (reached
  via the ribbon's **Developer** tab) — that's for legacy VBA/`.dotm`
  macro templates and will never show an Office.js add-in.
- **Task pane opens but shows a blank/broken page** ("might be
  temporarily down") when using the local dev server: this is usually the
  browser not yet trusting the self-signed dev certificate for that exact
  host/port — an iframe (which is what the task pane is) can't show a
  "proceed anyway" certificate warning. Fix: open
  `https://localhost:3000/src/taskpane/taskpane.html` directly in a normal
  browser tab first, accept any certificate warning there, then reopen the
  task pane in Word. (This is only relevant to the local-dev `manifest.xml`
  — `manifest-share.xml`, served over GitHub Pages with a real certificate,
  doesn't have this problem at all.)

## Design principles

- **Read-only by default.** The manifest requests `ReadDocument`
  permission only. No code path in this version writes to the Word
  document.
- **Fully offline / client-side.** All checks run in the task pane's own
  JavaScript. No external API calls, no telemetry, no data leaves the
  machine. (Office.js itself is loaded once from Microsoft's CDN, as is
  standard/required for every Word add-in to talk to the host application —
  this is the add-in platform framework, not a data-sending API call. If you
  need a fully air-gapped setup, you can download `office.js` once and
  reference it from `assets/` instead of the CDN URL in `taskpane.html`.)
- **Framework-agnostic rule engine.** `src/rules/ruleEngine.js` has no
  dependency on Office.js or the DOM — it's a pure module that takes an
  array of paragraph strings/objects and returns issues, so it is fully
  unit-testable and reusable.
- **Plainly worded for non-specialists.** Every flagged issue explains the
  category, the snippet, *why* it's flagged in plain language, and which
  rule it comes from — aimed at a media team that is not expected to be
  editorial or subject-matter experts.

## Extending the rule engine

Add a new check as a small pure function in `src/rules/ruleEngine.js` that
returns an array of issues built with the internal `makeIssue(...)` helper,
then call it from `checkDocument(...)`. Add a corresponding test case in
`tests/ruleEngine.test.js` with both a clean example and a deliberately
flagged example.
