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
   numeral reference is surfaced in its own category so the media team can
   see at a glance what might break if the module is reordered later.
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
│   └── taskpane/
│       ├── taskpane.html
│       ├── taskpane.css
│       └── taskpane.js          # Office.js glue: reads doc, calls ruleEngine, renders UI
├── tests/
│   └── ruleEngine.test.js       # Vitest unit tests for the rule engine
└── assets/
    └── icon-32.png / icon-64.png / icon-80.png
```

## A note on how this project was delivered

This project was authored in a sandboxed session whose write access was
restricted to flat files in a scratch folder (no subdirectories, no
`node`/`npm` execution were permitted there). The code above is complete and
was carefully hand-traced against the test suite, but could not be executed
in that session. A `reconstruct.sh` script is included alongside these flat
files — **run it locally** (where you have normal shell access) to lay out
the directory tree shown above, decode the placeholder icons, and then run
`npm install` / `npm test` yourself to verify everything passes:

```bash
cd /path/to/the/folder/containing/these/flat/files
chmod +x reconstruct.sh
./reconstruct.sh
cd idea-lab-script-checker
npm install
npm test
```

If you are instead looking at an already-assembled copy of this project
(i.e. you can already see the `src/`, `tests/` folders above), you can skip
straight to `npm install && npm test` below.

## Running the tests

```bash
npm install
npm test
```

This runs the Vitest suite in `tests/ruleEngine.test.js` against
`src/rules/ruleEngine.js`, covering realistic example scripts (clean scripts,
and scripts deliberately containing each class of flagged issue) for every
rule category above.

## Sideloading in Word on a Mac (for local testing)

Office Add-ins must be served over **HTTPS** with a certificate Word trusts
— `file://` URLs are not supported for the task pane. The included
`server.js` uses `office-addin-dev-certs` to generate and trust a local
development certificate, so you don't need any external hosting.

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
