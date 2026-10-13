/**
 * Service worker: injects the rule engine + content script into the
 * active tab only when the user explicitly clicks the toolbar icon
 * (activeTab permission — no broad host permissions, nothing runs
 * automatically, mirroring the "explicit Run Check" model of the
 * Word add-in version of this tool).
 *
 * Scope for the Canvas Editorial Checker: house style/editorial
 * compliance and inclusivity (both from the House Style and
 * Conventions Guide) plus general obvious-error proofreading — see
 * ALLOWED_CATEGORIES in content/panel.js. Read-aloud flow, bullet
 * points, and session/activity numbering are script-specific checks
 * and are intentionally excluded here.
 */
// Chrome never allows ANY extension to run on these kinds of pages
// (chrome://, the Web Store, other extensions' pages, etc). Checking
// up front lets us give the user a visible nudge instead of only a
// silent console error on the (normally hidden) service worker
// console.
var RESTRICTED_URL_PATTERN =
  /^(chrome|chrome-extension|edge|about|devtools|view-source):|^https:\/\/chrome\.google\.com\/webstore/i;

function flashBadge(text, title) {
  chrome.action.setBadgeBackgroundColor({ color: "#c0392b" });
  chrome.action.setBadgeText({ text: text });
  if (title) chrome.action.setTitle({ title: title });
  setTimeout(function () {
    chrome.action.setBadgeText({ text: "" });
    chrome.action.setTitle({ title: "Run IDEA Lab Canvas Editorial Checker on this page" });
  }, 4000);
}

chrome.action.onClicked.addListener(async (tab) => {
  if (!tab.id) return;

  if (!tab.url || RESTRICTED_URL_PATTERN.test(tab.url)) {
    flashBadge(
      "!",
      "Can't run here — open a Canvas Page, Assignment, Discussion, or Syllabus first."
    );
    return;
  }

  try {
    await chrome.scripting.insertCSS({
      target: { tabId: tab.id },
      files: ["content/panel.css"],
    });
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ["lib/ruleEngine.js", "content/domExtractor.js", "content/panel.js"],
    });
  } catch (err) {
    console.error("IDEA Lab Canvas Editorial Checker: failed to inject into tab", err);
    flashBadge("!", "Couldn't run the checker on this page: " + (err && err.message ? err.message : err));
  }
});
