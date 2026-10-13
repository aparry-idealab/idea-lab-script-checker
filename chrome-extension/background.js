/**
 * Service worker: injects the rule engine + content script into the
 * active tab when the user clicks the toolbar icon (activeTab
 * permission — the default behaviour needs no special host
 * permissions and nothing runs automatically).
 *
 * Optionally, from inside the panel, the user can turn on
 * "Auto-open on this site" for a specific Canvas domain. That grants
 * a one-time, site-scoped host permission (requested via
 * chrome.permissions.request — Chrome shows exactly which site is
 * being granted, never "all websites") and registers a persistent
 * dynamic content script for that origin, so the panel re-injects
 * and refreshes itself automatically every time you navigate to a
 * new page on that site. The toolbar icon still just opens/closes
 * the panel on the current page either way.
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

var CONTENT_SCRIPT_FILES = ["lib/ruleEngine.js", "content/domExtractor.js", "content/panel.js"];
var CONTENT_SCRIPT_CSS = ["content/panel.css"];
var AUTO_MODE_STORAGE_KEY = "autoModeOrigins";
var DEFAULT_TITLE = "Run IDEA Lab Canvas Editorial Checker on this page";

function flashBadge(text, title) {
  chrome.action.setBadgeBackgroundColor({ color: "#c0392b" });
  chrome.action.setBadgeText({ text: text });
  if (title) chrome.action.setTitle({ title: title });
  setTimeout(function () {
    chrome.action.setBadgeText({ text: "" });
    chrome.action.setTitle({ title: DEFAULT_TITLE });
  }, 4000);
}

function registeredScriptIdFor(originPattern) {
  return "idea-lab-auto-" + originPattern;
}

async function getAutoModeOrigins() {
  var stored = await chrome.storage.local.get(AUTO_MODE_STORAGE_KEY);
  return stored[AUTO_MODE_STORAGE_KEY] || {};
}

async function setAutoModeOrigin(originPattern, enabled) {
  var origins = await getAutoModeOrigins();
  if (enabled) {
    origins[originPattern] = true;
  } else {
    delete origins[originPattern];
  }
  await chrome.storage.local.set({ [AUTO_MODE_STORAGE_KEY]: origins });
}

// Re-register dynamic content scripts for any origins the user
// previously enabled auto-mode for. chrome.scripting dynamic content
// scripts already persist across browser restarts on their own, but
// re-registering here is cheap, idempotent, and self-healing if the
// extension was updated/reloaded in a way that cleared them.
async function restoreAutoModeScripts() {
  var origins = await getAutoModeOrigins();
  var patterns = Object.keys(origins);
  if (patterns.length === 0) return;
  try {
    await chrome.scripting.registerContentScripts(
      patterns.map(function (originPattern) {
        return {
          id: registeredScriptIdFor(originPattern),
          matches: [originPattern],
          js: CONTENT_SCRIPT_FILES,
          css: CONTENT_SCRIPT_CSS,
          runAt: "document_idle",
        };
      })
    );
  } catch (err) {
    // Scripts with these IDs may already be registered (e.g. service
    // worker restarted without the extension restarting) — ignore.
  }
}
chrome.runtime.onStartup.addListener(restoreAutoModeScripts);
chrome.runtime.onInstalled.addListener(restoreAutoModeScripts);

chrome.runtime.onMessage.addListener(function (message, sender, sendResponse) {
  if (!message || !message.type) return false;

  if (message.type === "get-auto-mode-status") {
    getAutoModeOrigins().then(function (origins) {
      sendResponse({ enabled: !!origins[message.originPattern] });
    });
    return true; // keep the message channel open for the async response
  }

  if (message.type === "enable-auto-mode") {
    var originPattern = message.originPattern;
    chrome.permissions
      .request({ origins: [originPattern] })
      .then(function (granted) {
        if (!granted) {
          sendResponse({ ok: false, reason: "permission-denied" });
          return;
        }
        return chrome.scripting
          .registerContentScripts([
            {
              id: registeredScriptIdFor(originPattern),
              matches: [originPattern],
              js: CONTENT_SCRIPT_FILES,
              css: CONTENT_SCRIPT_CSS,
              runAt: "document_idle",
            },
          ])
          .then(function () {
            return setAutoModeOrigin(originPattern, true);
          })
          .then(function () {
            sendResponse({ ok: true });
          });
      })
      .catch(function (err) {
        sendResponse({ ok: false, reason: String(err) });
      });
    return true;
  }

  if (message.type === "disable-auto-mode") {
    var pattern = message.originPattern;
    chrome.scripting
      .unregisterContentScripts({ ids: [registeredScriptIdFor(pattern)] })
      .catch(function () {
        // Not registered — fine, continue to clean up storage/permission anyway.
      })
      .then(function () {
        return setAutoModeOrigin(pattern, false);
      })
      .then(function () {
        return chrome.permissions.remove({ origins: [pattern] }).catch(function () {});
      })
      .then(function () {
        sendResponse({ ok: true });
      });
    return true;
  }

  return false;
});

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
      files: CONTENT_SCRIPT_CSS,
    });
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: CONTENT_SCRIPT_FILES,
    });
  } catch (err) {
    console.error("IDEA Lab Canvas Editorial Checker: failed to inject into tab", err);
    flashBadge("!", "Couldn't run the checker on this page: " + (err && err.message ? err.message : err));
  }
});
