/**
 * Service worker: injects the rule engine + content script into the
 * active tab only when the user explicitly clicks the toolbar icon
 * (activeTab permission — no broad host permissions, nothing runs
 * automatically, mirroring the "explicit Run Check" model of the
 * Word add-in version of this tool).
 */
chrome.action.onClicked.addListener(async (tab) => {
  if (!tab.id) return;
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
    console.error("IDEA Lab Script Checker: failed to inject into tab", err);
  }
});
