/**
 * Script for permission.html — a small extension-owned popup window
 * used only to request the site-scoped "Auto-open on this site" host
 * permission.
 *
 * This has to live here (not in background.js, and not be triggered
 * via a message from the content script) because
 * chrome.permissions.request() only succeeds when called directly,
 * synchronously, inside a real user gesture (e.g. a click handler in
 * an extension page). Routing the click through runtime.sendMessage
 * to the background service worker breaks that gesture context, which
 * is why the toggle used to silently do nothing.
 */
(function () {
  var params = new URLSearchParams(location.search);
  var originPattern = params.get("origin") || "";
  var hostname = originPattern.replace(/\/\*$/, "").replace(/^https?:\/\//, "");

  var explainer = document.getElementById("explainer");
  var grantBtn = document.getElementById("grant");
  var cancelBtn = document.getElementById("cancel");
  var statusEl = document.getElementById("status");

  if (hostname) {
    explainer.textContent =
      "This lets the IDEA Lab Canvas Editorial Checker panel stay open and " +
      "refresh itself automatically as you move between pages on " +
      hostname +
      ". Chrome will remember your choice until you turn it off again.";
  }

  cancelBtn.addEventListener("click", function () {
    window.close();
  });

  grantBtn.addEventListener("click", function () {
    if (!originPattern) {
      window.close();
      return;
    }
    grantBtn.disabled = true;
    cancelBtn.disabled = true;
    statusEl.textContent = "Requesting permission…";

    // This call is the whole point of this page existing: it runs
    // directly inside this click handler, so Chrome recognises the
    // user gesture and actually shows the permission prompt.
    chrome.permissions
      .request({ origins: [originPattern] })
      .then(function (granted) {
        if (!granted) {
          statusEl.textContent = "Permission was not granted.";
          setTimeout(function () {
            window.close();
          }, 900);
          return;
        }
        statusEl.textContent = "Setting up…";
        chrome.runtime.sendMessage(
          { type: "finish-enable-auto-mode", originPattern: originPattern },
          function (response) {
            statusEl.textContent =
              response && response.ok ? "Done — this window will close." : "Something went wrong.";
            setTimeout(function () {
              window.close();
            }, 600);
          }
        );
      })
      .catch(function (err) {
        statusEl.textContent = "Error: " + (err && err.message ? err.message : err);
        setTimeout(function () {
          window.close();
        }, 1500);
      });
  });
})();
