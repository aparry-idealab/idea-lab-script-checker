/**
 * IDEA Lab Canvas Editorial Checker — panel
 * ---------------------------------------------------------------
 * Injected on toolbar-icon click. Extracts the current Canvas page's
 * rich-text content, runs it through the (unchanged) ScriptCheckerRules
 * engine, and renders a categorised, tallied, collapsible results
 * panel — the same read-only "quick checker, not a substitute for
 * editorial/subject review" model as the Word add-in. Clicking a
 * flagged issue scrolls to and highlights the corresponding element
 * on the page for as long as that issue stays selected. Never edits
 * page content.
 *
 * Scope: this tool is specifically for the media team's editorial
 * pass against the House Style and Conventions Guide, so only the
 * categories derived from that guide (plus general proofreading) are
 * shown here. Read-aloud flow, bullet-point-in-script, and
 * session/activity numbering are script-specific checks from the
 * other two style documents and are intentionally excluded.
 */
(function () {
  "use strict";

  var PANEL_ID = "idea-lab-canvas-editorial-checker-panel";
  var PANEL_WIDTH_PX = 360; // must match the panel width in panel.css
  var LAYOUT_SHIFT_ATTR = "data-idea-lab-orig-margin-right";
  var ALLOWED_CATEGORIES = ["obviousErrors", "houseStyle", "inclusivity"];
  // Match pattern for "Auto-open on this site": grants/revokes a
  // site-scoped (not "all websites") host permission and registers a
  // persistent content script so the panel re-injects itself on every
  // new page you navigate to on this origin, until turned off again.
  var ORIGIN_PATTERN = location.origin + "/*";

  /**
   * Shifts the page's own content left by the panel's width so text
   * isn't hidden underneath the fixed-position panel, instead of just
   * overlaying on top of it. Best-effort: elements the host page
   * itself fixes to the right edge of the viewport (independent of
   * document flow) may still sit under the panel — we can't rewrite
   * the host page's own CSS — but the main scrollable content area
   * reflows correctly in the common case.
   */
  function applyLayoutShift() {
    var html = document.documentElement;
    if (!html.hasAttribute(LAYOUT_SHIFT_ATTR)) {
      html.setAttribute(LAYOUT_SHIFT_ATTR, html.style.marginRight || "");
    }
    html.style.transition = "margin-right 0.2s ease";
    html.style.marginRight = PANEL_WIDTH_PX + "px";
  }

  function restoreLayoutShift() {
    var html = document.documentElement;
    if (html.hasAttribute(LAYOUT_SHIFT_ATTR)) {
      html.style.marginRight = html.getAttribute(LAYOUT_SHIFT_ATTR);
      html.removeAttribute(LAYOUT_SHIFT_ATTR);
    }
  }

  /** Removes any highlight left on the page, regardless of which script injection applied it. */
  function clearPageHighlight() {
    document.querySelectorAll(".idea-lab-highlight-target").forEach(function (el) {
      el.classList.remove("idea-lab-highlight-target");
    });
  }

  function closePanel(panelEl) {
    clearPageHighlight();
    restoreLayoutShift();
    panelEl.remove();
  }

  // Toggle: if the panel is already open, close it instead of
  // re-injecting (the icon acts as an open/close switch).
  var existing = document.getElementById(PANEL_ID);
  if (existing) {
    closePanel(existing);
    return;
  }

  /** Keeps only the categories relevant to the Canvas Editorial Checker's scope (see module comment above). */
  function restrictToAllowedCategories(result) {
    var categoryOrder = result.categoryOrder.filter(function (catId) {
      return ALLOWED_CATEGORIES.indexOf(catId) !== -1;
    });
    var categories = {};
    var countsByCategory = {};
    categoryOrder.forEach(function (catId) {
      categories[catId] = result.categories[catId];
      countsByCategory[catId] = result.countsByCategory[catId] || 0;
    });
    var issues = result.issues.filter(function (issue) {
      return ALLOWED_CATEGORIES.indexOf(issue.category) !== -1;
    });
    return {
      totalWords: result.totalWords,
      totalIssues: issues.length,
      categoryOrder: categoryOrder,
      categories: categories,
      countsByCategory: countsByCategory,
      issues: issues,
    };
  }

  var extraction = window.IdeaLabCanvasDomExtractor && window.IdeaLabCanvasDomExtractor.extract();

  var panel = document.createElement("div");
  panel.id = PANEL_ID;

  if (!extraction || extraction.paragraphs.length === 0) {
    panel.innerHTML =
      '<div class="idea-lab-header">' +
      "<span>IDEA Lab Canvas Editorial Checker</span>" +
      '<button class="idea-lab-automode-toggle" title="Keep this panel open and auto-refreshing as you browse this site">📌 Auto-open: off</button>' +
      '<button class="idea-lab-close" aria-label="Close">×</button>' +
      "</div>" +
      '<div class="idea-lab-empty">Couldn\u2019t find any checkable content on this page. This works on Canvas Page, Assignment, Discussion, and Syllabus <em>view</em> pages (not yet the Rich Content Editor while actively editing).</div>';
    document.body.appendChild(panel);
    applyLayoutShift();
    panel.querySelector(".idea-lab-close").addEventListener("click", function () {
      closePanel(panel);
    });
    wireUpAutoModeToggle();
    return;
  }

  var paragraphs = extraction.paragraphs;
  var rawResult = ScriptCheckerRules.checkDocument(
    paragraphs.map(function (p) {
      return { text: p.text, isListItem: p.isListItem, style: p.style };
    })
  );
  var result = restrictToAllowedCategories(rawResult);

  render(result, paragraphs);

  function escapeHtml(str) {
    return String(str == null ? "" : str).replace(/[&<>"']/g, function (ch) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch];
    });
  }

  function render(result, paragraphs) {
    var html =
      '<div class="idea-lab-header">' +
      "<span>IDEA Lab Canvas Editorial Checker</span>" +
      '<button class="idea-lab-automode-toggle" title="Keep this panel open and auto-refreshing as you browse this site">📌 Auto-open: off</button>' +
      '<button class="idea-lab-close" aria-label="Close">×</button>' +
      "</div>" +
      '<div class="idea-lab-banner">\u26A0\uFE0F Quick checker only — not a substitute for editorial or subject review.</div>' +
      '<div class="idea-lab-summary">' +
      result.totalIssues +
      " issue" +
      (result.totalIssues === 1 ? "" : "s") +
      " flagged across " +
      paragraphs.length +
      " paragraph" +
      (paragraphs.length === 1 ? "" : "s") +
      " (" +
      result.totalWords +
      " words)." +
      "</div>" +
      '<div class="idea-lab-tally">';

    result.categoryOrder.forEach(function (catId) {
      var cat = result.categories[catId];
      var count = result.countsByCategory[catId] || 0;
      html +=
        '<div class="idea-lab-tally-badge' +
        (count === 0 ? " zero" : "") +
        '" data-cat="' +
        catId +
        '"><span class="count">' +
        count +
        "</span><span>" +
        escapeHtml(cat.label) +
        "</span></div>";
    });
    html += "</div>" + '<div class="idea-lab-results">';

    if (result.totalIssues === 0) {
      html += '<div class="idea-lab-empty">No issues flagged.</div>';
    } else {
      result.categoryOrder.forEach(function (catId) {
        var cat = result.categories[catId];
        var issues = result.issues.filter(function (i) {
          return i.category === catId;
        });
        html +=
          '<section class="idea-lab-category' +
          (issues.length === 0 ? " collapsed" : "") +
          '" id="idea-lab-category-' +
          catId +
          '">' +
          '<button class="idea-lab-category-header">' +
          escapeHtml(cat.label) +
          " (" +
          issues.length +
          ")</button>" +
          '<div class="idea-lab-category-body">';
        issues.forEach(function (issue, i) {
          html +=
            '<div class="idea-lab-issue" data-para-index="' +
            issue.paragraphIndex +
            '" data-issue-index="' +
            i +
            '" data-cat="' +
            catId +
            '">' +
            '<div class="idea-lab-issue-top">' +
            '<span class="idea-lab-severity idea-lab-severity-' +
            issue.severity +
            '"></span>' +
            '<span class="idea-lab-issue-loc">' +
            (typeof issue.paragraphIndex === "number" ? "Paragraph " + issue.paragraphIndex : "Document-wide") +
            "</span>" +
            "</div>" +
            '<div class="idea-lab-issue-snippet">' +
            escapeHtml(issue.snippet) +
            "</div>" +
            '<div class="idea-lab-issue-message">' +
            escapeHtml(issue.message) +
            "</div>" +
            '<div class="idea-lab-issue-rule">Rule: ' +
            escapeHtml(issue.rule) +
            "</div>" +
            "</div>";
        });
        html += "</div></section>";
      });
    }

    html += "</div>";
    panel.innerHTML = html;
    document.body.appendChild(panel);
    applyLayoutShift();
    wireUpEvents(result);
  }

  function wireUpEvents(result) {
    panel.querySelector(".idea-lab-close").addEventListener("click", function () {
      closePanel(panel);
    });
    wireUpAutoModeToggle();

    panel.querySelectorAll(".idea-lab-tally-badge").forEach(function (badge) {
      badge.addEventListener("click", function () {
        var section = document.getElementById("idea-lab-category-" + badge.getAttribute("data-cat"));
        if (section) {
          section.classList.remove("collapsed");
          section.scrollIntoView({ behavior: "smooth", block: "start" });
        }
      });
    });

    panel.querySelectorAll(".idea-lab-category-header").forEach(function (header) {
      header.addEventListener("click", function () {
        header.parentElement.classList.toggle("collapsed");
      });
    });

    panel.querySelectorAll(".idea-lab-issue").forEach(function (issueEl) {
      issueEl.addEventListener("click", function () {
        var idx = parseInt(issueEl.getAttribute("data-para-index"), 10);
        highlightParagraph(idx, issueEl);
      });
    });
  }

  // The currently-highlighted page element, if any — kept highlighted
  // for as long as its issue stays selected (no auto-clearing timeout).
  var currentHighlightedEl = null;

  function deselectCurrentIssue() {
    if (currentHighlightedEl) {
      currentHighlightedEl.classList.remove("idea-lab-highlight-target");
      currentHighlightedEl = null;
    }
    panel.querySelectorAll(".idea-lab-issue.active").forEach(function (n) {
      n.classList.remove("active");
    });
  }

  function highlightParagraph(paragraphIndex, issueEl) {
    var wasAlreadySelected = issueEl.classList.contains("active");
    deselectCurrentIssue();

    // Clicking an already-selected issue again deselects it (and
    // clears its highlight) rather than re-selecting it.
    if (wasAlreadySelected) return;

    var para = paragraphs[paragraphIndex];
    if (!para || !para.el) return;

    issueEl.classList.add("active");
    para.el.scrollIntoView({ behavior: "smooth", block: "center" });
    para.el.classList.add("idea-lab-highlight-target");
    currentHighlightedEl = para.el;
  }
  function updateAutoModeButton(button, enabled) {
    button.setAttribute("aria-pressed", enabled ? "true" : "false");
    button.classList.toggle("on", enabled);
    button.textContent = enabled ? "📌 Auto-open: on" : "📌 Auto-open: off";
    button.title = enabled
      ? "Auto-open is on for " + location.hostname + " — click to turn off."
      : "Keep this panel open and auto-refreshing as you browse this site";
  }

  function wireUpAutoModeToggle() {
    var button = panel.querySelector(".idea-lab-automode-toggle");
    if (!button) return;

    chrome.runtime.sendMessage({ type: "get-auto-mode-status", originPattern: ORIGIN_PATTERN }, function (response) {
      updateAutoModeButton(button, !!(response && response.enabled));
    });

    // Turning auto-mode ON opens a small extension popup window
    // (permission.html) where the actual browser permission prompt is
    // triggered — Chrome requires that call to happen inside a direct
    // click in an extension page, which a message from this content
    // script can't provide. This listener picks up the outcome once
    // that window closes (granted, denied, or just closed).
    chrome.runtime.onMessage.addListener(function (message) {
      if (!message || message.type !== "auto-mode-result" || message.originPattern !== ORIGIN_PATTERN) {
        return;
      }
      button.disabled = false;
      updateAutoModeButton(button, !!message.enabled);
    });

    button.addEventListener("click", function () {
      var currentlyOn = button.classList.contains("on");
      if (currentlyOn) {
        button.disabled = true;
        chrome.runtime.sendMessage({ type: "disable-auto-mode", originPattern: ORIGIN_PATTERN }, function (response) {
          button.disabled = false;
          updateAutoModeButton(button, response && response.ok ? false : currentlyOn);
        });
        return;
      }

      // Don't disable the button here — the popup window may stay
      // open for a while, and the result arrives asynchronously via
      // the "auto-mode-result" message above.
      button.textContent = "📌 Check the popup window…";
      chrome.runtime.sendMessage({ type: "open-permission-window", originPattern: ORIGIN_PATTERN });
    });
  }

})();
