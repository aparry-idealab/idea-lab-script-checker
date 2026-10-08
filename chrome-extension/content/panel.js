/**
 * IDEA Lab Script Checker — Canvas panel
 * ---------------------------------------------------------------
 * Injected on toolbar-icon click. Extracts the current Canvas page's
 * rich-text content, runs it through the (unchanged) ScriptCheckerRules
 * engine, and renders a categorised, tallied, collapsible results
 * panel — the same read-only "quick checker, not a substitute for
 * editorial/subject review" model as the Word add-in. Clicking a
 * flagged issue scrolls to and briefly highlights the corresponding
 * element on the page. Never edits page content.
 */
(function () {
  "use strict";

  var PANEL_ID = "idea-lab-script-checker-panel";

  // Toggle: if the panel is already open, close it instead of
  // re-injecting (the icon acts as an open/close switch).
  var existing = document.getElementById(PANEL_ID);
  if (existing) {
    existing.remove();
    return;
  }

  var extraction = window.IdeaLabCanvasDomExtractor && window.IdeaLabCanvasDomExtractor.extract();

  var panel = document.createElement("div");
  panel.id = PANEL_ID;

  if (!extraction || extraction.paragraphs.length === 0) {
    panel.innerHTML =
      '<div class="idea-lab-header">' +
      "<span>IDEA Lab Script Checker</span>" +
      '<button class="idea-lab-close" aria-label="Close">×</button>' +
      "</div>" +
      '<div class="idea-lab-empty">Couldn\u2019t find any checkable content on this page. This works on Canvas Page, Assignment, Discussion, and Syllabus <em>view</em> pages (not yet the Rich Content Editor while actively editing).</div>';
    document.body.appendChild(panel);
    panel.querySelector(".idea-lab-close").addEventListener("click", function () {
      panel.remove();
    });
    return;
  }

  var paragraphs = extraction.paragraphs;
  var result = ScriptCheckerRules.checkDocument(
    paragraphs.map(function (p) {
      return { text: p.text, isListItem: p.isListItem, style: p.style };
    })
  );

  render(result, paragraphs);

  function escapeHtml(str) {
    return String(str == null ? "" : str).replace(/[&<>"']/g, function (ch) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch];
    });
  }

  function render(result, paragraphs) {
    var html =
      '<div class="idea-lab-header">' +
      "<span>IDEA Lab Script Checker</span>" +
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
    wireUpEvents(result);
  }

  function wireUpEvents(result) {
    panel.querySelector(".idea-lab-close").addEventListener("click", function () {
      panel.remove();
    });

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

  var highlightTimer = null;
  function highlightParagraph(paragraphIndex, issueEl) {
    panel.querySelectorAll(".idea-lab-issue.active").forEach(function (n) {
      n.classList.remove("active");
    });
    if (issueEl) issueEl.classList.add("active");

    var para = paragraphs[paragraphIndex];
    if (!para || !para.el) return;

    para.el.scrollIntoView({ behavior: "smooth", block: "center" });
    para.el.classList.add("idea-lab-highlight-target");
    if (highlightTimer) clearTimeout(highlightTimer);
    highlightTimer = setTimeout(function () {
      para.el.classList.remove("idea-lab-highlight-target");
    }, 2500);
  }
})();
