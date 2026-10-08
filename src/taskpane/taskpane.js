/* global Office, Word, ScriptCheckerRules */

/**
 * Task pane glue code: reads the open Word document (read-only), runs it
 * through the ScriptCheckerRules engine, and renders a categorised, tallied
 * list of flagged issues. Never writes to the document.
 */
(function () {
  "use strict";

  var runBtn, refreshBtn, statusText, tallyEl, resultsEl;

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Word) {
      return;
    }
    runBtn = document.getElementById("run-check-btn");
    refreshBtn = document.getElementById("refresh-btn");
    statusText = document.getElementById("status-text");
    tallyEl = document.getElementById("tally");
    resultsEl = document.getElementById("results");

    runBtn.addEventListener("click", runCheck);
    refreshBtn.addEventListener("click", runCheck);

    // Auto-run once on load, as requested (auto-run + a manual refresh button).
    runCheck();
  });

  function setStatus(text) {
    if (statusText) statusText.textContent = text;
  }

  /**
   * Reads all body paragraphs from the active Word document, including
   * Word's own list-item metadata and paragraph style name, which the rule
   * engine uses for list and heading detection. Read-only: no edits.
   */
  function readDocumentParagraphs() {
    return Word.run(function (context) {
      var paragraphs = context.document.body.paragraphs;
      paragraphs.load("text,style");
      return context.sync().then(function () {
        // listItemOrNullObject needs to be requested per-paragraph.
        paragraphs.items.forEach(function (p) {
          p.listItemOrNullObject.load("isNullObject");
        });
        return context.sync().then(function () {
          return paragraphs.items.map(function (p) {
            return {
              text: p.text || "",
              style: p.style || "",
              isListItem: !p.listItemOrNullObject.isNullObject,
            };
          });
        });
      });
    });
  }

  function runCheck() {
    setStatus("Checking…");
    runBtn.disabled = true;
    refreshBtn.disabled = true;

    readDocumentParagraphs()
      .then(function (paragraphs) {
        var result = ScriptCheckerRules.checkDocument(paragraphs);
        renderResults(result);
        setStatus(
          "Checked " +
            paragraphs.length +
            " paragraph" +
            (paragraphs.length === 1 ? "" : "s") +
            " (" +
            result.totalWords +
            " words) — " +
            result.totalIssues +
            " issue" +
            (result.totalIssues === 1 ? "" : "s") +
            " flagged."
        );
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Something went wrong reading the document: " + (err && err.message ? err.message : err));
      })
      .then(function () {
        runBtn.disabled = false;
        refreshBtn.disabled = false;
      });
  }

  function renderResults(result) {
    renderTally(result);
    renderCategories(result);
  }

  function renderTally(result) {
    tallyEl.innerHTML = "";
    result.categoryOrder.forEach(function (catId) {
      var cat = result.categories[catId];
      var count = result.countsByCategory[catId] || 0;
      var badge = document.createElement("div");
      badge.className = "tally-badge" + (count === 0 ? " zero" : "");
      badge.innerHTML = '<span class="count">' + count + "</span><span>" + escapeHtml(cat.label) + "</span>";
      badge.addEventListener("click", function () {
        var section = document.getElementById("category-" + catId);
        if (section) {
          section.classList.remove("collapsed");
          section.scrollIntoView({ behavior: "smooth", block: "start" });
        }
      });
      tallyEl.appendChild(badge);
    });
  }

  function renderCategories(result) {
    resultsEl.innerHTML = "";

    if (result.totalIssues === 0) {
      var empty = document.createElement("div");
      empty.className = "empty-state";
      empty.textContent = "No issues flagged. Remember: this is a quick checker, not a substitute for editorial or subject review.";
      resultsEl.appendChild(empty);
      return;
    }

    result.categoryOrder.forEach(function (catId) {
      var cat = result.categories[catId];
      var issues = result.issues.filter(function (i) {
        return i.category === catId;
      });

      var section = document.createElement("section");
      section.className = "category" + (issues.length === 0 ? " collapsed" : "");
      section.id = "category-" + catId;

      var header = document.createElement("button");
      header.className = "category-header";
      header.setAttribute("aria-expanded", issues.length > 0 ? "true" : "false");
      header.innerHTML =
        "<span>" +
        escapeHtml(cat.label) +
        ' <span class="category-count">(' +
        issues.length +
        ')</span><div class="meta">' +
        escapeHtml(cat.description) +
        "</div></span>" +
        '<span aria-hidden="true">▾</span>';
      header.addEventListener("click", function () {
        section.classList.toggle("collapsed");
      });

      var body = document.createElement("div");
      body.className = "category-body";

      if (issues.length === 0) {
        var none = document.createElement("div");
        none.className = "issue-message";
        none.textContent = "No issues flagged in this category.";
        body.appendChild(none);
      } else {
        issues.forEach(function (issue) {
          body.appendChild(renderIssue(issue));
        });
      }

      section.appendChild(header);
      section.appendChild(body);
      resultsEl.appendChild(section);
    });
  }

  function renderIssue(issue) {
    var el = document.createElement("div");
    el.className = "issue";

    var location =
      issue.paragraphIndex === null || issue.paragraphIndex === undefined
        ? "Document-wide"
        : "Paragraph " + issue.paragraphIndex;

    el.innerHTML =
      '<div class="issue-top-row">' +
      '<span class="severity-dot severity-' +
      issue.severity +
      '" title="' +
      issue.severity +
      '"></span>' +
      '<span class="issue-location">' +
      escapeHtml(location) +
      "</span>" +
      "</div>" +
      '<div class="issue-snippet">' +
      escapeHtml(issue.snippet) +
      "</div>" +
      '<div class="issue-message">' +
      escapeHtml(issue.message) +
      "</div>" +
      '<div class="issue-rule">Rule: ' +
      escapeHtml(issue.rule) +
      "</div>";

    return el;
  }

  function escapeHtml(str) {
    return String(str == null ? "" : str).replace(/[&<>"']/g, function (ch) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch];
    });
  }
})();
