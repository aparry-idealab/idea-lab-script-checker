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
    var isLocatable = typeof issue.paragraphIndex === "number";
    el.className = "issue" + (isLocatable ? " issue-clickable" : "");

    var location = isLocatable ? "Paragraph " + issue.paragraphIndex : "Document-wide";

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
      "</div>" +
      (isLocatable
        ? '<div class="issue-locate-hint">📍 Click to highlight this in the document</div>'
        : "");

    if (isLocatable) {
      el.setAttribute("role", "button");
      el.setAttribute("tabindex", "0");
      el.setAttribute("aria-label", "Highlight this issue in the document (paragraph " + issue.paragraphIndex + ")");
      el.addEventListener("click", function () {
        highlightIssueInDocument(issue, el);
      });
      el.addEventListener("keydown", function (evt) {
        if (evt.key === "Enter" || evt.key === " ") {
          evt.preventDefault();
          highlightIssueInDocument(issue, el);
        }
      });
    }

    return el;
  }

  /**
   * Strips the leading/trailing ellipsis markers added by the rule engine's
   * snippet truncation (see makeSnippet in ruleEngine.js) so the remaining
   * text can be used as a literal search string against the live document.
   * Read-only: only ever used with Range.select(), never to edit text.
   */
  function extractSearchTextFromSnippet(snippet) {
    if (!snippet) return "";
    var cleaned = snippet.replace(/^…\s*/, "").replace(/\s*…$/, "").trim();
    // Word's search API caps search strings at 255 characters.
    return cleaned.slice(0, 250);
  }

  /**
   * Selects (highlights) the paragraph/text a flagged issue refers to in
   * the open document. Purely a selection — never modifies the document.
   * Falls back to selecting the whole paragraph if the exact snippet text
   * can no longer be found (e.g. the document was edited since the last
   * check), and reports that back in the status line rather than failing
   * silently.
   */
  function highlightIssueInDocument(issue, el) {
    document.querySelectorAll(".issue.active").forEach(function (n) {
      n.classList.remove("active");
    });
    if (el) el.classList.add("active");

    Word.run(function (context) {
      var paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        var idx = issue.paragraphIndex;
        if (idx < 0 || idx >= paragraphs.items.length) {
          throw new Error("paragraph-not-found");
        }
        var paragraph = paragraphs.items[idx];
        var searchText = extractSearchTextFromSnippet(issue.snippet);

        if (searchText) {
          var range = paragraph.getRange();
          var results = range.search(searchText, {
            matchCase: false,
            matchWholeWord: false,
            ignoreSpace: true,
          });
          results.load("items");
          return context.sync().then(function () {
            if (results.items.length > 0) {
              results.items[0].select();
            } else {
              paragraph.select();
            }
            return context.sync().then(function () {
              return "exact";
            });
          });
        }

        paragraph.select();
        return context.sync().then(function () {
          return "paragraph";
        });
      });
    })
      .then(function (precision) {
        setStatus(
          precision === "exact"
            ? "Highlighted the flagged text in paragraph " + issue.paragraphIndex + "."
            : "Highlighted paragraph " + issue.paragraphIndex + " (exact text wasn't found - the document may have changed since the last check)."
        );
      })
      .catch(function (err) {
        console.error(err);
        setStatus(
          "Couldn't locate paragraph " +
            issue.paragraphIndex +
            " in the document — it may have changed since the last check. Try Refresh and click the issue again."
        );
      });
  }

  function escapeHtml(str) {
    return String(str == null ? "" : str).replace(/[&<>"']/g, function (ch) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch];
    });
  }
})();
