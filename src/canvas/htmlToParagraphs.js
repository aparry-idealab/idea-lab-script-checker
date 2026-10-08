/**
 * HTML -> paragraph converter for Canvas LMS content.
 * ---------------------------------------------------------------
 * Canvas Pages, Assignments, the Syllabus, and Discussion Topics all
 * store their body as rich-text HTML (from the Canvas RCE), not plain
 * paragraph text like a Word document. This module turns that HTML
 * into the same { text, isListItem, style } paragraph shape the
 * rule engine (src/rules/ruleEngine.js) already expects, so the exact
 * same deterministic rule engine can be reused unchanged for Canvas
 * course content.
 *
 * Pure, framework-agnostic, dependency-free (no DOM, no network) so
 * it can run in plain Node and be unit tested in isolation.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.CanvasHtmlToParagraphs = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var NAMED_ENTITIES = {
    amp: "&",
    lt: "<",
    gt: ">",
    quot: '"',
    apos: "'",
    nbsp: " ",
    rsquo: "\u2019",
    lsquo: "\u2018",
    rdquo: "\u201d",
    ldquo: "\u201c",
    ndash: "\u2013",
    mdash: "\u2014",
    hellip: "\u2026",
  };

  function decodeEntities(str) {
    return String(str || "")
      .replace(/&#(\d+);/g, function (_, code) {
        return String.fromCharCode(parseInt(code, 10));
      })
      .replace(/&#x([0-9a-fA-F]+);/g, function (_, code) {
        return String.fromCharCode(parseInt(code, 16));
      })
      .replace(/&([a-zA-Z]+);/g, function (match, name) {
        return Object.prototype.hasOwnProperty.call(NAMED_ENTITIES, name) ? NAMED_ENTITIES[name] : match;
      });
  }

  /** Strip all remaining inline tags (b, strong, em, a, span, img, br, etc.) and collapse whitespace. */
  function stripInlineTags(html) {
    return decodeEntities(
      String(html || "")
        .replace(/<br\s*\/?>/gi, " ")
        .replace(/<[^>]+>/g, "")
    )
      .replace(/\s+/g, " ")
      .trim();
  }

  var HEADING_RE = /^h[1-6]$/i;

  /**
   * Converts a Canvas rich-text HTML body into an array of paragraph
   * objects compatible with ScriptCheckerRules.checkDocument:
   *   { text, isListItem, style }
   *
   * Block-level elements (p, li, headings, div, blockquote) each
   * become their own paragraph. Table cells are treated the same way
   * as plain paragraphs (a conservative choice — table content is
   * rarely "read-aloud prose", but still worth scanning for the same
   * style/inclusivity issues).
   */
  function htmlToParagraphs(html) {
    if (!html || typeof html !== "string") return [];

    var paragraphs = [];
    // Match each block-level element's opening tag, its tag name, and
    // its (possibly nested-tag-containing) inner content up to the
    // matching closing tag. This is a deliberately simple, non-DOM
    // regex-based splitter — sufficient for Canvas's typically simple
    // RCE-authored markup, not a general-purpose HTML parser.
    var blockRe = /<(p|li|h[1-6]|div|blockquote|td|th)\b[^>]*>([\s\S]*?)<\/\1>/gi;
    var match;
    var consumedAny = false;

    while ((match = blockRe.exec(html))) {
      consumedAny = true;
      var tag = match[1].toLowerCase();
      var inner = match[2];
      var isListItem = tag === "li";
      var style = HEADING_RE.test(tag) ? "Heading" : "Normal";

      // A <li> or <div>/<td> can itself contain nested block tags
      // (e.g. a list item containing a <p>, or a wrapper <div>
      // containing several <p>s). Recurse into those first; only
      // treat the remainder as this element's own paragraph text if
      // it has no nested block children of its own.
      var nested = htmlToParagraphs(inner);
      if (nested.length > 0 && /<(p|li|h[1-6]|div|blockquote|td|th)\b/i.test(inner)) {
        if (isListItem) {
          nested.forEach(function (p) {
            p.isListItem = true;
          });
        }
        paragraphs = paragraphs.concat(nested);
        continue;
      }

      var text = stripInlineTags(inner);
      if (text) {
        paragraphs.push({ text: text, isListItem: isListItem, style: style });
      }
    }

    if (!consumedAny) {
      // No recognised block tags at all (e.g. a bare text fragment, or
      // only inline markup) — treat the whole thing as one paragraph.
      var plain = stripInlineTags(html);
      if (plain) {
        paragraphs.push({ text: plain, isListItem: false, style: "Normal" });
      }
    }

    return paragraphs;
  }

  return { htmlToParagraphs: htmlToParagraphs };
});
