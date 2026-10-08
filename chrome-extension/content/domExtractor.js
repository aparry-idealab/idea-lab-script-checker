/**
 * IDEA Lab Script Checker — Canvas DOM extractor
 * ---------------------------------------------------------------
 * Finds the main rich-text content container for whatever Canvas
 * page is currently open (a Page, Assignment, Discussion Topic, or
 * the Syllabus), and walks its live DOM into the same paragraph
 * shape the rule engine expects: { text, isListItem, style }, plus
 * (extension-only) a reference to the actual DOM element each
 * paragraph came from, so a flagged issue can be scrolled to and
 * highlighted in place.
 *
 * Read-only: only ever reads text content and toggles a temporary
 * CSS class for highlighting. Never edits page/document content.
 */
(function (window) {
  "use strict";

  // Tried in priority order; the first that matches something with
  // actual text content wins. Covers the common Canvas "show" (view)
  // pages — not the Rich Content Editor while actively editing.
  var CONTENT_SELECTORS = [
    "#course_syllabus", // Syllabus page
    ".assignment-description.user_content", // Assignment show
    "#assignment_show .description.user_content",
    ".discussion-topic .message.user_content", // Discussion topic show
    ".discussion_topic .discussion-section .message",
    ".show-content.user_content", // generic Canvas Page show
    "#wiki_page_show .show-content",
    ".user_content", // fallback: any rendered rich-content block
  ];

  var LEAF_BLOCK_TAGS = ["P", "LI", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE", "TD", "TH"];

  function findContentContainer() {
    for (var i = 0; i < CONTENT_SELECTORS.length; i++) {
      var els = document.querySelectorAll(CONTENT_SELECTORS[i]);
      for (var j = 0; j < els.length; j++) {
        if (els[j].textContent && els[j].textContent.trim().length > 0) {
          return els[j];
        }
      }
    }
    return null;
  }

  function containsBlockChild(el) {
    for (var i = 0; i < el.children.length; i++) {
      if (LEAF_BLOCK_TAGS.indexOf(el.children[i].tagName) !== -1) return true;
    }
    return false;
  }

  function cleanText(el) {
    return (el.textContent || "").replace(/\s+/g, " ").trim();
  }

  /**
   * Recursively walks `root`, treating each "leaf" block-level element
   * (one with no further block-level children of its own) as one
   * paragraph. Mirrors the splitting rules of src/canvas/htmlToParagraphs.js
   * but operates on the live DOM so each paragraph keeps a real element
   * reference for highlight-on-click.
   */
  function walk(root, inList) {
    var paragraphs = [];
    var children = root.children ? Array.prototype.slice.call(root.children) : [];

    if (children.length === 0) {
      return paragraphs;
    }

    children.forEach(function (child) {
      var isBlock = LEAF_BLOCK_TAGS.indexOf(child.tagName) !== -1;
      var isListItem = child.tagName === "LI" || inList;

      if (isBlock && !containsBlockChild(child)) {
        var text = cleanText(child);
        if (text) {
          paragraphs.push({
            text: text,
            isListItem: isListItem,
            style: /^H[1-6]$/.test(child.tagName) ? "Heading" : "Normal",
            el: child,
          });
        }
        return;
      }

      // Not a leaf block itself (e.g. a <div>, <ul>, <ol>, <blockquote>
      // wrapping further block children, or a non-block wrapper) — recurse.
      var childIsList = child.tagName === "UL" || child.tagName === "OL";
      paragraphs = paragraphs.concat(walk(child, inList || childIsList));
    });

    return paragraphs;
  }

  /**
   * @returns {{ container: Element, paragraphs: Array<{text,isListItem,style,el}> } | null}
   */
  function extract() {
    var container = findContentContainer();
    if (!container) return null;
    var paragraphs = walk(container, false);
    if (paragraphs.length === 0) {
      // No recognised block structure at all — fall back to the whole
      // container's text as a single paragraph so a flat/plain page
      // still gets checked.
      var text = cleanText(container);
      if (text) {
        paragraphs = [{ text: text, isListItem: false, style: "Normal", el: container }];
      }
    }
    return { container: container, paragraphs: paragraphs };
  }

  window.IdeaLabCanvasDomExtractor = { extract: extract };
})(window);
