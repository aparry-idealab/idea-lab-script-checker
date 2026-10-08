import { describe, it, expect } from "vitest";
import CanvasHtmlToParagraphs from "../src/canvas/htmlToParagraphs.js";

var htmlToParagraphs = CanvasHtmlToParagraphs.htmlToParagraphs;

describe("Canvas HTML -> paragraphs", function () {
  it("splits simple <p> tags into separate paragraphs", function () {
    var html = "<p>First paragraph.</p><p>Second paragraph.</p>";
    var result = htmlToParagraphs(html);
    expect(result.length).toBe(2);
    expect(result[0].text).toBe("First paragraph.");
    expect(result[1].text).toBe("Second paragraph.");
    expect(result[0].isListItem).toBe(false);
  });

  it("strips inline formatting tags and decodes entities", function () {
    var html = "<p>This is <strong>bold</strong> &amp; <em>italic</em> text &ndash; nice.</p>";
    var result = htmlToParagraphs(html);
    expect(result.length).toBe(1);
    expect(result[0].text).toBe("This is bold & italic text \u2013 nice.");
  });

  it("marks <li> items as isListItem and flattens nested lists", function () {
    var html = "<ul><li>First item</li><li>Second item</li></ul>";
    var result = htmlToParagraphs(html);
    expect(result.length).toBe(2);
    expect(result[0].isListItem).toBe(true);
    expect(result[1].isListItem).toBe(true);
    expect(result[0].text).toBe("First item");
  });

  it("marks headings with a Heading style", function () {
    var html = "<h2>Section title</h2><p>Body text.</p>";
    var result = htmlToParagraphs(html);
    expect(result[0].style).toBe("Heading");
    expect(result[1].style).toBe("Normal");
  });

  it("treats <br> as a space rather than a paragraph break", function () {
    var html = "<p>Line one<br>Line two</p>";
    var result = htmlToParagraphs(html);
    expect(result.length).toBe(1);
    expect(result[0].text).toBe("Line one Line two");
  });

  it("recurses into a wrapping <div> containing multiple <p> tags", function () {
    var html = "<div><p>Inside first.</p><p>Inside second.</p></div>";
    var result = htmlToParagraphs(html);
    expect(result.length).toBe(2);
    expect(result[0].text).toBe("Inside first.");
  });

  it("falls back to one paragraph for bare inline-only HTML", function () {
    var html = "<strong>Just bold text, no block tags.</strong>";
    var result = htmlToParagraphs(html);
    expect(result.length).toBe(1);
    expect(result[0].text).toBe("Just bold text, no block tags.");
  });

  it("returns an empty array for empty/null/non-string input", function () {
    expect(htmlToParagraphs("")).toEqual([]);
    expect(htmlToParagraphs(null)).toEqual([]);
    expect(htmlToParagraphs(undefined)).toEqual([]);
  });

  it("skips empty paragraphs produced by empty tags", function () {
    var html = "<p>Real content.</p><p>&nbsp;</p><p></p>";
    var result = htmlToParagraphs(html);
    expect(result.length).toBe(1);
    expect(result[0].text).toBe("Real content.");
  });
});
