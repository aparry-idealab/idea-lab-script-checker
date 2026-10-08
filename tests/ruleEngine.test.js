/**
 * Unit tests for the Script Checker rule engine.
 * Run with: npm test  (Vitest)
 */
import { describe, it, expect } from "vitest";
import ScriptCheckerRules from "../src/rules/ruleEngine.js";

var checkDocument = ScriptCheckerRules.checkDocument;

function categoryIssues(result, category) {
  return result.issues.filter(function (i) {
    return i.category === category;
  });
}

describe("Clean script", function () {
  it("flags nothing on a well-formed, house-style-compliant paragraph", function () {
    var paragraphs = [
      "Welcome to session five. In this short video we will look at how to plan an interview, and why preparation matters for a good recording.",
    ];
    var result = checkDocument(paragraphs);
    var obviousAndInclusivity = categoryIssues(result, "obviousErrors").concat(categoryIssues(result, "inclusivity"));
    expect(obviousAndInclusivity.length).toBe(0);
  });
});

describe("1) Obvious errors", function () {
  it("flags doubled words", function () {
    var result = checkDocument(["This is the the best approach."]);
    var hits = categoryIssues(result, "obviousErrors").filter(function (i) {
      return i.rule.indexOf("Doubled word") !== -1;
    });
    expect(hits.length).toBe(1);
    expect(hits[0].paragraphIndex).toBe(0);
  });

  it("flags repeated punctuation but not a genuine ellipsis", function () {
    var result = checkDocument(["Really?? That is surprising!! Let's wait... and see."]);
    var hits = categoryIssues(result, "obviousErrors").filter(function (i) {
      return i.rule.indexOf("punctuation") !== -1;
    });
    expect(hits.length).toBe(2); // "??" and "!!" only, not the "..."
  });

  it("flags space before a comma", function () {
    var result = checkDocument(["We will cover this topic , and then move on."]);
    var hits = categoryIssues(result, "obviousErrors").filter(function (i) {
      return i.message.indexOf("comma") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("flags unmatched straight quotes", function () {
    var result = checkDocument(['She said "welcome to the module and let us begin.']);
    var hits = categoryIssues(result, "obviousErrors").filter(function (i) {
      return i.rule.indexOf("quotation") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("flags unmatched brackets", function () {
    var result = checkDocument(["This refers to the reading list (see Appendix A for details."]);
    var hits = categoryIssues(result, "obviousErrors").filter(function (i) {
      return i.rule.indexOf("brackets") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("flags common possible typos", function () {
    var result = checkDocument(["Please recieve teh materials adn read them before the session."]);
    var hits = categoryIssues(result, "obviousErrors").filter(function (i) {
      return i.rule.indexOf("typo") !== -1;
    });
    expect(hits.length).toBe(3);
  });
});

describe("2) Read-aloud flow", function () {
  it("flags sentences over ~30 words with a word count", function () {
    var longSentence =
      "In this video we are going to explore a wide range of topics including planning, scripting, filming and editing because each of these stages matters enormously for the final quality of the piece.";
    var result = checkDocument([longSentence]);
    var hits = categoryIssues(result, "readAloudFlow").filter(function (i) {
      return i.rule.indexOf("Long sentence") !== -1;
    });
    expect(hits.length).toBe(1);
    expect(hits[0].message).toMatch(/\d+ words/);
  });

  it("flags AI filler phrases", function () {
    var result = checkDocument(["It's important to note that we should delve into this topic further."]);
    var hits = categoryIssues(result, "readAloudFlow").filter(function (i) {
      return i.rule.indexOf("filler") !== -1;
    });
    expect(hits.length).toBe(2); // "it's important to note" and "delve into"
  });

  it("flags scripts far outside the ~400-750 word pacing guidance", function () {
    var shortScript = ["This video is very short indeed."];
    var result = checkDocument(shortScript);
    var hits = categoryIssues(result, "readAloudFlow").filter(function (i) {
      return i.rule.indexOf("pacing") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("does not flag pacing for a script within the guidance range", function () {
    var words = [];
    for (var i = 0; i < 500; i++) words.push("word");
    var result = checkDocument([words.join(" ") + "."]);
    var hits = categoryIssues(result, "readAloudFlow").filter(function (i) {
      return i.rule.indexOf("pacing") !== -1;
    });
    expect(hits.length).toBe(0);
  });
});

describe("3) Bullet points in scripts", function () {
  it("flags Word list paragraphs", function () {
    var paragraphs = [
      { text: "Here is what we will cover:", isListItem: false },
      { text: "Planning your shoot", isListItem: true },
      { text: "Recording good audio", isListItem: true },
    ];
    var result = checkDocument(paragraphs);
    var hits = categoryIssues(result, "bulletPoints");
    expect(hits.length).toBe(2);
  });

  it("flags markdown-style bullet characters", function () {
    var paragraphs = ["- Plan your shoot", "* Record good audio", "• Edit the final cut"];
    var result = checkDocument(paragraphs);
    var hits = categoryIssues(result, "bulletPoints");
    expect(hits.length).toBe(3);
  });

  it("flags numbered list patterns at line start", function () {
    var result = checkDocument(["1. Introduce the topic", "2) Explain the method"]);
    var hits = categoryIssues(result, "bulletPoints");
    expect(hits.length).toBe(2);
  });
});

describe("4) House style / editorial compliance", function () {
  it("flags missing hyphenation", function () {
    var result = checkDocument(["This is a cost effective and part time approach to full time study."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Hyphenation") !== -1;
    });
    expect(hits.length).toBe(3);
  });

  it("flags plain hyphen used as a parenthetical aside", function () {
    var result = checkDocument(["The module leader - who wrote the script - will review it."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("dash") !== -1;
    });
    expect(hits.length).toBeGreaterThanOrEqual(1);
  });

  it("flags number ranges written with a hyphen", function () {
    var result = checkDocument(["The course ran from 1950-1980 in its original form."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.message.indexOf("en dash") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("flags forward slashes with no surrounding spaces", function () {
    var result = checkDocument(["You can submit a video and/or a written reflection."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("slash") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("flags Latin abbreviations missing full stops, and a trailing comma", function () {
    var result = checkDocument(["Bring your own equipment, eg a laptop, and other items etc, we can begin.", "This includes e.g., tripods."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Latin") !== -1;
    });
    expect(hits.length).toBe(3); // "eg", "etc", and "e.g.," comma
  });

  it("flags standalone digits 1-10 that should be spelled out, but not measurements", function () {
    var result = checkDocument(["We will record 3 takes and then review 5 kg of equipment weight."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Numbers") !== -1 && i.message.indexOf("spelled out") !== -1;
    });
    expect(hits.length).toBe(1); // only "3", not "5 kg"
  });

  it("flags a sentence starting with a numeral", function () {
    var result = checkDocument(["12 students attended the session."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.message.indexOf("start with a numeral") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("flags mixed use of percent forms across the document", function () {
    var result = checkDocument(["Attendance rose by 10 per cent this year.", "Around 25% of students responded."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Percentage") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("flags numeric slash/dash dates and dates missing an ordinal suffix", function () {
    var result = checkDocument(["The deadline is 01/01/2024.", "We will meet on 2 March 2024 to discuss."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Date") !== -1;
    });
    expect(hits.length).toBe(2);
  });

  it("flags American spellings", function () {
    var result = checkDocument(["We need to organize and recognize the color scheme whilst we analyze behavior at the center."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Spelling") !== -1;
    });
    expect(hits.length).toBe(7);
  });

  it("flags practice used as a verb and practise used as a noun", function () {
    var result = checkDocument(["Students need to practice their delivery.", "This is a good practise to adopt."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("practice vs practise") !== -1;
    });
    expect(hits.length).toBe(2);
  });

  it('flags "vs." with a full stop', function () {
    var result = checkDocument(["This compares online vs. in-person learning."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("vs") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it("flags an Oxford comma as a low-confidence note", function () {
    var result = checkDocument(["Bring a laptop, a notebook, and a pen to the session."]);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Oxford") !== -1;
    });
    expect(hits.length).toBe(1);
  });

  it('flags "Prof" and full-stopped Dr/Mr/Mrs titles, and "Syndicate group"', function () {
    var paragraphs = ["Prof Smith will introduce the session.", "Dr. Jones and Mrs. Patel will co-host.", "Join your Syndicate group afterwards."];
    var result = checkDocument(paragraphs);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Professor") !== -1 || i.rule.indexOf("no full stop") !== -1 || i.rule.indexOf("Syndicate") !== -1;
    });
    expect(hits.length).toBe(4); // Prof, Dr., Mrs., Syndicate group
  });

  it("flags list stem without a colon and inconsistent item formatting", function () {
    var paragraphs = [
      { text: "Here is what to bring", isListItem: false },
      { text: "laptop.", isListItem: true },
      { text: "Notebook", isListItem: true },
    ];
    var result = checkDocument(paragraphs);
    var hits = categoryIssues(result, "houseStyle").filter(function (i) {
      return i.rule.indexOf("Bullet/list formatting") !== -1;
    });
    // Missing colon on stem, lowercase start + stray full stop on first item.
    expect(hits.length).toBeGreaterThanOrEqual(2);
  });
});

describe("5) Session numbering", function () {
  it("flags numeral session references and specific activity references separately", function () {
    var result = checkDocument(["Please revisit session 5 before attempting activity 3.3."]);
    var hits = categoryIssues(result, "sessionNumbering");
    expect(hits.length).toBe(2);
    expect(hits.some(function (h) { return h.message.indexOf("spell out") !== -1; })).toBe(true);
    expect(hits.some(function (h) { return h.message.indexOf("avoid these entirely") !== -1; })).toBe(true);
  });

  it("does not flag correctly spelled-out session references", function () {
    var result = checkDocument(["Please revisit session five before we continue."]);
    var hits = categoryIssues(result, "sessionNumbering");
    expect(hits.length).toBe(0);
  });
});

describe("6) Inclusivity / diversity", function () {
  it("flags avoid-terms with a suggested alternative", function () {
    var result = checkDocument(["The policeman spoke to the handicapped man before the chairman arrived."]);
    var hits = categoryIssues(result, "inclusivity");
    expect(hits.length).toBe(3);
  });

  it("flags he/she and s/he suggesting they", function () {
    var result = checkDocument(["Each student should submit his or her work, so he/she must check the deadline."]);
    var hits = categoryIssues(result, "inclusivity").filter(function (i) {
      return i.message.indexOf("they") !== -1;
    });
    expect(hits.length).toBeGreaterThanOrEqual(1);
  });

  it("flags generic gendered pronouns following a non-specific subject", function () {
    var result = checkDocument(["If a student is unsure, he should ask his tutor for help."]);
    var hits = categoryIssues(result, "inclusivity").filter(function (i) {
      return i.rule.indexOf("Generic gendered pronoun") !== -1;
    });
    expect(hits.length).toBeGreaterThanOrEqual(1);
  });
});

describe("Aggregate result shape", function () {
  it("returns tallied counts per category and total word count", function () {
    var result = checkDocument(["The the quick video covers session 5 and the handicapped access policy."]);
    expect(result.countsByCategory.obviousErrors).toBeGreaterThanOrEqual(1);
    expect(result.countsByCategory.sessionNumbering).toBeGreaterThanOrEqual(1);
    expect(result.countsByCategory.inclusivity).toBeGreaterThanOrEqual(1);
    expect(result.totalIssues).toBe(result.issues.length);
    expect(typeof result.totalWords).toBe("number");
    expect(result.categoryOrder.length).toBe(6);
  });
});
