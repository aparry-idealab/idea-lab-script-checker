/* AUTO-GENERATED — do not edit directly.
 * Copied from src/rules/ruleEngine.js by scripts/copy-rule-engine.js.
 * Run `npm run build:extension` after changing the source file. */
/**
 * IDEA Lab Script Checker — Rule Engine
 * ---------------------------------------------------------------
 * A framework-agnostic, pure-JS module that scans video script text
 * (an array of paragraph strings, or paragraph objects with extra
 * Word metadata) and returns a categorised, tallied list of flagged
 * issues. No network calls, no AI, no document mutation — every
 * function here is a pure read -> analyse -> report function.
 *
 * Works both as:
 *   - a CommonJS/ESM module for Node-based tests (Vitest), and
 *   - a plain <script> global (window.ScriptCheckerRules) inside the
 *     Word task pane, with no build step required.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.ScriptCheckerRules = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---------------------------------------------------------------
  // Categories
  // ---------------------------------------------------------------
  var CATEGORIES = {
    obviousErrors: {
      id: "obviousErrors",
      label: "Obvious errors",
      description: "Typos, doubled words, stray punctuation, unmatched quotes/brackets.",
    },
    readAloudFlow: {
      id: "readAloudFlow",
      label: "Read-aloud flow",
      description: "Sentence length, overall pacing, AI-sounding filler phrases, jargon/readability.",
    },
    bulletPoints: {
      id: "bulletPoints",
      label: "Bullet points in script",
      description: "Scripts should be prose meant to be read aloud, not bulleted/numbered lists.",
    },
    houseStyle: {
      id: "houseStyle",
      label: "House style & editorial compliance",
      description: "Hyphenation, dashes, numbers, dates, abbreviations, spelling and formatting conventions.",
    },
    sessionNumbering: {
      id: "sessionNumbering",
      label: "Session / activity numbering",
      description: "References to specific session or activity numbers that may break if the module is reordered.",
    },
    inclusivity: {
      id: "inclusivity",
      label: "Inclusivity & diversity language",
      description: "Terms flagged by the house style inclusive-language guidance, with suggested alternatives.",
    },
  };

  var CATEGORY_ORDER = [
    "obviousErrors",
    "readAloudFlow",
    "bulletPoints",
    "houseStyle",
    "sessionNumbering",
    "inclusivity",
  ];

  // ---------------------------------------------------------------
  // Small generic helpers
  // ---------------------------------------------------------------

  function countWords(text) {
    var trimmed = (text || "").trim();
    if (!trimmed) return 0;
    var matches = trimmed.match(/[A-Za-z0-9'-]+/g);
    return matches ? matches.length : 0;
  }

  /** Very rough sentence splitter: splits on ./!/? followed by whitespace. */
  function splitSentences(text) {
    var trimmed = (text || "").trim();
    if (!trimmed) return [];
    var parts = trimmed.match(/[^.!?]+[.!?]+(\s+|$)|[^.!?]+$/g);
    if (!parts) return [trimmed];
    return parts.map(function (s) {
      return s.trim();
    }).filter(Boolean);
  }

  /** Approximate syllable counter for a Flesch-Kincaid style estimate. */
  function estimateSyllables(word) {
    var w = (word || "").toLowerCase().replace(/[^a-z]/g, "");
    if (!w) return 0;
    var groups = w.match(/[aeiouy]+/g);
    var count = groups ? groups.length : 1;
    if (w.length > 2 && w.slice(-1) === "e" && !/[aeiouy]e$/.test(w)) {
      count = Math.max(count - 1, 1);
    }
    return Math.max(count, 1);
  }

  /** Build a short quoted snippet around a match for display in the task pane. */
  function makeSnippet(text, index, length, context) {
    context = typeof context === "number" ? context : 40;
    var start = Math.max(0, index - context);
    var end = Math.min(text.length, index + length + context);
    var prefix = start > 0 ? "…" : "";
    var suffix = end < text.length ? "…" : "";
    return (prefix + text.slice(start, end).trim() + suffix).replace(/\s+/g, " ");
  }

  var issueCounter = 0;
  function makeIssue(opts) {
    issueCounter += 1;
    return {
      id: "issue-" + issueCounter,
      category: opts.category,
      severity: opts.severity || "warning", // 'error' | 'warning' | 'info'
      paragraphIndex: typeof opts.paragraphIndex === "number" ? opts.paragraphIndex : null,
      snippet: opts.snippet || "",
      message: opts.message || "",
      rule: opts.rule || "",
    };
  }

  /** Normalise input paragraphs: accept plain strings or {text, isListItem, style} objects. */
  function normaliseParagraphs(paragraphs) {
    return (paragraphs || []).map(function (p) {
      if (typeof p === "string") {
        return { text: p, isListItem: false, style: "" };
      }
      return {
        text: p && typeof p.text === "string" ? p.text : "",
        isListItem: !!(p && p.isListItem),
        style: (p && p.style) || "",
      };
    });
  }

  // =================================================================
  // 1) OBVIOUS ERRORS
  // =================================================================

  function checkDoubledWords(text, paragraphIndex) {
    var issues = [];
    var re = /\b([A-Za-z]+)\s+\1\b/gi;
    var match;
    while ((match = re.exec(text))) {
      issues.push(
        makeIssue({
          category: "obviousErrors",
          severity: "error",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: 'The word "' + match[1] + '" appears to be repeated twice in a row.',
          rule: "Obvious Errors: Doubled word",
        })
      );
    }
    return issues;
  }

  function checkRepeatedPunctuation(text, paragraphIndex) {
    var issues = [];
    var patterns = [
      { re: /\?{2,}/g, label: "Repeated question marks" },
      { re: /!{2,}/g, label: "Repeated exclamation marks" },
      // Exactly two dots (not a genuine three-dot ellipsis, and not four+ dots)
      { re: /(^|[^.])\.\.([^.]|$)/g, label: "Stray double full stop (not an ellipsis)" },
      { re: /\s,/g, label: "Space before a comma" },
    ];
    patterns.forEach(function (p) {
      var match;
      p.re.lastIndex = 0;
      while ((match = p.re.exec(text))) {
        issues.push(
          makeIssue({
            category: "obviousErrors",
            severity: "error",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message: p.label + " found — check this is not a typo.",
            rule: "Obvious Errors: Repeated/stray punctuation",
          })
        );
      }
    });
    return issues;
  }

  function checkUnbalancedQuotesAndBrackets(text, paragraphIndex) {
    var issues = [];

    function countChar(ch) {
      var re = new RegExp("\\" + ch, "g");
      var m = text.match(re);
      return m ? m.length : 0;
    }

    var straight = countChar('"');
    if (straight % 2 !== 0) {
      issues.push(
        makeIssue({
          category: "obviousErrors",
          severity: "error",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, Math.max(0, text.indexOf('"')), 1),
          message: 'This paragraph has an odd number of straight double quotes (") — a closing or opening quote may be missing.',
          rule: "Obvious Errors: Unmatched quotation marks",
        })
      );
    }

    var curlyOpen = (text.match(/\u201C/g) || []).length;
    var curlyClose = (text.match(/\u201D/g) || []).length;
    if (curlyOpen !== curlyClose) {
      issues.push(
        makeIssue({
          category: "obviousErrors",
          severity: "error",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, 0, Math.min(40, text.length)),
          message: "Curly (smart) quotation marks are unbalanced in this paragraph (" + curlyOpen + " opening vs " + curlyClose + " closing).",
          rule: "Obvious Errors: Unmatched quotation marks",
        })
      );
    }

    var pairs = { "(": ")", "[": "]", "{": "}" };
    var stack = [];
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (pairs[ch]) {
        stack.push({ ch: ch, index: i });
      } else if (ch === ")" || ch === "]" || ch === "}") {
        var last = stack.pop();
        if (!last || pairs[last.ch] !== ch) {
          issues.push(
            makeIssue({
              category: "obviousErrors",
              severity: "error",
              paragraphIndex: paragraphIndex,
              snippet: makeSnippet(text, i, 1),
              message: "Unmatched bracket \"" + ch + "\" found — brackets should always be opened and closed in pairs.",
              rule: "Obvious Errors: Unmatched brackets",
            })
          );
        }
      }
    }
    if (stack.length) {
      stack.forEach(function (open) {
        issues.push(
          makeIssue({
            category: "obviousErrors",
            severity: "error",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, open.index, 1),
            message: 'Opening bracket "' + open.ch + '" is never closed in this paragraph.',
            rule: "Obvious Errors: Unmatched brackets",
          })
        );
      });
    }

    return issues;
  }

  // Small, clearly-labelled "possible typo" list — not a full spellchecker.
  var TYPO_MAP = {
    teh: "the",
    adn: "and",
    recieve: "receive",
    recieved: "received",
    recieving: "receiving",
    definately: "definitely",
    seperate: "separate",
    seperately: "separately",
    occured: "occurred",
    untill: "until",
    wich: "which",
    thier: "their",
    becuase: "because",
    truely: "truly",
  };

  function checkTypos(text, paragraphIndex) {
    var issues = [];
    Object.keys(TYPO_MAP).forEach(function (wrong) {
      var re = new RegExp("\\b" + wrong + "\\b", "gi");
      var match;
      while ((match = re.exec(text))) {
        issues.push(
          makeIssue({
            category: "obviousErrors",
            severity: "warning",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message:
              'Possible typo: "' + match[0] + '" — did you mean "' + TYPO_MAP[wrong] + '"? (This is a small heuristic list, not a full spellcheck.)',
            rule: "Obvious Errors: Possible typo",
          })
        );
      }
    });
    return issues;
  }

  // =================================================================
  // 2) READ-ALOUD FLOW ISSUES
  // =================================================================

  var MAX_SENTENCE_WORDS = 30;

  function checkLongSentences(text, paragraphIndex) {
    var issues = [];
    splitSentences(text).forEach(function (sentence) {
      var words = countWords(sentence);
      if (words > MAX_SENTENCE_WORDS) {
        issues.push(
          makeIssue({
            category: "readAloudFlow",
            severity: "warning",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(sentence, 0, Math.min(60, sentence.length)),
            message: "This sentence is " + words + " words long — long sentences are hard to read aloud in a single breath. Consider splitting it up.",
            rule: "Read-Aloud Flow: Long sentence (>" + MAX_SENTENCE_WORDS + " words)",
          })
        );
      }
    });
    return issues;
  }

  var FILLER_PHRASES = [
    "it's important to note",
    "in today's rapidly changing landscape",
    "let's dive in",
    "in conclusion",
    "delve into",
    "navigate the complexities of",
    "at the end of the day",
    "needless to say",
  ];

  function checkFillerPhrases(text, paragraphIndex) {
    var issues = [];
    FILLER_PHRASES.forEach(function (phrase) {
      var re = new RegExp(phrase.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/'/g, "['’]"), "gi");
      var match;
      while ((match = re.exec(text))) {
        issues.push(
          makeIssue({
            category: "readAloudFlow",
            severity: "info",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message:
              '"' + match[0] + '" is a common AI-generated filler phrase. Consider rewriting in a more natural, spoken style.',
            rule: "Read-Aloud Flow: AI filler phrase",
          })
        );
      }
    });
    return issues;
  }

  function checkReadability(text, paragraphIndex) {
    var issues = [];
    var sentences = splitSentences(text);
    var words = text.match(/[A-Za-z'-]+/g) || [];
    if (!sentences.length || !words.length) return issues;
    var syllables = words.reduce(function (sum, w) {
      return sum + estimateSyllables(w);
    }, 0);
    var score = 206.835 - 1.015 * (words.length / sentences.length) - 84.6 * (syllables / words.length);
    if (words.length >= 15 && score < 50) {
      issues.push(
        makeIssue({
          category: "readAloudFlow",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, 0, Math.min(60, text.length)),
          message:
            "This paragraph scores roughly " +
            Math.round(score) +
            " on a Flesch Reading Ease style estimate (lower = harder to read). Consider simpler words/shorter sentences — informational only.",
          rule: "Read-Aloud Flow: Readability / jargon (informational)",
        })
      );
    }
    return issues;
  }

  function checkPacing(paragraphs) {
    var issues = [];
    var totalWords = paragraphs.reduce(function (sum, p) {
      return sum + countWords(p.text);
    }, 0);
    var MIN_WORDS = 400;
    var MAX_WORDS = 750;
    if (totalWords > 0 && (totalWords < MIN_WORDS || totalWords > MAX_WORDS)) {
      issues.push(
        makeIssue({
          category: "readAloudFlow",
          severity: "info",
          paragraphIndex: null,
          snippet: totalWords + " words total",
          message:
            "At ~125–150 words per minute, this script (" +
            totalWords +
            " words) would run roughly " +
            (totalWords / 150).toFixed(1) +
            "–" +
            (totalWords / 125).toFixed(1) +
            " minutes, which is outside the typical ~" +
            MIN_WORDS +
            "–" +
            MAX_WORDS +
            " word range for a 3–5 minute video. This is an informational pacing note, not an error.",
          rule: "Read-Aloud Flow: Overall pacing (informational)",
        })
      );
    }
    return { issues: issues, totalWords: totalWords };
  }

  // =================================================================
  // 3) BULLET POINTS IN SCRIPTS
  // =================================================================

  var MARKDOWN_BULLET_RE = /^\s*[-*•]\s+\S/;
  var MARKDOWN_NUMBERED_RE = /^\s*\d+[.)]\s+\S/;

  function checkBulletPoints(paragraphs) {
    var issues = [];
    paragraphs.forEach(function (p, index) {
      var trimmed = (p.text || "").trim();
      if (!trimmed) return;
      var isMarkdownList = MARKDOWN_BULLET_RE.test(trimmed) || MARKDOWN_NUMBERED_RE.test(trimmed);
      if (p.isListItem || isMarkdownList) {
        issues.push(
          makeIssue({
            category: "bulletPoints",
            severity: "warning",
            paragraphIndex: index,
            snippet: makeSnippet(trimmed, 0, Math.min(60, trimmed.length)),
            message:
              "This looks like a bulleted/numbered list item. Scripts are meant to be read aloud as flowing prose — bullet points should generally be rewritten as sentences.",
            rule: "Scripting Principle: No bullet points in scripts",
          })
        );
      }
    });
    return issues;
  }

  /**
   * For genuine Word list items (not markdown bullets), check the editorial
   * formatting rules for any lists intentionally kept: the stem must end in
   * a colon, items start with a capital letter, and full stops are applied
   * consistently (all-but-last unless every item is a full sentence/question).
   */
  function checkListFormatting(paragraphs) {
    var issues = [];
    var i = 0;
    while (i < paragraphs.length) {
      if (paragraphs[i].isListItem) {
        var start = i;
        var end = i;
        while (end + 1 < paragraphs.length && paragraphs[end + 1].isListItem) {
          end += 1;
        }
        var items = paragraphs.slice(start, end + 1);

        // Stem check: the paragraph immediately before the list should end in ':'
        if (start > 0) {
          var stem = (paragraphs[start - 1].text || "").trim();
          if (stem && !/:$/.test(stem)) {
            issues.push(
              makeIssue({
                category: "houseStyle",
                severity: "info",
                paragraphIndex: start - 1,
                snippet: makeSnippet(stem, Math.max(0, stem.length - 40), Math.min(40, stem.length)),
                message: "The sentence introducing this list should end with a colon.",
                rule: "House Style: Bullet/list formatting",
              })
            );
          }
        }

        // Capitalisation check
        items.forEach(function (item, offset) {
          var t = (item.text || "").trim();
          if (t && /^[a-z]/.test(t)) {
            issues.push(
              makeIssue({
                category: "houseStyle",
                severity: "info",
                paragraphIndex: start + offset,
                snippet: makeSnippet(t, 0, Math.min(40, t.length)),
                message: "List items should start with a capital letter.",
                rule: "House Style: Bullet/list formatting",
              })
            );
          }
        });

        // Full stop / question mark consistency
        var allQuestions = items.every(function (item) {
          return /\?$/.test((item.text || "").trim());
        });
        var anyMultiSentence = items.some(function (item) {
          return splitSentences(item.text).length > 1;
        });
        items.forEach(function (item, offset) {
          var t = (item.text || "").trim();
          if (!t) return;
          var isLast = offset === items.length - 1;
          if (allQuestions) {
            if (!/\?$/.test(t)) {
              issues.push(
                makeIssue({
                  category: "houseStyle",
                  severity: "info",
                  paragraphIndex: start + offset,
                  snippet: makeSnippet(t, Math.max(0, t.length - 40), Math.min(40, t.length)),
                  message: "All items in this list are questions, so each item should end with a question mark.",
                  rule: "House Style: Bullet/list formatting",
                })
              );
            }
          } else if (anyMultiSentence) {
            if (!/[.!?]$/.test(t)) {
              issues.push(
                makeIssue({
                  category: "houseStyle",
                  severity: "info",
                  paragraphIndex: start + offset,
                  snippet: makeSnippet(t, Math.max(0, t.length - 40), Math.min(40, t.length)),
                  message: "At least one list item has multiple sentences, so every item should end with a full stop.",
                  rule: "House Style: Bullet/list formatting",
                })
              );
            }
          } else if (!isLast && /[.!?]$/.test(t)) {
            issues.push(
              makeIssue({
                category: "houseStyle",
                severity: "info",
                paragraphIndex: start + offset,
                snippet: makeSnippet(t, Math.max(0, t.length - 40), Math.min(40, t.length)),
                message: "Only the last item in a simple list should end with a full stop.",
                rule: "House Style: Bullet/list formatting",
              })
            );
          }
        });

        i = end + 1;
      } else {
        i += 1;
      }
    }
    return issues;
  }

  // =================================================================
  // 4) HOUSE STYLE / EDITORIAL COMPLIANCE  (+ 5) SESSION NUMBERING
  // =================================================================

  var NUMBER_WORDS = [
    "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
    "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen",
    "eighteen", "nineteen", "twenty",
  ];

  var FUTURE_PROOFING_NOTE =
    " Avoid specific session/activity numbers in scripts in case sessions are reordered later; use words like 'session five' sparingly and only if numbering is stable, or avoid entirely.";

  function checkSessionActivityNumbers(text, paragraphIndex) {
    var issues = [];
    // Matches both digit forms ("session 5", "activity 3.3") and spelled-out
    // word forms ("session one", "Session One") so every numbered
    // session/activity reference is surfaced for future-proofing, not just
    // the digit ones.
    var re = new RegExp(
      "\\b(session|activity)\\s+(\\d+(?:\\.\\d+)*|" + NUMBER_WORDS.join("|") + ")\\b",
      "gi"
    );
    var match;
    while ((match = re.exec(text))) {
      var keyword = match[1];
      var keywordLower = keyword.toLowerCase();
      var number = match[2];
      var isDigitForm = /^\d/.test(number);
      var isSpecificSubReference = (isDigitForm && number.indexOf(".") !== -1) || keywordLower === "activity";
      var isExactlyCorrectCase = match[0] === match[0].toLowerCase();

      var message;
      var severity;
      if (isSpecificSubReference) {
        message =
          'Reference to a specific activity/page number ("' +
          match[0] +
          '") — avoid these entirely, as they may move during production.';
        severity = "warning";
      } else if (isDigitForm) {
        message =
          'Scripts should spell out session numbers in lowercase words (e.g. "session five"), not digits ("' +
          match[0] +
          '").';
        severity = "warning";
      } else if (!isExactlyCorrectCase) {
        message =
          'Session references should use a lowercase "s" and a lowercase spelled-out number, e.g. "session five", not "' +
          match[0] +
          '".';
        severity = "warning";
      } else {
        // Correctly-formatted "session five" — not a style violation, just
        // surfaced as an informational future-proofing note.
        message = 'Session reference found ("' + match[0] + '").';
        severity = "info";
      }

      issues.push(
        makeIssue({
          category: "sessionNumbering",
          severity: severity,
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: message + FUTURE_PROOFING_NOTE,
          rule: "Session Numbering: Numeral/word number used with session/activity",
        })
      );
    }
    return issues;
  }

  var HYPHEN_TERMS = [
    "co-operate",
    "co-ordinate",
    "cost-effective",
    "decision-making",
    "director-general",
    "first-year",
    "full-time",
    "high-risk",
    "HIV-positive",
    "in-house",
    "long-standing",
    "long-term",
    "on-campus",
    "part-time",
    "policy-making",
    "re-examine",
  ];

  function checkHyphenation(text, paragraphIndex) {
    var issues = [];
    HYPHEN_TERMS.forEach(function (hyphenated) {
      var unhyphenated = hyphenated.replace(/-/g, " ");
      var re = new RegExp("\\b" + unhyphenated.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "gi");
      var match;
      while ((match = re.exec(text))) {
        issues.push(
          makeIssue({
            category: "houseStyle",
            severity: "warning",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message: '"' + match[0] + '" should be hyphenated: "' + hyphenated + '".',
            rule: "House Style: Hyphenation",
          })
        );
      }
    });
    return issues;
  }

  function checkDashUsage(text, paragraphIndex) {
    var issues = [];

    var asideRe = /\S \- \S/g;
    var match;
    while ((match = asideRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: 'A plain hyphen is being used as a parenthetical aside (" - ") — house style prefers a spaced em dash ( — ) for this.',
          rule: "House Style: En/em dash usage",
        })
      );
    }

    var rangeRe = /\b\d{2,4}-\d{2,4}\b/g;
    while ((match = rangeRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: 'Number range "' + match[0] + '" should use an en dash (–) rather than a hyphen (-), e.g. "1950–1980".',
          rule: "House Style: En/em dash usage",
        })
      );
    }

    return issues;
  }

  function checkSlashUsage(text, paragraphIndex) {
    var issues = [];
    var re = /\b([A-Za-z]+)\/([A-Za-z]+)\b/g;
    var match;
    while ((match = re.exec(text))) {
      var before = text.slice(Math.max(0, match.index - 8), match.index);
      if (/https?:$|www\.$|^\s*[\\/]/.test(before)) continue; // skip URLs/paths
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: '"' + match[0] + '" — house style adds spaces around a forward slash, e.g. "' + match[1] + " / " + match[2] + '".',
          rule: "House Style: Forward slash spacing",
        })
      );
    }
    return issues;
  }

  function checkLatinAbbreviations(text, paragraphIndex) {
    var issues = [];
    var noDotPatterns = [
      { re: /\beg\b(?!\.)/gi, correct: "e.g." },
      { re: /\bie\b(?!\.)/gi, correct: "i.e." },
      { re: /\betc\b(?!\.)/gi, correct: "etc." },
    ];
    noDotPatterns.forEach(function (p) {
      var match;
      p.re.lastIndex = 0;
      while ((match = p.re.exec(text))) {
        issues.push(
          makeIssue({
            category: "houseStyle",
            severity: "warning",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message: '"' + match[0] + '" should be written as "' + p.correct + '".',
            rule: "House Style: Latin abbreviations",
          })
        );
      }
    });

    var commaRe = /\b(e\.g\.|i\.e\.),/gi;
    var match;
    while ((match = commaRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "warning",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: 'British house style does not use a comma after "' + match[1] + '".',
          rule: "House Style: Latin abbreviations",
        })
      );
    }
    return issues;
  }

  var NUMBER_WORDS = ["one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"];
  var UNIT_WORDS_RE = /^(kg|km|cm|mm|g|lb|lbs|ft|hr|hrs|min|mins|sec|secs|am|pm|°c|°f|%)\b/i;

  function checkNumberStyle(text, paragraphIndex) {
    var issues = [];

    // Standalone digits 1-10 not touching a unit/currency symbol.
    var re = /(^|[^\w.$£€%])\b([1-9]|10)\b([^\w.$£€%]|$)/g;
    var match;
    while ((match = re.exec(text))) {
      var numStart = match.index + match[1].length;
      var afterIndex = numStart + match[2].length;
      var rest = text.slice(afterIndex).replace(/^\s+/, "");
      if (UNIT_WORDS_RE.test(rest)) continue; // "5 kg" etc. — measurement context
      // Skip numbered-list starts like "1. " or "1)"
      var beforeTrim = text.slice(0, match.index + match[1].length).trimEnd();
      if (beforeTrim === "" && /^[.)]/.test(text.slice(afterIndex))) continue;
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, numStart, match[2].length),
          message:
            'The number "' +
            match[2] +
            '" should usually be spelled out in prose (e.g. "' +
            NUMBER_WORDS[parseInt(match[2], 10) - 1] +
            '"), unless it is a measurement, currency amount or part of a list/table.',
          rule: "House Style: Numbers (spell out 1–10)",
        })
      );
    }

    // Sentence starting with a numeral.
    splitSentences(text).forEach(function (sentence) {
      var m = /^\s*(\d+)/.exec(sentence);
      if (m) {
        issues.push(
          makeIssue({
            category: "houseStyle",
            severity: "warning",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(sentence, 0, Math.min(40, sentence.length)),
            message: "Sentences should not start with a numeral — spell the number out instead, regardless of its size.",
            rule: "House Style: Numbers (never start a sentence with a digit)",
          })
        );
      }
    });

    return issues;
  }

  function checkPercentageConsistency(paragraphs) {
    var issues = [];
    var wordFormIndex = -1;
    var symbolFormIndex = -1;
    paragraphs.forEach(function (p, index) {
      if (wordFormIndex === -1 && /\bper\s?cent\b/i.test(p.text)) wordFormIndex = index;
      if (symbolFormIndex === -1 && /\d\s?%/.test(p.text)) symbolFormIndex = index;
    });
    if (wordFormIndex !== -1 && symbolFormIndex !== -1) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: wordFormIndex,
          snippet: "Both word and symbol forms found (e.g. paragraphs " + wordFormIndex + " and " + symbolFormIndex + ")",
          message: 'This document mixes "per cent"/"percent" with the "%" symbol. House style should use one form consistently throughout.',
          rule: "House Style: Percentage consistency",
        })
      );
    }
    return issues;
  }

  var MONTHS = "January|February|March|April|May|June|July|August|September|October|November|December";

  function checkDateFormats(text, paragraphIndex) {
    var issues = [];

    var numericRe = /\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b/g;
    var match;
    while ((match = numericRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "warning",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: 'Numeric date "' + match[0] + '" should be written in full house style, e.g. "1st January 2024".',
          rule: "House Style: Date format",
        })
      );
    }

    var wordDateRe = new RegExp("\\b(\\d{1,2})(st|nd|rd|th)?\\s+(" + MONTHS + ")\\s+(\\d{4})\\b", "gi");
    while ((match = wordDateRe.exec(text))) {
      if (!match[2]) {
        issues.push(
          makeIssue({
            category: "houseStyle",
            severity: "info",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message: 'Date "' + match[0] + '" is missing an ordinal suffix — house style uses e.g. "' + match[1] + getOrdinalSuffix(match[1]) + " " + match[3] + " " + match[4] + '".',
            rule: "House Style: Date format",
          })
        );
      }
    }

    return issues;
  }

  function getOrdinalSuffix(dayStr) {
    var day = parseInt(dayStr, 10);
    if (day % 10 === 1 && day !== 11) return "st";
    if (day % 10 === 2 && day !== 12) return "nd";
    if (day % 10 === 3 && day !== 13) return "rd";
    return "th";
  }

  var AMERICAN_SPELLING_MAP = {
    organize: "organise",
    organizes: "organises",
    organized: "organised",
    organizing: "organising",
    organization: "organisation",
    organizations: "organisations",
    recognize: "recognise",
    recognizes: "recognises",
    recognized: "recognised",
    recognizing: "recognising",
    color: "colour",
    colors: "colours",
    colored: "coloured",
    coloring: "colouring",
    center: "centre",
    centers: "centres",
    centered: "centred",
    analyze: "analyse",
    analyzes: "analyses",
    analyzed: "analysed",
    analyzing: "analysing",
    behavior: "behaviour",
    behaviors: "behaviours",
    whilst: "while",
    learnt: "learned",
  };

  function checkAmericanSpelling(text, paragraphIndex) {
    var issues = [];
    Object.keys(AMERICAN_SPELLING_MAP).forEach(function (wrong) {
      var re = new RegExp("\\b" + wrong + "\\b", "gi");
      var match;
      while ((match = re.exec(text))) {
        issues.push(
          makeIssue({
            category: "houseStyle",
            severity: "warning",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message: '"' + match[0] + '" should use house spelling: "' + AMERICAN_SPELLING_MAP[wrong] + '".',
            rule: "House Style: Spelling (British vs American)",
          })
        );
      }
    });

    // practice / practise heuristic
    var verbPatterns = [/\bto practice\b/gi, /\bpracticing\b/gi, /\bpracticed\b/gi];
    verbPatterns.forEach(function (re) {
      var match;
      while ((match = re.exec(text))) {
        var suggestion = match[0].replace(/practic/i, function (m) {
          return m === "Practic" ? "Practis" : "practis";
        });
        issues.push(
          makeIssue({
            category: "houseStyle",
            severity: "info",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message: '"' + match[0] + '" is a verb use, so house style expects "' + suggestion + '" (practise = verb, practice = noun).',
            rule: "House Style: practice vs practise",
          })
        );
      }
    });
    var nounRe = /\b(a|the|this|that|good|best|common)\s+practise\b/gi;
    var nMatch;
    while ((nMatch = nounRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, nMatch.index, nMatch[0].length),
          message: '"' + nMatch[0] + '" is a noun use, so house style expects "practice" (practice = noun, practise = verb).',
          rule: "House Style: practice vs practise",
        })
      );
    }

    var vsRe = /\bvs\.(?=\s|$)/gi;
    var vMatch;
    while ((vMatch = vsRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, vMatch.index, vMatch[0].length),
          message: '"vs." should be written without a full stop: "vs".',
          rule: "House Style: vs (no full stop)",
        })
      );
    }

    var ampRe = /&/g;
    var aMatch;
    while ((aMatch = ampRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, aMatch.index, 1),
          message: '"&" found — house style reserves the ampersand for company/brand names; otherwise spell out "and".',
          rule: "House Style: Ampersand usage",
        })
      );
    }

    return issues;
  }

  function checkOxfordComma(text, paragraphIndex) {
    var issues = [];
    // Each list "item" may be more than one word (e.g. "a laptop", "a notebook").
    var re = /\b[\w'-]+(?:\s+[\w'-]+)*,\s+[\w'-]+(?:\s+[\w'-]+)*,\s+(and|or)\s+[\w'-]+(?:\s+[\w'-]+)*/gi;
    var match;
    while ((match = re.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: "Possible Oxford/serial comma — house style generally avoids a comma before \"" + match[1] + '" in a simple list unless needed for clarity.',
          rule: "House Style: Oxford comma (note, low confidence)",
        })
      );
    }
    return issues;
  }

  function checkTitleCapitalisation(paragraph, paragraphIndex) {
    var issues = [];
    if (!/heading|title/i.test(paragraph.style || "")) return issues;
    var text = (paragraph.text || "").trim();
    if (!text) return issues;
    var words = text.split(/\s+/);
    if (words.length < 2) return issues;
    var capitalisedAfterFirst = words.slice(1).filter(function (w) {
      return /^[A-Z][a-z]/.test(w);
    }).length;
    if (capitalisedAfterFirst >= 2) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, 0, Math.min(60, text.length)),
          message:
            'This title looks like it is in Title Case. House style uses initial-capital-then-lowercase for titles (e.g. "Getting started with research methods"), except for proper nouns, including after a colon.',
          rule: "House Style: Title capitalisation",
        })
      );
    }
    return issues;
  }

  function checkHouseAbbreviations(text, paragraphIndex) {
    var issues = [];

    var profRe = /\bProf\.?\s+[A-Z]/g;
    var match;
    while ((match = profRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "error",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: '"Professor" must never be abbreviated to "Prof" in house style.',
          rule: "House Style: Professor must not be abbreviated",
        })
      );
    }

    [/\bDr\.(?=\s)/g, /\bMr\.(?=\s)/g, /\bMrs\.(?=\s)/g].forEach(function (re) {
      var m;
      while ((m = re.exec(text))) {
        var title = m[0].replace(".", "");
        issues.push(
          makeIssue({
            category: "houseStyle",
            severity: "warning",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, m.index, m[0].length),
            message: 'House style does not use a full stop after "' + title + '" — write "' + title + '" with no full stop.',
            rule: "House Style: Titles (Dr/Mr/Mrs) have no full stop",
          })
        );
      }
    });

    var syndicateRe = /\bsyndicate group\b/gi;
    while ((match = syndicateRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "houseStyle",
          severity: "warning",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, match.index, match[0].length),
          message: '"Syndicate group" is out of use in house style — use "study group" instead.',
          rule: "House Style: Syndicate group -> study group",
        })
      );
    }

    return issues;
  }

  // =================================================================
  // 6) INCLUSIVITY / DIVERSITY
  // =================================================================

  var INCLUSIVITY_TERMS = [
    { re: /\bhandicapped\b/gi, suggestion: "person/people with disabilities (or 'disabled person/people')" },
    { re: /\b(an|the)\s+invalid\b/gi, suggestion: "person/people with disabilities" },
    { re: /\bthe disabled\b/gi, suggestion: "disabled people / people with disabilities" },
    { re: /\b(a|an|the)\s+(diabetics?|epileptics?|anorexics?|bulimics?|schizophrenics?)\b/gi, suggestion: "a person with [condition]" },
    { re: /\b(diabetics|epileptics|anorexics|bulimics|schizophrenics)\b/gi, suggestion: "people with [condition]" },
    { re: /\bschizo\b/gi, suggestion: "a person with schizophrenia" },
    { re: /\bable-bodied\b/gi, suggestion: "non-disabled" },
    { re: /\bmankind\b/gi, suggestion: "humankind" },
    { re: /\bmanpower\b/gi, suggestion: "workforce" },
    { re: /\bbusinessm[ae]n\b/gi, suggestion: "businessperson/businesspeople" },
    { re: /\bactress(es)?\b/gi, suggestion: "actor(s)" },
    { re: /\bchairman\b/gi, suggestion: "chair" },
    { re: /\bauthoress(es)?\b/gi, suggestion: "author(s)" },
    { re: /\bpolice(man|woman|men|women)\b/gi, suggestion: "police officer(s)" },
    { re: /\bheadmaster\b/gi, suggestion: "head teacher" },
    { re: /\bheadmistress\b/gi, suggestion: "head teacher" },
    { re: /\b(woman doctor|female doctor)\b/gi, suggestion: "doctor (only mention gender if relevant to the content)" },
    { re: /\bhe\/she\b/gi, suggestion: "they" },
    { re: /\bs\/he\b/gi, suggestion: "they" },
    { re: /\bforeign\b/gi, suggestion: "international / overseas" },
    { re: /\bthird world\b/gi, suggestion: "Global South / developing world" },
    { re: /\b(oriental|far east)\b/gi, suggestion: "Asia, or name the specific region/country" },
    { re: /\bdevelop(ed|ing) countries\b/gi, suggestion: "a specific measure of development (used loosely here)" },
  ];

  function checkInclusivity(text, paragraphIndex) {
    var issues = [];
    INCLUSIVITY_TERMS.forEach(function (term) {
      var match;
      term.re.lastIndex = 0;
      while ((match = term.re.exec(text))) {
        issues.push(
          makeIssue({
            category: "inclusivity",
            severity: "warning",
            paragraphIndex: paragraphIndex,
            snippet: makeSnippet(text, match.index, match[0].length),
            message: '"' + match[0] + '" — house inclusive-language guidance suggests: ' + term.suggestion + ".",
            rule: "Inclusivity: Avoid-term lookup",
          })
        );
      }
    });

    var genericSubjects = "(a|an|the)\\s+(student|manager|learner|teacher|doctor|nurse|engineer|user|customer|participant|member)|someone|anyone|everybody|everyone|somebody|anybody";
    var pronounRe = new RegExp("\\b(" + genericSubjects + ")\\b([^.?!]{0,40})\\b(he|his|him|she|her)\\b", "gi");
    var pMatch;
    while ((pMatch = pronounRe.exec(text))) {
      issues.push(
        makeIssue({
          category: "inclusivity",
          severity: "info",
          paragraphIndex: paragraphIndex,
          snippet: makeSnippet(text, pMatch.index, pMatch[0].length),
          message: "Generic pronoun used after a non-specific subject — consider \"they/their/them\" instead of \"" + pMatch[4] + '".',
          rule: "Inclusivity: Generic gendered pronoun (heuristic, low confidence)",
        })
      );
    }

    return issues;
  }

  // =================================================================
  // Main entry point
  // =================================================================

  function checkDocument(rawParagraphs) {
    issueCounter = 0;
    var paragraphs = normaliseParagraphs(rawParagraphs);
    var issues = [];

    paragraphs.forEach(function (p, index) {
      var text = p.text || "";
      if (!text.trim()) return;

      issues = issues.concat(checkDoubledWords(text, index));
      issues = issues.concat(checkRepeatedPunctuation(text, index));
      issues = issues.concat(checkUnbalancedQuotesAndBrackets(text, index));
      issues = issues.concat(checkTypos(text, index));

      issues = issues.concat(checkLongSentences(text, index));
      issues = issues.concat(checkFillerPhrases(text, index));
      issues = issues.concat(checkReadability(text, index));

      issues = issues.concat(checkSessionActivityNumbers(text, index));

      issues = issues.concat(checkHyphenation(text, index));
      issues = issues.concat(checkDashUsage(text, index));
      issues = issues.concat(checkSlashUsage(text, index));
      issues = issues.concat(checkLatinAbbreviations(text, index));
      issues = issues.concat(checkNumberStyle(text, index));
      issues = issues.concat(checkDateFormats(text, index));
      issues = issues.concat(checkAmericanSpelling(text, index));
      issues = issues.concat(checkOxfordComma(text, index));
      issues = issues.concat(checkTitleCapitalisation(p, index));
      issues = issues.concat(checkHouseAbbreviations(text, index));

      issues = issues.concat(checkInclusivity(text, index));
    });

    issues = issues.concat(checkBulletPoints(paragraphs));
    issues = issues.concat(checkListFormatting(paragraphs));
    issues = issues.concat(checkPercentageConsistency(paragraphs));

    var pacing = checkPacing(paragraphs);
    issues = issues.concat(pacing.issues);

    var countsByCategory = {};
    CATEGORY_ORDER.forEach(function (cat) {
      countsByCategory[cat] = 0;
    });
    issues.forEach(function (issue) {
      countsByCategory[issue.category] = (countsByCategory[issue.category] || 0) + 1;
    });

    return {
      issues: issues,
      countsByCategory: countsByCategory,
      totalIssues: issues.length,
      totalWords: pacing.totalWords,
      categories: CATEGORIES,
      categoryOrder: CATEGORY_ORDER,
    };
  }

  return {
    checkDocument: checkDocument,
    CATEGORIES: CATEGORIES,
    CATEGORY_ORDER: CATEGORY_ORDER,
    // exported for unit testing of internals
    _internal: {
      countWords: countWords,
      splitSentences: splitSentences,
      estimateSyllables: estimateSyllables,
      checkDoubledWords: checkDoubledWords,
      checkRepeatedPunctuation: checkRepeatedPunctuation,
      checkUnbalancedQuotesAndBrackets: checkUnbalancedQuotesAndBrackets,
      checkTypos: checkTypos,
      checkLongSentences: checkLongSentences,
      checkFillerPhrases: checkFillerPhrases,
      checkReadability: checkReadability,
      checkPacing: checkPacing,
      checkBulletPoints: checkBulletPoints,
      checkListFormatting: checkListFormatting,
      checkSessionActivityNumbers: checkSessionActivityNumbers,
      checkHyphenation: checkHyphenation,
      checkDashUsage: checkDashUsage,
      checkSlashUsage: checkSlashUsage,
      checkLatinAbbreviations: checkLatinAbbreviations,
      checkNumberStyle: checkNumberStyle,
      checkPercentageConsistency: checkPercentageConsistency,
      checkDateFormats: checkDateFormats,
      checkAmericanSpelling: checkAmericanSpelling,
      checkOxfordComma: checkOxfordComma,
      checkTitleCapitalisation: checkTitleCapitalisation,
      checkHouseAbbreviations: checkHouseAbbreviations,
      checkInclusivity: checkInclusivity,
    },
  };
});
