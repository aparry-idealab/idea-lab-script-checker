/**
 * Copies the framework-agnostic rule engine (src/rules/ruleEngine.js)
 * into chrome-extension/lib/ so the Chrome extension can load it as a
 * plain classic script (no bundler). src/rules/ruleEngine.js remains
 * the single source of truth — re-run this (npm run build:extension)
 * after editing it.
 */
"use strict";

const fs = require("fs");
const path = require("path");

const src = path.join(__dirname, "..", "src", "rules", "ruleEngine.js");
const destDir = path.join(__dirname, "..", "chrome-extension", "lib");
const dest = path.join(destDir, "ruleEngine.js");

const banner =
  "/* AUTO-GENERATED — do not edit directly.\n" +
  " * Copied from src/rules/ruleEngine.js by scripts/copy-rule-engine.js.\n" +
  " * Run `npm run build:extension` after changing the source file. */\n";

fs.mkdirSync(destDir, { recursive: true });
fs.writeFileSync(dest, banner + fs.readFileSync(src, "utf8"));
console.log("Copied src/rules/ruleEngine.js -> chrome-extension/lib/ruleEngine.js");
