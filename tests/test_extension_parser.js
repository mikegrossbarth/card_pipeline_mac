const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const contentPath = path.join(__dirname, "..", "cardladder-autocomp", "extension", "content.js");
const source = fs.readFileSync(contentPath, "utf8");

function functionSource(name) {
  const start = source.indexOf(`function ${name}(`);
  if (start < 0) throw new Error(`Missing function ${name}`);
  const bodyStart = source.indexOf("{", start);
  let depth = 0;
  for (let index = bodyStart; index < source.length; index += 1) {
    const char = source[index];
    if (char === "{") depth += 1;
    if (char === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(start, index + 1);
    }
  }
  throw new Error(`Could not extract function ${name}`);
}

const parserCode = `
const COMP_SOURCE_LABELS = ["eBay", "Fanatics", "Card Ladder"];
${functionSource("sourceLabelToPattern")}
const COMP_SOURCE_PATTERN_TEXT = COMP_SOURCE_LABELS
  .map(sourceLabelToPattern)
  .sort((a, b) => b.length - a.length)
  .join("|");
const COMP_SOURCE_PATTERN = new RegExp("\\\\b(" + COMP_SOURCE_PATTERN_TEXT + ")\\\\b", "i");
${functionSource("compDatePattern")}
${functionSource("sourceLineMatch")}
${functionSource("currentCompChunkOnly")}
${functionSource("cleanCompTitle")}
${functionSource("parseCompChunk")}
globalThis.parseCompChunk = parseCompChunk;
globalThis.sourceLineMatch = sourceLineMatch;
`;

const context = {};
vm.createContext(context);
vm.runInContext(parserCode, context);

const chunk = [
  "eBay 2014 Panini Flawless Greats Patches Autographs Gold #22 Joe Montana BGS 9",
  "Jun 1, 2026 Auction $25.00",
  "Fanatics 2022 Panini Donruss Chet Holmgren PSA 9",
  "Jun 2, 2026 Buy It Now $90.00",
].join(" ");

const comp = context.parseCompChunk(chunk, context.sourceLineMatch("eBay"));

assert(comp, "Expected first comp to parse");
assert.strictEqual(comp.source, "EBAY");
assert.strictEqual(comp.date_sold, "Jun 1, 2026");
assert.strictEqual(comp.price, "$25.00");
assert.match(comp.title, /Joe Montana/);
assert.doesNotMatch(comp.title, /Chet Holmgren/);

console.log("extension parser regression ok");
