const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');
const mainJs = fs.readFileSync(path.join(repoRoot, 'main.js'), 'utf8');

function extractMethodSource(source, methodName) {
  const marker = `${methodName}(markdown) {`;
  const start = source.indexOf(marker);
  assert.notStrictEqual(start, -1, `Could not find ${methodName} in main.js`);

  let braceDepth = 0;
  const bodyStart = source.indexOf('{', start);
  let i = bodyStart;
  for (; i < source.length; i++) {
    const ch = source[i];
    if (ch === '{') braceDepth++;
    if (ch === '}') {
      braceDepth--;
      if (braceDepth === 0) break;
    }
  }

  const methodText = source.slice(start, i + 1);
  return `({ ${methodText} })`;
}

function loadNormalizer() {
  const ctx = vm.createContext({ RegExp, String, Array, Object });
  const detection = vm.runInContext(extractMethodSource(mainJs, 'isMalformedSummaryMarkdown'), ctx);
  const normalizer = vm.runInContext(extractMethodSource(mainJs, 'normalizeMalformedSummaryMarkdown'), ctx);
  return {
    isMalformedSummaryMarkdown: detection.isMalformedSummaryMarkdown,
    normalizeMalformedSummaryMarkdown: normalizer.normalizeMalformedSummaryMarkdown,
  };
}

const malformedFixture = fs.readFileSync(
  path.join(__dirname, 'fixtures', 'malformed-morey-summary.md'),
  'utf8'
);
const wellFormedFixture = fs.readFileSync(
  path.join(__dirname, 'fixtures', 'well-formed-trillium-summary.md'),
  'utf8'
);

test('detects malformed collapsed markdown', () => {
  const { isMalformedSummaryMarkdown } = loadNormalizer();
  assert.equal(isMalformedSummaryMarkdown(malformedFixture), true);
  assert.equal(isMalformedSummaryMarkdown(wellFormedFixture), false);
});

test('normalizes collapsed metadata and section boundaries', () => {
  const { normalizeMalformedSummaryMarkdown } = loadNormalizer();
  const normalized = normalizeMalformedSummaryMarkdown(malformedFixture);

  assert.match(normalized, /^### Metadata$/m);
  assert.match(normalized, /^```json$/m);
  assert.match(normalized, /^```$/m);
  assert.match(normalized, /^### Current Test Station Infrastructure$/m);
  assert.match(normalized, /^### Data Collection and Ticketing System$/m);
  assert.match(normalized, /^- 300-400 unique test stations spanning decades of technology$/m);
  assert.match(normalized, /^- Slack-based automated ticketing system$/m);
});

test('leaves well-formed markdown unchanged', () => {
  const { normalizeMalformedSummaryMarkdown } = loadNormalizer();
  assert.equal(normalizeMalformedSummaryMarkdown(wellFormedFixture), wellFormedFixture.trim());
});
