// Run with: node tests/test_launch_review_ui.js
// Checks the 2026-10-12 UI fixes: date entry (A2), zero amounts (B4), blank % Paid (C).
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('web/app.js', 'utf8');

function fn(name) {
  const start = source.indexOf(`  function ${name}(`);
  assert.notEqual(start, -1, `Missing function ${name}`);
  let depth = 0;
  for (let i = source.indexOf('{', start); i < source.length; i++) {
    if (source[i] === '{') depth++;
    else if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(`Unbalanced function ${name}`);
}

const toasts = [];
const context = { showToast: (msg, kind) => toasts.push([msg, kind]) };
vm.createContext(context);
for (const name of ['maskDateString', 'normalizeDateInput', 'commitDateInput', 'isBlankAmount', 'parseAmount', 'formatPercent']) {
  vm.runInContext(fn(name), context);
}

// A2: typing with separators is never reshuffled.
assert.equal(context.maskDateString('1/5/2026'), '1/5/2026');
assert.equal(context.maskDateString('2026-10-12'), '2026-10-12');
assert.equal(context.maskDateString('01052026'), '01/05/2026');

// A2: every common way of writing a date ends as MM/DD/YYYY.
for (const [input, expected] of [
  ['1/5/2026', '01/05/2026'], ['01/05/2026', '01/05/2026'], ['2026-10-12', '10/12/2026'],
  ['10/12/26', '10/12/2026'], ['01052026', '01/05/2026'], ['25/12/2026', '12/25/2026'],
  ['46307', '10/12/2026'], ['', ''], ['-', ''],
]) {
  assert.equal(context.normalizeDateInput(input), expected, input);
}
// ...and impossible dates are refused, never turned into something else.
for (const bad of ['15/20/26', '02/30/2026', '13/13/2026', 'abc', '1/5']) {
  assert.equal(context.normalizeDateInput(bad), null, bad);
}
const input = { value: '15/20/26' };
assert.equal(context.commitDateInput(input, '01/05/2026'), null);
assert.equal(input.value, '01/05/2026');
assert.equal(toasts.at(-1)[1], 'err');

// B4: 0 is an amount, blank is not.
assert.equal(context.isBlankAmount('0'), false);
assert.equal(context.isBlankAmount('₱0.00'), false);
assert.equal(context.isBlankAmount(''), true);
assert.equal(context.isBlankAmount(' - '), true);

// C: blank % Paid stays blank.
assert.equal(context.formatPercent(''), '');
assert.equal(context.formatPercent(null), '');
assert.equal(context.formatPercent('33.3'), '33.30%');

console.log('Date entry, zero amount and % Paid UI checks passed.');
