// Markdown helpers shared by the label scripts' summaries. Label names, colors,
// paths and API error text come from the manifest, the live repository or the
// network, so each one is rendered here instead of being pasted into Markdown.

import assert from "node:assert/strict";

// One line: CR, LF and CRLF each become a single space, so a value cannot
// start a new paragraph or split a table row.
export function singleLine(value) {
  return String(value).replace(/\r\n?|\n/g, " ");
}

// Plain text in a table cell: one line, with pipes escaped.
export function markdownEscape(value) {
  return singleLine(value).replaceAll("|", "\\|");
}

// A code span that the value cannot close: its fence is one backtick longer
// than the longest run of backticks in the value. A space pads a value that
// starts or ends with a backtick, which would otherwise join the fence, and a
// value that starts and ends with a space, from which Markdown strips one
// space on each side. In a table cell, pipes are escaped as well; GitHub
// removes that backslash inside a code span.
export function inlineCode(value, { table = false } = {}) {
  let text = singleLine(value);
  if (table) text = text.replaceAll("|", "\\|");
  if (text.length === 0) return "` `";
  const longest = Math.max(0, ...(text.match(/`+/g) ?? []).map(run => run.length));
  const fence = "`".repeat(longest + 1);
  const padded = /^`|`$/.test(text) || (/^ /.test(text) && / $/.test(text) && /[^ ]/.test(text));
  const padding = padded ? " " : "";
  return `${fence}${padding}${text}${padding}${fence}`;
}

// The helpers' own cases; each label script runs them in its self-test.
export function selfTestMarkdown() {
  for (const [value, expected] of [
    ["plain", "`plain`"],
    ["a`b", "``a`b``"],
    ["wip `x`", "`` wip `x` ``"],
    ["`x` wip", "`` `x` wip ``"],
    ["a ``b`` c", "```a ``b`` c```"],
    ["a `b` ```c```", "```` a `b` ```c``` ````"],
    ["`edge", "`` `edge ``"],
    ["edge`", "`` edge` ``"],
    ["`", "`` ` ``"],
    [" spaced ", "`  spaced  `"],
    ["   ", "`   `"],
    ["", "` `"],
    ["a|b", "`a|b`"],
    ["one\r\ntwo\rthree\nfour", "`one two three four`"],
    ["\r\n", "` `"],
    [404, "`404`"]
  ]) {
    assert.equal(inlineCode(value), expected, JSON.stringify(value));
  }
  assert.equal(inlineCode("a|b", { table: true }), "`a\\|b`");
  assert.equal(inlineCode("`a|b`\r\nc", { table: true }), "`` `a\\|b` c ``");
  assert.equal(markdownEscape("x|y\r\nz\rw\nv"), "x\\|y z w v");
  assert.equal(markdownEscape("a\r\n\r\nb"), "a  b");
  assert.equal(singleLine("a\r\nb"), "a b");
}
