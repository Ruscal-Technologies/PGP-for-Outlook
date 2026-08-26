import { describe, it, expect, beforeEach } from 'vitest';

// MessageRead.js calls Office.onReady(...) at module load time, so it needs a
// minimal Office stub in place before import (same requirement documented in
// tests/message-read-popout.test.js). buildQuotedReplyHtml itself is pure —
// it only additionally needs a DOMParser, which Node doesn't provide, so we
// stub a minimal hand-rolled one rather than pull in jsdom (this repo's
// vitest.config.js runs with environment: 'node').
// Strips tags by re-scanning until a pass makes no further change, so nested
// constructs (e.g. `<<script>script>`) can't survive a single regex pass.
function stripTags(html) {
  let text = html;
  let prev;
  do {
    prev = text;
    text = text.replace(/<[^>]+>/g, '');
  } while (text !== prev);
  return text;
}

function installDomParserStub() {
  // Good enough for these tests: split a `<body>...</body>`-wrapped fragment
  // out of the HTML string, and derive textContent by stripping tags.
  global.DOMParser = class {
    parseFromString(html) {
      const match = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
      const innerHTML = match ? match[1] : html;
      const textContent = stripTags(innerHTML);
      return { body: { innerHTML, textContent } };
    }
  };
}

let buildQuotedReplyHtml;

beforeEach(async () => {
  global.Office = { onReady: () => {} };
  installDomParserStub();
  ({ buildQuotedReplyHtml } = await import('../web/MessageRead.js'));
});

describe('buildQuotedReplyHtml', () => {
  it('wraps a short HTML body unchanged, in the <div> quote style', () => {
    const html = '<html><body><p>Hi there</p></body></html>';
    const result = buildQuotedReplyHtml(html, true, 'Alice', 'Jan 1, 2026');

    expect(result).toContain('<div style="border-left:2px solid #888;padding-left:8px;margin-left:4px;">');
    expect(result).toContain('<p>Hi there</p>');
    expect(result).toContain('--- Original message from Alice on Jan 1, 2026 ---');
    expect(result.length).toBeLessThanOrEqual(31000);
  });

  it('wraps a short plain-text body unchanged, in the <blockquote> quote style, escaped', () => {
    const text = 'Hi there\n<script>alert(1)</script>';
    const result = buildQuotedReplyHtml(text, false, 'Bob', '');

    expect(result).toContain('<blockquote style="border-left:2px solid #888;padding-left:8px;margin-left:4px;">');
    expect(result).toContain('Hi there<br>&lt;script&gt;alert(1)&lt;/script&gt;');
    expect(result).not.toContain('<script>');
  });

  it('escapes sender name and sent date in the quote header', () => {
    const result = buildQuotedReplyHtml('hi', false, '<b>Eve</b>', '<i>today</i>');
    expect(result).toContain('--- Original message from &lt;b&gt;Eve&lt;/b&gt; on &lt;i&gt;today&lt;/i&gt; ---');
  });

  it('falls back to a plain-text quote when the HTML body would exceed maxLength', () => {
    const bigHtml = `<html><body><p style="color:red">${'a'.repeat(500)}</p></body></html>`;
    const result = buildQuotedReplyHtml(bigHtml, true, 'Alice', '', 300);

    // Must not use the HTML wrapper/style-laden markup that overflowed.
    expect(result).toContain('<blockquote');
    expect(result).not.toContain('<p style="color:red">');
    expect(result.length).toBeLessThanOrEqual(300);
  });

  it('truncates and appends a notice when even the plain-text quote exceeds maxLength', () => {
    const bigText = 'a'.repeat(500);
    const result = buildQuotedReplyHtml(bigText, false, '', '', 300);

    expect(result).toContain('[Original message truncated — too large to quote in full]');
    expect(result.length).toBeLessThanOrEqual(300);
  });

  it('produces a result that never exceeds maxLength for large HTML input', () => {
    const bigHtml = `<html><body>${'<div>x</div>'.repeat(2000)}</body></html>`;
    const result = buildQuotedReplyHtml(bigHtml, true, 'Alice', 'today', 300);

    expect(result.length).toBeLessThanOrEqual(300);
  });

  it('never exceeds maxLength even when maxLength is smaller than the fixed wrap+notice overhead', () => {
    const result = buildQuotedReplyHtml('a'.repeat(500), false, '', '', 50);

    expect(result.length).toBeLessThanOrEqual(50);
  });
});
