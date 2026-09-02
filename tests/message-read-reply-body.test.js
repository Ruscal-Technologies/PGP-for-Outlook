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
  // Good enough for these tests: split `<head>...</head>` and `<body>...</body>`
  // fragments out of the HTML string, derive textContent by stripping tags,
  // and expose head.querySelectorAll('style') the way a real Document does.
  global.DOMParser = class {
    parseFromString(html) {
      const bodyMatch = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
      const innerHTML = bodyMatch ? bodyMatch[1] : html;
      const textContent = stripTags(innerHTML);

      const headMatch = /<head[^>]*>([\s\S]*)<\/head>/i.exec(html);
      const styleTags = headMatch
        ? [...headMatch[1].matchAll(/<style[^>]*>[\s\S]*?<\/style>/gi)].map(m => ({ outerHTML: m[0] }))
        : [];

      return {
        head: { querySelectorAll: (sel) => (sel === 'style' ? styleTags : []) },
        body: { innerHTML, textContent },
      };
    }
  };
}

let buildQuotedReplyHtml;
let REPLY_TRUNCATION_NOTICE;

beforeEach(async () => {
  global.Office = { onReady: () => {} };
  installDomParserStub();
  ({ buildQuotedReplyHtml, REPLY_TRUNCATION_NOTICE } = await import('../web/MessageRead.js'));
});

describe('buildQuotedReplyHtml', () => {
  it('wraps a short HTML body unchanged, in the <div> quote style, and reports truncated:false', () => {
    const html = '<html><body><p>Hi there</p></body></html>';
    const { html: result, truncated } = buildQuotedReplyHtml(html, true, 'Alice', 'Jan 1, 2026');

    expect(result).toContain('<div style="border-left:2px solid #888;padding-left:8px;margin-left:4px;">');
    expect(result).toContain('<p>Hi there</p>');
    expect(result).toContain('--- Original message from Alice on Jan 1, 2026 ---');
    expect(result.length).toBeLessThanOrEqual(31000);
    expect(truncated).toBe(false);
  });

  it('wraps a short plain-text body unchanged, in the <blockquote> quote style, escaped', () => {
    const text = 'Hi there\n<script>alert(1)</script>';
    const { html: result, truncated } = buildQuotedReplyHtml(text, false, 'Bob', '');

    expect(result).toContain('<blockquote style="border-left:2px solid #888;padding-left:8px;margin-left:4px;">');
    expect(result).toContain('Hi there<br>&lt;script&gt;alert(1)&lt;/script&gt;');
    expect(result).not.toContain('<script>');
    expect(truncated).toBe(false);
  });

  it('escapes sender name and sent date in the quote header', () => {
    const { html: result } = buildQuotedReplyHtml('hi', false, '<b>Eve</b>', '<i>today</i>');
    expect(result).toContain('--- Original message from &lt;b&gt;Eve&lt;/b&gt; on &lt;i&gt;today&lt;/i&gt; ---');
  });

  it('falls back to a plain-text quote when the HTML body would exceed maxLength, and reports truncated:true', () => {
    const bigHtml = `<html><body><p style="color:red">${'a'.repeat(500)}</p></body></html>`;
    const { html: result, truncated } = buildQuotedReplyHtml(bigHtml, true, 'Alice', '', 300);

    // Must not use the HTML wrapper/style-laden markup that overflowed.
    expect(result).toContain('<blockquote');
    expect(result).not.toContain('<p style="color:red">');
    expect(result.length).toBeLessThanOrEqual(300);
    expect(truncated).toBe(true);
  });

  it('truncates and appends a notice when even the plain-text quote exceeds maxLength', () => {
    const bigText = 'a'.repeat(500);
    const { html: result, truncated } = buildQuotedReplyHtml(bigText, false, '', '', 300);

    expect(result).toContain('[Original message truncated — too large to quote in full]');
    expect(result.length).toBeLessThanOrEqual(300);
    expect(truncated).toBe(true);
  });

  it('exports REPLY_TRUNCATION_NOTICE matching the notice actually appended on truncation', () => {
    expect(REPLY_TRUNCATION_NOTICE).toBe('<br><em>[Original message truncated — too large to quote in full]</em>');

    const { html: result, truncated } = buildQuotedReplyHtml('a'.repeat(500), false, '', '', 300);
    expect(result).toContain(REPLY_TRUNCATION_NOTICE);
    expect(truncated).toBe(true);

    const { html: shortResult, truncated: shortTruncated } = buildQuotedReplyHtml('hi', false, '', '');
    expect(shortResult).not.toContain(REPLY_TRUNCATION_NOTICE);
    expect(shortTruncated).toBe(false);
  });

  it('reports truncated:false even when the decrypted message itself legitimately contains the notice text (no false positive)', () => {
    // e.g. quoting an earlier reply that itself got truncated -- the literal
    // notice text can appear in a perfectly normal-sized message. truncated
    // must reflect whether *this* call had to shorten anything, not whether
    // the substring happens to appear somewhere in the output.
    const text = `Earlier in this thread: ${REPLY_TRUNCATION_NOTICE}`;
    const { truncated } = buildQuotedReplyHtml(text, false, '', '');
    expect(truncated).toBe(false);
  });

  it('produces a result that never exceeds maxLength for large HTML input', () => {
    const bigHtml = `<html><body>${'<div>x</div>'.repeat(2000)}</body></html>`;
    const { html: result, truncated } = buildQuotedReplyHtml(bigHtml, true, 'Alice', 'today', 300);

    expect(result.length).toBeLessThanOrEqual(300);
    expect(truncated).toBe(true);
  });

  it('never exceeds maxLength even when maxLength is smaller than the fixed wrap+notice overhead', () => {
    const { html: result } = buildQuotedReplyHtml('a'.repeat(500), false, '', '', 50);

    expect(result.length).toBeLessThanOrEqual(50);
  });

  it('carries <style> rules from <head> into the quote, so paragraph spacing renders like the preview/pop-out', () => {
    // Mirrors Outlook Desktop's Word-based HTML export, which relies on a
    // <style> rule to collapse paragraph margins -- without it, the reply
    // compose window falls back to default browser <p> margins and grows
    // extra blank lines the decrypt preview/pop-out never showed.
    const html = '<html><head><style>p.MsoNormal { margin: 0; }</style></head>' +
      '<body><p class="MsoNormal">Line 1</p><p class="MsoNormal">Line 2</p></body></html>';
    const { html: result } = buildQuotedReplyHtml(html, true, '', '');

    expect(result).toContain('<style>p.MsoNormal { margin: 0; }</style>');
    expect(result).toContain('<p class="MsoNormal">Line 1</p>');
    // The style block must precede the quoted content it applies to.
    expect(result.indexOf('<style>')).toBeLessThan(result.indexOf('Line 1'));
  });

  it('does not reintroduce a nested <head> element (Office rejects nested <html> tags in htmlBody)', () => {
    const html = '<html><head><style>body { color: red; }</style></head><body><p>Hi</p></body></html>';
    const { html: result } = buildQuotedReplyHtml(html, true, '', '');

    expect(result).not.toContain('<head>');
    expect(result).not.toContain('<html>');
  });

  it('produces an empty style prefix when the HTML has no <head>/<style>', () => {
    const { html: result } = buildQuotedReplyHtml('<div>Hi there</div>', true, '', '');
    expect(result).toContain('<div>Hi there</div>');
    expect(result).not.toContain('<style>');
  });
});
