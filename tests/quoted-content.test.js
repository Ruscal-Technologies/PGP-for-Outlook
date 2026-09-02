import { describe, it, expect, beforeEach } from 'vitest';

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

let formatDecryptedContentAsHtml;
let formatDecryptedContentAsPlainTextHtml;

beforeEach(async () => {
  installDomParserStub();
  ({ formatDecryptedContentAsHtml, formatDecryptedContentAsPlainTextHtml } =
    await import('../web/js/pgp/quoted-content.js'));
});

describe('formatDecryptedContentAsHtml', () => {
  it('returns HTML content unchanged (body innerHTML) for isHtml=true', () => {
    const html = '<html><body><p>Hi there</p></body></html>';
    expect(formatDecryptedContentAsHtml(html, true)).toBe('<p>Hi there</p>');
  });

  it('carries <style> rules from <head> ahead of the body content', () => {
    const html = '<html><head><style>p.MsoNormal { margin: 0; }</style></head>' +
      '<body><p class="MsoNormal">Line 1</p></body></html>';
    const result = formatDecryptedContentAsHtml(html, true);

    expect(result).toContain('<style>p.MsoNormal { margin: 0; }</style>');
    expect(result).toContain('<p class="MsoNormal">Line 1</p>');
    expect(result.indexOf('<style>')).toBeLessThan(result.indexOf('Line 1'));
  });

  it('produces no style prefix when there is no <head>/<style>', () => {
    const result = formatDecryptedContentAsHtml('<div>Hi there</div>', true);
    expect(result).toBe('<div>Hi there</div>');
  });

  it('HTML-escapes plain text and converts \\n to <br> for isHtml=false', () => {
    const text = 'Hi there\n<script>alert(1)</script>';
    const result = formatDecryptedContentAsHtml(text, false);
    expect(result).toBe('Hi there<br>&lt;script&gt;alert(1)&lt;/script&gt;');
  });
});

describe('formatDecryptedContentAsPlainTextHtml', () => {
  it('extracts textContent (strips tags) for isHtml=true, rather than escaping raw markup', () => {
    const html = '<html><body><p style="color:red">Hi there</p></body></html>';
    const result = formatDecryptedContentAsPlainTextHtml(html, true);

    expect(result).not.toContain('<p');
    expect(result).not.toContain('style=');
    expect(result).toBe('Hi there');
  });

  it('HTML-escapes plain text and converts \\n to <br> for isHtml=false', () => {
    const text = 'Hi there\n<script>alert(1)</script>';
    const result = formatDecryptedContentAsPlainTextHtml(text, false);
    expect(result).toBe('Hi there<br>&lt;script&gt;alert(1)&lt;/script&gt;');
  });
});
