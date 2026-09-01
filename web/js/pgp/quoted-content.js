/**
 * quoted-content.js
 * Turns decrypted message content into HTML ready to embed elsewhere (a
 * reply's quoted block, a spliced-in reply body). Pure string/DOM logic only
 * — no Office.js, no OpenPGP.js, no module-level state. Standalone: imported
 * by MessageRead.js (buildQuotedReplyHtml, small/normal messages) and
 * MessageCompose.js (the native-reply armor splice, large messages), which
 * otherwise have no dependency on each other.
 */

/**
 * Formats decrypted message content as HTML ready to embed in a reply body.
 *
 * For HTML content: extracts <body> innerHTML plus any <head><style>
 * block(s). Outlook Desktop's Word-based HTML export commonly relies on a
 * <style> rule (e.g. `p.MsoNormal { margin:0 }`) to render single-spaced
 * lines correctly — dropping it (as a bare doc.body.innerHTML would) lets
 * default browser paragraph margins reappear, inserting visible blank lines
 * that don't exist wherever the full original document is rendered (the
 * decrypt preview iframe, the pop-out window).
 *
 * For plain text: HTML-escapes it and converts \n to <br>.
 *
 * @param {string} decryptedText
 * @param {boolean} decryptedIsHtml
 * @returns {string}
 */
export function formatDecryptedContentAsHtml(decryptedText, decryptedIsHtml) {
  if (!decryptedIsHtml) return escapePlainTextAsHtml(decryptedText);

  // Office rejects nested <html> tags in htmlBody, and a full document isn't
  // appropriate to splice into an existing body either — body content plus
  // any <head><style> is what both callers actually want.
  const doc = new DOMParser().parseFromString(decryptedText, 'text/html');
  const styleBlocks = doc.head
    ? Array.from(doc.head.querySelectorAll('style')).map(s => s.outerHTML).join('')
    : '';
  return styleBlocks + (doc.body ? doc.body.innerHTML : decryptedText);
}

/**
 * Plain-text rendering of decrypted content, for when formatted HTML doesn't
 * fit a size budget. For originally-HTML content this extracts textContent
 * (stripping tags) rather than treating the raw markup as literal text, so
 * the fallback never leaves visible tag soup or unbalanced tags.
 *
 * @param {string} decryptedText
 * @param {boolean} decryptedIsHtml
 * @returns {string}
 */
export function formatDecryptedContentAsPlainTextHtml(decryptedText, decryptedIsHtml) {
  if (!decryptedIsHtml) return escapePlainTextAsHtml(decryptedText);
  const doc = new DOMParser().parseFromString(decryptedText, 'text/html');
  const plainText = doc.body ? doc.body.textContent : decryptedText;
  return escapePlainTextAsHtml(plainText);
}

function escapePlainTextAsHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
}
