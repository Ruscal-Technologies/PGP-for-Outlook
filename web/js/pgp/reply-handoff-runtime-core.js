/**
 * reply-handoff-runtime-core.js
 * Shared reply-handoff BroadcastChannel listener + armor-splice logic, used
 * by both MessageCompose.js's own listener (armed when the user manually
 * opens the pane) and web/ReplyHandoffPane.js (a dedicated minimal pane,
 * opened automatically via an InsightMessage notification from the
 * OnNewMessageCompose event handler on hosts where BroadcastChannel isn't
 * available in that handler's own runtime — see #22). Extracted so both
 * callers share the exact same splice logic, provably, not by copy-paste.
 *
 * Standalone except for the two other standalone pgp/ modules it imports —
 * no Office.js state is cached at module scope, so this is safe to import
 * from any page/runtime that has Office.context available.
 */

import { formatDecryptedContentAsHtml } from './quoted-content.js';
import { getReplyHandoffChannelName } from './reply-handoff-channel.js';

// Matches (with margin) MessageRead.js's REPLY_HANDOFF_TIMEOUT_MS, so a
// reply window that's genuinely waiting on a handoff isn't cut off before
// the sender itself gives up and falls back.
const REPLY_HANDOFF_LISTEN_TIMEOUT_MS = 12000;

function getComposeTypeAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getComposeTypeAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value.composeType);
      else reject(new Error(result.error.message));
    });
  });
}

function getBodyAsync(coercionType) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      coercionType,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error.message));
      }
    );
  });
}

function setBodyHtmlAsync(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error.message));
      }
    );
  });
}

// Internal splitting delimiter for stripPgpArmorBlock() — a Private Use Area
// character sequence that can never appear in real email text, and survives
// HTML (de)serialization unescaped (no &, <, >, or ").
const ARMOR_SPLICE_MARKER_BASE = '__PGP_ARMOR_SPLICE__';

/**
 * Returns a marker guaranteed not to already appear in `html` — the input is
 * attacker-influenceable PGP message content, so the base marker alone can't
 * be assumed unique, however unlikely a literal collision is in practice.
 * Exported for tests only, so the collision path can be exercised directly
 * (constructing a real collision from outside would otherwise require
 * hardcoding the exact marker text, including its non-printing characters).
 */
export function pickSpliceMarker(html) {
  let marker = ARMOR_SPLICE_MARKER_BASE;
  for (let i = 0; html.includes(marker); i++) {
    marker = `${ARMOR_SPLICE_MARKER_BASE}${i}`;
  }
  return marker;
}

/**
 * Locates the PGP armor block (`-----BEGIN PGP MESSAGE-----` through
 * `-----END PGP MESSAGE-----`, inclusive) inside an HTML body string and
 * splits the HTML around it, so the caller can splice something else in at
 * that location.
 *
 * The armor block may be split across multiple sibling text nodes/lines the
 * way Outlook's native reply-quoting renders a quoted message (including a
 * `<pre>`-wrapped block — this add-in's own setBodyAsync() in
 * MessageCompose.js sends the armor that way, so a reply to one of its own
 * messages commonly quotes it back inside a `<pre>`). This walks a detached
 * DOM the same way MessageRead.js's extractArmorFromHtml() does (same
 * BLOCK-element/`<br>`/`<pre>` handling), but additionally tracks which text
 * node each character came from, so the located range can be mapped back to
 * specific nodes and removed — rather than only extracted, like
 * extractArmorFromHtml() does.
 *
 * Removal itself works by replacing the range with a unique marker inside
 * the DOM, serializing once, then splitting the resulting HTML string on
 * that marker — this guarantees valid, unmangled surrounding HTML regardless
 * of how many nodes the range spans, without needing to reconstruct partial
 * DOM structure by hand.
 *
 * @param {string} html
 * @returns {{found: boolean, before?: string, after?: string}}
 */
export function stripPgpArmorBlock(html) {
  const div = document.createElement('div');
  div.innerHTML = html;

  const BLOCK = new Set([
    'div', 'p', 'li', 'blockquote', 'tr',
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'article', 'section', 'header', 'footer', 'html', 'body',
  ]);
  const SKIP = new Set(['style', 'script', 'head', 'title', 'noscript']);

  // { node, start, end } — [start, end) is this node's exact character range
  // within `flat`, always equal to flat.slice(start, end) === node.textContent.
  const segments = [];
  let flat = '';

  function walk(node) {
    if (node.nodeType === Node.TEXT_NODE) {
      const start = flat.length;
      flat += node.textContent;
      segments.push({ node, start, end: flat.length });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const tag = node.tagName.toLowerCase();
    if (SKIP.has(tag)) return;
    if (tag === 'br') { flat += '\n'; return; }
    if (tag === 'pre') {
      flat += '\n';
      const start = flat.length;
      flat += node.textContent;
      segments.push({ node, start, end: flat.length });
      flat += '\n';
      return;
    }
    for (const child of Array.from(node.childNodes)) walk(child);
    if (BLOCK.has(tag)) flat += '\n';
  }
  walk(div);

  const beginIdx = flat.indexOf('-----BEGIN PGP MESSAGE-----');
  if (beginIdx === -1) return { found: false };
  const endMarkerIdx = flat.indexOf('-----END PGP MESSAGE-----', beginIdx);
  if (endMarkerIdx === -1) return { found: false };
  const endIdx = endMarkerIdx + '-----END PGP MESSAGE-----'.length;

  const marker = pickSpliceMarker(html);
  let markerPlaced = false;
  for (const seg of segments) {
    if (seg.end <= beginIdx || seg.start >= endIdx) continue; // no overlap

    const localBegin = Math.max(0, beginIdx - seg.start);
    const localEnd = Math.min(seg.end - seg.start, endIdx - seg.start);
    const text = seg.node.textContent;
    seg.node.textContent = text.slice(0, localBegin) + (markerPlaced ? '' : marker) + text.slice(localEnd);
    markerPlaced = true;
  }

  const spliced = div.innerHTML;
  const parts = spliced.split(marker);
  // Anything other than exactly 2 parts means the marker either never made
  // it into the serialized output, or (despite pickSpliceMarker's check
  // against the raw input) something produced more copies of it than
  // expected -- either way, splicing on an assumption that doesn't hold
  // would corrupt the reply body, so bail out the same safe way as "not
  // found" rather than guess.
  if (parts.length !== 2) return { found: false };
  const [before, after] = parts;
  return { found: true, before, after };
}

/**
 * Reads the current (Outlook-native-quoted) body, strips out the PGP armor
 * block Outlook quoted from the original encrypted message, and splices the
 * decrypted plaintext in at that location.
 *
 * Does not itself show any status — the caller decides when to surface
 * `message` (see armReplyHandoffListener): a failed attempt here can be
 * retried on a later re-broadcast, and showing a fresh warning on every one
 * of those (up to ~25 times over the retry window) would flicker the status
 * bar rather than inform the user.
 *
 * @param {string} text - Decrypted payload from MessageRead.js
 * @param {boolean} isHtml - True when the payload is HTML
 * @returns {Promise<{success: boolean, message: string}>} `success` is true
 *   only if the splice actually wrote; never guesses at a partial edit.
 */
async function applyReplyHandoff(text, isHtml) {
  try {
    const bodyHtml = await getBodyAsync(Office.CoercionType.Html);
    const { found, before, after } = stripPgpArmorBlock(bodyHtml);
    if (!found) {
      return { success: false, message: 'Could not find the encrypted message in this reply to replace — please verify the body before sending.' };
    }
    const formattedContent = formatDecryptedContentAsHtml(text, isHtml);
    await setBodyHtmlAsync(before + formattedContent + after);
    return { success: true, message: 'Decrypted message inserted into this reply.' };
  } catch (e) {
    return { success: false, message: `Could not automatically insert the decrypted message into this reply: ${e.message} — please verify the body before sending.` };
  }
}

/**
 * Arms a BroadcastChannel listener for the large-message reply-handoff
 * splice (see MessageRead.js's openNativeReplyWithHandoff). Resolves
 * immediately once armed (or once it's determined there's nothing to arm —
 * not a reply, no scoping ID, BroadcastChannel unavailable) — it does NOT
 * wait for the handoff to actually settle, so callers that only need the
 * listener running (MessageCompose.js's own task pane) can await this and
 * continue immediately, matching this function's pre-extraction behavior
 * exactly. Callers that need to react once the handoff settles (e.g.
 * web/ReplyHandoffPane.js, which auto-closes itself afterward) pass
 * `onSettled`, invoked once with `{success, message}` when the handoff acks,
 * times out, or (defensively) if applyReplyHandoff throws unexpectedly.
 *
 * @param {object} opts
 * @param {boolean} opts.has110 - Mailbox 1.10 supported (getComposeTypeAsync)
 * @param {boolean} opts.has114 - Mailbox 1.14 supported (item.inReplyTo fallback)
 * @param {(message: string, type: string) => void} [opts.onStatus] - Called
 *   for user-visible status, same points a caller's own showStatus() would be.
 * @param {(result: {success: boolean, message: string}) => void} [opts.onSettled]
 */
export async function armReplyHandoffListener({ has110, has114, onStatus, onSettled } = {}) {
  if (typeof BroadcastChannel !== 'function') {
    if (onSettled) onSettled({ success: false, message: 'BroadcastChannel is not available in this window.' });
    return;
  }

  if (has110) {
    // Office.MailboxEnums.ComposeType has exactly three values -- Reply,
    // NewMail, Forward -- Reply All is NOT a distinct value; getComposeTypeAsync
    // reports 'reply' for both.
    const composeType = await getComposeTypeAsync().catch((e) => {
      console.error('Reply handoff: getComposeTypeAsync failed', e);
      return null; // unknown -- fall through and listen anyway, see below
    });
    if (composeType !== null && composeType !== Office.MailboxEnums.ComposeType.Reply) {
      console.log('Reply handoff: not listening -- composeType is', composeType);
      if (onSettled) onSettled({ success: false, message: 'This isn\'t a reply window.' });
      return;
    }
  }
  // _has110 false: can't confirm compose type, so listen anyway (broader
  // exposure on older hosts only, still bounded by the timeout below).

  // Prefer conversationId, but fall back to inReplyTo (Mailbox 1.14 -- the
  // internet message ID of the message being replied to) when it's missing.
  // MessageRead.js derives the matching scoping ID the same way, from its
  // own item.conversationId / item.internetMessageId. See
  // openNativeReplyWithHandoff.
  const conversationId = Office.context.mailbox.item.conversationId;
  const inReplyTo = has114 ? Office.context.mailbox.item.inReplyTo : undefined;
  const scopingId = conversationId || inReplyTo;
  if (!scopingId) {
    // No way to scope the channel to this specific conversation/message at
    // all -- falling back to the shared base channel name would let any
    // same-origin page listen for this (and every other) large reply's
    // decrypted plaintext. Treat this the same as "handoff unavailable" and
    // don't listen at all; MessageRead.js makes the matching decision on its
    // side (see openNativeReplyWithHandoff).
    console.log('Reply handoff: not listening -- no conversationId or inReplyTo available');
    if (onSettled) onSettled({ success: false, message: 'No conversation/message ID available to scope the handoff.' });
    return;
  }

  const channelName = getReplyHandoffChannelName(scopingId);
  let channel;
  try {
    channel = new BroadcastChannel(channelName);
  } catch (e) {
    console.error('Reply handoff: BroadcastChannel construction failed', e);
    if (onSettled) onSettled({ success: false, message: 'Could not set up the reply handoff listener.' });
    return;
  }
  console.log('Reply handoff: listening on', channelName);

  let consumed = false;
  const idleTimer = setTimeout(() => {
    if (!consumed) {
      channel.close();
      if (onSettled) onSettled({ success: false, message: 'Timed out waiting for a reply handoff.' });
    }
  }, REPLY_HANDOFF_LISTEN_TIMEOUT_MS);

  let handoffInFlight = false;
  // Shown once, not on every retry: a splice that fails once (e.g. the
  // armor block sits in a DOM shape stripPgpArmorBlock() doesn't recognize)
  // fails identically on every re-broadcast, and MessageRead.js retries
  // every ~400ms for up to ~10s -- re-showing the same warning that often
  // would flicker the status bar instead of informing the user.
  let hasShownFailureWarning = false;

  channel.onmessage = async (event) => {
    const data = event.data;
    if (!data || data.type !== 'pgp-reply-handoff' || consumed || handoffInFlight) return;
    handoffInFlight = true;

    const { success, message } = await applyReplyHandoff(data.text, data.isHtml);
    console.log('Reply handoff: applyReplyHandoff result', { success, message });
    // Only ack on confirmed success -- an ack that arrives despite a failed
    // splice would make MessageRead.js treat this as done and never trigger
    // its own fallback, leaving the user with a still-armored body and no
    // backup window. Staying silent here lets that timeout-based fallback
    // fire instead.
    if (success) {
      consumed = true;
      clearTimeout(idleTimer);
      channel.postMessage({ type: 'pgp-reply-handoff-ack', token: data.token });
      channel.close();
      if (onStatus) onStatus(message, 'success');
      if (onSettled) onSettled({ success: true, message });
      return;
    }
    if (!hasShownFailureWarning) {
      hasShownFailureWarning = true;
      if (onStatus) onStatus(message, 'warning');
    }
    handoffInFlight = false;
  };
}
