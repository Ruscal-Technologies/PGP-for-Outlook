/**
 * reply-handoff-channel.js
 * Derives the BroadcastChannel name for the large-message Reply/Reply All
 * native-reply handoff between MessageRead.js (sender) and
 * MessageCompose.js (receiver). Standalone, no imports — both files import
 * this same function so their derivation can never drift apart (a mismatch
 * would silently break every large-message reply, with no error).
 *
 * NOT a secrecy boundary: BroadcastChannel is same-origin-only already, and
 * this is a plain (non-cryptographic) hash of the conversation ID, not a
 * secret — a co-located attacker who already knows which conversation to
 * target could recompute it. Its purpose is narrowing blast radius: a
 * listener needs to already know which conversation to target, rather than
 * one fixed name every large reply in the product shares. The real
 * mitigations for a same-origin eavesdropper are MessageCompose.js's
 * compose-type gating and bounded listener lifetime (see its own comments).
 */

const BASE_CHANNEL_NAME = 'pgp_reply_handoff';

/**
 * @param {string} [conversationId] - Office.context.mailbox.item.conversationId
 * @returns {string}
 */
export function getReplyHandoffChannelName(conversationId) {
  if (!conversationId) return BASE_CHANNEL_NAME;
  return `${BASE_CHANNEL_NAME}_${hashString(conversationId)}`;
}

// FNV-1a — fast, deterministic, no crypto dependency; see module docblock
// for why this doesn't need to be cryptographically strong.
function hashString(text) {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16);
}
