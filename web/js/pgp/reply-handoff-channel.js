/**
 * reply-handoff-channel.js
 * Derives the BroadcastChannel name for the large-message Reply/Reply All
 * native-reply handoff between MessageRead.js (sender) and
 * MessageCompose.js (receiver). Standalone, no imports — both files import
 * this same function so their derivation can never drift apart (a mismatch
 * would silently break every large-message reply, with no error).
 *
 * Uses the conversation ID directly (URI-encoded), not a hash of it: a fixed
 * 32-bit (or any bounded-width) hash has a real collision probability at
 * scale, which would make two unrelated conversations share a channel name —
 * cross-talk that could splice one conversation's decrypted plaintext into
 * another's reply, or leak it there. The conversation ID itself isn't secret
 * (anyone with access to the message already has it), so encoding it
 * directly costs nothing and can't collide.
 *
 * NOT a secrecy boundary: BroadcastChannel is same-origin-only already, and
 * the conversation ID isn't a secret — a co-located attacker who already
 * knows which conversation to target could reconstruct this name outright.
 * Its purpose is narrowing blast radius: a listener needs to already know
 * which conversation to target, rather than one fixed name every large reply
 * in the product shares. The real mitigations for a same-origin eavesdropper
 * are MessageCompose.js's compose-type gating and bounded listener lifetime
 * (see its own comments).
 */

const BASE_CHANNEL_NAME = 'pgp_reply_handoff';

/**
 * @param {string} [conversationId] - Office.context.mailbox.item.conversationId
 * @returns {string}
 */
export function getReplyHandoffChannelName(conversationId) {
  if (!conversationId) return BASE_CHANNEL_NAME;
  return `${BASE_CHANNEL_NAME}_${encodeURIComponent(conversationId)}`;
}
