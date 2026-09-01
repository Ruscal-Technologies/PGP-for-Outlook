import { describe, it, expect } from 'vitest';
import { getReplyHandoffChannelName } from '../web/js/pgp/reply-handoff-channel.js';

describe('getReplyHandoffChannelName', () => {
  it('falls back to the base name when no conversationId is given', () => {
    expect(getReplyHandoffChannelName()).toBe('pgp_reply_handoff');
    expect(getReplyHandoffChannelName('')).toBe('pgp_reply_handoff');
  });

  it('derives a distinct, deterministic name per conversationId', () => {
    const a = getReplyHandoffChannelName('conversation-A');
    const b = getReplyHandoffChannelName('conversation-B');

    expect(a).toMatch(/^pgp_reply_handoff_[0-9a-f]+$/);
    expect(b).toMatch(/^pgp_reply_handoff_[0-9a-f]+$/);
    expect(a).not.toBe(b);
    expect(getReplyHandoffChannelName('conversation-A')).toBe(a);
  });
});
