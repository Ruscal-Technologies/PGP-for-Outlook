import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';

// Regression test tying together the three places that must agree on the
// literal string "encryptSend": manifest.xml's two
// composeEncryptSendTaskPaneUrl resources (?mode=encryptSend),
// manifest.json's ComposeEncryptSendTaskPane runtime code.page URL, and
// MessageCompose.js's own comparison against
// URLSearchParams(window.location.search).get('mode'). A typo in any one of
// these would silently break the Encrypt & Send button (it would open the
// pane but never trigger the forced flow) with nothing failing loudly.
// Follows the read-only source-text-comparison pattern in
// tests/message-read-reply-button-wiring.test.js -- no imports of the actual
// modules, no Office/document stubbing needed.
const manifestXmlPath = fileURLToPath(new URL('../manifest/manifest.xml', import.meta.url));
const manifestJsonPath = fileURLToPath(new URL('../manifest/manifest.json', import.meta.url));
const jsPath = fileURLToPath(new URL('../web/MessageCompose.js', import.meta.url));

const manifestXmlSource = readFileSync(manifestXmlPath, 'utf8');
const manifestJsonSource = readFileSync(manifestJsonPath, 'utf8');
const jsSource = readFileSync(jsPath, 'utf8');

describe('Encrypt & Send ?mode=encryptSend wiring agreement', () => {
  it('manifest.xml declares ?mode=encryptSend at least twice (once per VersionOverrides block)', () => {
    const matches = manifestXmlSource.match(/\?mode=encryptSend/g);
    expect(matches).not.toBeNull();
    expect(matches.length).toBeGreaterThanOrEqual(2);
  });

  it('manifest.json declares ?mode=encryptSend for its ComposeEncryptSendTaskPane runtime', () => {
    expect(manifestJsonSource).toContain('?mode=encryptSend');
  });

  it("MessageCompose.js compares the parsed mode query param against the literal 'encryptSend'", () => {
    expect(jsSource).toContain(
      "new URLSearchParams(window.location.search).get('mode') === 'encryptSend'"
    );
  });
});
