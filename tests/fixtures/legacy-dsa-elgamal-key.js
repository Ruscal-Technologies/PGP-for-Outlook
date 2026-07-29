// A real, freshly-generated (throwaway, no security value) legacy PGP key —
// DSA-1024 primary + ElGamal-1024 encryption subkey, self-signed with SHA-256
// (this is GnuPG 2.4's default; genuinely reproducing an old SHA-1
// self-signature would need forging packets or an ancient GnuPG version, so
// that specific retry path in pgp-core.js's _isLegacySelfSigError handling
// remains untested — see CLAUDE.md's test-suite section).
//
// This still gives real coverage of the ElGamal/DSA "weak key" detection and
// rejection paths (_isWeakKeyError, hasWeakEncryptionKey) and the modern-
// subkey-augmentation flow (hasModernSubkeys / addModernSubkeys), none of
// which had any fixture to run against before.
//
// Generated with: gpg --batch --gen-key (Key-Type: DSA/1024, Subkey-Type:
// ELG-E/1024), then `gpg --armor --export[-secret-keys] legacy@example.com`.

export const LEGACY_PASSPHRASE = 'legacypassphrase';

export const LEGACY_PUBLIC_KEY = `-----BEGIN PGP PUBLIC KEY BLOCK-----

mQGiBGppIxsRBADXza+rpr75OX1zQCqCNaSjnKHv9tulwtgw3gjxHSpwHApbfAaV
dAuXHElawBLgwHSV2FPCvmbLPYz9WOEc1H34rlVCAFMxgFWCZD8kU6/E7gGkbMnv
gbvwJ7QldjLRfS3doPZGu47JNitDaN7fAUU0/XkJKc9dYkdb8dWPAOMFLwCgw/6y
2n/eHtIdUkkXSv0HNuyllN8EAKQWsVXbwlkd7RVKnk35i9Iow8AG1+wYeQmp7fVK
3qRmvDZxYEigwqgqe9LuiJtVdNT+CVIIuJIszbj2VNJ0AXsIVDCJd8ntiIa5JCFP
Tbifn9xLSb/GRjePD2sXu5aoiKDymUklp28sRzU9UYKp2Nzb1ahJl6NB4QQRUdEe
J5hdA/9fqt9mjQZeKBuHVDa+cJ7LJn8bWylgScI9rC4x+koQYAIB39dT3EGXd+Ab
RgyyWHGRa7gnrK/AgtZAJANZKrklwgHgKWVGMzZlBO9TsN8urw4qc4KOSTcW0l9P
2Qg3MWuPyGyJE/Ug1l5qhIweSPDnsRMRq8NeX1IqauNxAvIKx7QgTGVnYWN5IFRl
c3QgPGxlZ2FjeUBleGFtcGxlLmNvbT6IeAQTEQIAOBYhBDfTl7ONO7ox6h4qKcVl
zrBS6Qn0BQJqaSMbAhsjBQsJCAcCBhUKCQgLAgQWAgMBAh4BAheAAAoJEMVlzrBS
6Qn0JosAn3JSuCEBOyst2H8EI/bmTvPMvGMlAJ9C8C2XTEWjNQ7aUGakDvqpv/vH
wLkBDQRqaSMbEAQAzhv9qwWJbzD3yNYGDkc9WGHLsr87dhw7t2GlX/W1646yTTL4
aa+IPe67466+vZgkJ7MUMnZjsJ1cqDlYsfXdUiIvIG04eMVGt8ylLrrilw8DWkvB
3ScFAVde2RctH2OJGsHwA2E0dMeYTE87zBP/9NGGyuwuyRcGDrWgNNgZBv8AAwUE
AK+AWb7QM/Nv37Yr9YXGlmpCUCnIrLWD21+ibyBLMCzd/OZjZPEWaSLv4TxVINOZ
4ph6MyG++Y3aQ8tpCsW79bo0pH42FTFDQmX01K9hm71htySGepqK0mhkz2r416QB
1mC+1Z6MvqKvdDEk5kwp7TMQoVfWVuP59SB1ZqWPlmPNiGAEGBECACAWIQQ305ez
jTu6MeoeKinFZc6wUukJ9AUCamkjGwIbDAAKCRDFZc6wUukJ9B9vAJ40Tyz6MgD/
KbmiHSKWp0AswqVC0gCdErEa92vv6EweGFwbDWrd3eUK8m0=
=6CCr
-----END PGP PUBLIC KEY BLOCK-----
`;

export const LEGACY_PRIVATE_KEY = `-----BEGIN PGP PRIVATE KEY BLOCK-----

lQHpBGppIxsRBADXza+rpr75OX1zQCqCNaSjnKHv9tulwtgw3gjxHSpwHApbfAaV
dAuXHElawBLgwHSV2FPCvmbLPYz9WOEc1H34rlVCAFMxgFWCZD8kU6/E7gGkbMnv
gbvwJ7QldjLRfS3doPZGu47JNitDaN7fAUU0/XkJKc9dYkdb8dWPAOMFLwCgw/6y
2n/eHtIdUkkXSv0HNuyllN8EAKQWsVXbwlkd7RVKnk35i9Iow8AG1+wYeQmp7fVK
3qRmvDZxYEigwqgqe9LuiJtVdNT+CVIIuJIszbj2VNJ0AXsIVDCJd8ntiIa5JCFP
Tbifn9xLSb/GRjePD2sXu5aoiKDymUklp28sRzU9UYKp2Nzb1ahJl6NB4QQRUdEe
J5hdA/9fqt9mjQZeKBuHVDa+cJ7LJn8bWylgScI9rC4x+koQYAIB39dT3EGXd+Ab
RgyyWHGRa7gnrK/AgtZAJANZKrklwgHgKWVGMzZlBO9TsN8urw4qc4KOSTcW0l9P
2Qg3MWuPyGyJE/Ug1l5qhIweSPDnsRMRq8NeX1IqauNxAvIKx/4HAwIgfcDfLuaG
7/dK0zo5kvO79zzaCJ/oaGj3T4zmVHHrNL4vDNLfM4p0Yc58nQDrjM6IWmKMrV8+
xP6D3JQi704ZGHn4tCBMZWdhY3kgVGVzdCA8bGVnYWN5QGV4YW1wbGUuY29tPoh4
BBMRAgA4FiEEN9OXs407ujHqHiopxWXOsFLpCfQFAmppIxsCGyMFCwkIBwIGFQoJ
CAsCBBYCAwECHgECF4AACgkQxWXOsFLpCfQmiwCfclK4IQE7Ky3YfwQj9uZO88y8
YyUAn0LwLZdMRaM1DtpQZqQO+qm/+8fAnQFgBGppIxsQBADOG/2rBYlvMPfI1gYO
Rz1YYcuyvzt2HDu3YaVf9bXrjrJNMvhpr4g97rvjrr69mCQnsxQydmOwnVyoOVix
9d1SIi8gbTh4xUa3zKUuuuKXDwNaS8HdJwUBV17ZFy0fY4kawfADYTR0x5hMTzvM
E//00YbK7C7JFwYOtaA02BkG/wADBQQAr4BZvtAz82/ftiv1hcaWakJQKcistYPb
X6JvIEswLN385mNk8RZpIu/hPFUg05nimHozIb75jdpDy2kKxbv1ujSkfjYVMUNC
ZfTUr2GbvWG3JIZ6morSaGTPavjXpAHWYL7Vnoy+oq90MSTmTCntMxChV9ZW4/n1
IHVmpY+WY83+BwMCeU4687kPDEr3WJBaqTZxbDbBtxL+6/EGfPelYADtfTNXzpeW
rF5Y59PYOuUBhEO+A95jBcnQwS2261I4WgIOIbZjuhS6cJVIz5K5U0aQqIhgBBgR
AgAgFiEEN9OXs407ujHqHiopxWXOsFLpCfQFAmppIxsCGwwACgkQxWXOsFLpCfQf
bwCeNE8s+jIA/ym5oh0ilqdALMKlQtIAnRKxGvdr7+hMHhhcGw1q3d3lCvJt
=YW2G
-----END PGP PRIVATE KEY BLOCK-----
`;
