// wkd-client - A WKD client implementation in javascript
// Copyright (C) 2018 Wiktor Kwapisiewicz
//
// This library is free software; you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public
// License as published by the Free Software Foundation; either
// version 3.0 of the License, or (at your option) any later version.
//
// This library is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
// Lesser General Public License for more details.
//
// You should have received a copy of the GNU Lesser General Public
// License along with this library; if not, write to the Free Software
// Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA

/**
 * This class implements a client for the Web Key Directory (WKD) protocol
 * in order to lookup keys on designated servers.
 * @see https://datatracker.ietf.org/doc/draft-koch-openpgp-webkey-service/
 */
export default class WKD {
    /**
     * Initialize the WKD client
     */
    constructor() {
        this._fetch = typeof globalThis.fetch === 'function' ? globalThis.fetch : require('node-fetch');
        const { subtle } = globalThis.crypto || require('crypto').webcrypto || new (require('@peculiar/webcrypto').Crypto)();
        this._subtle = subtle;
    }

    /**
     * Search for a public key using Web Key Directory protocol.
     *
     * NOTE: this method deviates from the original wkd-client upstream
     * source (see file header) by adding an explicit 10-second timeout to
     * both fetch() calls below — the upstream code has none, which let a
     * hung network request block indefinitely (a real risk once this
     * add-in's auto-encrypt feature waits on key lookups to resolve).
     *
     * @param   {String}   options.email         User's email.
     * @returns {Uint8Array} The public key.
     * @async
     */
    async lookup(options) {
        const fetch = this._fetch;

        if (!options.email) {
            throw new Error('You must provide an email parameter!');
        }

        if (typeof options.email !== 'string' || !options.email.includes('@')) {
            throw new Error('Invalid e-mail address.');
        }

        const [localPart, domain] = options.email.split('@');
        const localPartEncoded = new TextEncoder().encode(localPart.toLowerCase());
        const localPartHashed = new Uint8Array(await this._subtle.digest('SHA-1', localPartEncoded));
        const localPartBase32 = encodeZBase32(localPartHashed);
        const localPartEscaped = encodeURIComponent(localPart);

        // CORS: WKD-compliant servers must serve Access-Control-Allow-Origin: *
        // AppDomains in the manifest does not restrict fetch() — only navigation.
        const urlAdvanced = `https://openpgpkey.${domain}/.well-known/openpgpkey/${domain}/hu/${localPartBase32}?l=${localPartEscaped}`;
        const urlDirect = `https://${domain}/.well-known/openpgpkey/hu/${localPartBase32}?l=${localPartEscaped}`;

        // AbortSignal.timeout requires Chrome 103+/Firefox 100+/Safari 16+,
        // above this add-in's documented floor — omit the signal there
        // rather than throwing synchronously and breaking WKD lookup outright.
        // A fresh signal per call: AbortSignal.timeout()'s clock starts at
        // creation, so reusing one instance across both fetches would leave
        // the direct-URL fallback with whatever budget the advanced lookup
        // didn't use (zero, if the advanced call is what timed out) instead
        // of its own full 10s.
        const fetchOptions = () => (typeof AbortSignal?.timeout === 'function'
            ? { signal: AbortSignal.timeout(10000) }
            : undefined);

        let response;
        try {
            response = await fetch(urlAdvanced, fetchOptions());
            if (response.status !== 200) {
                throw new Error('Advanced WKD lookup failed: ' + response.statusText);
            }
        } catch (err) {
            response = await fetch(urlDirect, fetchOptions());
            if (response.status !== 200) {
                throw new Error('Direct WKD lookup failed: ' + response.statusText);
            }
        }

        return new Uint8Array(await response.arrayBuffer());
    }
}

/**
 * Encode input buffer using Z-Base32 encoding.
 * See: https://tools.ietf.org/html/rfc6189#section-5.1.6
 *
 * @param {Uint8Array} data - The binary data to encode
 * @returns {String} Binary data encoded using Z-Base32.
 */
export function encodeZBase32(data) {
    if (data.length === 0) {
        return "";
    }
    const ALPHABET = "ybndrfg8ejkmcpqxot1uwisza345h769";
    const SHIFT = 5;
    const MASK = 31;
    let buffer = data[0];
    let index = 1;
    let bitsLeft = 8;
    let result = '';
    while (bitsLeft > 0 || index < data.length) {
        if (bitsLeft < SHIFT) {
            if (index < data.length) {
                buffer <<= 8;
                buffer |= data[index++] & 0xff;
                bitsLeft += 8;
            } else {
                const pad = SHIFT - bitsLeft;
                buffer <<= pad;
                bitsLeft += pad;
            }
        }
        bitsLeft -= SHIFT;
        result += ALPHABET[MASK & (buffer >> bitsLeft)];
    }
    return result;
}
