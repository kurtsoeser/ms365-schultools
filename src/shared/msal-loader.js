/**
 * Einziger Ladepfad für @azure/msal-browser. Die Bibliothek kommt aus dem
 * npm-Paket (Version in package.json) und wird mit der App ausgeliefert –
 * kein Nachladen von CDNs. Details: ./vendor-libs.js
 *
 * @returns {Promise<typeof import('@azure/msal-browser')>}
 */
export { loadMsalBrowser } from './vendor-libs.js';
