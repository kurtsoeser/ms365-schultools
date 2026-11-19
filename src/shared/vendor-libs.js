/**
 * Lädt Bibliotheken, die mit der App ausgeliefert werden (Ordner vendor/, erzeugt
 * von scripts/vendor-libs.mjs aus node_modules). Es wird nichts von fremden
 * Servern nachgeladen.
 *
 * Die `new URL('…', import.meta.url)`-Literale funktionieren in beiden Welten:
 * ungebündelt (dist/src/shared/ → dist/vendor/) und im Vite-Bundle (Vite schreibt
 * die Datei nach assets/ und setzt die URL passend um).
 */
const MSAL_BROWSER_URL = new URL('../../vendor/msal-browser/msal-browser.min.js', import.meta.url).href;
const JSZIP_URL = new URL('../../vendor/jszip/jszip.min.js', import.meta.url).href;
const MERMAID_URL = new URL('../../vendor/mermaid/mermaid.min.js', import.meta.url).href;
const PDFJS_URL = new URL('../../vendor/pdfjs/pdf.min.js', import.meta.url).href;
const PDFJS_WORKER_URL = new URL('../../vendor/pdfjs/pdf.worker.min.js', import.meta.url).href;

const REGISTRY_KEY = '__ms365VendorLibLoads';

/**
 * Lädt ein Skript einmalig und liefert das globale Objekt, das es anlegt.
 * Die Registry hängt am globalen Objekt, damit gebündelte und ungebündelte
 * Kopien dieses Moduls dieselbe Bibliothek nicht doppelt laden.
 *
 * @param {string} url
 * @param {string} globalName
 * @returns {Promise<any>}
 */
export function loadVendorGlobal(url, globalName) {
    if (globalThis[globalName]) return Promise.resolve(globalThis[globalName]);
    const registry = (globalThis[REGISTRY_KEY] = globalThis[REGISTRY_KEY] || {});
    if (registry[globalName]) return registry[globalName];
    const pending = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = url;
        script.async = true;
        script.onload = () => {
            const lib = globalThis[globalName];
            if (lib) resolve(lib);
            else reject(new Error(`Bibliothek "${globalName}" wurde geladen, ist aber nicht verfügbar.`));
        };
        script.onerror = () => reject(new Error(`Bibliothek "${globalName}" konnte nicht geladen werden (${url}).`));
        document.head.appendChild(script);
    }).catch((err) => {
        // Fehlversuch nicht merken – der nächste Aufruf darf es erneut probieren.
        delete registry[globalName];
        throw err;
    });
    registry[globalName] = pending;
    return pending;
}

/** @returns {Promise<typeof import('@azure/msal-browser')>} */
export function loadMsalBrowser() {
    return loadVendorGlobal(MSAL_BROWSER_URL, 'msal');
}

/** @returns {Promise<typeof import('jszip')>} */
export function loadJsZip() {
    return loadVendorGlobal(JSZIP_URL, 'JSZip');
}

/** @returns {Promise<import('mermaid').Mermaid>} */
export function loadMermaid() {
    return loadVendorGlobal(MERMAID_URL, 'mermaid');
}

/** pdf.js inklusive Worker-Pfad. */
export async function loadPdfJs() {
    const lib = await loadVendorGlobal(PDFJS_URL, 'pdfjsLib');
    lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
    return lib;
}
