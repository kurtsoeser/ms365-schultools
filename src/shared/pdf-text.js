/**
 * PDF-Text extrahieren (pdf.js aus vendor/, lazy – siehe vendor-libs.js).
 * Für WebUntis Class_*.pdf u. a.
 */
(function () {
    'use strict';

    let loadPromise = null;

    function ensurePdfJs() {
        if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
        if (loadPromise) return loadPromise;
        // Dynamischer Import: funktioniert als klassisches Skript (Pfad relativ
        // zu dieser Datei) und im Vite-Bundle.
        loadPromise = import('./vendor-libs.js').then(function (vendor) {
            return vendor.loadPdfJs();
        });
        return loadPromise;
    }

    /**
     * @param {ArrayBuffer|Uint8Array} data
     * @returns {Promise<{ text: string, words: { str: string, x: number, y: number }[] }>}
     */
    function extractPdfContent(data) {
        return ensurePdfJs().then(function (pdfjsLib) {
            const buf = data instanceof Uint8Array ? data : new Uint8Array(data);
            // isEvalSupported:false – pdf.js 3.x würde sonst Code aus der PDF-Datei
            // per `new Function` bauen (CVE-2024-4367); wir lesen nur Text.
            return pdfjsLib.getDocument({ data: buf, isEvalSupported: false }).promise.then(function (pdf) {
                const pagePromises = [];
                for (let p = 1; p <= pdf.numPages; p++) {
                    const pageNum = p;
                    pagePromises.push(
                        pdf.getPage(p).then(function (page) {
                            return page.getTextContent().then(function (content) {
                                const words = [];
                                const lines = [];
                                (content.items || []).forEach(function (item) {
                                    const str = String(item.str || '').trim();
                                    if (!str) return;
                                    const tr = item.transform || [1, 0, 0, 1, 0, 0];
                                    const x = Number(tr[4]) || 0;
                                    const y = Number(tr[5]) || 0;
                                    words.push({ str: str, x: x, y: y, page: pageNum });
                                    lines.push(str);
                                });
                                return { words: words, text: lines.join('\n') };
                            });
                        })
                    );
                }
                return Promise.all(pagePromises).then(function (pages) {
                    const words = [];
                    const textParts = [];
                    pages.forEach(function (pg) {
                        words.push.apply(words, pg.words);
                        textParts.push(pg.text);
                    });
                    return { text: textParts.join('\n'), words: words };
                });
            });
        });
    }

    /**
     * @param {File} file
     */
    function extractPdfFile(file) {
        return new Promise(function (resolve, reject) {
            if (!file) {
                reject(new Error('Keine Datei'));
                return;
            }
            const reader = new FileReader();
            reader.onload = function (e) {
                extractPdfContent(e.target.result).then(resolve, reject);
            };
            reader.onerror = function () {
                reject(new Error('PDF konnte nicht gelesen werden'));
            };
            reader.readAsArrayBuffer(file);
        });
    }

    window.ms365PdfText = {
        ensurePdfJs: ensurePdfJs,
        extractPdfContent: extractPdfContent,
        extractPdfFile: extractPdfFile
    };
})();
