/**
 * PDF-Text extrahieren (pdf.js von CDN, lazy).
 * Für WebUntis Class_*.pdf u. a.
 */
(function () {
    'use strict';

    const PDFJS_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    const PDFJS_WORKER = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    let loadPromise = null;

    function ensurePdfJs() {
        if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
        if (loadPromise) return loadPromise;
        loadPromise = new Promise(function (resolve, reject) {
            const s = document.createElement('script');
            s.src = PDFJS_CDN;
            s.async = true;
            s.onload = function () {
                const lib = window.pdfjsLib;
                if (!lib) {
                    reject(new Error('pdf.js nicht verfügbar'));
                    return;
                }
                lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
                resolve(lib);
            };
            s.onerror = function () {
                reject(new Error('pdf.js konnte nicht geladen werden'));
            };
            document.head.appendChild(s);
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
            return pdfjsLib.getDocument({ data: buf }).promise.then(function (pdf) {
                const pagePromises = [];
                for (let p = 1; p <= pdf.numPages; p++) {
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
                                    words.push({ str: str, x: x, y: y });
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
