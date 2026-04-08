(function () {
    'use strict';

    const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

    ns.normalizeImportedRowKeys = function normalizeImportedRowKeys(row) {
        const out = {};
        Object.keys(row).forEach(k => {
            const nk = k.replace(/^\uFEFF/, '').trim();
            out[nk] = row[k];
        });
        return out;
    };

    ns.splitKlassenCell = function splitKlassenCell(raw) {
        const s = String(raw || '').trim();
        if (!s) return [];
        return s.split(/[,;]+/).map(c => c.trim()).filter(Boolean);
    };

    ns.applyWebuntisRows = function applyWebuntisRows(rows) {
        ns.kursteamEntryMode = 'webuntis';
        ns.rawData = rows;
        ns.filteredData = [...ns.rawData];
        ns.invalidateTeams();
        document.getElementById('totalRecords').textContent = ns.rawData.length;
        document.getElementById('uniqueSubjects').textContent = new Set(ns.rawData.map(r => r.fach).filter(f => f)).size;
        document.getElementById('uniqueTeachers').textContent = new Set(ns.rawData.map(r => r.lehrer).filter(l => l)).size;
        document.getElementById('importStats').style.display = 'block';
    };

    /**
     * Eine Zeile aus Copy-Paste: Lehrer, Fach, Klasse (Tab, mehrere Leerzeichen oder einfache Leerzeichen).
     * @returns {{ lehrer: string, fach: string, klasse: string } | null}
     */
    ns.parseWebuntisPasteLine = function parseWebuntisPasteLine(line) {
        const t = String(line || '').trim();
        if (!t || t.startsWith('#')) return null;
        let parts;
        if (t.includes('|')) {
            parts = t.split(/\s*\|\s*/).map(s => s.trim()).filter(Boolean);
        } else if (t.includes('\t')) {
            parts = t.split(/\t+/).map(s => s.trim()).filter(Boolean);
        } else if (/\s{2,}/.test(t)) {
            parts = t.split(/\s{2,}/).map(s => s.trim()).filter(Boolean);
        } else {
            parts = t.split(/\s+/).filter(Boolean);
        }
        if (parts.length < 3) return null;
        if (parts.length === 3) {
            return { lehrer: parts[0], fach: parts[1], klasse: parts[2] };
        }
        return {
            lehrer: parts[0],
            fach: parts[1],
            klasse: parts.slice(2).join(' ').trim()
        };
    };

    ns.importWebuntisFromPaste = function importWebuntisFromPaste() {
        const ta = document.getElementById('webuntisPasteInput');
        const text = ta ? ta.value : '';
        const lines = String(text).split(/\r?\n/);
        const seen = new Set();
        const rows = [];
        let id = 0;
        let skipped = 0;
        let dup = 0;
        lines.forEach(line => {
            const p = ns.parseWebuntisPasteLine(line);
            if (!p) {
                if (String(line).trim()) skipped++;
                return;
            }
            const lehrer = p.lehrer.trim();
            const fach = p.fach.trim();
            const klasse = p.klasse.trim();
            if (!lehrer || !fach || !klasse) {
                skipped++;
                return;
            }
            const key = `${lehrer.toUpperCase()}|${fach.toUpperCase()}|${klasse.toUpperCase()}`;
            if (seen.has(key)) {
                dup++;
                return;
            }
            seen.add(key);
            rows.push({
                id: id++,
                klasse,
                fach,
                lehrer,
                gruppe: '',
                original: { paste: true, line }
            });
        });
        if (!rows.length) {
            ns.showToast('Keine gültigen Zeilen (je Zeile: Lehrer, Fach, Klasse – durch Tab oder Leerzeichen getrennt).');
            return;
        }
        ns.applyWebuntisRows(rows);
        ns.showToast(
            rows.length +
                ' eindeutige Zeile(n)' +
                (dup ? ', ' + dup + ' Duplikat(e) entfernt' : '') +
                (skipped ? ', ' + skipped + ' Zeile(n) übersprungen' : '') +
                '.'
        );
    };

    ns.processImportedData = function processImportedData(data) {
        const rows = [];
        let id = 0;
        data.forEach(origRaw => {
            const row = ns.normalizeImportedRowKeys(origRaw);
            const lehrer = (row.Lehrer || row.lehrer || row.Teacher || row.LehrerIn || '').toString().trim();
            const fach = (row.Fach || row.fach || row.Subject || row.Unterrichtsfach || '').toString().trim();
            const klasseRaw = (row['Klasse(n)'] || row.Klasse || row.klasse || row.Class || '').toString().trim();
            const gruppe = (row['Schülergruppe'] || row.Schülergruppe || row.Gruppe || row.gruppe || row.Group || '').toString().trim();
            if (!lehrer || !fach) return;

            const klassenParts = ns.splitKlassenCell(klasseRaw);
            const targets = klassenParts.length ? klassenParts : [''];

            targets.forEach(klasse => {
                rows.push({
                    id: id++,
                    klasse,
                    fach,
                    lehrer,
                    gruppe,
                    original: row
                });
            });
        });
        ns.applyWebuntisRows(rows);
    };

    function handleFile(file) {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const name = (file.name || '').toLowerCase();
                let jsonData;
                if (name.endsWith('.csv')) {
                    const buf = new Uint8Array(e.target.result);
                    let text = new TextDecoder('utf-8').decode(buf);
                    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
                    let workbook = XLSX.read(text, { type: 'string', FS: ';' });
                    let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    jsonData = XLSX.utils.sheet_to_json(firstSheet);
                    if (!jsonData.length || Object.keys(jsonData[0] || {}).length < 2) {
                        const wb2 = XLSX.read(text, { type: 'string', FS: ',' });
                        const sh2 = wb2.Sheets[wb2.SheetNames[0]];
                        const j2 = XLSX.utils.sheet_to_json(sh2);
                        if (j2.length) jsonData = j2;
                    }
                } else {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    jsonData = XLSX.utils.sheet_to_json(firstSheet);
                }
                ns.processImportedData(jsonData);
            } catch (error) {
                ns.showToast('Fehler beim Lesen der Datei: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // Upload wiring
    if (ns.dom && ns.dom.uploadArea && ns.dom.fileInput) {
        ns.dom.uploadArea.addEventListener('click', () => ns.dom.fileInput.click());
        ns.dom.uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            ns.dom.uploadArea.classList.add('dragover');
        });
        ns.dom.uploadArea.addEventListener('dragleave', () => ns.dom.uploadArea.classList.remove('dragover'));
        ns.dom.uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            ns.dom.uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
        });
        ns.dom.fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) handleFile(e.target.files[0]);
        });
    }

    // Export in global scope for HTML onclick
    window.importWebuntisFromPaste = ns.importWebuntisFromPaste;
})();

