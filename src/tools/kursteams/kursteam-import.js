
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

function normStr(v) {
    return String(v ?? '').trim();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

ns.seedWebuntisPasteIfEmpty = function seedWebuntisPasteIfEmpty() {
    const ta = document.getElementById('webuntisPasteInput');
    if (!ta) return false;
    if (normStr(ta.value)) return false;
    if (Array.isArray(ns.rawData) && ns.rawData.length) return false;

    let tenant = null;
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') tenant = window.ms365TenantSettingsLoad();
    } catch {
        tenant = null;
    }

    const teachers = Array.isArray(tenant?.teachers) ? tenant.teachers : [];
    const subjects = Array.isArray(tenant?.subjects) ? tenant.subjects : [];
    const classes = Array.isArray(tenant?.classes) ? tenant.classes : [];

    const teacherCodes = teachers.map((t) => normCode(t?.code)).filter(Boolean);
    const subjectCodes = subjects.map((s) => normCode(s?.code)).filter(Boolean);
    const classCodes = classes.map((c) => normCode(c?.code) || normStr(c?.name)).filter(Boolean);

    const tA = teacherCodes[0] || 'LEH';
    const tB = teacherCodes[1] || teacherCodes[0] || 'MUS';
    const s1 = subjectCodes[0] || 'M';
    const s2 = subjectCodes[1] || 'D';
    const s3 = subjectCodes[2] || 'E';
    const c1 = classCodes[0] || '1A';
    const c2 = classCodes[1] || '1B';
    const c3 = classCodes[2] || '2A';

    // 6 Beispielzeilen (Lehrer, Fach, Klasse)
    const lines = [
        `${tA}\t${s1}\t${c1}`,
        `${tA}\t${s2}\t${c2}`,
        `${tA}\t${s3}\t${c3}`,
        `${tB}\t${s1}\t${c2}`,
        `${tB}\t${s2}\t${c3}`,
        `${tB}\t${s3}\t${c1}`
    ];
    ta.value = lines.join('\n');
    return true;
};

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
    if (typeof ns.markAutoSaveDirty === 'function') ns.markAutoSaveDirty();
    document.getElementById('totalRecords').textContent = ns.rawData.length;
    document.getElementById('uniqueSubjects').textContent = new Set(ns.rawData.map(r => r.fach).filter(f => f)).size;
    document.getElementById('uniqueTeachers').textContent = new Set(ns.rawData.map(r => r.lehrer).filter(l => l)).size;
    document.getElementById('importStats').style.display = 'block';
    if (typeof ns.setContinueButton === 'function') {
        ns.setContinueButton('continueBtn1', ns.rawData.length > 0, '');
    }
    if (ns.rawData.length > 0 && typeof ns.scrollToContinue === 'function') {
        ns.scrollToContinue('continueBtn1');
    }
    if (typeof ns.refreshSubjectFilterUI === 'function') ns.refreshSubjectFilterUI();
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

/**
 * Erkennt ob ein JSON-Datensatz aus einem Sokrates-Export stammt und mappt die Felder.
 * Sokrates-Exporte haben typischerweise Spalten wie:
 *   "Unterrichtsgegenstand", "Lehrperson" / "Lehrer/in", "Klasse" / "Klassen"
 * Gibt { lehrer, fach, klasseRaw, gruppe } zurück oder null wenn kein Sokrates-Muster.
 */
ns.trySocratesMapping = function trySocratesMapping(row) {
    const fach =
        (row['Unterrichtsgegenstand'] || row['Unterrichtsgegenstand (Abkürzung)'] ||
         row['UG'] || row['UG-Kürzel'] || row['Gegenstand'] || row['Gegenstandskürzel'] || '').toString().trim();
    const lehrer =
        (row['Lehrer/in'] || row['Lehrperson'] || row['Lehrerin/Lehrer'] || row['LehrerIn'] ||
         row['Lehrer Kürzel'] || row['LehrerKürzel'] || row['Lehrkraft'] || '').toString().trim();
    const klasseRaw =
        (row['Klassen'] || row['Schülerklasse'] || row['Klasse'] || row['Schulklasse'] || '').toString().trim();
    const gruppe =
        (row['Gruppe'] || row['Schülergruppe'] || row['Teilungsgruppe'] || '').toString().trim();

    if (!fach && !lehrer) return null;
    return { lehrer, fach, klasseRaw, gruppe };
};

/**
 * WebUntis „ExportLessons“ (xls/xlsx/csv): Spalten subject, teacher, klassen
 * (optional periods, room, foreignKey). Gibt Mapping oder null.
 */
ns.tryWebuntisLessonsMapping = function tryWebuntisLessonsMapping(row) {
    const hasExportShape =
        Object.prototype.hasOwnProperty.call(row, 'subject') ||
        Object.prototype.hasOwnProperty.call(row, 'teacher') ||
        Object.prototype.hasOwnProperty.call(row, 'klassen');
    if (!hasExportShape) return null;

    const fach = (row.subject || row.Subject || '').toString().trim();
    const lehrer = (row.teacher || row.Teacher || '').toString().trim();
    const klasseRaw = (row.klassen || row.Klassen || row.class || row.Class || '').toString().trim();
    const gruppe = (row.gruppe || row.group || row.Group || '').toString().trim();

    if (!fach && !lehrer) return null;
    return { lehrer, fach, klasseRaw, gruppe };
};

/**
 * Einheitliche Spaltenauflösung für Importzeilen (WebUntis Lessons, Sokrates, Vorlage).
 * @returns {{ lehrer: string, fach: string, klasseRaw: string, gruppe: string, profile: string } | null}
 */
ns.mapImportedLessonRow = function mapImportedLessonRow(rowRaw) {
    const row = ns.normalizeImportedRowKeys(rowRaw || {});

    const wu = ns.tryWebuntisLessonsMapping(row);
    if (wu && (wu.lehrer || wu.fach)) {
        return { ...wu, profile: 'webuntis-lessons' };
    }

    const soc = ns.trySocratesMapping(row);
    if (soc && (soc.lehrer || soc.fach)) {
        return { ...soc, profile: 'sokrates' };
    }

    const lehrer = (row.Lehrer || row.lehrer || row.Teacher || row.teacher || row.LehrerIn || '').toString().trim();
    const fach = (row.Fach || row.fach || row.Subject || row.subject || row.Unterrichtsfach || '').toString().trim();
    const klasseRaw = (
        row['Klasse(n)'] ||
        row.Klasse ||
        row.klasse ||
        row.Klassen ||
        row.klassen ||
        row.Class ||
        row.class ||
        ''
    )
        .toString()
        .trim();
    const gruppe = (
        row['Schülergruppe'] ||
        row.Schülergruppe ||
        row.Gruppe ||
        row.gruppe ||
        row.Group ||
        row.group ||
        ''
    )
        .toString()
        .trim();

    if (!lehrer && !fach) return null;
    return { lehrer, fach, klasseRaw, gruppe, profile: 'generic' };
};

/**
 * Zeigt eine Diagnose-Meldung wenn der Import 0 Zeilen liefert.
 * Gibt dem Nutzer Hinweis auf erkannte Spalten und empfohlene Spaltenbezeichnungen.
 */
ns.showImportDiagnosis = function showImportDiagnosis(jsonData) {
    if (!jsonData || !jsonData.length) {
        ns.showToast('Datei leer oder Format nicht erkannt – bitte XLSX-Vorlage oder WebUntis-ExportLessons verwenden.');
        return;
    }
    const firstRow = ns.normalizeImportedRowKeys(jsonData[0]);
    const erkannt = Object.keys(firstRow).join(', ') || '(keine)';
    const erwartet =
        'WebUntis ExportLessons: "subject", "teacher", "klassen" – oder "Lehrer", "Fach", "Klasse(n)" – oder Sokrates: "Lehrperson"/"Lehrer/in", "Unterrichtsgegenstand", "Klasse"';
    ns.openModal(
        'Import: Keine Daten erkannt',
        '<p style="margin-bottom:10px;">Die Datei enthält <strong>0 verwertbare Zeilen</strong>.</p>' +
        '<p style="margin-bottom:6px;"><strong>Erkannte Spalten:</strong></p>' +
        '<code style="display:block;padding:8px;background:var(--soft);border-radius:6px;word-break:break-all;font-size:0.88em;">' +
            ns.escapeHtml(erkannt) +
        '</code>' +
        '<p style="margin-top:10px;margin-bottom:4px;"><strong>Erwartete Spalten:</strong></p>' +
        '<code style="display:block;padding:8px;background:var(--soft);border-radius:6px;font-size:0.88em;">' +
            ns.escapeHtml(erwartet) +
        '</code>' +
        '<p style="margin-top:12px;color:var(--text-secondary);font-size:0.92em;">Tipp: In WebUntis <strong>ExportLessons</strong> (xls) direkt hochladen – Spalten <code>subject</code>, <code>teacher</code>, <code>klassen</code> werden erkannt. Alternativ XLSX-Vorlage oder Sokrates-Export.</p>',
        null
    );
};

ns.processImportedData = function processImportedData(data) {
    const rows = [];
    let id = 0;
    let socratesHits = 0;
    let webuntisHits = 0;
    data.forEach(origRaw => {
        const mapped = ns.mapImportedLessonRow(origRaw);
        if (!mapped) return;

        const { lehrer, fach, klasseRaw, gruppe, profile } = mapped;
        if (!lehrer || !fach) return;

        if (profile === 'sokrates') socratesHits++;
        if (profile === 'webuntis-lessons') webuntisHits++;

        const klassenParts = ns.splitKlassenCell(klasseRaw);
        const targets = klassenParts.length ? klassenParts : [''];

        targets.forEach(klasse => {
            rows.push({
                id: id++,
                klasse,
                fach,
                lehrer,
                gruppe,
                original: ns.normalizeImportedRowKeys(origRaw)
            });
        });
    });

    if (!rows.length) {
        ns.showImportDiagnosis(data);
        return;
    }

    if (webuntisHits > 0 && webuntisHits > rows.length / 2) {
        ns.showToast('WebUntis-ExportLessons erkannt – ' + rows.length + ' Zeile(n) importiert.');
    } else if (socratesHits > 0 && socratesHits > rows.length / 2) {
        ns.showToast('Sokrates-Export erkannt – ' + rows.length + ' Zeile(n) importiert.');
    }

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
                // Latin-1 / Windows-1252 Fallback für österreichische Sonderzeichen
                let text;
                try {
                    text = new TextDecoder('utf-8', { fatal: true }).decode(buf);
                } catch {
                    text = new TextDecoder('windows-1252').decode(buf);
                }
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
                // Tab-getrennte CSV (z.B. WebUntis-Exporte)
                if (!jsonData.length || Object.keys(jsonData[0] || {}).length < 2) {
                    const wb3 = XLSX.read(text, { type: 'string', FS: '\t' });
                    const sh3 = wb3.Sheets[wb3.SheetNames[0]];
                    const j3 = XLSX.utils.sheet_to_json(sh3);
                    if (j3.length) jsonData = j3;
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

function wireUploadArea() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    if (!uploadArea || !fileInput) return;
    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleFile(e.target.files[0]);
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireUploadArea);
} else {
    wireUploadArea();
}

ns.downloadKursteamImportTemplateXlsx = function downloadKursteamImportTemplateXlsx() {
    if (typeof XLSX === 'undefined' || !XLSX.utils || !XLSX.writeFile) {
        ns.showToast('Excel-Bibliothek nicht geladen – Seite neu laden.');
        return;
    }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
        ['U-Nr', 'Lehrer', 'Fach', 'Klasse(n)', 'Schülergruppe'],
        [1, 'LEH', 'M', '1A', ''],
        [2, 'LEH', 'D', '1B', ''],
        [3, 'LEH', 'E', '2A', ''],
        [4, 'MUS', 'M', '1B', ''],
        [5, 'MUS', 'D', '2A', ''],
        [6, 'MUS', 'E', '1A', 'Gruppe A']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Kursteams');
    XLSX.writeFile(wb, 'Kursteams-Import-Vorlage.xlsx');
    ns.showToast('Excel-Vorlage heruntergeladen.');
};

// Export in global scope for HTML onclick
window.importWebuntisFromPaste = ns.importWebuntisFromPaste;
window.downloadKursteamImportTemplateXlsx = ns.downloadKursteamImportTemplateXlsx;

