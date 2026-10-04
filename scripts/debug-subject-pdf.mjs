import { readFileSync } from 'node:fs';
import { createContext, runInContext } from 'node:vm';
import { execSync } from 'node:child_process';

const sandbox = { console };
sandbox.window = sandbox;
createContext(sandbox);
runInContext(readFileSync('src/shared/webuntis-export-import.js', 'utf8'), sandbox, {
    filename: 'webuntis-export-import.js'
});
const wu = sandbox.ms365WebuntisExportImport;

const pdf = 'c:/Users/KurtSöser/Downloads/Subject_20261002_1320.pdf';
const exported = JSON.parse(
    execSync(`python scripts/export-subject-pdf-json.py "${pdf}"`, { encoding: 'utf8', maxBuffer: 30 * 1024 * 1024 })
);
const words = exported.words;

const pw = wu.parseSubjectsFromPdfWords(words);
console.log('words count', pw.meta.subjectCount);
const garbage = pw.subjects.filter((s) => s.code.length > 15);
console.log('long codes', garbage.length, garbage.slice(0, 3));

const text = exported.text;
const pt = wu.parseSubjectsFromPdfText(text);
console.log('text count', pt.meta.subjectCount);
const garbageT = pt.subjects.filter((s) => s.code.length > 15);
console.log('text long codes', garbageT.length, garbageT.slice(0, 2).map((s) => s.code));

const rows = new Map();
for (const w of words) {
    const y = Math.round(w.y);
    if (y < 125 || y > 780) continue;
    if (!rows.has(y)) rows.set(y, []);
    rows.get(y).push(w);
}
for (const [y, cells] of [...rows.entries()].sort((a, b) => a[0] - b[0])) {
    const kurz = cells.filter((c) => c.x < 75).map((c) => c.str);
    const joined = kurz.join('');
    if (joined.includes('D1') && joined.length > 8) {
        console.log('bad y', y, 'kurz parts', kurz.length, kurz.slice(0, 12));
    }
}
