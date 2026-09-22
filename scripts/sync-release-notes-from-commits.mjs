/**
 * Hängt feat:/fix:-Commits an public/release-notes.json an (Dedup per gitSha).
 *
 *   node scripts/sync-release-notes-from-commits.mjs
 *   node scripts/sync-release-notes-from-commits.mjs --since=HEAD~30
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const filePath = path.join(root, 'public', 'release-notes.json');
const sinceArg = process.argv.find((a) => a.startsWith('--since='));
const since = sinceArg ? sinceArg.slice('--since='.length) : 'HEAD~40';

function readNotes() {
    try {
        const raw = fs.readFileSync(filePath, 'utf8');
        const data = JSON.parse(raw);
        return Array.isArray(data) ? data : [];
    } catch {
        return [];
    }
}

function escapeHtml(s) {
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function bodyToHtml(body, subject) {
    const text = String(body || '').trim();
    if (!text) return `<p>${escapeHtml(subject)}</p>`;
    return text
        .split(/\n{2,}/)
        .map((p) => `<p>${escapeHtml(p).replace(/\n/g, '<br>')}</p>`)
        .join('');
}

function classify(subject) {
    const m = String(subject || '').match(/^(feat|fix)(\(.+\))?!?:\s*(.+)$/i);
    if (!m) return null;
    const kind = m[1].toLowerCase() === 'fix' ? 'fix' : 'feature';
    const title = String(m[3] || '').trim();
    if (!title) return null;
    if (/^merge\b/i.test(title)) return null;
    if (/release.?notes|sync-release-notes|\[skip notes\]/i.test(subject)) return null;
    return { kind, title };
}

const raw = execFileSync(
    'git',
    ['log', since, '--pretty=format:%H%x09%cI%x09%s%x09%b%x1e'],
    { encoding: 'utf8' }
);

const commits = String(raw)
    .split('\x1e')
    .map((chunk) => chunk.trim())
    .filter(Boolean)
    .map((chunk) => {
        const [sha, date, subject, ...bodyParts] = chunk.split('\t');
        return {
            sha: String(sha || '').trim(),
            date: String(date || '').trim(),
            subject: String(subject || '').trim(),
            body: bodyParts.join('\t').trim()
        };
    })
    .filter((c) => c.sha && c.subject);

const notes = readNotes();
const knownSha = new Set(
    notes.map((n) => String(n.gitSha || '').toLowerCase()).filter(Boolean)
);

const newestNoteMs = notes.reduce((acc, n) => {
    const t = Date.parse(String(n.at || ''));
    return Number.isFinite(t) && t > acc ? t : acc;
}, 0);

// Beim ersten Befüllen nur die neuesten wenigen Commits – sonst alles nach dem letzten Note-Datum.
const bootstrapLimit = notes.length ? 100 : 5;
let candidates = 0;

let added = 0;
for (const c of commits) {
    const cls = classify(c.subject);
    if (!cls) continue;
    if (knownSha.has(c.sha.toLowerCase())) continue;

    const commitMs = Date.parse(c.date);
    if (newestNoteMs && Number.isFinite(commitMs) && commitMs <= newestNoteMs) continue;

    if (!newestNoteMs) {
        candidates += 1;
        if (candidates > bootstrapLimit) continue;
    }

    const id = 'rn_git_' + c.sha.slice(0, 12);
    if (notes.some((n) => n.id === id)) continue;

    notes.unshift({
        id,
        at: c.date || new Date().toISOString(),
        title: cls.title,
        kind: cls.kind,
        source: 'github',
        gitSha: c.sha,
        bodyHtml: bodyToHtml(c.body, cls.title),
        images: []
    });
    knownSha.add(c.sha.toLowerCase());
    added += 1;
}

notes.sort((a, b) => String(b.at).localeCompare(String(a.at)));
fs.mkdirSync(path.dirname(filePath), { recursive: true });
fs.writeFileSync(filePath, JSON.stringify(notes, null, 2) + '\n', 'utf8');
console.log(
    added
        ? `sync-release-notes: ${added} Eintrag/Einträge ergänzt → ${path.relative(root, filePath)}`
        : `sync-release-notes: keine neuen feat/fix-Commits (${path.relative(root, filePath)})`
);
