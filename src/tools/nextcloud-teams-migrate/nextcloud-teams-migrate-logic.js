/**
 * Pure Logik: nummerierte Ordner → Kanäle, Dry-Run-Plan, Validierung.
 */

/**
 * @param {string} s
 */
export function norm(s) {
    return String(s == null ? '' : s).trim();
}

/**
 * Erkennt „01 Fachkonferenz“, „02_Protokolle“, „3. Sitzungen“.
 * @param {string} folderName
 * @returns {{ number: string, title: string, channelName: string } | null}
 */
export function parseNumberedFolderName(folderName) {
    const name = norm(folderName);
    if (!name) return null;
    const m = name.match(/^(\d{1,2})[.\s_\-–—]+(.+)$/);
    if (!m) return null;
    const number = m[1].padStart(2, '0');
    const title = norm(m[2]);
    if (!title) return null;
    const channelName = `${number} ${title}`.slice(0, 50);
    return { number, title, channelName };
}

/**
 * Nur aus der Auswahl (nicht rekursiv aus Unterordnern).
 * @param {Array<{ name: string, isFolder?: boolean, relPath?: string }>} items
 * @returns {Array<{ sourceFolder: string, sourceRelPath: string, number: string, title: string, channelName: string, enabled: boolean }>}
 */
export function suggestChannelMappings(items) {
    const list = Array.isArray(items) ? items : [];
    /** @type {ReturnType<typeof suggestChannelMappings>} */
    const out = [];
    /** @type {Set<string>} */
    const seen = new Set();
    for (const it of list) {
        if (it && it.isFolder === false) continue;
        const name = norm(it && it.name);
        const rel = norm(it && it.relPath) || name;
        const parsed = parseNumberedFolderName(name);
        if (!parsed) continue;
        const key = rel.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        out.push({
            sourceFolder: name,
            sourceRelPath: rel,
            number: parsed.number,
            title: parsed.title,
            channelName: parsed.channelName,
            enabled: true
        });
    }
    out.sort((a, b) => a.number.localeCompare(b.number, 'de') || a.sourceFolder.localeCompare(b.sourceFolder, 'de'));
    return out;
}

/**
 * @param {number} n
 */
export function formatBytes(n) {
    const v = Number(n) || 0;
    if (v < 1024) return `${v} B`;
    if (v < 1024 * 1024) return `${(v / 1024).toFixed(1)} KB`;
    if (v < 1024 * 1024 * 1024) return `${(v / (1024 * 1024)).toFixed(2)} MB`;
    return `${(v / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

/**
 * Relativpfad innerhalb eines Basisordners.
 * @param {string} baseRel
 * @param {string} fileRel
 */
export function relativeInside(baseRel, fileRel) {
    const base = norm(baseRel).replace(/^\/+|\/+$/g, '');
    const file = norm(fileRel).replace(/^\/+|\/+$/g, '');
    if (!base) return file;
    if (file === base) return '';
    if (file.startsWith(base + '/')) return file.slice(base.length + 1);
    return file;
}

/**
 * Top-Level-Ordner eines Relativpfads (unter der Auswahlwurzel).
 * @param {string} relInside
 */
export function topLevelSegment(relInside) {
    const rel = norm(relInside).replace(/^\/+/, '');
    if (!rel) return '';
    return rel.split('/')[0] || '';
}

/**
 * @typedef {{
 *   sourceRelPath: string,
 *   fileName: string,
 *   size: number,
 *   destKind: 'channel' | 'teamRoot',
 *   channelName: string,
 *   destFolderRel: string
 * }} MigrationPlanItem
 */

/**
 * Baut Dry-Run-/Upload-Plan.
 * Kanal-Zuordnung nur über die gewählten Ordner (Pfad-Präfix) – Unterordner wie 21_22
 * bleiben Ordner im Kanal, werden keine eigenen Kanäle.
 * @param {{
 *   sourceRootRel: string,
 *   files: Array<{ relPath: string, name?: string, size?: number }>,
 *   channelMaps: Array<{ sourceFolder?: string, sourceRelPath?: string, channelName: string, enabled: boolean }>,
 *   destRootFolder: string,
 *   mapNumberedToChannels: boolean
 * }} input
 * @returns {{ items: MigrationPlanItem[], summary: { fileCount: number, totalBytes: number, channelCount: number, teamRootCount: number } }}
 */
export function buildMigrationPlan(input) {
    const sourceRootRel = norm(input.sourceRootRel).replace(/^\/+|\/+$/g, '');
    const destRootFolder = norm(input.destRootFolder) || 'Import aus Nextcloud';
    const mapChannels = !!input.mapNumberedToChannels;

    /** @type {Array<{ sourceRelPath: string, channelName: string }>} */
    const maps = [];
    if (mapChannels) {
        for (const m of input.channelMaps || []) {
            if (!m || m.enabled === false) continue;
            const sourceRelPath = norm(m.sourceRelPath || m.sourceFolder).replace(/^\/+|\/+$/g, '');
            const channelName = norm(m.channelName);
            if (sourceRelPath && channelName) maps.push({ sourceRelPath, channelName });
        }
        maps.sort((a, b) => b.sourceRelPath.length - a.sourceRelPath.length);
    }

    /** @type {MigrationPlanItem[]} */
    const items = [];
    let totalBytes = 0;
    let channelCount = 0;
    let teamRootCount = 0;

    for (const f of input.files || []) {
        const fileRel = norm(f.relPath).replace(/^\/+|\/+$/g, '');
        if (!fileRel) continue;
        if (sourceRootRel && fileRel !== sourceRootRel && !fileRel.startsWith(sourceRootRel + '/')) {
            continue;
        }

        const matched = maps.find(
            (m) => fileRel === m.sourceRelPath || fileRel.startsWith(m.sourceRelPath + '/')
        );

        /** @type {MigrationPlanItem} */
        let planItem;
        if (matched) {
            const under = relativeInside(matched.sourceRelPath, fileRel);
            if (!under) continue;
            const parts = under.split('/').filter(Boolean);
            const fileName = parts[parts.length - 1] || norm(f.name) || 'datei';
            const destFolderRel = parts.slice(0, -1).join('/');
            planItem = {
                sourceRelPath: fileRel,
                fileName,
                size: Number(f.size) || 0,
                destKind: 'channel',
                channelName: matched.channelName,
                destFolderRel
            };
            channelCount++;
        } else {
            const relInside = sourceRootRel ? relativeInside(sourceRootRel, fileRel) : fileRel;
            if (!relInside) continue;
            const parts = relInside.split('/').filter(Boolean);
            const fileName = parts[parts.length - 1] || norm(f.name) || 'datei';
            const parentParts = parts.slice(0, -1);
            const rootName = sourceRootRel.split('/').filter(Boolean).pop() || 'Import';
            const underRoot = [destRootFolder, rootName, ...parentParts].filter(Boolean);
            planItem = {
                sourceRelPath: fileRel,
                fileName,
                size: Number(f.size) || 0,
                destKind: 'teamRoot',
                channelName: '',
                destFolderRel: underRoot.join('/')
            };
            teamRootCount++;
        }
        totalBytes += planItem.size;
        items.push(planItem);
    }

    return {
        items,
        summary: {
            fileCount: items.length,
            totalBytes,
            channelCount,
            teamRootCount
        }
    };
}

/**
 * @param {{
 *   sourceMode?: 'helper' | 'zip' | 'folder' | 'webdav',
 *   hasLocalFs?: boolean,
 *   hasHelper?: boolean,
 *   nextcloudUrl?: string,
 *   nextcloudUser?: string,
 *   hasPassword?: boolean,
 *   destGroupId?: string,
 *   selectedRelPaths?: string[]
 * }} input
 */
export function validateNcMigrateSelection(input) {
    /** @type {string[]} */
    const issues = [];
    const mode = norm(input.sourceMode) || 'helper';
    if (mode === 'helper') {
        if (!input.hasHelper) issues.push('Lokaler Helfer nicht verbunden (Start-Helfer.cmd starten)');
    } else if (mode === 'webdav') {
        if (!norm(input.nextcloudUrl)) issues.push('Nextcloud-URL fehlt');
        if (!norm(input.nextcloudUser)) issues.push('Nextcloud-Benutzer fehlt');
        if (!input.hasPassword) issues.push('App-Passwort fehlt');
    } else if (!input.hasLocalFs) {
        issues.push(mode === 'folder' ? 'Bitte einen Ordner vom PC wählen' : 'Bitte eine ZIP-Datei wählen');
    }
    if (!norm(input.destGroupId)) issues.push('Ziel-Team fehlt');
    const sel = Array.isArray(input.selectedRelPaths) ? input.selectedRelPaths.filter(Boolean) : [];
    if (!sel.length) issues.push('Keine Quelle ausgewählt (Ordner/Dateien)');
    return { ok: issues.length === 0, issues };
}

/**
 * WebDAV-Basis ohne Slash.
 * @param {string} baseUrl
 */
export function normalizeNextcloudBase(baseUrl) {
    return norm(baseUrl).replace(/\/+$/, '').replace(/\/index\.php$/i, '');
}
