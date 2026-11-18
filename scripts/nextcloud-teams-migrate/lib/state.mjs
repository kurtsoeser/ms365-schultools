/**
 * Resume-State: erfolgreich übertragene Relativpfade.
 */
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

/**
 * @param {string} mappingId
 * @param {string} nextcloudPath
 * @param {string} groupId
 */
export function stateKey(mappingId, nextcloudPath, groupId) {
    const raw = `${mappingId}::${nextcloudPath}::${groupId}`;
    return crypto.createHash('sha1').update(raw).digest('hex').slice(0, 16);
}

/**
 * @param {string} stateDir
 * @param {string} key
 */
export function stateFilePath(stateDir, key) {
    return path.join(stateDir, `done-${key}.json`);
}

/**
 * @param {string} stateDir
 * @param {string} key
 * @returns {{ done: Record<string, { size: number, at: string }> }}
 */
export function loadState(stateDir, key) {
    const fp = stateFilePath(stateDir, key);
    if (!fs.existsSync(fp)) return { done: {} };
    try {
        const data = JSON.parse(fs.readFileSync(fp, 'utf8'));
        return { done: data.done && typeof data.done === 'object' ? data.done : {} };
    } catch {
        return { done: {} };
    }
}

/**
 * @param {string} stateDir
 * @param {string} key
 * @param {string} relPathWithinSource relative to source folder root
 * @param {number} size
 * @param {{ done: Record<string, { size: number, at: string }> }} state
 */
export function markDone(stateDir, key, relPathWithinSource, size, state) {
    if (!fs.existsSync(stateDir)) fs.mkdirSync(stateDir, { recursive: true });
    state.done[relPathWithinSource] = { size: Number(size) || 0, at: new Date().toISOString() };
    const fp = stateFilePath(stateDir, key);
    fs.writeFileSync(
        fp,
        JSON.stringify(
            {
                key,
                updatedAt: new Date().toISOString(),
                doneCount: Object.keys(state.done).length,
                done: state.done
            },
            null,
            2
        ),
        'utf8'
    );
}
