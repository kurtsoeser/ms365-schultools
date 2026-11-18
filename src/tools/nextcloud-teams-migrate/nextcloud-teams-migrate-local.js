/**
 * Lokale Quelle: Ordner (webkitdirectory) oder ZIP → virtueller Dateibaum.
 */
import { norm } from './nextcloud-teams-migrate-logic.js';

/**
 * @typedef {{ relPath: string, name: string, isFolder: boolean, size: number, file?: File, asyncBlob?: () => Promise<Blob> }} VirtualNode
 * @typedef {{ byPath: Map<string, VirtualNode>, children: Map<string, string[]> }} VirtualFs
 */

/**
 * @param {string} p
 */
function cleanPath(p) {
    return String(p || '')
        .replace(/\\/g, '/')
        .replace(/^\/+|\/+$/g, '')
        .replace(/\/+/g, '/');
}

/**
 * @param {VirtualFs} vfs
 * @param {string} folderRel
 */
function ensureFolder(vfs, folderRel) {
    const rel = cleanPath(folderRel);
    if (!rel) return;
    if (!vfs.byPath.has(rel)) {
        vfs.byPath.set(rel, {
            relPath: rel,
            name: rel.split('/').pop() || rel,
            isFolder: true,
            size: 0
        });
    }
    const parent = rel.includes('/') ? rel.split('/').slice(0, -1).join('/') : '';
    const list = vfs.children.get(parent) || [];
    if (!list.includes(rel)) {
        list.push(rel);
        vfs.children.set(parent, list);
    }
    if (parent) ensureFolder(vfs, parent);
}

/**
 * @param {VirtualFs} vfs
 * @param {VirtualNode} node
 */
function addFileNode(vfs, node) {
    const rel = cleanPath(node.relPath);
    if (!rel || rel.endsWith('/')) return;
    node.relPath = rel;
    node.name = rel.split('/').pop() || node.name || 'datei';
    vfs.byPath.set(rel, node);
    const parent = rel.includes('/') ? rel.split('/').slice(0, -1).join('/') : '';
    ensureFolder(vfs, parent);
    const list = vfs.children.get(parent) || [];
    if (!list.includes(rel)) {
        list.push(rel);
        vfs.children.set(parent, list);
    }
}

/**
 * @returns {VirtualFs}
 */
export function createVirtualFs() {
    return { byPath: new Map(), children: new Map([['', []]]) };
}

/**
 * Dateiliste aus &lt;input webkitdirectory&gt;.
 * @param {FileList|File[]} fileList
 * @returns {VirtualFs}
 */
export function buildFsFromFileList(fileList) {
    const vfs = createVirtualFs();
    const files = Array.from(fileList || []);
    for (const file of files) {
        const rel = cleanPath(
            /** @type {any} */ (file).webkitRelativePath || file.name
        );
        if (!rel || /(^|\/)\.DS_Store$/i.test(rel) || /(^|\/)__MACOSX\//i.test(rel)) continue;
        addFileNode(vfs, {
            relPath: rel,
            name: file.name,
            isFolder: false,
            size: file.size || 0,
            file,
            asyncBlob: () => Promise.resolve(file)
        });
    }
    return vfs;
}

/**
 * ZIP (JSZip-Instanz mit .files).
 * @param {{ files: Record<string, { dir?: boolean, name: string, async: (type: string) => Promise<any> }> }} zip
 * @returns {Promise<VirtualFs>}
 */
export async function buildFsFromJsZip(zip) {
    const vfs = createVirtualFs();
    const entries = Object.keys(zip.files || {});
    for (const raw of entries) {
        const entry = zip.files[raw];
        if (!entry) continue;
        let rel = cleanPath(raw);
        if (!rel || /(^|\/)\.DS_Store$/i.test(rel) || /(^|\/)__MACOSX(\/|$)/i.test(rel)) continue;
        if (entry.dir || raw.endsWith('/')) {
            ensureFolder(vfs, rel);
            continue;
        }
        addFileNode(vfs, {
            relPath: rel,
            name: rel.split('/').pop() || rel,
            isFolder: false,
            size: 0,
            asyncBlob: async () => {
                const blob = await entry.async('blob');
                return blob;
            }
        });
    }
    // Größen nachziehen (optional, kann bei großen ZIPs dauern – nur Metadaten wenn möglich)
    return vfs;
}

/**
 * @param {VirtualFs} vfs
 * @param {string} relPath
 */
export function listVirtualChildren(vfs, relPath) {
    const parent = cleanPath(relPath);
    const ids = vfs.children.get(parent) || [];
    return ids
        .map((id) => vfs.byPath.get(id))
        .filter(Boolean)
        .sort((a, b) => {
            if (a.isFolder !== b.isFolder) return a.isFolder ? -1 : 1;
            return a.name.localeCompare(b.name, 'de', { sensitivity: 'base' });
        });
}

/**
 * @param {VirtualFs} vfs
 * @param {string} relPath Ordner oder Datei
 */
export function inventoryVirtual(vfs, relPath) {
    const base = cleanPath(relPath);
    /** @type {VirtualNode[]} */
    const files = [];
    const baseNode = base ? vfs.byPath.get(base) : null;
    if (baseNode && !baseNode.isFolder) {
        return {
            baseRel: base,
            files: [baseNode],
            fileCount: 1,
            folderCount: 0,
            totalBytes: baseNode.size || 0
        };
    }
    /** @type {string[]} */
    const queue = [base];
    const seen = new Set();
    let folderCount = 0;
    while (queue.length) {
        const cur = queue.shift() || '';
        if (seen.has(cur)) continue;
        seen.add(cur);
        const kids = listVirtualChildren(vfs, cur);
        for (const it of kids) {
            if (it.isFolder) {
                folderCount++;
                queue.push(it.relPath);
            } else {
                files.push(it);
            }
        }
    }
    files.sort((a, b) => a.relPath.localeCompare(b.relPath, 'de'));
    return {
        baseRel: base,
        files,
        fileCount: files.length,
        folderCount,
        totalBytes: files.reduce((s, f) => s + (f.size || 0), 0)
    };
}

/**
 * @param {VirtualFs} vfs
 * @param {string} relPath
 * @returns {Promise<Blob>}
 */
export async function readVirtualBlob(vfs, relPath) {
    const node = vfs.byPath.get(cleanPath(relPath));
    if (!node || node.isFolder) throw new Error('Datei nicht gefunden: ' + relPath);
    if (typeof node.asyncBlob === 'function') return node.asyncBlob();
    if (node.file) return node.file;
    throw new Error('Kein Dateiinhalt: ' + relPath);
}

/**
 * Nur die ausgewählten Ordner selbst (keine Unterordner) – für Kanal-Vorschläge.
 * @param {VirtualFs} _vfs
 * @param {Iterable<string>} selectedRels
 */
export function folderSamplesFromSelection(_vfs, selectedRels) {
    /** @type {Array<{ name: string, isFolder: boolean, relPath: string }>} */
    const out = [];
    for (const rel of selectedRels) {
        const r = cleanPath(rel);
        if (!r) continue;
        const name = r.split('/').pop() || r;
        out.push({ name, isFolder: true, relPath: r });
    }
    return out;
}

/**
 * @param {string} mode
 */
export function isLocalSourceMode(mode) {
    const m = norm(mode);
    return m === 'zip' || m === 'folder';
}
