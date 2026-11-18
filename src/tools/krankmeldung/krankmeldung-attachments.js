/**
 * Anhänge Krankmeldung (IndexedDB) – z. B. ärztliche Bestätigung / Foto.
 */

import { UPLOAD_MAX_BYTES } from './krankmeldung-model.js';

const DB_NAME = 'ms365-krankmeldung-attachments';
const STORE = 'files';

function openDb() {
    return new Promise((resolve, reject) => {
        if (typeof indexedDB === 'undefined') {
            reject(new Error('IndexedDB nicht verfügbar.'));
            return;
        }
        const req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = () => {
            const db = req.result;
            if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error || new Error('IDB-Fehler'));
    });
}

/**
 * @param {File|Blob} file
 */
export function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ''));
        reader.onerror = () => reject(reader.error || new Error('Datei lesen fehlgeschlagen.'));
        reader.readAsDataURL(file);
    });
}

/**
 * @param {File} file
 */
export function validateUploadFile(file) {
    if (!file) return { ok: false, error: 'Keine Datei.' };
    if (file.size > UPLOAD_MAX_BYTES) {
        return {
            ok: false,
            error: 'Datei zu groß (max. ' + Math.round(UPLOAD_MAX_BYTES / (1024 * 1024)) + ' MB).'
        };
    }
    const name = String(file.name || '').toLowerCase();
    const mime = String(file.type || '').toLowerCase();
    const okMime =
        !mime ||
        mime.startsWith('image/') ||
        mime === 'application/pdf' ||
        mime.includes('word') ||
        mime.includes('officedocument');
    const okExt = /\.(pdf|png|jpe?g|gif|webp|doc|docx)$/i.test(name);
    if (!okMime && !okExt) {
        return { ok: false, error: 'Erlaubt: PDF, Bild oder Word-Dokument.' };
    }
    return { ok: true };
}

/**
 * @param {string} key
 * @param {{ name: string, mime: string, size: number, dataUrl: string }} payload
 */
export async function saveAttachment(key, payload) {
    const id = String(key || '').trim();
    if (!id) throw new Error('Schlüssel fehlt.');
    const db = await openDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, 'readwrite');
        tx.objectStore(STORE).put(
            {
                name: String(payload.name || 'datei'),
                mime: String(payload.mime || 'application/octet-stream'),
                size: Number(payload.size) || 0,
                dataUrl: String(payload.dataUrl || ''),
                savedAt: new Date().toISOString()
            },
            id
        );
        tx.oncomplete = () => resolve(id);
        tx.onerror = () => reject(tx.error || new Error('Speichern fehlgeschlagen.'));
    });
}

/**
 * @param {string} key
 */
export async function loadAttachment(key) {
    const id = String(key || '').trim();
    if (!id) return null;
    try {
        const db = await openDb();
        return await new Promise((resolve, reject) => {
            const req = db.transaction(STORE, 'readonly').objectStore(STORE).get(id);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => reject(req.error);
        });
    } catch {
        return null;
    }
}

/**
 * @param {string} key
 */
export async function deleteAttachment(key) {
    const id = String(key || '').trim();
    if (!id) return;
    try {
        const db = await openDb();
        await new Promise((resolve, reject) => {
            const tx = db.transaction(STORE, 'readwrite');
            tx.objectStore(STORE).delete(id);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    } catch {
        /* ignore */
    }
}
