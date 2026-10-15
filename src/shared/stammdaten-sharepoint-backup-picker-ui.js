/**
 * Modal: Backup-Datei in der IT-Bibliothek wählen (nicht nur „aktuell“).
 */
import { DEFAULT_FOLDER, CURRENT_FILE } from './stammdaten-sharepoint-sync-logic.js';
import { listImportableBackupFiles, downloadBackupByItem } from './stammdaten-sharepoint-sync-api.js';

const OVERLAY_ID = 'ms365BackupFilePicker';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function ensureOverlay() {
    let root = document.getElementById(OVERLAY_ID);
    if (root) return root;
    root = document.createElement('div');
    root.id = OVERLAY_ID;
    root.className = 'modal-overlay ms365-backup-picker-overlay';
    root.setAttribute('aria-hidden', 'true');
    root.innerHTML =
        '<div class="modal-box ms365-backup-picker-box" role="dialog" aria-modal="true" aria-labelledby="ms365BpickTitle" tabindex="-1">' +
        '<div class="ms365-bpick-scroll" data-bpick-mount></div>' +
        '</div>';
    document.body.appendChild(root);
    root.addEventListener('click', function (e) {
        if (e.target === root) {
            root.dispatchEvent(new CustomEvent('ms365-bpick-cancel', { bubbles: true }));
        }
    });
    return root;
}

/**
 * @param {import('./stammdaten-sharepoint-sync-logic.js').BackupPickerEntry[]} entries
 * @param {{ title?: string, folder?: string, selectedId?: string }} opts
 */
function renderPickerBody(entries, opts) {
    const options = opts || {};
    const title = options.title || 'Backup-Datei wählen';
    const folder = options.folder || DEFAULT_FOLDER;
    const preferred =
        options.selectedId ||
        (entries.find(function (e) {
            return e.isCurrent;
        }) ||
            entries[0] ||
            {}).id ||
        '';

    const rows = entries
        .map(function (entry) {
            const checked = String(entry.id) === String(preferred) ? ' checked' : '';
            const badge = entry.badgeLabel
                ? '<span class="ms365-bpick-badge">' + escapeHtml(entry.badgeLabel) + '</span>'
                : '';
            return (
                '<label class="ms365-bpick-row' +
                (entry.isCurrent ? ' ms365-bpick-row--current' : '') +
                '">' +
                '<input type="radio" name="ms365-bpick-file" value="' +
                escapeHtml(entry.id) +
                '"' +
                checked +
                ' data-bpick-id="' +
                escapeHtml(entry.id) +
                '" data-bpick-name="' +
                escapeHtml(entry.name) +
                '" />' +
                '<span class="ms365-bpick-row__body">' +
                '<span class="ms365-bpick-row__name">' +
                escapeHtml(entry.name) +
                badge +
                '</span>' +
                '<span class="ms365-bpick-row__meta muted">' +
                escapeHtml(entry.subtitle || '–') +
                '</span>' +
                '</span>' +
                '</label>'
            );
        })
        .join('');

    return (
        '<div class="ms365-bpick-head">' +
        '<h3 class="ms365-bpick-title" id="ms365BpickTitle">' +
        escapeHtml(title) +
        '</h3>' +
        '<p class="muted ms365-bpick-intro">Wählen Sie eine JSON-Sicherung aus „' +
        escapeHtml(folder) +
        '“. Vorschlag: <code>' +
        escapeHtml(CURRENT_FILE) +
        '</code> (aktuell).</p>' +
        '</div>' +
        (entries.length
            ? '<div class="ms365-bpick-list" role="radiogroup" aria-label="Backup-Dateien">' +
              rows +
              '</div>'
            : '<p class="ms365-bpick-empty">Keine importierbaren Backup-Dateien in „' +
              escapeHtml(folder) +
              '“ gefunden.</p>') +
        '<div class="ms365-bpick-actions">' +
        '<button type="button" class="btn" data-bpick-action="cancel">Abbrechen</button>' +
        '<button type="button" class="btn btn-success" data-bpick-action="confirm"' +
        (entries.length ? '' : ' disabled') +
        '><i class="bi bi-cloud-arrow-down" aria-hidden="true"></i>Datei laden</button>' +
        '</div>'
    );
}

/**
 * Zeigt die Dateiliste und liefert den gewählten Eintrag oder null (Abbruch).
 * @param {{
 *   folder?: string,
 *   title?: string,
 *   entries?: import('./stammdaten-sharepoint-sync-logic.js').BackupPickerEntry[],
 *   selectedId?: string
 * }} [opts]
 * @returns {Promise<import('./stammdaten-sharepoint-sync-logic.js').BackupPickerEntry|null>}
 */
export async function promptPickSharePointBackupFile(opts) {
    const options = opts || {};
    const folder = String(options.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER;
    const entries =
        Array.isArray(options.entries) && options.entries.length
            ? options.entries
            : await listImportableBackupFiles({ folder: folder });

    if (typeof document === 'undefined' || !document.body) {
        return entries.find(function (e) {
            return e.isCurrent;
        }) || entries[0] || null;
    }

    return new Promise(function (resolve) {
        const overlay = ensureOverlay();
        const mount = overlay.querySelector('[data-bpick-mount]');
        if (!mount) {
            resolve(null);
            return;
        }

        mount.innerHTML = renderPickerBody(entries, {
            title: options.title,
            folder: folder,
            selectedId: options.selectedId
        });

        function finish(entry) {
            overlay.classList.remove('open');
            overlay.setAttribute('aria-hidden', 'true');
            document.removeEventListener('keydown', onKey, true);
            mount.removeEventListener('click', onClick);
            overlay.removeEventListener('ms365-bpick-cancel', onCancel);
            resolve(entry);
        }

        function onCancel() {
            finish(null);
        }

        function onKey(ev) {
            if (ev.key === 'Escape') {
                ev.preventDefault();
                finish(null);
            }
            if (ev.key === 'Enter') {
                const conf = mount.querySelector('[data-bpick-action="confirm"]');
                if (conf && !conf.disabled) {
                    ev.preventDefault();
                    conf.click();
                }
            }
        }

        function onClick(ev) {
            const btn = ev.target && ev.target.closest ? ev.target.closest('[data-bpick-action]') : null;
            if (!btn || !mount.contains(btn)) return;
            const action = btn.getAttribute('data-bpick-action');
            if (action === 'cancel') {
                finish(null);
                return;
            }
            if (action === 'confirm') {
                const checked = mount.querySelector('input[name="ms365-bpick-file"]:checked');
                if (!checked) {
                    finish(null);
                    return;
                }
                const id = checked.getAttribute('data-bpick-id') || checked.value;
                const found =
                    entries.find(function (e) {
                        return String(e.id) === String(id);
                    }) || null;
                finish(found);
            }
        }

        mount.addEventListener('click', onClick);
        document.addEventListener('keydown', onKey, true);
        overlay.addEventListener('ms365-bpick-cancel', onCancel, { once: true });

        overlay.classList.add('open');
        overlay.setAttribute('aria-hidden', 'false');
        const box = overlay.querySelector('.ms365-backup-picker-box');
        if (box) box.focus();
    });
}

/**
 * Datei wählen und herunterladen (ohne Import, außer apply=true).
 * @param {{ folder?: string, title?: string, apply?: boolean, recordMeta?: boolean }} [opts]
 * @returns {Promise<null|{ payload: object, item: object, meta: object, entry: object }>}
 */
export async function pickAndDownloadSharePointBackup(opts) {
    const options = opts || {};
    const folder = String(options.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER;
    const entries = await listImportableBackupFiles({ folder: folder });
    if (!entries.length) {
        throw new Error(
            'Keine importierbaren Backup-Dateien in „' + folder + '“ gefunden (erwartet z. B. ' + CURRENT_FILE + ').'
        );
    }
    const picked = await promptPickSharePointBackupFile({
        folder: folder,
        title: options.title || 'Backup-Datei wählen',
        entries: entries
    });
    if (!picked) return null;
    const downloaded = await downloadBackupByItem({
        folder: folder,
        itemId: picked.id,
        item: {
            id: picked.id,
            name: picked.name,
            webUrl: picked.webUrl,
            lastModifiedDateTime: picked.lastModifiedDateTime,
            size: picked.size
        },
        fileName: picked.name,
        apply: options.apply === true,
        recordMeta: options.recordMeta
    });
    return Object.assign({}, downloaded, { entry: picked });
}

export default {
    promptPickSharePointBackupFile,
    pickAndDownloadSharePointBackup
};
