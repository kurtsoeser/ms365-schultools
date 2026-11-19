/**
 * Öffentliche Eltern-Antwortseite (Magic-Link) inkl. optionaler Datei-/Bildabgabe.
 */

import {
    fileToDataUrl,
    saveAttachment,
    validateUploadFile
} from './einsammler-attachments.js';
import { pushInboxSlip } from './einsammler-inbox.js';
import { ANTWORT_STATUS, MODUS, TYP_LABEL, UPLOAD_MAX_BYTES } from './einsammler-model.js';
import { allowedAntwortStatusesForModus } from './einsammler-logic.js';
import { einsammlerRemoteConfigured, remoteRespond, remoteView } from './einsammler-remote.js';
import {
    buildResponseSlip,
    decodePublicToken,
    encodeResponseSlip,
    readParentLinkToken
} from './einsammler-token.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatDeDate(iso) {
    const s = String(iso || '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return s || '–';
    const [y, m, d] = s.split('-');
    return d + '.' + m + '.' + y;
}

function readTokenFromLocation() {
    return readParentLinkToken(window.location);
}

/**
 * @param {HTMLElement} host
 */
export async function mountPublicForm(host) {
    if (!host) return;
    const token = readTokenFromLocation();
    if (!token) {
        host.innerHTML =
            '<div class="ein-public-card"><h2>Link fehlt</h2><p>Bitte öffnen Sie den Link, den Sie von der Schule erhalten haben.</p></div>';
        return;
    }

    host.innerHTML = '<p class="ein-public-meta">Laden …</p>';

    let payload = null;
    if (einsammlerRemoteConfigured()) {
        try {
            const data = await remoteView(token);
            payload = data.payload || data;
            if (data.answered) {
                host.innerHTML =
                    '<div class="ein-public-card ein-public-card--ok"><h2>Bereits beantwortet</h2><p>Für diesen Link liegt schon eine Rückmeldung vor. Danke!</p></div>';
                return;
            }
        } catch (e) {
            const msg = e && e.message ? String(e.message) : '';
            if (/bereits/i.test(msg)) {
                host.innerHTML =
                    '<div class="ein-public-card ein-public-card--ok"><h2>Bereits beantwortet</h2><p>Für diesen Link liegt schon eine Rückmeldung vor. Danke!</p></div>';
                return;
            }
        }
    }
    if (!payload) {
        const decoded = decodePublicToken(token);
        if (!decoded.ok) {
            host.innerHTML =
                '<div class="ein-public-card ein-public-card--err"><h2>Ungültiger Link</h2><p>' +
                esc(decoded.error || 'Fehler') +
                '</p></div>';
            return;
        }
        payload = decoded.payload;
    }

    if (!payload || !payload.eid || !payload.aid || !payload.titel) {
        host.innerHTML =
            '<div class="ein-public-card ein-public-card--err"><h2>Fehler</h2><p>Link unvollständig.</p></div>';
        return;
    }
    if (payload.exp && Date.now() > Number(payload.exp)) {
        host.innerHTML =
            '<div class="ein-public-card ein-public-card--err"><h2>Abgelaufen</h2><p>Dieser Link ist nicht mehr gültig.</p></div>';
        return;
    }
    const kopf = String(payload.statusKopf || '').toLowerCase();
    if (kopf && kopf !== 'aktiv' && kopf !== 'entwurf') {
        host.innerHTML =
            '<div class="ein-public-card ein-public-card--err"><h2>Geschlossen</h2><p>Dieser Einsammler ist nicht mehr offen.</p></div>';
        return;
    }

    renderForm(host, payload, token);
}

/**
 * @param {HTMLElement} host
 * @param {object} payload
 * @param {string} token
 */
function renderForm(host, payload, token) {
    const modus = String(payload.modus || MODUS.EINWILLIGUNG);
    const isRead = modus === MODUS.LESEBESTAETIGUNG;
    const isAbgabe = modus === MODUS.ABGABE;
    const wantUpload = !!(payload.uploadErlaubt || payload.uploadPflicht || isAbgabe);
    const mustUpload = !!(payload.uploadPflicht || isAbgabe);
    const title = isAbgabe ? 'Datei-/Bildabgabe' : isRead ? 'Lesebestätigung' : 'Einwilligung';
    const actions = isAbgabe
        ? [{ id: ANTWORT_STATUS.ABGEGEBEN, label: 'Datei absenden', primary: true }]
        : isRead
          ? [{ id: ANTWORT_STATUS.GELESEN, label: 'Zur Kenntnis genommen', primary: true }]
          : [
                { id: ANTWORT_STATUS.ZUGESTIMMT, label: 'Zustimmen', primary: true },
                { id: ANTWORT_STATUS.ABGELEHNT, label: 'Ablehnen', primary: false }
            ];

    host.innerHTML =
        '<div class="ein-public-card">' +
        (payload.schoolName
            ? '<p class="ein-public-school">' + esc(payload.schoolName) + '</p>'
            : '') +
        '<p class="ein-public-kicker">' +
        esc(title) +
        ' · ' +
        esc(TYP_LABEL[payload.typ] || payload.typ) +
        '</p>' +
        '<h2>' +
        esc(payload.titel) +
        '</h2>' +
        '<p class="ein-public-meta">Für <strong>' +
        esc(payload.schuelerName || 'Schüler:in') +
        '</strong>' +
        (payload.klasseCode ? ' (' + esc(payload.klasseCode) + ')' : '') +
        (payload.faelligAm ? ' · Frist ' + esc(formatDeDate(payload.faelligAm)) : '') +
        '</p>' +
        '<div class="ein-public-body">' +
        esc(payload.textInhalt || '').replace(/\n/g, '<br>') +
        '</div>' +
        (wantUpload
            ? '<label class="ein-public-label">Datei / Foto' +
              (mustUpload ? ' *' : ' (optional)') +
              '<input type="file" id="einPubFile" accept="image/*,.pdf,.doc,.docx,.png,.jpg,.jpeg">' +
              '<span class="ein-public-meta">Max. ' +
              Math.round(UPLOAD_MAX_BYTES / (1024 * 1024)) +
              ' MB · Bild oder PDF</span></label>'
            : '') +
        '<label class="ein-public-label">Kommentar (optional)<textarea id="einPubKommentar" rows="3" placeholder="Optionaler Hinweis an die Schule"></textarea></label>' +
        '<div class="ein-public-actions" id="einPubActions"></div>' +
        '<p class="ein-public-note">Kein Microsoft-Konto nötig. Der Link ist persönlich – bitte nicht weiterleiten.</p>' +
        '</div>';

    const box = host.querySelector('#einPubActions');
    for (const a of actions) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = a.primary ? 'btn btn-success' : 'btn';
        btn.textContent = a.label;
        btn.addEventListener('click', () =>
            submit(
                host,
                payload,
                token,
                a.id,
                String(host.querySelector('#einPubKommentar')?.value || ''),
                host.querySelector('#einPubFile')?.files?.[0] || null,
                mustUpload
            )
        );
        box.appendChild(btn);
    }

    void allowedAntwortStatusesForModus(modus);
}

/**
 * @param {HTMLElement} host
 * @param {object} payload
 * @param {string} token
 * @param {string} status
 * @param {string} kommentar
 * @param {File|null} file
 * @param {boolean} mustUpload
 */
async function submit(host, payload, token, status, kommentar, file, mustUpload) {
    if (mustUpload && !file) {
        host.querySelector('.ein-public-note') &&
            (host.querySelector('.ein-public-note').textContent =
                'Bitte eine Datei oder ein Foto auswählen.');
        return;
    }
    if (file) {
        const v = validateUploadFile(file);
        if (!v.ok) {
            host.innerHTML =
                '<div class="ein-public-card ein-public-card--err"><h2>Datei</h2><p>' +
                esc(v.error || '') +
                '</p><button type="button" class="btn" id="einPubBack">Zurück</button></div>';
            host.querySelector('#einPubBack')?.addEventListener('click', () =>
                renderForm(host, payload, token)
            );
            return;
        }
    }

    host.innerHTML = '<p class="ein-public-meta">Wird gespeichert …</p>';

    /** @type {{ anhangName?: string, anhangMime?: string, anhangSize?: number, anhangKey?: string, anhangDataUrl?: string }} */
    let anhang = {};
    if (file) {
        try {
            const dataUrl = await fileToDataUrl(file);
            const key = String(payload.aid || '');
            await saveAttachment(key, {
                name: file.name,
                mime: file.type || 'application/octet-stream',
                size: file.size,
                dataUrl
            });
            anhang = {
                anhangName: file.name,
                anhangMime: file.type || 'application/octet-stream',
                anhangSize: file.size,
                anhangKey: key,
                anhangDataUrl: dataUrl
            };
        } catch (e) {
            host.innerHTML =
                '<div class="ein-public-card ein-public-card--err"><h2>Upload fehlgeschlagen</h2><p>' +
                esc(e && e.message ? e.message : String(e)) +
                '</p></div>';
            return;
        }
    }

    const slip = buildResponseSlip({
        eid: payload.eid,
        aid: payload.aid,
        status,
        kommentar,
        schuelerName: payload.schuelerName,
        ...anhang
    });

    let remoteOk = false;
    if (einsammlerRemoteConfigured()) {
        try {
            await remoteRespond({
                token,
                status,
                kommentar,
                anhangName: anhang.anhangName || '',
                anhangMime: anhang.anhangMime || '',
                anhangSize: anhang.anhangSize || 0,
                // API speichert Base64 nur wenn vorhanden (gleiche Geräte-Grenze sonst)
                anhangDataUrl: anhang.anhangDataUrl || ''
            });
            remoteOk = true;
        } catch (e) {
            void e;
        }
    }

    try {
        pushInboxSlip(slip);
    } catch {
        /* ignore */
    }

    const code = encodeResponseSlip({
        ...slip,
        // Code ohne große Datei – nur Metadaten
        anhangDataUrl: undefined
    });
    const label =
        status === ANTWORT_STATUS.ZUGESTIMMT
            ? 'Zustimmung gespeichert'
            : status === ANTWORT_STATUS.ABGELEHNT
              ? 'Ablehnung gespeichert'
              : status === ANTWORT_STATUS.GELESEN
                ? 'Lesebestätigung gespeichert'
                : status === ANTWORT_STATUS.ABGEGEBEN
                  ? 'Datei abgegeben'
                  : 'Antwort gespeichert';

    host.innerHTML =
        '<div class="ein-public-card ein-public-card--ok">' +
        '<h2>' +
        esc(label) +
        '</h2>' +
        '<p>Vielen Dank' +
        (payload.schuelerName ? ', Antwort für <strong>' + esc(payload.schuelerName) + '</strong>' : '') +
        '.' +
        (anhang.anhangName ? ' Anhang: <strong>' + esc(anhang.anhangName) + '</strong>.' : '') +
        '</p>' +
        (remoteOk
            ? '<p class="ein-public-meta">Die Schule hat die Antwort bereits erhalten.</p>'
            : '<p class="ein-public-meta">Falls die Schule einen <strong>Antwortcode</strong> erbittet, kopieren Sie ihn hier (ohne Dateiinhalt – Datei liegt lokal im Browser bzw. auf dem Server):</p>' +
              '<textarea class="ein-public-code" readonly rows="4">' +
              esc(code) +
              '</textarea>' +
              '<div class="ein-public-actions"><button type="button" class="btn" id="einPubCopy">Code kopieren</button></div>') +
        '</div>';

    const copyBtn = host.querySelector('#einPubCopy');
    if (copyBtn) {
        copyBtn.addEventListener('click', async () => {
            try {
                await navigator.clipboard.writeText(code);
                copyBtn.textContent = 'Kopiert';
            } catch {
                copyBtn.textContent = 'Bitte manuell markieren';
            }
        });
    }
}
