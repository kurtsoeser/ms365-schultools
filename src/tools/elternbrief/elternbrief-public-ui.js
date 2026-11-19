/**
 * Öffentliche Elternbrief-Leseseite (Magic-Link) inkl. Lesebestätigung + Anhang-Download.
 */

import { loadAttachment } from './elternbrief-attachments.js';
import { pushLeseSlip } from './elternbrief-inbox.js';
import {
    buildLeseSlip,
    decodePublicToken,
    encodeLeseSlip,
    readParentLinkToken
} from './elternbrief-token.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function readTokenFromLocation() {
    return readParentLinkToken(window.location);
}

/**
 * @param {HTMLElement} host
 */
export async function mountPublicLesen(host) {
    if (!host) return;
    const token = readTokenFromLocation();
    if (!token) {
        host.innerHTML =
            '<div class="eb-public-card"><h2>Link fehlt</h2><p>Bitte öffnen Sie den Link, den Sie von der Schule erhalten haben.</p></div>';
        return;
    }

    const decoded = decodePublicToken(token);
    if (!decoded.ok) {
        host.innerHTML =
            '<div class="eb-public-card eb-public-card--err"><h2>Ungültiger Link</h2><p>' +
            esc(decoded.error || 'Fehler') +
            '</p></div>';
        return;
    }
    renderLesen(host, decoded.payload);
}

/**
 * @param {HTMLElement} host
 * @param {object} payload
 */
function renderLesen(host, payload) {
    const wantConfirm = payload.lesebestaetigung !== false;
    const hasAnhang = !!(payload.hasAnhang || payload.anhangName);

    host.innerHTML =
        '<div class="eb-public-card">' +
        (payload.schoolName
            ? '<p class="eb-public-school">' + esc(payload.schoolName) + '</p>'
            : '') +
        '<p class="eb-public-kicker">Elternbrief' +
        (payload.absenderName ? ' · ' + esc(payload.absenderName) : '') +
        '</p>' +
        '<h2>' +
        esc(payload.betreff) +
        '</h2>' +
        '<p class="eb-public-meta">Für <strong>' +
        esc(payload.schuelerName || 'Schüler:in') +
        '</strong>' +
        (payload.klasseCode ? ' (' + esc(payload.klasseCode) + ')' : '') +
        '</p>' +
        '<div class="eb-public-body">' +
        esc(payload.inhalt || '').replace(/\n/g, '<br>') +
        '</div>' +
        (hasAnhang
            ? '<div class="eb-public-actions"><button type="button" class="btn" id="ebDlAnhang"><i class="bi bi-download"></i> Anhang herunterladen' +
              (payload.anhangName ? ' (' + esc(payload.anhangName) + ')' : '') +
              '</button></div>' +
              '<p class="eb-public-note" id="ebAnhangHint">Der Anhang ist verfügbar, wenn Sie denselben Link auf dem Gerät der Schule öffnen bzw. die Datei dort hinterlegt wurde.</p>'
            : '') +
        (wantConfirm
            ? '<div class="eb-public-actions">' +
              '<button type="button" class="btn btn-success" id="ebConfirmRead">Zur Kenntnis genommen</button>' +
              '</div>' +
              '<p class="eb-public-note">Nach der Bestätigung erhalten Sie einen Code – bei Bedarf der Schule weitergeben (oder er wird lokal eingelesen).</p>'
            : '<p class="eb-public-note">Keine Lesebestätigung erforderlich – Sie können diese Seite schließen.</p>') +
        '<div id="ebConfirmResult" hidden></div>' +
        '</div>';

    const dl = host.querySelector('#ebDlAnhang');
    if (dl) {
        dl.addEventListener('click', async () => {
            const file = await loadAttachment(String(payload.bid || ''));
            if (!file || !file.dataUrl) {
                const hint = host.querySelector('#ebAnhangHint');
                if (hint) {
                    hint.textContent =
                        'Anhang hier nicht gefunden (anderer Browser/Gerät). Bitte die Schule um erneute Zusendung bitten oder den Link später erneut öffnen.';
                }
                return;
            }
            const a = document.createElement('a');
            a.href = file.dataUrl;
            a.download = file.name || payload.anhangName || 'anhang';
            document.body.appendChild(a);
            a.click();
            a.remove();
        });
    }

    const btn = host.querySelector('#ebConfirmRead');
    if (btn) {
        btn.addEventListener('click', () => {
            const slip = buildLeseSlip({ bid: payload.bid, eid: payload.eid });
            try {
                pushLeseSlip(slip);
            } catch {
                /* ignore */
            }
            const code = encodeLeseSlip(slip);
            const box = host.querySelector('#ebConfirmResult');
            if (box) {
                box.hidden = false;
                box.innerHTML =
                    '<div class="eb-public-card eb-public-card--ok" style="margin-top:12px;padding:14px;">' +
                    '<h2 style="font-size:1.1rem;margin:0 0 8px;">Danke – gelesen</h2>' +
                    '<p class="eb-public-meta">Ihre Bestätigung wurde lokal vorgemerkt. Bestätigungscode (falls die Schule ihn braucht):</p>' +
                    '<textarea class="eb-public-code" readonly rows="3">' +
                    esc(code) +
                    '</textarea></div>';
            }
            btn.disabled = true;
            btn.textContent = 'Bereits bestätigt';
        });
    }
}
