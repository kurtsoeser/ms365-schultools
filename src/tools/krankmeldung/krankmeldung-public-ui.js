/**
 * Öffentliches Krankmeldungs-Formular – Klassenauswahl + optionaler Anhang.
 */

import {
    fileToDataUrl,
    saveAttachment,
    validateUploadFile
} from './krankmeldung-attachments.js';
import { resolvePublicKlassen } from './krankmeldung-klassen.js';
import { newId, UPLOAD_MAX_BYTES, toIsoDateOnly } from './krankmeldung-model.js';
import { KATEGORIE_LABEL } from './krankmeldung-model.js';
import {
    krankmeldungRemoteConfigured,
    remoteSubmitMeldung
} from './krankmeldung-remote.js';
import { loadKrankmeldungSettings } from './krankmeldung-settings.js';
import { submitPublicMeldung } from './krankmeldung-store.js';
import { pushKrankmeldungMeldung } from './krankmeldung-sync.js';

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

/**
 * @param {HTMLElement} host
 */
export function mountPublicForm(host) {
    if (!host) return;
    const settings = loadKrankmeldungSettings();
    const today = toIsoDateOnly(new Date());
    const klassen = resolvePublicKlassen();
    const allowFree = settings.allowFreeKlasse !== false;
    const allowAnhang = settings.allowAnhang !== false;
    const anhangPflicht = !!(settings.anhangPflicht && allowAnhang);

    const klasseField = klassen.length
        ? '<div><span class="km-field-label">Klasse wählen</span>' +
          '<div class="km-class-pick" id="kmPublicClassPick" role="group" aria-label="Klasse">' +
          klassen
              .map(
                  (code) =>
                      '<button type="button" class="km-class-chip" data-km-pub-klasse="' +
                      esc(code) +
                      '">' +
                      esc(code) +
                      '</button>'
              )
              .join('') +
          '</div>' +
          '<label>Gewählte Klasse<input name="klasseCode" id="kmPublicKlasse" required readonly placeholder="Bitte oben eine Klasse antippen"></label>' +
          (allowFree
              ? '<p class="km-hint">Klasse nicht gefunden? <button type="button" class="km-linkish" id="kmPublicKlasseFree">Andere Klasse eintippen</button></p>'
              : '') +
          '</div>'
        : '<label>Klasse<input name="klasseCode" id="kmPublicKlasse" required placeholder="z. B. 3AHW" autocomplete="off"></label>' +
          '<p class="km-hint">Tipp an die Schule: Klassenliste unter Einstellungen veröffentlichen oder den Formular-Link mit <code>?klassen=…</code> teilen.</p>';

    const anhangField = allowAnhang
        ? '<label>Anhang' +
          (anhangPflicht ? ' *' : ' (optional)') +
          '<input type="file" name="anhang" id="kmPublicAnhang" accept=".pdf,.png,.jpg,.jpeg,.gif,.webp,.doc,.docx,application/pdf,image/*"' +
          (anhangPflicht ? ' required' : '') +
          '>' +
          '<span class="km-hint">PDF, Foto oder Word – max. ' +
          Math.round(UPLOAD_MAX_BYTES / (1024 * 1024)) +
          ' MB' +
          (anhangPflicht ? ' · Pflicht' : ' · z. B. ärztliche Bestätigung') +
          '</span></label>'
        : '';

    host.innerHTML =
        '<div class="km-public-card">' +
        (settings.schoolName
            ? '<p class="km-public-school">' + esc(settings.schoolName) + '</p>'
            : '') +
        '<p class="km-public-kicker">Elternmeldung · ohne Anmeldung</p>' +
        '<h2>Krankmeldung</h2>' +
        '<p class="km-public-meta">Melden Sie Ihr Kind für die passende Klasse krank – ohne Anruf im Sekretariat.</p>' +
        '<form class="km-form" id="kmPublicForm" style="border:0;padding:0;max-width:none;">' +
        '<label>Name der Schülerin / des Schülers<input name="schuelerName" required autocomplete="name"></label>' +
        klasseField +
        '<div class="km-form__row">' +
        '<label>Von<input type="date" name="von" id="kmPubVon" required value="' +
        esc(today) +
        '"></label>' +
        '<label>Bis<input type="date" name="bis" id="kmPubBis" required value="' +
        esc(today) +
        '"></label></div>' +
        '<p class="km-hint">Tipp: Bei einem Tag reichen gleiche Von-/Bis-Daten. Lange Zeiträume bitte mit der Schule abstimmen.</p>' +
        '<label>Grund<select name="kategorie">' +
        Object.keys(KATEGORIE_LABEL)
            .map((k) => '<option value="' + esc(k) + '">' + esc(KATEGORIE_LABEL[k]) + '</option>')
            .join('') +
        '</select></label>' +
        anhangField +
        '<label>Ihr Name (optional)<input name="gemeldetVonName" autocomplete="name"></label>' +
        '<label>Ihre E-Mail (optional)<input type="email" name="gemeldetVonEmail" autocomplete="email"></label>' +
        '<label>Hinweis an die Schule<textarea name="notiz" rows="3" placeholder="Optional"></textarea></label>' +
        '<div class="km-form__actions"><button type="submit" class="btn btn-success">Absenden</button></div>' +
        '</form>' +
        '<p class="km-public-note">Einfache Meldung an die Schule – kein Ersatz für Untis-Entschuldigungen, wo die Schule das verlangt.</p>' +
        '</div>';

    const form = host.querySelector('#kmPublicForm');
    if (!form) return;

    const klasseInput = /** @type {HTMLInputElement|null} */ (host.querySelector('#kmPublicKlasse'));
    const vonEl = /** @type {HTMLInputElement|null} */ (host.querySelector('#kmPubVon'));
    const bisEl = /** @type {HTMLInputElement|null} */ (host.querySelector('#kmPubBis'));

    if (vonEl && bisEl) {
        vonEl.addEventListener('change', () => {
            if (!bisEl.value || bisEl.value < vonEl.value) bisEl.value = vonEl.value;
        });
    }

    host.querySelectorAll('[data-km-pub-klasse]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const code = String(btn.getAttribute('data-km-pub-klasse') || '');
            if (klasseInput) {
                klasseInput.value = code;
                klasseInput.readOnly = true;
            }
            host.querySelectorAll('[data-km-pub-klasse]').forEach((b) => {
                b.classList.toggle('is-on', b === btn);
            });
        });
    });

    const freeBtn = host.querySelector('#kmPublicKlasseFree');
    if (freeBtn && klasseInput) {
        freeBtn.addEventListener('click', () => {
            klasseInput.readOnly = false;
            klasseInput.value = '';
            klasseInput.focus();
            host.querySelectorAll('[data-km-pub-klasse]').forEach((b) => b.classList.remove('is-on'));
        });
    }

    form.addEventListener('submit', async (ev) => {
        ev.preventDefault();
        const fd = new FormData(form);
        const id = newId('km');
        /** @type {object} */
        const draft = {
            id,
            schuelerName: String(fd.get('schuelerName') || ''),
            klasseCode: String(fd.get('klasseCode') || (klasseInput && klasseInput.value) || ''),
            von: String(fd.get('von') || ''),
            bis: String(fd.get('bis') || ''),
            kategorie: String(fd.get('kategorie') || 'krankheit'),
            gemeldetVonName: String(fd.get('gemeldetVonName') || ''),
            gemeldetVonEmail: String(fd.get('gemeldetVonEmail') || ''),
            notiz: String(fd.get('notiz') || '')
        };

        form.querySelectorAll('.km-msg').forEach((el) => el.remove());
        if (!draft.klasseCode) {
            const warn = document.createElement('p');
            warn.className = 'km-msg km-msg--err';
            warn.setAttribute('role', 'alert');
            warn.textContent = 'Bitte eine Klasse wählen oder eintragen.';
            form.prepend(warn);
            return;
        }

        const fileInput = /** @type {HTMLInputElement|null} */ (host.querySelector('#kmPublicAnhang'));
        const file = fileInput && fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
        if (anhangPflicht && !file) {
            const warn = document.createElement('p');
            warn.className = 'km-msg km-msg--err';
            warn.textContent = 'Bitte einen Anhang hochladen.';
            form.prepend(warn);
            return;
        }
        if (file) {
            const v = validateUploadFile(file);
            if (!v.ok) {
                const warn = document.createElement('p');
                warn.className = 'km-msg km-msg--err';
                warn.textContent = v.error || 'Anhang ungültig.';
                form.prepend(warn);
                return;
            }
            try {
                const dataUrl = await fileToDataUrl(file);
                await saveAttachment(id, {
                    name: file.name,
                    mime: file.type || 'application/octet-stream',
                    size: file.size,
                    dataUrl
                });
                draft.anhangKey = id;
                draft.anhangName = file.name;
                draft.anhangMime = file.type || '';
                draft.anhangSize = file.size;
            } catch {
                const warn = document.createElement('p');
                warn.className = 'km-msg km-msg--err';
                warn.textContent = 'Anhang konnte nicht gespeichert werden.';
                form.prepend(warn);
                return;
            }
        }

        const submitBtn = /** @type {HTMLButtonElement|null} */ (
            form.querySelector('button[type="submit"]')
        );
        if (submitBtn) {
            submitBtn.disabled = true;
            submitBtn.textContent = 'Wird gesendet …';
        }

        let remoteOk = false;
        let remoteTried = false;
        let remoteErr = '';
        if (krankmeldungRemoteConfigured()) {
            remoteTried = true;
            try {
                await remoteSubmitMeldung(draft);
                remoteOk = true;
            } catch (e) {
                const raw = e && e.message ? String(e.message) : 'Server nicht erreichbar';
                remoteErr = /403|401|Forbidden|Unauthorized|Zugriff/i.test(raw)
                    ? 'Zugriff auf den Server nicht möglich'
                    : /Failed to fetch|NetworkError|Load failed/i.test(raw)
                      ? 'Server nicht erreichbar'
                      : raw;
            }
        }

        const local = submitPublicMeldung(draft);
        if (!local.ok && !remoteOk) {
            if (submitBtn) {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Absenden';
            }
            host.innerHTML =
                '<div class="km-public-card km-public-card--err"><h2>Nicht gesendet</h2><p>' +
                esc((local.errors || []).join(' ') || 'Bitte Angaben prüfen.') +
                (remoteErr ? '<br>' + esc(remoteErr) : '') +
                '</p>' +
                '<div class="km-form__actions"><button type="button" class="btn" id="kmPubBack">Zurück</button></div></div>';
            host.querySelector('#kmPubBack')?.addEventListener('click', () => mountPublicForm(host));
            return;
        }

        // Optional SPO (Kiosk mit angemeldeter Schul-Session / Graph-Hilfen).
        let spoOk = false;
        let spoTried = false;
        if (local.ok && local.item && typeof window !== 'undefined' && window.ms365SpoGraph) {
            spoTried = true;
            try {
                const push = await pushKrankmeldungMeldung(local.item);
                spoOk = !!(push.ok && !push.skipped);
            } catch {
                spoOk = false;
            }
        }

        const warnParts = [];
        if (local.duplicate) {
            warnParts.push(
                'Es gibt bereits eine ähnliche Meldung für diesen Zeitraum – die Schule sieht beides.'
            );
        }
        for (const w of local.warnings || []) warnParts.push(w);
        if (remoteTried && !remoteOk) {
            warnParts.push(
                'Hinweis: Server-Übermittlung fehlgeschlagen (' +
                    (remoteErr || 'unbekannt') +
                    ') – lokal gespeichert, falls das Sekretariat denselben Browser nutzt.'
            );
        }

        const deliveryHint = spoOk
            ? 'Die Meldung liegt in SharePoint – die Schule hat sie erhalten.'
            : remoteOk
              ? 'Die Schule hat die Meldung über den Server erhalten.'
              : remoteTried
                ? ''
                : spoTried
                  ? 'SharePoint-Sync nicht möglich (kein Login) – Sekretariat sieht die Meldung in diesem Browser bzw. nach Sync.'
                  : 'Kein Server hinterlegt – Sekretariat sieht die Meldung in diesem Browser bzw. nach SharePoint-Sync.';

        host.innerHTML =
            '<div class="km-public-card km-public-card--ok">' +
            '<h2>Danke – Meldung eingegangen</h2>' +
            '<p>Krankmeldung für <strong>' +
            esc(draft.schuelerName) +
            '</strong> in Klasse <strong>' +
            esc(draft.klasseCode) +
            '</strong><br>Zeitraum: ' +
            esc(formatDeDate(draft.von)) +
            (draft.bis !== draft.von ? ' bis ' + esc(formatDeDate(draft.bis)) : '') +
            (draft.anhangName ? '<br>Anhang: ' + esc(draft.anhangName) : '') +
            '.</p>' +
            (warnParts.length
                ? '<p class="km-public-meta">' + esc(warnParts.join(' ')) + '</p>'
                : '') +
            (deliveryHint
                ? '<p class="km-public-meta">' + esc(deliveryHint) + '</p>'
                : '') +
            '<div class="km-form__actions"><button type="button" class="btn" id="kmPubAgain">Weitere Meldung</button></div>' +
            '</div>';
        host.querySelector('#kmPubAgain')?.addEventListener('click', () => mountPublicForm(host));
    });
}
