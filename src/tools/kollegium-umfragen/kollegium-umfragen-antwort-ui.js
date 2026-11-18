/**
 * Lehrkräfte-Antwort (MSAL) – inkl. Datei-Fragen / optionalem Anhang.
 */

import {
    fileToDataUrl,
    saveAttachment,
    validateUploadFile
} from './kollegium-umfragen-attachments.js';
import {
    teacherInAudience,
    umfrageNochOffen
} from './kollegium-umfragen-logic.js';
import { FRAGE_TYP, STATUS, UPLOAD_MAX_BYTES, newId } from './kollegium-umfragen-model.js';
import {
    getUmfrage,
    listAntworten,
    loadState,
    seedDemoIfEmpty,
    submitAntwort
} from './kollegium-umfragen-store.js';
import { loadKollegiumSettings } from './kollegium-umfragen-settings.js';
import {
    pullKollegiumFromSharePoint,
    pushKollegiumAntwort
} from './kollegium-umfragen-sync.js';
import {
    kollegiumTeacherCtxFromCache,
    resolveKollegiumRoles
} from './kollegium-umfragen-entra-role.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function accountInfo() {
    if (typeof window.ms365AuthGetAccountInfo === 'function') {
        return window.ms365AuthGetAccountInfo() || null;
    }
    return null;
}

function isLoggedIn() {
    return typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
}

function login() {
    if (typeof window.ms365AuthLogin === 'function') window.ms365AuthLogin();
    else if (typeof window.ms365AuthLoginRedirect === 'function') window.ms365AuthLoginRedirect();
}

/**
 * @param {HTMLElement} host
 * @param {{ umfrageId?: string }} [opts]
 */
export function mountAntwortApp(host, opts = {}) {
    if (!host) return;
    seedDemoIfEmpty();

    const params = new URLSearchParams(location.search);
    const preferId = String(opts.umfrageId || params.get('id') || '').trim();

    function render() {
        const settings = loadKollegiumSettings();
        if (!isLoggedIn()) {
            host.innerHTML =
                '<div class="ku-antwort-card">' +
                '<h2>Kollegiumsumfrage</h2>' +
                '<p class="ku-antwort-meta">Bitte mit dem schulischen Microsoft-Konto anmelden – kein öffentlicher Link.</p>' +
                '<button type="button" class="btn btn-success" id="kuLoginBtn">Anmelden</button>' +
                '</div>';
            host.querySelector('#kuLoginBtn')?.addEventListener('click', () => login());
            return;
        }

        const acc = accountInfo() || {};
        const email = String(acc.username || acc.upn || '')
            .trim()
            .toLowerCase();
        const name = String(acc.name || '').trim();
        const state = loadState();
        const teacherCtx = kollegiumTeacherCtxFromCache(email);
        const aktiv = state.umfragen.filter((u) => u.status === STATUS.AKTIV && umfrageNochOffen(u));
        let current = preferId ? getUmfrage(preferId) : null;
        if (current && (current.status !== STATUS.AKTIV || !umfrageNochOffen(current))) {
            current = null;
        }
        if (!current) {
            current = aktiv.find((u) => teacherInAudience(u, teacherCtx)) || null;
        }

        if (!current) {
            host.innerHTML =
                '<div class="ku-antwort-card"><h2>Keine aktive Umfrage</h2>' +
                '<p class="ku-antwort-meta">Aktuell gibt es nichts Offenes für Sie' +
                (email ? ' (' + esc(email) + ')' : '') +
                (teacherCtx.isKv ? ' · KV' : '') +
                '.</p></div>';
            return;
        }

        const inAudience = teacherInAudience(current, teacherCtx);
        if (!inAudience) {
            host.innerHTML =
                '<div class="ku-antwort-card"><h2>' +
                esc(current.titel) +
                '</h2><p class="ku-antwort-meta">Diese Umfrage ist nicht für Ihr Konto freigeschaltet' +
                (current.zielgruppe === 'klassenvorstand' && !teacherCtx.isKv
                    ? ' (nur Klassenvorstände – in Stammdaten als KV hinterlegen).'
                    : '.') +
                '</p></div>';
            return;
        }

        const existing = listAntworten(current.id).find((a) => a.lehrerEmail === email);
        if (existing && existing.status === 'abgegeben') {
            host.innerHTML =
                '<div class="ku-antwort-card"><h2>Bereits abgegeben</h2>' +
                '<p class="ku-antwort-meta">Danke – Ihre Antwort zu „' +
                esc(current.titel) +
                '“ wurde gespeichert' +
                (current.anonym ? ' (anonyme Auswertung).' : '.') +
                (existing.anhangName ? '<br>Anhang: ' + esc(existing.anhangName) : '') +
                '</p></div>';
            return;
        }

        const picker =
            aktiv.length > 1
                ? '<label style="display:block;margin-bottom:12px;font-size:0.85rem;">Andere Umfrage ' +
                  '<select id="kuPick">' +
                  aktiv
                      .map(
                          (u) =>
                              '<option value="' +
                              esc(u.id) +
                              '"' +
                              (u.id === current.id ? ' selected' : '') +
                              '>' +
                              esc(u.titel) +
                              '</option>'
                      )
                      .join('') +
                  '</select></label>'
                : '';

        const fragen = current.fragen
            .map((f) => {
                let body = '';
                if (f.typ === FRAGE_TYP.SINGLE) {
                    body = (f.optionen || [])
                        .map(
                            (o) =>
                                '<label class="ku-check"><input type="radio" name="q_' +
                                esc(f.id) +
                                '" value="' +
                                esc(o) +
                                '"' +
                                (f.pflicht ? ' required' : '') +
                                '> ' +
                                esc(o) +
                                '</label>'
                        )
                        .join('');
                } else if (f.typ === FRAGE_TYP.MULTI) {
                    body = (f.optionen || [])
                        .map(
                            (o) =>
                                '<label class="ku-check"><input type="checkbox" name="q_' +
                                esc(f.id) +
                                '" value="' +
                                esc(o) +
                                '"> ' +
                                esc(o) +
                                '</label>'
                        )
                        .join('');
                } else if (f.typ === FRAGE_TYP.SKALA) {
                    const opts = [];
                    for (let n = f.skalaMin; n <= f.skalaMax; n++) opts.push(n);
                    body =
                        '<p class="ku-antwort-meta" style="margin-bottom:6px;">' +
                        esc(f.skalaLabelMin || String(f.skalaMin)) +
                        ' → ' +
                        esc(f.skalaLabelMax || String(f.skalaMax)) +
                        '</p>' +
                        opts
                            .map(
                                (n) =>
                                    '<label class="ku-check"><input type="radio" name="q_' +
                                    esc(f.id) +
                                    '" value="' +
                                    n +
                                    '"' +
                                    (f.pflicht ? ' required' : '') +
                                    '> ' +
                                    n +
                                    '</label>'
                            )
                            .join('');
                } else if (f.typ === FRAGE_TYP.DATEI) {
                    body =
                        '<input type="file" name="q_' +
                        esc(f.id) +
                        '" accept=".pdf,.png,.jpg,.jpeg,.gif,.webp,.doc,.docx,application/pdf,image/*,.xls,.xlsx,.ppt,.pptx"' +
                        (f.pflicht ? ' required' : '') +
                        ' style="width:100%;font:inherit;">' +
                        '<p class="ku-antwort-meta">Max. ' +
                        Math.round(UPLOAD_MAX_BYTES / (1024 * 1024)) +
                        ' MB</p>';
                } else {
                    body =
                        '<textarea name="q_' +
                        esc(f.id) +
                        '" rows="3"' +
                        (f.pflicht ? ' required' : '') +
                        ' style="width:100%;font:inherit;padding:8px;border:1px solid #cbd5e1;border-radius:8px;"></textarea>';
                }
                return (
                    '<div class="ku-frage-block"><fieldset><legend>' +
                    esc(f.text) +
                    (f.pflicht ? ' *' : '') +
                    '</legend>' +
                    body +
                    '</fieldset></div>'
                );
            })
            .join('');

        const extraAnhang = current.anhangErlaubt
            ? '<label style="display:grid;gap:4px;font-size:0.85rem;font-weight:600;margin:12px 0;">Zusätzlicher Anhang (optional)' +
              '<input type="file" name="extraAnhang" accept=".pdf,.png,.jpg,.jpeg,.gif,.webp,.doc,.docx,application/pdf,image/*">' +
              '</label>'
            : '';

        host.innerHTML =
            '<div class="ku-antwort-card">' +
            (settings.schoolName
                ? '<p class="ku-antwort-meta" style="color:#0369a1;font-weight:600;">' +
                  esc(settings.schoolName) +
                  '</p>'
                : '') +
            picker +
            '<h2>' +
            esc(current.titel) +
            '</h2>' +
            '<p class="ku-antwort-meta">' +
            esc(current.beschreibung || 'Bitte beantworten Sie die Fragen.') +
            '<br>Angemeldet als ' +
            esc(name || email) +
            (current.anonym ? ' · Auswertung anonym' : '') +
            (current.faelligAm ? ' · Frist ' + esc(current.faelligAm) : '') +
            '</p>' +
            '<form id="kuAntwortForm">' +
            fragen +
            extraAnhang +
            '<div class="ku-form__actions"><button type="submit" class="btn btn-success">Absenden</button></div>' +
            '</form></div>';

        host.querySelector('#kuPick')?.addEventListener('change', (ev) => {
            const id = ev.target.value;
            history.replaceState({}, '', '?id=' + encodeURIComponent(id));
            mountAntwortApp(host, { umfrageId: id });
        });

        host.querySelector('#kuAntwortForm')?.addEventListener('submit', async (ev) => {
            ev.preventDefault();
            const form = /** @type {HTMLFormElement} */ (ev.target);
            const antwortId = newId('ka');
            /** @type {Record<string, unknown>} */
            const werte = {};
            /** @type {{ name: string, mime: string, size: number, key: string }|null} */
            let firstFileMeta = null;

            try {
                for (const f of current.fragen) {
                    if (f.typ === FRAGE_TYP.MULTI) {
                        werte[f.id] = [
                            ...form.querySelectorAll('[name="q_' + f.id + '"]:checked')
                        ].map((el) => /** @type {HTMLInputElement} */ (el).value);
                    } else if (f.typ === FRAGE_TYP.TEXT) {
                        werte[f.id] = String(
                            /** @type {HTMLInputElement|null} */ (
                                form.querySelector('[name="q_' + f.id + '"]')
                            )?.value || ''
                        );
                    } else if (f.typ === FRAGE_TYP.DATEI) {
                        const input = /** @type {HTMLInputElement|null} */ (
                            form.querySelector('[name="q_' + f.id + '"]')
                        );
                        const file = input && input.files && input.files[0] ? input.files[0] : null;
                        if (!file) {
                            werte[f.id] = '';
                            continue;
                        }
                        const v = validateUploadFile(file);
                        if (!v.ok) throw new Error(v.error || 'Datei ungültig.');
                        const key = antwortId + '-' + f.id;
                        const dataUrl = await fileToDataUrl(file);
                        await saveAttachment(key, {
                            name: file.name,
                            mime: file.type || 'application/octet-stream',
                            size: file.size,
                            dataUrl
                        });
                        const meta = {
                            name: file.name,
                            mime: file.type || '',
                            size: file.size,
                            key
                        };
                        werte[f.id] = meta;
                        if (!firstFileMeta) firstFileMeta = meta;
                    } else {
                        const el = form.querySelector('[name="q_' + f.id + '"]:checked');
                        werte[f.id] = el ? /** @type {HTMLInputElement} */ (el).value : '';
                        if (f.typ === FRAGE_TYP.SKALA && werte[f.id] !== '') {
                            werte[f.id] = Number(werte[f.id]);
                        }
                    }
                }

                const extra = /** @type {HTMLInputElement|null} */ (
                    form.querySelector('[name="extraAnhang"]')
                );
                if (extra && extra.files && extra.files[0]) {
                    const file = extra.files[0];
                    const v = validateUploadFile(file);
                    if (!v.ok) throw new Error(v.error || 'Anhang ungültig.');
                    const dataUrl = await fileToDataUrl(file);
                    await saveAttachment(antwortId, {
                        name: file.name,
                        mime: file.type || 'application/octet-stream',
                        size: file.size,
                        dataUrl
                    });
                    firstFileMeta = {
                        name: file.name,
                        mime: file.type || '',
                        size: file.size,
                        key: antwortId
                    };
                }
            } catch (e) {
                host.innerHTML =
                    '<div class="ku-antwort-card"><h2>Nicht gespeichert</h2><p>' +
                    esc(e && e.message ? e.message : String(e)) +
                    '</p><button type="button" class="btn" id="kuBack">Zurück</button></div>';
                host.querySelector('#kuBack')?.addEventListener('click', () =>
                    mountAntwortApp(host, { umfrageId: current.id })
                );
                return;
            }

            const res = submitAntwort({
                id: antwortId,
                umfrageId: current.id,
                lehrerEmail: email,
                lehrerName: name,
                werte,
                anhangName: firstFileMeta ? firstFileMeta.name : '',
                anhangMime: firstFileMeta ? firstFileMeta.mime : '',
                anhangSize: firstFileMeta ? firstFileMeta.size : 0,
                anhangKey: firstFileMeta ? firstFileMeta.key : ''
            });
            if (!res.ok) {
                host.innerHTML =
                    '<div class="ku-antwort-card"><h2>Nicht gespeichert</h2><p>' +
                    esc((res.errors || []).join(' ')) +
                    '</p><button type="button" class="btn" id="kuBack">Zurück</button></div>';
                host.querySelector('#kuBack')?.addEventListener('click', () =>
                    mountAntwortApp(host, { umfrageId: current.id })
                );
                return;
            }
            host.innerHTML =
                '<div class="ku-antwort-card"><h2>Danke!</h2><p class="ku-antwort-meta">Ihre Antwort wurde lokal gespeichert' +
                (firstFileMeta ? ' inkl. Datei „' + esc(firstFileMeta.name) + '“' : '') +
                ' – Sync nach SharePoint …</p></div>';
            void pushKollegiumAntwort(res.item).then((push) => {
                host.innerHTML =
                    '<div class="ku-antwort-card"><h2>Danke!</h2><p class="ku-antwort-meta">' +
                    (push.ok
                        ? 'Ihre Antwort wurde in SharePoint gespeichert' +
                          (firstFileMeta ? ' inkl. Datei „' + esc(firstFileMeta.name) + '“' : '') +
                          '.'
                        : 'Lokal gespeichert. ' +
                          esc(push.error || 'SharePoint-Sync später erneut versuchen.')) +
                    '</p></div>';
            });
        });
    }

    void pullKollegiumFromSharePoint().then(() => {
        render();
        const acc0 = accountInfo() || {};
        const em0 = String(acc0.username || acc0.upn || '')
            .trim()
            .toLowerCase();
        void resolveKollegiumRoles(em0).then(() => render());
    });
    render();
}
