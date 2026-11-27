/**
 * Schüler-Primärpfad: Wunschreihenfolge für Wahlfach-Runde (schlank, wahl-portal).
 */

import { hydrateStammdatenFromLiveSite } from '../../shared/school-stammdaten-spo-lists.js';
import {
    assertWzOwnership,
    isAnmeldefensterOffen,
    resolveWzOwnership,
    resolveWzPhase
} from './wahlfach-zuteilung-logic.js';
import { isWzDemoSeedState } from './wahlfach-zuteilung-demo.js';
import { RUNDE_STATUS_LABEL } from './wahlfach-zuteilung-model.js';
import {
    loadState,
    saveState,
    seedDemoIfEmpty,
    upsertAnmeldung
} from './wahlfach-zuteilung-store.js';
import { softPullWz, softPushWzAnmeldung } from './wahlfach-zuteilung-sync.js';
import { announceAppRole } from '../../shared/app-role-chrome.js';
import {
    canonicalKlasse,
    classCodesFromStammdaten,
    findOwnAnmeldung,
    loadWzStammdaten,
    readWzAccount,
    resolveWzSiteUrl,
    wantsWzDemoFromUrl
} from './wahlfach-zuteilung-state.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function angebotOptions(angebote, selected, emptyLabel) {
    const opts = ['<option value="">' + esc(emptyLabel) + '</option>'];
    (angebote || [])
        .filter((a) => a.aktiv)
        .forEach((a) => {
            opts.push(
                '<option value="' +
                    esc(a.id) +
                    '"' +
                    (selected === a.id ? ' selected' : '') +
                    '>' +
                    esc(a.bezeichnung) +
                    (a.zielJahrgaenge ? ' (' + esc(a.zielJahrgaenge) + ')' : '') +
                    '</option>'
            );
        });
    return opts.join('');
}

function phaseBadge(runde) {
    const p = resolveWzPhase(runde);
    const cls =
        p.tone === 'ok' ? 'wz-badge--ok' : p.tone === 'warn' ? 'wz-badge--warn' : '';
    return '<span class="wz-badge ' + cls + '">' + esc(p.label) + '</span>';
}

/**
 * @param {HTMLElement} root
 */
export function mountWahlfachAnmeldung(root) {
    if (!root) return;

    /** @type {{ type: string, text: string }|null} */
    let message = null;
    let softError = '';
    let localDemoOnly = false;
    let stammdaten = loadWzStammdaten();
    let prefill = {
        schuelerName: '',
        klasse: '',
        schuelerEmail: '',
        matchedStudent: false,
        wunsch1Id: '',
        wunsch2Id: '',
        wunsch3Id: ''
    };
    let submitting = false;
    let editing = false;

    function accountCtx() {
        const a = readWzAccount();
        return { accountEmail: a.email, accountName: a.name, students: stammdaten.students || [] };
    }

    function applyOwnershipPrefill() {
        const account = readWzAccount();
        const saved = findOwnAnmeldung(loadState(), account.email);
        const own = resolveWzOwnership(
            {
                schuelerName: (saved && saved.schuelerName) || prefill.schuelerName,
                klasse: (saved && saved.klasse) || prefill.klasse,
                schuelerEmail: account.email
            },
            accountCtx()
        );
        const codes = classCodesFromStammdaten(stammdaten);
        prefill = {
            schuelerName: own.schuelerName || prefill.schuelerName,
            klasse: canonicalKlasse(own.klasse || prefill.klasse, codes),
            schuelerEmail: own.schuelerEmail || prefill.schuelerEmail,
            matchedStudent: !!own.matchedStudent,
            wunsch1Id: (saved && saved.wunsch1Id) || prefill.wunsch1Id || '',
            wunsch2Id: (saved && saved.wunsch2Id) || prefill.wunsch2Id || '',
            wunsch3Id: (saved && saved.wunsch3Id) || prefill.wunsch3Id || ''
        };
    }

    function offerName(state, id) {
        if (!id) return '';
        const hit = (state.angebote || []).find((a) => a && a.id === id);
        return hit ? hit.bezeichnung : '';
    }

    async function hydrateStammdaten(siteId) {
        stammdaten = loadWzStammdaten();
        if (!siteId || localDemoOnly) return;
        const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
        if (!api || typeof api.getGraphToken !== 'function') return;
        try {
            const tok = await api.getGraphToken([
                'https://graph.microsoft.com/User.Read',
                'https://graph.microsoft.com/Sites.Read.All'
            ]);
            const result = await hydrateStammdatenFromLiveSite({
                token: tok,
                siteId: String(siteId),
                graph: api,
                local: loadWzStammdaten()
            });
            stammdaten = result.stammdaten;
        } catch {
            /* local-first */
        }
    }

    async function bootData() {
        if (wantsWzDemoFromUrl()) {
            seedDemoIfEmpty();
            localDemoOnly = true;
            softError = '';
            return;
        }
        const local = loadState();
        if (isWzDemoSeedState(local)) {
            localDemoOnly = true;
            return;
        }
        const siteUrl = resolveWzSiteUrl();
        if (!siteUrl) return;
        const pull = await softPullWz({ localDemoOnly: false, siteUrl });
        if (pull.ok && pull.state) {
            /* Schüler: SPO-Runde/Angebote übernehmen, lokale Anmeldungen nicht blind wipen wenn SPO leer */
            const cur = loadState();
            const next = pull.state;
            if (!(next.anmeldungen && next.anmeldungen.length) && cur.anmeldungen.length) {
                next.anmeldungen = cur.anmeldungen;
                next.zuteilungen = cur.zuteilungen;
                saveState(next);
            }
            await hydrateStammdaten(pull.siteId || '');
        } else if (pull.softError) {
            softError = pull.softError;
        }
    }

    function klasseField() {
        const codes = classCodesFromStammdaten(stammdaten);
        if (!codes.length) {
            return (
                '<label class="wz-field">Klasse *' +
                '<input name="klasse" id="wzSchKlasse" required placeholder="z. B. 3AK" autocomplete="organization-title" value="' +
                esc(prefill.klasse) +
                '">' +
                '</label>'
            );
        }
        const opts = ['<option value="">– bitte wählen –</option>'].concat(
            codes.map(
                (c) =>
                    '<option value="' +
                    esc(c) +
                    '"' +
                    (prefill.klasse === c ? ' selected' : '') +
                    '>' +
                    esc(c) +
                    '</option>'
            )
        );
        return (
            '<label class="wz-field">Klasse *' +
            '<select name="klasse" id="wzSchKlasse" required>' +
            opts.join('') +
            '</select></label>'
        );
    }

    function render() {
        const s = loadState();
        const r = s.runde;
        const offen = r && isAnmeldefensterOffen(r);
        const angebote = (s.angebote || []).filter((a) => a.aktiv);
        const phase = resolveWzPhase(r);

        const account = readWzAccount();
        announceAppRole({
            appId: 'wahlfach-anmeldung',
            role: 'schueler',
            source: localDemoOnly ? 'demo' : '',
            email: account.email,
            demoMode: localDemoOnly
        });

        root.innerHTML =
            '<div class="wz-shell wz-shell--form"><div class="wz-main"><section class="wz-student">' +
            '<header class="wz-student__hero">' +
            '<p class="wz-student__kicker">Wahlfach-Anmeldung</p>' +
            '<h1 class="wz-student__title">' +
            esc((r && r.titel) || 'Wahlfächer') +
            '</h1>' +
            (r
                ? '<p class="wz-student__meta">Schuljahr ' +
                  esc(r.schuljahr) +
                  ' · ' +
                  esc(RUNDE_STATUS_LABEL[r.status] || r.status) +
                  '</p><p class="wz-student__meta">' +
                  phaseBadge(r) +
                  '</p>'
                : '<p class="wz-student__meta">Noch keine Runde eingerichtet.</p>') +
            '</header>' +
            (softError
                ? '<div class="wz-alert wz-alert--info" role="status">' + esc(softError) + '</div>'
                : '') +
            (message
                ? '<p class="wz-msg ' +
                  (message.type === 'ok' ? 'wz-msg--ok' : 'wz-msg--err') +
                  '" role="status">' +
                  esc(message.text) +
                  '</p>'
                : '') +
            studentBody(s, r, offen, angebote, phase) +
            '</section></div></div>';

        const form = root.querySelector('#wzSchuelerForm');
        if (form) {
            form.addEventListener('submit', (ev) => {
                ev.preventDefault();
                void submitForm(form);
            });
            const w1 = form.querySelector('#wzSchW1');
            const w2 = form.querySelector('#wzSchW2');
            const w3 = form.querySelector('#wzSchW3');
            const clearDup = (sel, others) => {
                if (!sel) return;
                sel.addEventListener('change', () => {
                    if (sel.value && others.some((o) => o && o.value === sel.value)) sel.value = '';
                });
            };
            clearDup(w2, [w1]);
            clearDup(w3, [w1, w2]);
            clearDup(w1, [w2, w3]);
        }
        const edit = root.querySelector('#wzEditWishes');
        if (edit) {
            edit.addEventListener('click', () => {
                editing = true;
                message = null;
                render();
            });
        }
        const cancel = root.querySelector('#wzCancelEdit');
        if (cancel) {
            cancel.addEventListener('click', () => {
                editing = false;
                message = null;
                render();
            });
        }
    }

    function identityFields() {
        if (prefill.matchedStudent && prefill.schuelerName && prefill.klasse) {
            return (
                '<div class="wz-student__grid">' +
                '<p class="wz-field">Name<strong>' +
                esc(prefill.schuelerName) +
                '</strong></p>' +
                '<p class="wz-field">Klasse<strong>' +
                esc(prefill.klasse) +
                '</strong></p>' +
                '<input type="hidden" name="schuelerName" value="' +
                esc(prefill.schuelerName) +
                '">' +
                '<input type="hidden" name="klasse" value="' +
                esc(prefill.klasse) +
                '">' +
                '</div>'
            );
        }
        return (
            '<div class="wz-student__grid">' +
            '<label class="wz-field">Name *' +
            '<input name="schuelerName" id="wzSchName" required placeholder="Vor- und Nachname" autocomplete="name" value="' +
            esc(prefill.schuelerName) +
            '">' +
            '</label>' +
            klasseField() +
            '</div>'
        );
    }

    function statusHtml(s, own, offen) {
        const z = (s.zuteilungen || []).find((x) => x && x.anmeldungId === own.id);
        const zugeteilt = z && z.angebotId ? offerName(s, z.angebotId) : '';
        const line = (n, id) =>
            '<li><span>' + n + '. Wunsch</span> ' + esc(offerName(s, id) || '–') + '</li>';
        return (
            '<div class="wz-student__status">' +
            '<p class="wz-student__kicker">Deine Anmeldung</p>' +
            '<ol class="wz-student__wishes-done">' +
            line('1', own.wunsch1Id) +
            line('2', own.wunsch2Id) +
            line('3', own.wunsch3Id) +
            '</ol>' +
            '<p class="wz-student__result"><strong>' +
            (zugeteilt ? 'Zugeteilt: ' + esc(zugeteilt) : 'Noch nicht zugeteilt') +
            '</strong></p>' +
            (offen
                ? '<div class="wz-actions"><button type="button" class="fb-btn fb-btn--ghost" id="wzEditWishes">Wünsche ändern</button></div>'
                : '<p class="wz-student__foot">Das Anmeldefenster ist zu. Die Schulleitung teilt zu.</p>') +
            '</div>'
        );
    }

    function studentBody(s, r, offen, angebote, phase) {
        const own = findOwnAnmeldung(s, prefill.schuelerEmail || readWzAccount().email);
        if (!r) {
            return '<p class="wz-quiet-empty" role="status"><i class="bi bi-hourglass" aria-hidden="true"></i><span>Sobald die Schule eine Runde freigibt, erscheint hier das Formular.</span></p>';
        }
        if (own && !editing) return statusHtml(s, own, offen);
        if (!offen) {
            return (
                '<div class="wz-student__closed"><i class="bi bi-lock" aria-hidden="true"></i><div><p><strong>' +
                esc(phase.label) +
                '</strong></p><p>Das Anmeldefenster ist zu.</p></div></div>'
            );
        }
        if (!angebote.length) return '<p class="wz-quiet-empty">Noch keine aktiven Angebote.</p>';
        return (
            '<form class="wz-student__form" id="wzSchuelerForm">' +
            '<fieldset class="wz-student__block">' +
            '<legend>Ihre Angaben</legend>' +
            identityFields() +
            '</fieldset>' +
            '<fieldset class="wz-student__block">' +
            '<legend>Wunschreihenfolge</legend>' +
            '<p class="wz-student__hint">Der 1. Wunsch ist Pflicht. 2. und 3. Wunsch helfen, wenn ein Fach voll ist.</p>' +
            '<div class="wz-student__wishes">' +
            '<label class="wz-field wz-field--wish"><span class="wz-wish-n">1</span> 1. Wunsch *' +
            '<select name="wunsch1Id" id="wzSchW1" required>' +
            angebotOptions(angebote, prefill.wunsch1Id, '– bitte wählen –') +
            '</select></label>' +
            '<label class="wz-field wz-field--wish"><span class="wz-wish-n">2</span> 2. Wunsch' +
            '<select name="wunsch2Id" id="wzSchW2">' +
            angebotOptions(angebote, prefill.wunsch2Id, '– optional –') +
            '</select></label>' +
            '<label class="wz-field wz-field--wish"><span class="wz-wish-n">3</span> 3. Wunsch' +
            '<select name="wunsch3Id" id="wzSchW3">' +
            angebotOptions(angebote, prefill.wunsch3Id, '– optional –') +
            '</select></label>' +
            '</div></fieldset>' +
            '<div class="wz-actions">' +
            '<button type="submit" class="fb-btn" id="wzSchSubmit"' +
            (submitting ? ' disabled' : '') +
            '><i class="bi bi-send"></i>' +
            (submitting ? 'Sende …' : own ? 'Wünsche speichern' : 'Anmeldung absenden') +
            '</button>' +
            (own
                ? '<button type="button" class="fb-btn fb-btn--ghost" id="wzCancelEdit">Zurück</button>'
                : '') +
            '</div>' +
            '<p class="wz-student__foot">Die Schulleitung teilt danach zu.</p>' +
            '</form>'
        );
    }

    async function submitForm(form) {
        const fd = new FormData(form);
        const w1 = String(fd.get('wunsch1Id') || '');
        const w2 = String(fd.get('wunsch2Id') || '');
        const w3 = String(fd.get('wunsch3Id') || '');
        if (w2 && w2 === w1) {
            message = { type: 'err', text: '2. Wunsch darf nicht gleich dem 1. Wunsch sein.' };
            render();
            return;
        }
        if (w3 && (w3 === w1 || w3 === w2)) {
            message = { type: 'err', text: 'Wünsche müssen unterschiedlich sein.' };
            render();
            return;
        }

        let ownership;
        try {
            ownership = resolveWzOwnership(
                {
                    schuelerName: fd.get('schuelerName'),
                    klasse: fd.get('klasse'),
                    schuelerEmail: prefill.schuelerEmail
                },
                accountCtx()
            );
            assertWzOwnership(ownership, { requireEmail: !localDemoOnly });
        } catch (e) {
            message = { type: 'err', text: (e && e.message) || String(e) };
            render();
            return;
        }

        const draft = {
            schuelerName: ownership.schuelerName || fd.get('schuelerName'),
            klasse: canonicalKlasse(
                ownership.klasse || fd.get('klasse'),
                classCodesFromStammdaten(stammdaten)
            ),
            schuelerEmail: ownership.schuelerEmail,
            wunsch1Id: w1,
            wunsch2Id: w2 || null,
            wunsch3Id: w3 || null,
            status: 'eingereicht'
        };

        submitting = true;
        render();

        const res = upsertAnmeldung(draft);
        if (!res.ok) {
            submitting = false;
            message = { type: 'err', text: res.errors.join(' ') };
            render();
            return;
        }

        if (!localDemoOnly) {
            const push = await softPushWzAnmeldung(res.anmeldung, { localDemoOnly: false });
            if (!push.ok && !push.skipped) {
                softError =
                    push.softError ||
                    'Lokal gespeichert. SharePoint-Sync nicht möglich – Daten bleiben erhalten.';
                /* kein Wipe */
            }
        }

        submitting = false;
        editing = false;
        applyOwnershipPrefill();
        message = {
            type: 'ok',
            text: 'Anmeldung gespeichert. Die Schulleitung teilt danach zu.'
        };
        prefill = {
            schuelerName: ownership.schuelerName,
            klasse: ownership.klasse,
            schuelerEmail: ownership.schuelerEmail
        };
        render();
    }

    root.innerHTML =
        '<div class="wz-shell wz-shell--form"><div class="wz-main"><p class="wz-nav__hint">Lade Anmeldung …</p></div></div>';
    void bootData()
        .then(() => {
            applyOwnershipPrefill();
            render();
        })
        .catch(() => {
            softError = 'Laden eingeschränkt – lokaler Stand wird genutzt.';
            applyOwnershipPrefill();
            render();
        });
}
