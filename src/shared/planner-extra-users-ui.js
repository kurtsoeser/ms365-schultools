/**
 * Einzelpersonen-Listen für Planer-Entra-Berechtigungen (Setup + Administration).
 */
import { pickEntraUser } from './entra-user-picker.js';
import {
    normalizePlannerUsers,
    mergePlannerUser,
    direktionUsersFromTenantStammdaten
} from '../tools/freistellung-planer/freistellung-planer-direktion-users.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatUserLine(u) {
    const name = String(u.displayName || '').trim();
    const mail = String(u.mail || '').trim();
    if (name && mail && name.toLowerCase() !== mail.toLowerCase()) {
        return escapeHtml(name) + ' <span class="fr-setup-user-row__mail">(' + escapeHtml(mail) + ')</span>';
    }
    return escapeHtml(name || mail);
}

/**
 * @param {Array<{ configKey: string, listId: string, addBtnId: string, stammdatenBtnId?: string, pickTitle: string, pickHint: string, emptyHint: string }>} specs
 */
export function createPlannerExtraUsersController(specs) {
    /** @type {Record<string, Array<{ id: string, displayName: string, mail: string }>>} */
    const draft = {};
    let removeChangeListener = false;
    specs.forEach((s) => {
        draft[s.configKey] = [];
    });

    function renderList(spec) {
        const ul = document.getElementById(spec.listId);
        if (!ul) return;
        const users = normalizePlannerUsers(draft[spec.configKey]);
        draft[spec.configKey] = users;
        if (!users.length) {
            ul.innerHTML =
                '<li class="fr-setup-user-row fr-setup-user-row--empty muted">' + escapeHtml(spec.emptyHint) + '</li>';
            return;
        }
        ul.innerHTML = users
            .map(
                (u, idx) =>
                    `<li class="fr-setup-user-row">` +
                    `<span class="fr-setup-user-row__label">${formatUserLine(u)}</span>` +
                    `<button type="button" class="fr-setup-user-row__rm btn btn-sm alt" data-pep-rm="${escapeHtml(spec.configKey)}" data-pep-idx="${idx}" title="Entfernen" aria-label="Entfernen"><i class="bi bi-x"></i></button>` +
                    `</li>`
            )
            .join('');
        ul.querySelectorAll('[data-pep-rm]').forEach((btn) => {
            btn.addEventListener('click', () => {
                const key = btn.getAttribute('data-pep-rm');
                const i = Number(btn.getAttribute('data-pep-idx'));
                const list = normalizePlannerUsers(draft[key]);
                list.splice(i, 1);
                draft[key] = list;
                const sp = specs.find((x) => x.configKey === key);
                if (sp) renderList(sp);
            });
        });
    }

    return {
        /**
         * @param {Record<string, unknown>} cfg normalized permissions config
         * @param {{ adminUsers?: boolean, kvUsers?: boolean, schuelerUsers?: boolean, lehrerUsers?: boolean, direktionUsers?: boolean }} [explicitSaved]
         */
        loadFromConfig(cfg, explicitSaved) {
            specs.forEach((spec) => {
                let users = normalizePlannerUsers(cfg[spec.configKey]);
                const wasSaved = explicitSaved && explicitSaved[spec.configKey];
                if (
                    !users.length &&
                    !wasSaved &&
                    spec.configKey === 'adminUsers' &&
                    spec.stammdatenBtnId
                ) {
                    users = direktionUsersFromTenantStammdaten();
                }
                draft[spec.configKey] = users;
                renderList(spec);
            });
        },
        readPatch() {
            const out = {};
            specs.forEach((spec) => {
                out[spec.configKey] = normalizePlannerUsers(draft[spec.configKey]);
            });
            return out;
        },
        /**
         * @param {() => void} onChange
         */
        wire(onChange) {
            specs.forEach((spec) => {
                const addBtn = document.getElementById(spec.addBtnId);
                if (addBtn && !addBtn.dataset.pepWired) {
                    addBtn.dataset.pepWired = '1';
                    addBtn.addEventListener('click', () => {
                        pickEntraUser({ title: spec.pickTitle, hint: spec.pickHint })
                            .then((sel) => {
                                if (!sel || !sel.mail) return;
                                draft[spec.configKey] = mergePlannerUser(draft[spec.configKey], {
                                    id: sel.id,
                                    displayName: sel.displayName,
                                    mail: sel.mail
                                });
                                renderList(spec);
                                if (typeof onChange === 'function') onChange();
                            })
                            .catch((e) => {
                                const msg = e && e.message ? e.message : String(e);
                                if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
                            });
                    });
                }
                if (spec.stammdatenBtnId) {
                    const stBtn = document.getElementById(spec.stammdatenBtnId);
                    if (stBtn && !stBtn.dataset.pepWired) {
                        stBtn.dataset.pepWired = '1';
                        stBtn.addEventListener('click', () => {
                            const fromStamm = direktionUsersFromTenantStammdaten();
                            if (!fromStamm.length) {
                                if (typeof window.ms365ToastOrAlert === 'function') {
                                    window.ms365ToastOrAlert(
                                        'Keine Verwaltungs-Personen in den Stammdaten – bitte unter Stammdaten → Verwaltung pflegen.'
                                    );
                                }
                                return;
                            }
                            fromStamm.forEach((u) => {
                                draft[spec.configKey] = mergePlannerUser(draft[spec.configKey], u);
                            });
                            renderList(spec);
                            if (typeof onChange === 'function') onChange();
                            if (typeof window.ms365ToastOrAlert === 'function') {
                                window.ms365ToastOrAlert(fromStamm.length + ' Person(en) aus Stammdaten ergänzt.');
                            }
                        });
                    }
                }
            });
            if (!removeChangeListener) {
                removeChangeListener = true;
                document.addEventListener('click', (ev) => {
                    const rm = ev.target.closest('[data-pep-rm]');
                    if (!rm) return;
                    const key = rm.getAttribute('data-pep-rm');
                    if (!specs.some((s) => s.configKey === key)) return;
                    if (typeof onChange === 'function') onChange();
                });
            }
        },
        renderAll() {
            specs.forEach((spec) => renderList(spec));
        }
    };
}
