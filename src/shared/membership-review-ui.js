/**
 * Wiederverwendbare Mitglieder-Abgleich-UI (Review-Panel + Mismatch-Leiste).
 * @file
 */
import { diffClassMemberships, diffMemberships, indexGraphMembersByEmail, memberEmailFromGraph } from './membership-reconcile.js';
import {
    attachMembershipReviewSectionActions,
    buildMembershipReviewBody,
    collectMembershipReviewActionButtons
} from './membership-review-render.js';

function normEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

/**
 * @param {object} cfg
 * @returns {object}
 */
export function createMembershipReview(cfg) {
    const c = cfg && typeof cfg === 'object' ? cfg : {};
    const ids = c.ids || {};
    const id = function (key, fallback) {
        return ids[key] || fallback || '';
    };

    /** @type {object|null} */
    let state = null;

    function gug() {
        const G = window.ms365GraphUnifiedGroups;
        if (!G) throw new Error('graph-unified-groups.js fehlt.');
        return G;
    }

    function el(key, fallback) {
        const eid = id(key, fallback);
        return eid ? document.getElementById(eid) : null;
    }

    function getSelectedEmails(sectionKey) {
        const out = [];
        document.querySelectorAll('input[data-mr-section="' + sectionKey + '"]:checked').forEach(function (cb) {
            const em = normEmail(cb.getAttribute('data-mr-email'));
            if (em.indexOf('@') !== -1) out.push(em);
        });
        return out;
    }

    function renderPanel() {
        const body = el('body', 'slgDeviationBody');
        const summaryEl = el('summary', 'slgDeviationSummary');
        const actions = el('actions', 'slgDeviationActions');
        if (!body || !summaryEl || !state) return;

        const diff = state.diff;
        const labels = c.labels || {};
        summaryEl.textContent =
            'Markieren Sie Personen und wählen Sie unten in der passenden Spalte die Aktion.';
        if (actions) collectMembershipReviewActionButtons(actions);

        body.replaceChildren();
        body.appendChild(
            buildMembershipReviewBody({
                diff: diff,
                graphByEmail: state.graphByEmail,
                listCount: state.listCount,
                labels: labels
            })
        );
        attachMembershipReviewSectionActions(body, {
            onlyLocal: diff.onlyLocal.length,
            onlyGraph: diff.onlyGraph.length
        });
    }

    function hidePanel() {
        const panel = el('panel', 'slgDeviationPanel');
        if (panel) panel.hidden = true;
        state = null;
    }

    async function loadReview(reviewKey) {
        const gid = typeof c.getGroupId === 'function' ? c.getGroupId(reviewKey) : '';
        const localEmails =
            typeof c.getLocalEmails === 'function' ? c.getLocalEmails(reviewKey) : [];
        const panel = el('panel', 'slgDeviationPanel');
        const body = el('body', 'slgDeviationBody');
        const titleEl = el('title', 'slgDeviationTitle');
        const summaryEl = el('summary', 'slgDeviationSummary');
        const actions = el('actions', 'slgDeviationActions');

        if (!gid) {
            if (typeof c.toast === 'function') c.toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        if (!panel || !body || !titleEl || !summaryEl) return;

        const reviewTitle =
            typeof c.getReviewTitle === 'function' ? c.getReviewTitle(reviewKey) : 'Mitglieder-Abgleich';
        titleEl.textContent = reviewTitle;
        summaryEl.textContent = 'Vergleich wird geladen …';
        if (actions) collectMembershipReviewActionButtons(actions);
        body.replaceChildren();
        panel.hidden = false;
        panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

        try {
            const token = await c.getGraphToken();
            const mem = await gug().fetchGroupMembers(token, gid);
            const graphEmails = (mem.items || [])
                .map(function (m) {
                    return memberEmailFromGraph(m);
                })
                .filter(function (em) {
                    return em.indexOf('@') !== -1;
                });
            let diff;
            if (c.mode === 'class') {
                const allStudents =
                    typeof c.getAllStudentEmails === 'function' ? c.getAllStudentEmails() : [];
                diff = diffClassMemberships(localEmails, allStudents, graphEmails);
            } else {
                diff = diffMemberships(localEmails, graphEmails);
            }
            state = {
                reviewKey: reviewKey,
                gid: gid,
                listCount: localEmails.length,
                diff: diff,
                graphByEmail: indexGraphMembersByEmail(mem.items || [])
            };
            renderPanel();
        } catch (e) {
            state = null;
            summaryEl.textContent = '';
            const p = document.createElement('p');
            p.className = 'slg-deviation-panel__error';
            p.textContent = 'Abgleich konnte nicht geladen werden: ' + (e.message || e);
            body.appendChild(p);
        }
    }

    async function applyAddToGroup() {
        if (!state) return;
        const emails = getSelectedEmails('onlyLocal');
        if (!emails.length) {
            if (c.toast) c.toast('Keine Einträge aus „Nur in der Schul-Liste“ ausgewählt.');
            return;
        }
        const ok = await c.dlgConfirm(
            emails.length +
                (emails.length === 1 ? ' Person in' : ' Personen in') +
                ' die Microsoft-365-Gruppe aufnehmen?',
            { title: 'In Gruppe aufnehmen' }
        );
        if (!ok) return;
        try {
            const token = await c.getGraphToken();
            const label = c.syncLabel || 'Mitglied';
            const r = await gug().syncEmailsToGroup(token, state.gid, emails, label, c.appendSyncLog || null);
            if (typeof c.logAction === 'function') {
                c.logAction('membership-add', state.gid, label + ': +' + r.ok + ' in Gruppe');
            }
            if (c.live && c.live.invalidateMembership) c.live.invalidateMembership();
            if (c.live && c.live.loadMembers) await c.live.loadMembers();
            if (typeof c.refreshCounts === 'function') await c.refreshCounts();
            if (c.toast) c.toast(r.ok + ' in Gruppe aufgenommen.');
            await loadReview(state.reviewKey);
            if (typeof c.onAfterChange === 'function') await c.onAfterChange();
        } catch (e) {
            if (c.toast) c.toast('Fehler: ' + (e.message || e));
        }
    }

    async function applyRemoveFromGroup() {
        if (!state) return;
        const emails = getSelectedEmails('onlyGraph');
        if (!emails.length) {
            if (c.toast) c.toast('Keine Einträge aus „Nur in der Microsoft-365-Gruppe“ ausgewählt.');
            return;
        }
        const ok = await c.dlgConfirm(
            emails.length +
                (emails.length === 1 ? ' Person aus' : ' Personen aus') +
                ' der Microsoft-365-Gruppe entfernen?',
            { title: 'Aus Gruppe entfernen', danger: true }
        );
        if (!ok) return;
        try {
            const token = await c.getGraphToken();
            const label = c.syncLabel || 'Mitglied';
            const r = await gug().removeEmailsFromGroup(token, state.gid, emails, label, c.appendSyncLog || null);
            if (typeof c.logAction === 'function') {
                c.logAction('membership-remove', state.gid, label + ': −' + r.ok + ' aus Gruppe');
            }
            if (c.live && c.live.invalidateMembership) c.live.invalidateMembership();
            if (c.live && c.live.loadMembers) await c.live.loadMembers();
            if (typeof c.refreshCounts === 'function') await c.refreshCounts();
            if (c.toast) c.toast(r.ok + ' aus Gruppe entfernt.');
            await loadReview(state.reviewKey);
            if (typeof c.onAfterChange === 'function') await c.onAfterChange();
        } catch (e) {
            if (c.toast) c.toast('Fehler: ' + (e.message || e));
        }
    }

    async function applyImportLocal() {
        if (!state) {
            if (c.toast) c.toast('Abgleich ist nicht aktiv – bitte erneut öffnen.');
            return;
        }
        if (typeof c.openImport !== 'function') {
            if (c.toast) c.toast('Import ist für dieses Werkzeug nicht verfügbar.');
            return;
        }
        const emails = getSelectedEmails('onlyGraph');
        if (!emails.length) {
            if (c.toast) c.toast('Keine Einträge aus „Nur in der Microsoft-365-Gruppe“ ausgewählt.');
            return;
        }
        try {
            const result = await c.openImport(emails, state);
            if (result && !result.cancelled && typeof c.refreshCounts === 'function') {
                await c.refreshCounts();
            }
        } catch (e) {
            if (c.toast) c.toast('Fehler: ' + (e.message || e));
        }
    }

    function updateMismatchBar(items) {
        const bar = el('mismatchBar', 'slgMismatchBar');
        const actions = el('mismatchActions', 'slgMismatchActions');
        if (!bar || !actions) return;
        const list = Array.isArray(items) ? items.filter(function (it) {
            return it && it.gid && it.listN !== null && it.groupN !== null && it.listN !== it.groupN;
        }) : [];
        if (!list.length) {
            bar.hidden = true;
            actions.replaceChildren();
            return;
        }
        bar.hidden = false;
        actions.replaceChildren();
        list.forEach(function (it) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-sm slg-mismatch-bar__btn';
            btn.innerHTML =
                '<i class="bi bi-intersect"></i>' +
                it.label +
                ': Liste ' +
                it.listN +
                ' / Gruppe ' +
                it.groupN +
                ' – Abgleich öffnen';
            btn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                void loadReview(it.key);
            });
            actions.appendChild(btn);
        });
    }

    function wire() {
        function on(idKey, fallback, fn) {
            const node = el(idKey, fallback);
            if (node && !node.dataset.mrBound) {
                node.dataset.mrBound = '1';
                node.addEventListener('click', fn);
            }
        }
        on('addBtn', 'slgMrAddToGroup', function () {
            void applyAddToGroup();
        });
        on('importBtn', 'slgMrImportLocal', function () {
            void applyImportLocal();
        });
        on('removeBtn', 'slgMrRemoveFromGroup', function () {
            void applyRemoveFromGroup();
        });
        on('reloadBtn', 'slgMrReload', function () {
            if (state) void loadReview(state.reviewKey);
        });
        on('closeBtn', 'slgDeviationClose', hidePanel);
        on('reviewBtn', 'slgBtnMembershipReview', function () {
            if (typeof c.getActiveReviewKey === 'function') {
                void loadReview(c.getActiveReviewKey());
            }
        });
    }

    return {
        loadReview: loadReview,
        hidePanel: hidePanel,
        updateMismatchBar: updateMismatchBar,
        wire: wire,
        getState: function () {
            return state;
        }
    };
}

const api = { createMembershipReview: createMembershipReview };

if (typeof window !== 'undefined') {
    window.ms365MembershipReviewUi = api;
}

export default api;
