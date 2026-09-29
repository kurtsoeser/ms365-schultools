/**
 * Wizard-Bindings OneNote-Verteilung (Analyse 02 Phase B).
 */
import { escapeHtml } from '../../shared/utils/strings.js';
import { ui } from './onenote-verteilung-state.js';
import { searchTeams, searchSites, listOnenoteNotebooks, listCentralTemplateNotebooks, loadOnenoteNotebookTree, copySectionToGroupSectionGroup, pickClassNotebook, publishCentralNotebookSnapshot, CENTRAL_TEMPLATE_NOTEBOOK_NAME } from './onenote-verteilung-graph.js';

function $(id) {
    return document.getElementById(id);
}

function formatNotebookWhen(iso) {
    const raw = String(iso || '').trim();
    if (!raw) return '';
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return '';
    try {
        return new Intl.DateTimeFormat('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        }).format(d);
    } catch {
        return d.toLocaleString('de-AT');
    }
}

/** @type {any} */
let api = null;
export function initOnenoteSteps(a) {
    api = a;
}

export function bindWizardNav() {
    document.querySelectorAll('[data-onv-goto]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const n = Number(btn.getAttribute('data-onv-goto'));
            if (!Number.isFinite(n)) return;
            if (n < ui.step || api.canEnterStep(n)) api.showStep(n);
            else api.toast('Bitte zuerst die vorherigen Schritte ausfüllen.');
        });
    });
    $('onvBtnNext1')?.addEventListener('click', () => api.showStep(2));
    $('onvBtnNext2')?.addEventListener('click', () => api.showStep(3));
    $('onvBtnNext3')?.addEventListener('click', () => api.showStep(4));
    $('onvBtnBack2')?.addEventListener('click', () => api.showStep(1));
    $('onvBtnBack3')?.addEventListener('click', () => api.showStep(2));
    $('onvBtnBack4')?.addEventListener('click', () => api.showStep(3));
}

export function bindStep1() {
    document.querySelectorAll('[data-onv-src-mode]').forEach((btn) => {
        btn.addEventListener('click', () => {
            api.setSrcMode(btn.getAttribute('data-onv-src-mode') || 'central');
        });
    });

    $('onvNotebookShelf')?.addEventListener('click', (e) => {
        if (e.target.closest('[data-nb-publish-pick], .onv-nb-pick')) return;
        const cover = e.target.closest('[data-nb-id]');
        if (!cover) return;
        api.selectNotebook(cover.getAttribute('data-nb-id') || '');
    });

    $('onvNotebookShelf')?.addEventListener('change', (e) => {
        const input = e.target.closest('[data-nb-publish-pick]');
        if (!input) return;
        const id = input.getAttribute('data-nb-publish-pick') || '';
        if (!id) return;
        if (input.checked) ui.publishPickIds.add(id);
        else ui.publishPickIds.delete(id);
        api.updatePublishBtn();
    });

    async function runSrcTeamSearch() {
        const q = ($('onvSrcTeamQuery') && $('onvSrcTeamQuery').value) || '';
        const ul = $('onvSrcTeamHits');
        if (!ul) return;
        try {
            const hits = await searchTeams(q);
            if (!hits.length) {
                ul.innerHTML = '<li class="muted" style="padding:10px;">Keine Treffer.</li>';
                ul.hidden = false;
                return;
            }
            ul.innerHTML = hits
                .map(
                    (t) =>
                        '<li><button type="button" data-src-team-id="' +
                        escapeHtml(t.id) +
                        '" data-src-team-name="' +
                        escapeHtml(t.displayName || '') +
                        '" data-src-team-nick="' +
                        escapeHtml(t.mailNickname || '') +
                        '"><strong>' +
                        escapeHtml(t.displayName || t.id) +
                        '</strong><span class="muted">' +
                        escapeHtml(t.mailNickname || t.mail || '') +
                        '</span></button></li>'
                )
                .join('');
            ul.hidden = false;
        } catch (err) {
            api.toast((err && err.message) || String(err));
            ul.hidden = true;
        }
    }

    async function runSrcSiteSearch() {
        const q = ($('onvSrcSiteQuery') && $('onvSrcSiteQuery').value) || '';
        const ul = $('onvSrcSiteHits');
        if (!ul) return;
        try {
            const hits = await searchSites(q);
            if (!hits.length) {
                ul.innerHTML = '<li class="muted" style="padding:10px;">Keine Sites gefunden.</li>';
                ul.hidden = false;
                return;
            }
            ul.innerHTML = hits
                .map(
                    (s) =>
                        '<li><button type="button" data-src-site-id="' +
                        escapeHtml(s.id) +
                        '" data-src-site-name="' +
                        escapeHtml(s.displayName || '') +
                        '" data-src-site-url="' +
                        escapeHtml(s.webUrl || '') +
                        '"><strong>' +
                        escapeHtml(s.displayName || s.id) +
                        '</strong><span class="muted" style="font-size:0.75rem;word-break:break-all;">' +
                        escapeHtml(s.webUrl || '') +
                        '</span></button></li>'
                )
                .join('');
            ul.hidden = false;
        } catch (err) {
            api.toast((err && err.message) || String(err));
            ul.hidden = true;
        }
    }

    $('onvBtnSrcSearchTeam')?.addEventListener('click', () => {
        runSrcTeamSearch().catch(() => {});
    });
    $('onvSrcTeamQuery')?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            runSrcTeamSearch().catch(() => {});
        }
    });
    $('onvSrcTeamHits')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-src-team-id]');
        if (!btn) return;
        ui.srcGroup = {
            id: btn.getAttribute('data-src-team-id') || '',
            displayName: btn.getAttribute('data-src-team-name') || '',
            mailNickname: btn.getAttribute('data-src-team-nick') || ''
        };
        const ul = $('onvSrcTeamHits');
        if (ul) ul.hidden = true;
        api.renderSrcTeamSelected();
        api.toast('Team gewählt: ' + (ui.srcGroup.displayName || ui.srcGroup.id));
    });

    $('onvBtnSrcSearchSite')?.addEventListener('click', () => {
        runSrcSiteSearch().catch(() => {});
    });
    $('onvSrcSiteQuery')?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            runSrcSiteSearch().catch(() => {});
        }
    });
    $('onvSrcSiteHits')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-src-site-id]');
        if (!btn) return;
        ui.srcSite = {
            id: btn.getAttribute('data-src-site-id') || '',
            displayName: btn.getAttribute('data-src-site-name') || '',
            webUrl: btn.getAttribute('data-src-site-url') || ''
        };
        const ul = $('onvSrcSiteHits');
        if (ul) ul.hidden = true;
        api.renderSrcSiteSelected();
        api.toast('Site gewählt: ' + (ui.srcSite.displayName || ui.srcSite.id));
    });

    $('onvBtnLoadSrc')?.addEventListener('click', async () => {
        try {
            ui.onSrcTree = null;
            ui.onSrcChecked = new Set();
            ui.onSrcExpanded = new Set();
            ui.previewSectionId = '';
            api.clearPreview();

            if (ui.srcMode === 'central') {
                api.log('Lade zentrale Vorlagen (Katalog-API bevorzugt) …');
                const { site, notebooks, via } = await listCentralTemplateNotebooks();
                ui.srcSite = site;
                ui.srcVia =
                    via === 'site-graph'
                        ? 'site-graph'
                        : via === 'catalog-snapshot'
                          ? 'catalog-snapshot'
                          : 'catalog-api';
                ui.onSrcNotebooks = notebooks;
                api.fillOnSelect($('onvSrcNotebook'), ui.onSrcNotebooks, '— Notizbuch wählen —');
                api.renderSrcTree();
                const prefer = ui.onSrcNotebooks.find(
                    (n) =>
                        n.displayName.toLowerCase() ===
                        CENTRAL_TEMPLATE_NOTEBOOK_NAME.toLowerCase()
                );
                if (prefer) api.selectNotebook(prefer.id);
                const viaLabel =
                    ui.srcVia === 'catalog-api'
                        ? 'via Katalog-API'
                        : ui.srcVia === 'catalog-snapshot'
                          ? 'via Snapshot (Schul-Tenants)'
                          : 'via Site-Graph';
                api.toast(
                    notebooks.length
                        ? notebooks.length + ' Notizbuch/Notizbücher geladen (' + viaLabel + ').'
                        : 'Keine Notizbücher im Katalog.'
                );
                api.log(
                    'Zentrale Vorlagen (' +
                        viaLabel +
                        ') „' +
                        (site && site.displayName ? site.displayName : 'Katalog') +
                        '“: ' +
                        notebooks.map((n) => n.displayName).join(', ')
                );
                const modeLabel = $('onvSrcModeLabel');
                if (modeLabel) {
                    if (ui.srcVia === 'catalog-api' || ui.srcVia === 'catalog-snapshot') {
                        modeLabel.textContent =
                            'Zentrale Vorlagen · Katalog (' + viaLabel + ', schulübergreifend)';
                    } else {
                        modeLabel.textContent =
                            'Zentrale Vorlagen · Site-Graph (kurtrocks) – für Schulen bitte Snapshot veröffentlichen';
                    }
                }
                api.updatePublishBtn();
                api.refreshPublishedMarks().catch(() => {});
                return;
            }

            if (ui.srcMode === 'team') {
                if (!ui.srcGroup || !ui.srcGroup.id) {
                    api.toast('Bitte zuerst ein Team suchen und wählen.');
                    return;
                }
                api.log('Lade Notizbücher aus Team „' + ui.srcGroup.displayName + '“ …');
                ui.srcVia = '';
                ui.onSrcNotebooks = await listOnenoteNotebooks({
                    kind: 'group',
                    id: ui.srcGroup.id
                });
                api.fillOnSelect($('onvSrcNotebook'), ui.onSrcNotebooks, '— Notizbuch wählen —');
                api.renderSrcTree();
                const auto = pickClassNotebook(ui.onSrcNotebooks, ui.srcGroup.displayName);
                if (auto) api.selectNotebook(auto.id);
                api.toast(
                    ui.onSrcNotebooks.length
                        ? ui.onSrcNotebooks.length + ' Team-Notizbuch/Notizbücher.'
                        : 'Keine Notizbücher in diesem Team.'
                );
                api.log(
                    'Team-Quelle „' +
                        ui.srcGroup.displayName +
                        '“: ' +
                        ui.onSrcNotebooks.map((n) => n.displayName).join(', ')
                );
                return;
            }

            if (ui.srcMode === 'site') {
                if (!ui.srcSite || !ui.srcSite.id) {
                    api.toast('Bitte zuerst eine SharePoint-Site suchen/wählen oder URL auflösen.');
                    return;
                }
                api.log('Lade Notizbücher von Site „' + ui.srcSite.displayName + '“ …');
                ui.srcVia = '';
                ui.onSrcNotebooks = await listOnenoteNotebooks({
                    kind: 'site',
                    id: ui.srcSite.id
                });
                api.fillOnSelect($('onvSrcNotebook'), ui.onSrcNotebooks, '— Notizbuch wählen —');
                api.renderSrcTree();
                api.toast(
                    ui.onSrcNotebooks.length
                        ? ui.onSrcNotebooks.length + ' Site-Notizbuch/Notizbücher.'
                        : 'Keine Notizbücher auf dieser Site.'
                );
                api.log(
                    'Site-Quelle „' +
                        ui.srcSite.displayName +
                        '“: ' +
                        ui.onSrcNotebooks.map((n) => n.displayName).join(', ')
                );
                return;
            }

            api.log('Lade eigene OneNote-Notizbücher (OneDrive) …');
            ui.srcSite = null;
            ui.srcGroup = null;
            ui.srcVia = '';
            ui.onSrcNotebooks = await listOnenoteNotebooks({ kind: 'me' });
            api.fillOnSelect($('onvSrcNotebook'), ui.onSrcNotebooks, '— Notizbuch wählen —');
            api.renderSrcTree();
            api.toast(ui.onSrcNotebooks.length + ' Notizbuch/Notizbücher gefunden.');
            api.log('Quelle (OneDrive): ' + ui.onSrcNotebooks.length + ' Notizbuch/Notizbücher.');
        } catch (err) {
            api.toast((err && err.message) || String(err));
            api.log('OneNote Quelle: ' + ((err && err.message) || err));
        }
    });

    $('onvBtnPublishSnapshot')?.addEventListener('click', async () => {
        if (!api.canPublishSnapshots()) {
            api.toast('Veröffentlichen nur im Betreiber-Tenant (kurtrocks) mit Site-Zugriff.');
            return;
        }
        let ids = Array.from(ui.publishPickIds);
        if (!ids.length) {
            const cur = ($('onvSrcNotebook') && $('onvSrcNotebook').value) || '';
            if (cur) ids = [cur];
        }
        const notebooks = ids
            .map((id) => (ui.onSrcNotebooks || []).find((n) => n.id === id))
            .filter(Boolean);
        if (!notebooks.length) {
            api.toast('Bitte Notizbücher anhaken oder eines auswählen.');
            return;
        }
        const btn = $('onvBtnPublishSnapshot');
        if (btn) btn.disabled = true;
        try {
            api.log(
                'Veröffentliche ' +
                    notebooks.length +
                    ' Notizbuch/Notizbücher für Schulen: ' +
                    notebooks.map((n) => n.displayName).join(', ')
            );
            api.toast(
                notebooks.length === 1
                    ? 'Snapshot wird erstellt …'
                    : notebooks.length + ' Bücher werden veröffentlicht …'
            );
            const result = await publishCentralNotebookSnapshot(
                notebooks,
                { kind: 'site', id: ui.srcSite.id },
                (info) => {
                    if (info && (info.phase === 'section' || info.phase === 'upload') && info.detail) {
                        api.log('Snapshot: ' + info.detail);
                    }
                }
            );
            notebooks.forEach((n) => {
                ui.publishedNotebookIds.add(n.id);
                if (result && result.updatedAt) {
                    ui.publishedAtById.set(n.id, result.updatedAt);
                }
            });
            ui.publishPickIds = new Set();
            const media = result && result.media;
            const mediaHint = media
                ? ' · Bilder ' +
                  (media.imagesInlined || 0) +
                  ' eingebettet' +
                  (media.imagesSkipped ? ', ' + media.imagesSkipped + ' übersprungen' : '') +
                  (media.filesInlined ? ', Dateien ' + media.filesInlined : '') +
                  (media.embedsReplaced
                      ? ', ' + media.embedsReplaced + ' Embeds → Platzhalter'
                      : '')
                : '';
            const emptyHint =
                result && result.pagesWithoutHtml
                    ? ' · ' + result.pagesWithoutHtml + ' Seiten ohne HTML'
                    : '';
            const pubLabel = formatNotebookWhen(result && result.updatedAt);
            api.toast(
                notebooks.length +
                    ' Buch/Bücher veröffentlicht' +
                    (pubLabel ? ' · ' + pubLabel : '') +
                    ' · Snapshot insgesamt ' +
                    (result.notebookCount || notebooks.length) +
                    ' Bücher.' +
                    mediaHint +
                    emptyHint
            );
            api.log(
                'Snapshot OK · updatedAt=' +
                    (result.updatedAt || '') +
                    ' · notebookCount=' +
                    (result.notebookCount || '') +
                    ' · pagesWithHtml=' +
                    (result.pagesWithHtml != null ? result.pagesWithHtml : '?') +
                    ' · pagesWithoutHtml=' +
                    (result.pagesWithoutHtml != null ? result.pagesWithoutHtml : '?') +
                    mediaHint
            );
            api.renderNotebookShelf(
                ui.onSrcNotebooks,
                ($('onvSrcNotebook') && $('onvSrcNotebook').value) || ''
            );
            api.updatePublishBtn();
            api.refreshPublishedMarks().catch(() => {});
        } catch (err) {
            const msg = (err && err.message) || String(err);
            api.toast(msg);
            api.log('Snapshot-Publish: ' + msg);
        } finally {
            if (btn) btn.disabled = false;
        }
    });

    $('onvSrcNotebook')?.addEventListener('change', async (e) => {
        const id = e.target.value || '';
        const name =
            (e.target.selectedOptions &&
                e.target.selectedOptions[0] &&
                e.target.selectedOptions[0].textContent) ||
            'Vorlagen-Notizbuch';
        const title = $('onvSrcTitle');
        if (title) title.textContent = id ? name : 'Vorlagen-Notizbuch';
        if (
            id &&
            ui.loadedSrcNotebookId === id &&
            ui.onSrcTree &&
            (ui.onSrcTree.sections || ui.onSrcTree.groups)
        ) {
            api.renderSrcTree();
            api.updatePublishBtn();
            return;
        }
        ui.onSrcTree = null;
        ui.onSrcChecked = new Set();
        ui.onSrcExpanded = new Set();
        ui.previewSectionId = '';
        ui.loadedSrcNotebookId = '';
        api.clearPreview();
        api.renderSrcTree();
        api.updatePublishBtn();
        if (!id) return;
        const gen = ++ui.srcTreeLoadGen;
        try {
            api.log('Lade Struktur „' + name + '“ …');
            const tree = await loadOnenoteNotebookTree(id, api.sourceScope());
            if (gen !== ui.srcTreeLoadGen) return;
            ui.onSrcTree = tree;
            ui.loadedSrcNotebookId = id;
            (tree.groups || []).forEach((g) => ui.onSrcExpanded.add(g.id));
            api.renderSrcTree();
            const c = api.countTreeSections(tree);
            api.log('Quelle geladen: ' + c.sections + ' Abschnitte, ' + c.groups + ' Gruppen.');
            api.toast('Vorlagen-Struktur geladen.');
        } catch (err) {
            if (gen !== ui.srcTreeLoadGen) return;
            api.toast((err && err.message) || String(err));
            api.log('OneNote Abschnitte: ' + ((err && err.message) || err));
        }
    });

    $('onvSrcTree')?.addEventListener('click', (e) => {
        const toggle = e.target.closest('[data-on-toggle]');
        if (toggle) {
            e.preventDefault();
            e.stopPropagation();
            const gid = toggle.getAttribute('data-on-toggle') || '';
            if (ui.onSrcExpanded.has(gid)) ui.onSrcExpanded.delete(gid);
            else ui.onSrcExpanded.add(gid);
            api.renderSrcTree();
            return;
        }
        const previewBtn = e.target.closest('[data-on-preview]');
        if (previewBtn) {
            e.preventDefault();
            e.stopPropagation();
            api.loadPreview(previewBtn.getAttribute('data-on-preview') || '');
            return;
        }
        if (e.target.closest('[data-on-check]')) return;
        const row = e.target.closest('[data-on-section]');
        if (row) {
            const sid = row.getAttribute('data-on-section') || '';
            if (sid) api.loadPreview(sid);
        }
    });

    $('onvSrcTree')?.addEventListener('change', (e) => {
        const input = e.target.closest('[data-on-check]');
        if (!input) return;
        const id = input.getAttribute('data-on-check') || '';
        if (input.checked) ui.onSrcChecked.add(id);
        else ui.onSrcChecked.delete(id);
        api.renderSrcTree();
    });

    $('onvPreviewPages')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-on-page]');
        if (!btn) return;
        const pid = btn.getAttribute('data-on-page') || '';
        api.showPagePreview(pid, btn.textContent || '');
    });
}

export function bindStep2() {
    $('onvBtnSearchTeam')?.addEventListener('click', async () => {
        const q = $('onvTeamQuery')?.value || '';
        const ul = $('onvTeamHits');
        try {
            const hits = await searchTeams(q);
            if (!ul) return;
            if (!hits.length) {
                ul.hidden = false;
                ul.innerHTML = '<li class="muted" style="padding:8px;">Keine Treffer.</li>';
                return;
            }
            ul.hidden = false;
            ul.innerHTML = hits
                .map(
                    (h) =>
                        `<li><button type="button" data-team-id="${escapeHtml(h.id)}" data-team-name="${escapeHtml(h.displayName)}" data-team-nick="${escapeHtml(h.mailNickname || '')}">` +
                        `<strong>${escapeHtml(h.displayName)}</strong>` +
                        `<span>${escapeHtml(h.mailNickname || h.mail || h.id)}</span></button></li>`
                )
                .join('');
        } catch (err) {
            api.toast((err && err.message) || String(err));
        }
    });

    $('onvTeamQuery')?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            $('onvBtnSearchTeam')?.click();
        }
    });

    $('onvTeamHits')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-team-id]');
        if (!btn) return;
        api.addTeam({
            id: btn.getAttribute('data-team-id') || '',
            displayName: btn.getAttribute('data-team-name') || '',
            mailNickname: btn.getAttribute('data-team-nick') || ''
        });
        const ul = $('onvTeamHits');
        if (ul) ul.hidden = true;
    });

    $('onvTeamList')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-onv-remove-team]');
        if (!btn) return;
        api.removeTeam(Number(btn.getAttribute('data-onv-remove-team')));
    });
}

export function bindStep3() {
    document.querySelectorAll('input[name="onvDestKind"]').forEach((inp) => {
        inp.addEventListener('change', () => {
            ui.destKind = api.getDestKind();
            if (api.teamsWithNotebook().length) {
                api.buildTargets().catch(() => {});
            }
        });
    });

    $('onvBtnBuildTargets')?.addEventListener('click', () => {
        api.buildTargets().catch(() => {});
    });

    $('onvTargetList')?.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-onv-remove-target]');
        if (!btn) return;
        const idx = Number(btn.getAttribute('data-onv-remove-target'));
        if (!Number.isFinite(idx) || idx < 0) return;
        ui.onTargets.splice(idx, 1);
        api.renderTargets();
    });
}

export function bindStep4() {
    $('onvBtnCopy')?.addEventListener('click', async () => {
        if (!ui.onSrcChecked.size) return api.toast('Mindestens einen Quell-Abschnitt anhaken.');
        if (!ui.onTargets.length) return api.toast('Keine Ziele in der Verteilerliste.');

        const sections = [...ui.onSrcChecked].map((id) => ({
            id,
            name: api.findSectionName(ui.onSrcTree, id) || 'Abschnitt'
        }));
        const destinations = ui.onTargets.slice();
        const total = destinations.length * sections.length;
        let done = 0;
        let ok = 0;
        let fail = 0;

        api.setOnProgress({
            pct: 2,
            message:
                'Verteile ' +
                sections.length +
                ' Abschnitt(e) an ' +
                destinations.length +
                ' Ziel(e) …'
        });
        const btn = $('onvBtnCopy');
        if (btn) btn.disabled = true;

        ui.onTargets.forEach((t) => {
            t.status = '';
            t.statusKind = '';
        });
        api.renderTargets();

        try {
            for (let d = 0; d < destinations.length; d++) {
                const dest = destinations[d];
                let destOk = 0;
                let destFail = 0;
                const row = ui.onTargets.find((t) => t.key === dest.key);
                if (row) {
                    row.status = 'Läuft …';
                    row.statusKind = 'run';
                    api.renderTargets();
                }
                for (let i = 0; i < sections.length; i++) {
                    const sec = sections[i];
                    done++;
                    const pct = Math.round((done / total) * 100);
                    api.setOnProgress({
                        pct: Math.min(99, pct),
                        message:
                            dest.teamName +
                            ': „' +
                            sec.name +
                            '“ (' +
                            done +
                            '/' +
                            total +
                            ') …'
                    });
                    api.log(
                        'OneNote ' +
                            (api.sourceScope().kind === 'catalog' ? 'Snapshot-Rebuild' : '1:1-Copy') +
                            ' → ' +
                            dest.teamName +
                            ' / ' +
                            dest.sectionGroupName +
                            ': ' +
                            sec.name
                    );
                    try {
                        await copySectionToGroupSectionGroup(
                            sec.id,
                            {
                                sectionGroupId: dest.sectionGroupId,
                                groupId: dest.teamId,
                                renameAs: sec.name
                            },
                            (op) => {
                                const st = (op && op.status) || '';
                                api.setOnProgress({
                                    pct: Math.min(99, pct),
                                    message: dest.teamName + ' · „' + sec.name + '“: ' + st
                                });
                            },
                            api.sourceScope()
                        );
                        ok++;
                        destOk++;
                        api.log('OK: ' + sec.name + ' → ' + dest.teamName);
                    } catch (err) {
                        fail++;
                        destFail++;
                        api.log(
                            'Fehler „' +
                                sec.name +
                                '“ → ' +
                                dest.teamName +
                                ': ' +
                                ((err && err.message) || err)
                        );
                    }
                }
                if (row) {
                    if (destFail && !destOk) {
                        row.status = 'Fehlgeschlagen (' + destFail + ')';
                        row.statusKind = 'err';
                    } else if (destFail) {
                        row.status = destOk + ' ok, ' + destFail + ' Fehler';
                        row.statusKind = 'err';
                    } else {
                        row.status = destOk + ' Abschnitt(e) kopiert';
                        row.statusKind = 'ok';
                    }
                    api.renderTargets();
                }
            }

            if (fail && !ok) {
                api.setOnProgress({
                    state: 'error',
                    pct: 100,
                    message: 'Alle Kopien fehlgeschlagen (' + fail + '). Details im Protokoll.'
                });
                api.toast('Verteilung fehlgeschlagen.');
            } else if (fail) {
                api.setOnProgress({
                    state: 'ok',
                    pct: 100,
                    message:
                        ok +
                        ' ok, ' +
                        fail +
                        ' Fehler · ' +
                        destinations.length +
                        ' Ziel(e). Protokoll prüfen.'
                });
                api.toast(ok + ' kopiert, ' + fail + ' Fehler.');
            } else {
                api.setOnProgress({
                    state: 'ok',
                    pct: 100,
                    message:
                        ok +
                        ' Kopie(n) auf ' +
                        destinations.length +
                        ' Ziel(e) verteilt.'
                });
                api.toast(
                    'Verteilt: ' +
                        sections.length +
                        ' Abschnitt(e) × ' +
                        destinations.length +
                        ' Ziel(e).'
                );
            }
        } finally {
            api.updateNavButtons();
        }
    });
}

