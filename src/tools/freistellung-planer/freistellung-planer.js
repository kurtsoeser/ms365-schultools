/**
 * Entry / Wiring: Freistellungs-Planer
 */
import {
    createInitialState,
    loadStammdaten,
    resolveRole,
    persistRole,
    persistSiteUrl,
    emptyForm,
    filterItems,
    scopeFromState,
    resolveKvForClass,
    loadSetupCfg
} from './freistellung-planer-state.js';
import {
    resolveFrContext,
    loadAllFreistellungen,
    createFreistellungItem,
    updateFreistellungStatus
} from './freistellung-planer-graph.js';
import { validateFreistellung } from './freistellung-planer-logic.js';
import {
    renderApp,
    readFormFromDom,
    readFiltersFromDom,
    applyKvFromClass
} from './freistellung-planer-ui.js';
import { downloadFreistellungCsv } from './freistellung-planer-export.js';
import {
    buildLocalDemoItems,
    getDemoSeedPackage,
    DEMO_SITE_DEFAULT,
    DEMO_SEED_TAG
} from './freistellung-planer-demo.js';
import {
    applyDemoStammdatenLocal,
    seedDemoFreistellungen,
    resetDemoFreistellungen
} from './freistellung-planer-demo-seed.js';

const state = createInitialState();
let root = null;

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function paint() {
    if (!root) return;
    renderApp(state, root);
    bindStatic();
}

function currentAccount() {
    try {
        if (typeof window.ms365GetAccountInfo === 'function') {
            const a = window.ms365GetAccountInfo();
            if (a) return a;
        }
    } catch {
        /* ignore */
    }
    try {
        const acc = window.msalAccount || null;
        if (acc) return { email: acc.username || acc.mail || '', name: acc.name || '' };
    } catch {
        /* ignore */
    }
    return { email: '', name: '' };
}

function syncAccount() {
    const a = currentAccount();
    state.accountEmail = String(a.email || a.username || '')
        .trim()
        .toLowerCase();
    state.accountName = String(a.name || '').trim();
    state.stammdaten = loadStammdaten();
    if (state.role === 'schueler' && state.accountName && !state.form.schuelerName) {
        state.form.schuelerName = state.accountName;
    }
}

async function refreshData() {
    state.loading = true;
    state.error = '';
    state.info = '';
    paint();
    try {
        syncAccount();
        const setup = loadSetupCfg();
        if (setup.listName) state.listName = setup.listName;
        if (setup.listId) state.listId = setup.listId;
        if (setup.emailDirektion) state.emailDirektion = setup.emailDirektion;

        const ctx = await resolveFrContext(state.siteUrl, {
            listName: state.listName,
            listId: state.listId
        });
        state.ctx = ctx;
        state.localDemoOnly = false;
        state.items = await loadAllFreistellungen(ctx);
        state.info =
            'Liste „' +
            (ctx.list.name || state.listName) +
            '“ · ' +
            state.items.length +
            ' Einträge';
    } catch (e) {
        state.error = String((e && e.message) || e || 'Laden fehlgeschlagen');
        state.ctx = null;
    } finally {
        state.loading = false;
        paint();
    }
}

function exportCsv() {
    const scope = scopeFromState(state, { scopeAll: state.role === 'direktion' });
    const items = filterItems(state.items, state.filters, scope);
    downloadFreistellungCsv(items);
    toast('CSV exportiert (' + items.length + ' Zeilen).');
}

function applyLocalDemo(pack) {
    syncAccount();
    const stamOk = applyDemoStammdatenLocal(pack.stammdaten);
    if (stamOk) state.stammdaten = loadStammdaten();
    state.localDemoOnly = true;
    state.ctx = null;
    state.error = '';
    state.items = buildLocalDemoItems({
        accountEmail: state.accountEmail,
        accountName: state.accountName,
        classes: state.stammdaten.classes
    });
    const c = pack.counts || {};
    state.info =
        'Demo SJ ' +
        (pack.schoolYear || '2026/27') +
        ' lokal (' +
        state.items.length +
        ' Anträge' +
        (c.ausstehend != null ? ', ' + c.ausstehend + ' offen' : '') +
        ').';
    return stamOk;
}

async function runDemoImport() {
    const pack = getDemoSeedPackage();
    const stamOk = applyLocalDemo(pack);
    paint();
    toast(
        'Demo geladen (' +
            state.items.length +
            ' Anträge)' +
            (stamOk ? ', Stammdaten übernommen' : '') +
            '.'
    );

    const siteInput = root && root.querySelector('#frSiteUrl');
    const siteUrl =
        String((siteInput && siteInput.value) || state.siteUrl || pack.siteUrl || DEMO_SITE_DEFAULT).trim();
    if (!siteUrl) return;

    const writeSp = window.confirm(
        'Demo auch auf SharePoint schreiben?\n\n' +
            'Site: ' +
            siteUrl +
            '\n' +
            state.items.length +
            ' Einträge (Upsert über Demo-ID, Tag ' +
            DEMO_SEED_TAG +
            ').\n\n' +
            'Hinweis: Bei aktivem Freistellungs-Flow können Approvals für NEUE Einträge starten. Flow vorher pausieren empfohlen.\n' +
            'Erneutes Demo aktualisiert bestehende Demo-Zeilen (kein neuer Trigger).'
    );
    if (!writeSp) return;

    state.siteUrl = siteUrl;
    persistSiteUrl(siteUrl);
    state.loading = true;
    state.info = 'Schreibe Demo auf SharePoint …';
    paint();
    try {
        const setup = loadSetupCfg();
        await seedDemoFreistellungen(siteUrl, (msg) => console.log('[fr-demo]', msg), {
            pack,
            listName: state.listName || setup.listName,
            listId: state.listId || setup.listId
        });
        state.localDemoOnly = false;
        await refreshData();
        toast('Demo auf SharePoint geschrieben und neu geladen.');
    } catch (e) {
        state.localDemoOnly = true;
        state.loading = false;
        state.error = String((e && e.message) || e);
        paint();
        toast('SharePoint-Seed fehlgeschlagen – lokale Demo bleibt: ' + state.error);
    }
}

async function runDemoReset() {
    const localOnly = state.localDemoOnly && !state.ctx;
    const siteUrl = String(state.siteUrl || DEMO_SITE_DEFAULT).trim();
    const ok = window.confirm(
        localOnly
            ? 'Lokale Demo-Daten leeren?'
            : 'Alle Demo-Einträge (Tag ' +
                  DEMO_SEED_TAG +
                  ') auf SharePoint löschen?\n\nSite: ' +
                  siteUrl +
                  '\n\nEchte (nicht-Demo) Einträge bleiben erhalten.'
    );
    if (!ok) return;

    if (localOnly) {
        state.items = [];
        state.localDemoOnly = false;
        state.info = 'Lokale Demo geleert.';
        paint();
        toast('Lokale Demo geleert.');
        return;
    }

    state.loading = true;
    paint();
    try {
        const setup = loadSetupCfg();
        const result = await resetDemoFreistellungen(
            siteUrl,
            (msg) => console.log('[fr-demo-reset]', msg),
            {
                listName: state.listName || setup.listName,
                listId: state.listId || setup.listId,
                seedTag: DEMO_SEED_TAG
            }
        );
        state.localDemoOnly = false;
        await refreshData();
        toast('Demo zurückgesetzt (' + (result.deleted || 0) + ' gelöscht).');
    } catch (e) {
        state.loading = false;
        state.error = String((e && e.message) || e);
        paint();
        toast(state.error);
    }
}

async function submitAntrag() {
    const draft = readFormFromDom(root);
    state.form = { ...state.form, ...draft };
    const check = validateFreistellung({ draft: state.form });
    if (!check.ok) {
        toast(check.errors[0] || 'Bitte Formular prüfen.');
        return;
    }
    if (state.localDemoOnly || !state.ctx) {
        const id = 'demo-' + Date.now();
        state.items = [
            {
                itemId: id,
                titel: draft.schuelerName + ' (' + draft.klasse + ')',
                schuelerName: draft.schuelerName,
                klasse: draft.klasse,
                beginn: draft.beginn,
                ende: draft.ende,
                status: 'Ausstehend',
                kategorie: draft.kategorie,
                beschreibung: draft.beschreibung,
                bemerkungen: '',
                kvEmail: draft.kvEmail,
                kvName: state.form.kvName || '',
                authorEmail: state.accountEmail,
                beantragtVon: state.accountEmail,
                dayCount: check.dayCount,
                multiDay: check.path.multiDay,
                approvalLabel: check.path.label
            },
            ...state.items
        ];
        state.form = emptyForm({ schuelerName: state.accountName });
        state.view = 'meine';
        toast('Demo: Antrag lokal gespeichert (nicht in SharePoint).');
        paint();
        return;
    }
    state.loading = true;
    paint();
    try {
        await createFreistellungItem(state.ctx, state.form);
        toast('Antrag eingereicht – Genehmigung startet über Microsoft Approvals.');
        state.form = emptyForm({ schuelerName: state.accountName });
        await refreshData();
        state.view = state.role === 'schueler' ? 'meine' : 'liste';
    } catch (e) {
        state.error = String((e && e.message) || e);
        state.loading = false;
        paint();
    }
}

async function setStatus(itemId, status) {
    if (state.localDemoOnly || !state.ctx) {
        state.items = state.items.map((it) =>
            String(it.itemId) === String(itemId) ? { ...it, status } : it
        );
        toast('Demo: Status → ' + status);
        paint();
        return;
    }
    try {
        await updateFreistellungStatus(state.ctx, itemId, status);
        toast('Status aktualisiert: ' + status);
        await refreshData();
    } catch (e) {
        toast(String((e && e.message) || e));
    }
}

function bindStatic() {
    root.querySelectorAll('[data-fr-view]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-fr-view') || 'dashboard';
            paint();
        });
    });
    root.querySelectorAll('[data-fr-view-jump]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.view = btn.getAttribute('data-fr-view-jump') || 'dashboard';
            paint();
        });
    });
    root.querySelectorAll('[data-fr-role]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const role = resolveRole(btn.getAttribute('data-fr-role'));
            state.role = role;
            persistRole(role);
            if (role === 'schueler' && (state.view === 'freigabe' || state.view === 'bericht')) {
                state.view = 'dashboard';
            }
            if ((role === 'kv' || role === 'direktion') && state.view === 'meine') {
                state.view = 'dashboard';
            }
            paint();
        });
    });

    const site = root.querySelector('#frSiteUrl');
    const btnLoad = root.querySelector('#frBtnLoad');
    if (btnLoad) {
        btnLoad.addEventListener('click', () => {
            if (site) {
                state.siteUrl = String(site.value || '').trim();
                persistSiteUrl(state.siteUrl);
            }
            refreshData();
        });
    }
    const btnCsv = root.querySelector('#frBtnCsv');
    if (btnCsv) btnCsv.addEventListener('click', exportCsv);
    const btnCsvInline = root.querySelector('#frBtnCsvInline');
    if (btnCsvInline) btnCsvInline.addEventListener('click', exportCsv);
    const btnDemo = root.querySelector('#frBtnDemo');
    if (btnDemo) btnDemo.addEventListener('click', () => runDemoImport().catch((e) => toast(String((e && e.message) || e))));
    const btnDemoReset = root.querySelector('#frBtnDemoReset');
    if (btnDemoReset) {
        btnDemoReset.addEventListener('click', () =>
            runDemoReset().catch((e) => toast(String((e && e.message) || e)))
        );
    }

    const form = root.querySelector('#frAntragForm');
    if (form) {
        form.addEventListener('submit', (ev) => {
            ev.preventDefault();
            submitAntrag();
        });
        const klasse = root.querySelector('#frFormKlasse');
        if (klasse) {
            klasse.addEventListener('change', () => {
                state.form = { ...state.form, ...readFormFromDom(root) };
                applyKvFromClass(state);
                paint();
            });
        }
        ['frFormBeginn', 'frFormEnde'].forEach((id) => {
            const el = root.querySelector('#' + id);
            if (el) {
                el.addEventListener('change', () => {
                    state.form = { ...state.form, ...readFormFromDom(root) };
                    paint();
                });
            }
        });
        const reset = root.querySelector('#frBtnFormReset');
        if (reset) {
            reset.addEventListener('click', () => {
                state.form = emptyForm({ schuelerName: state.accountName });
                paint();
            });
        }
    }

    const filters = root.querySelector('[data-fr-filters]');
    if (filters) {
        const apply = () => {
            state.filters = readFiltersFromDom(root);
            paint();
        };
        filters.querySelectorAll('select, input').forEach((el) => {
            el.addEventListener('change', apply);
            if (el.tagName === 'INPUT') {
                el.addEventListener('input', () => {
                    clearTimeout(el._frT);
                    el._frT = setTimeout(apply, 200);
                });
            }
        });
        const reset = root.querySelector('#frFilterReset');
        if (reset) {
            reset.addEventListener('click', () => {
                state.filters = { klasse: '', status: '', kategorie: '', multiDay: '', q: '' };
                paint();
            });
        }
    }

    root.querySelectorAll('[data-fr-detail]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.detailId = btn.getAttribute('data-fr-detail');
            paint();
        });
    });
    root.querySelectorAll('[data-fr-close-detail]').forEach((btn) => {
        btn.addEventListener('click', () => {
            state.detailId = null;
            paint();
        });
    });
    root.querySelectorAll('[data-fr-approve]').forEach((btn) => {
        btn.addEventListener('click', () => setStatus(btn.getAttribute('data-fr-approve'), 'Genehmigt'));
    });
    root.querySelectorAll('[data-fr-reject]').forEach((btn) => {
        btn.addEventListener('click', () => setStatus(btn.getAttribute('data-fr-reject'), 'Abgelehnt'));
    });
}

function boot() {
    root = document.getElementById('frApp');
    if (!root) return;
    syncAccount();
    const setup = loadSetupCfg();
    if (!state.siteUrl && setup.siteUrl) state.siteUrl = setup.siteUrl;
    if (state.form.klasse) {
        const kv = resolveKvForClass(state.stammdaten.classes, state.form.klasse);
        if (kv) {
            state.form.kvEmail = kv.email;
            state.form.kvName = kv.name;
        }
    }
    paint();
    if (state.siteUrl) {
        refreshData();
    } else {
        state.info =
            'Site-URL aus Freistellungen-Setup übernehmen oder eintragen, dann „Laden“. Alternativ Demo nutzen.';
        paint();
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
