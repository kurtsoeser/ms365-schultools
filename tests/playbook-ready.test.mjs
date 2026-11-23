import { describe, expect, it } from 'vitest';
import {
    formatReadyHint,
    gatesToReadyMap,
    normalizeReadyResult,
    readyForStep
} from '../src/shared/playbook-ready.js';
import { evaluateSchuljahresstartReadyChecks } from '../src/shared/playbook-schuljahresstart-gates.js';
import {
    blaupauseCountsFromSettings,
    evaluateBlaupauseReadyChecks
} from '../src/shared/playbook-blaupause-ready.js';
import {
    evaluateAppPlaybookReadyChecks,
    evaluateElternPlaybookReadyChecks,
    evaluateFreistellungenPlaybookReadyChecks,
    evaluateKursteamsPlaybookReadyChecks,
    evaluateSchularbeitenPlaybookReadyChecks
} from '../src/shared/playbook-tool-ready.js';
import {
    freistellungSetupDeepLink,
    PLAYBOOK_STAMMDATEN_HASH,
    schularbeitenPlanerDeepLink,
    stammdatenDeepLink
} from '../src/shared/playbook-deeplink.js';
import { paneIdFromLocationHash } from '../src/shared/tenant-stammdaten-setup-nav.js';
import {
    DASHBOARD_PLAYBOOKS,
    PLANNED_APP_PLAYBOOKS,
    allPlaybookDefs
} from '../src/shared/dashboard-playbooks-catalog.js';
import { defaultItWorkspaceTree } from '../src/shared/it-workspace-nav.js';

describe('playbook-ready', () => {
    it('normalisiert ok/missing aus Legacy-Gates', () => {
        expect(normalizeReadyResult({ ok: true, message: 'x' })).toEqual({
            status: 'ok',
            message: 'x'
        });
        expect(normalizeReadyResult({ ok: false, message: 'fehlt' }).status).toBe('missing');
        expect(gatesToReadyMap({ a: { ok: true, message: '' } }).a.status).toBe('ok');
    });

    it('formatiert Ready-Hinweise', () => {
        expect(formatReadyHint({ status: 'ok', message: 'fertig' })).toBe('Bereit: fertig');
        expect(formatReadyHint({ status: 'missing', message: 'Domain' })).toMatch(/^Fehlt:/);
        expect(readyForStep({ a: { status: 'warn', message: 'w' } }, 'a').status).toBe('warn');
    });

    it('liefert Schuljahresstart-Ready-Checks', () => {
        const map = evaluateSchuljahresstartReadyChecks({
            year: '2026/27',
            classes: 2,
            students: 5,
            teachers: 1,
            subjects: 3,
            domain: 'schule.at',
            schoolName: 'Demo'
        });
        expect(map.stammdaten.status).toBe('ok');
        expect(map.kursteams.status).toBe('ok');
        expect(map.cleanup.status).toBe('ok');
    });

    it('liefert Blaupause-Ready-Checks für Profil und Primär-Site', () => {
        const missing = evaluateBlaupauseReadyChecks(
            blaupauseCountsFromSettings({ schoolName: '', domain: '', classes: [] })
        );
        expect(missing.register.status).toBe('missing');
        expect(missing['tool-site'].status).toBe('missing');

        const ok = evaluateBlaupauseReadyChecks(
            blaupauseCountsFromSettings({
                schoolName: 'Demo',
                domain: 'schule.at',
                schoolYear: '2026/27',
                classes: [{ id: 1 }],
                students: [{ id: 1 }],
                teachers: [{ id: 1 }],
                subjects: [{ id: 1 }],
                planerDefaultSiteUrl: 'https://contoso.sharepoint.com/sites/tools',
                stammdatenItLibrary: { siteUrl: 'https://contoso.sharepoint.com/sites/it' }
            })
        );
        expect(ok.register.status).toBe('ok');
        expect(ok['tool-site'].status).toBe('ok');
        expect(ok.backup.status).toBe('ok');
    });

    it('liefert SA- und Freistellung-Ready-Checks', () => {
        const sa = evaluateSchularbeitenPlaybookReadyChecks({
            classes: [{ id: 1 }],
            subjects: [{ id: 1 }],
            planerDefaultSiteUrl: 'https://contoso.sharepoint.com/sites/tools'
        });
        expect(sa.stammdaten.status).toBe('ok');
        expect(sa.listen.status).toBe('ok');
        expect(sa.intranet.status).toBe('ok');

        const fr = evaluateFreistellungenPlaybookReadyChecks({
            classes: [],
            planerDefaultSiteUrl: ''
        });
        expect(fr.liste.status).toBe('warn');
        expect(fr['planer-config'].status).toBe('warn');
    });

    it('liefert App-/Kursteams-/Eltern-Ready-Checks (Wave 3)', () => {
        const app = evaluateAppPlaybookReadyChecks({
            planerDefaultSiteUrl: 'https://contoso.sharepoint.com/sites/tools'
        });
        expect(app.site.status).toBe('ok');
        expect(app.listen.status).toBe('ok');

        const kt = evaluateKursteamsPlaybookReadyChecks({
            classes: [{ id: 1 }],
            subjects: [{ id: 1 }],
            teachers: [{ id: 1 }],
            students: [{ id: 1 }]
        });
        expect(kt.kursteams.status).toBe('ok');

        const el = evaluateElternPlaybookReadyChecks({ students: [], classes: [] });
        expect(el.guardians.status).toBe('missing');
    });
});

describe('playbook-deeplink', () => {
    it('baut Stammdaten-, SA- und Freistellung-Links', () => {
        expect(stammdatenDeepLink('lehrer')).toBe('../tenant.html#lehrer');
        expect(PLAYBOOK_STAMMDATEN_HASH.planerApps).toBe('#spo-planer');
        expect(schularbeitenPlanerDeepLink('regeln')).toBe('schularbeiten-planer.html?view=regeln');
        expect(freistellungSetupDeepLink(3)).toBe('freistellung-setup.html?step=3');
        expect(freistellungSetupDeepLink(99)).toBe('freistellung-setup.html?step=7');
    });

    it('mapped Setup-Pane-Hashes für Planer & Apps', () => {
        expect(paneIdFromLocationHash('#spo-planer')).toBe('spo-planer');
        expect(paneIdFromLocationHash('#planer-apps')).toBe('spo-planer');
        expect(paneIdFromLocationHash('#profile')).toBe('profile');
    });
});

describe('playbook catalog wave 3', () => {
    it('führt App-Playbooks live im Dashboard (keine Stubs mehr)', () => {
        expect(DASHBOARD_PLAYBOOKS.length).toBe(16);
        expect(PLANNED_APP_PLAYBOOKS.length).toBe(0);
        expect(allPlaybookDefs().length).toBe(16);
        expect(DASHBOARD_PLAYBOOKS.some((p) => p.id === 'schuljahresstart')).toBe(true);
        expect(DASHBOARD_PLAYBOOKS.some((p) => p.id === 'schulaktivitaeten')).toBe(true);
        expect(DASHBOARD_PLAYBOOKS.some((p) => p.id === 'wahlen')).toBe(true);
        expect(DASHBOARD_PLAYBOOKS.filter((p) => p.categoryId === 'schulapps').length).toBe(7);
    });
});

describe('it-workspace-nav schuljahr', () => {
    it('verlinkt Schuljahresstart-Playbook statt totem schuljahreswechsel', () => {
        const tree = defaultItWorkspaceTree();
        const sj = tree.find((b) => b.id === 'dashTaskSchuljahr');
        expect(sj).toBeTruthy();
        expect(sj.tools[0].href).toBe('tools/playbook-schuljahresstart.html');
        expect(sj.tools[0].href).not.toMatch(/schuljahreswechsel/);
    });
});
