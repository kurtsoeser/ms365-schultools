import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function createMemoryStorage(initial) {
    const map = new Map(Object.entries(initial || {}));
    return {
        get length() {
            return map.size;
        },
        key(i) {
            return Array.from(map.keys())[i] ?? null;
        },
        getItem(k) {
            return map.has(k) ? map.get(k) : null;
        },
        setItem(k, v) {
            map.set(String(k), String(v));
        },
        removeItem(k) {
            map.delete(k);
        }
    };
}

function loadBackup(store, sessionStore) {
    const full = join(projectRoot, 'src/shared/browser-backup.js');
    const code = readFileSync(full, 'utf8');
    const sandbox = {
        console,
        document: {
            readyState: 'complete',
            getElementById() {
                return null;
            },
            querySelectorAll() {
                return [];
            },
            addEventListener() {}
        }
    };
    sandbox.window = sandbox;
    sandbox.localStorage = store;
    sandbox.sessionStorage = sessionStore || createMemoryStorage();
    createContext(sandbox);
    runInContext(code, sandbox, { filename: full });
    return sandbox;
}

function mountAppDataV2Stub(sandbox, store, v2, tenantSave) {
    sandbox.ms365AppDataV2 = {
        getContainer() {
            try {
                const raw = store.getItem('ms365-schooltool-data-v2');
                return raw ? JSON.parse(raw) : JSON.parse(JSON.stringify(v2));
            } catch {
                return JSON.parse(JSON.stringify(v2));
            }
        },
        setContainer(next) {
            store.setItem('ms365-schooltool-data-v2', JSON.stringify(next));
            return next;
        }
    };
    if (tenantSave) sandbox.ms365TenantSettingsSave = tenantSave;
}

describe('browser-backup', () => {
    let store;
    let sessionStore;
    let api;

    beforeEach(() => {
        store = createMemoryStorage({
            'ms365-schooltool-data-v2': JSON.stringify({ version: 3, core: { domain: 'schule.at' } }),
            'ms365-dashboard-favorites-v1': JSON.stringify(['kursteams']),
            'webuntis-teams-creator-state-v1': JSON.stringify({ step: 2 }),
            'ms365-dashboard-expert-open-v1': '1',
            'msal.token.keys.abc': 'secret',
            'acc-login.microsoftonline.com-idtoken-x': 'token',
            'unrelated-other-app': 'nope'
        });
        sessionStore = createMemoryStorage({
            'ms365-gaeste-verwalten-active-tab-v1': 'audit',
            'ms365-gast-zugaenge-snapshot-v1': JSON.stringify({ teams: [], invitations: [], savedAt: '2026-08-17T13:00:00.000Z' }),
            'ms365-access-granted-v1': '1',
            'ms365-post-login-url': 'https://example.invalid/index.html',
            'msal.interaction.status': 'busy'
        });
        api = loadBackup(store, sessionStore).ms365BrowserBackup;
    });

    it('erkennt App-Keys und blendet MSAL/fremde Keys aus', () => {
        expect(api.isAppStorageKey('ms365-schooltool-data-v2')).toBe(true);
        expect(api.isAppStorageKey('webuntis-teams-creator-state-v1')).toBe(true);
        expect(api.isAppStorageKey('msal.token.keys.abc')).toBe(false);
        expect(api.isAuthKey('acc-login.microsoftonline.com-idtoken-x')).toBe(true);
        expect(api.isAppStorageKey('unrelated-other-app')).toBe(false);
    });

    it('sammelt nur App-Werte und erhält Nicht-JSON-Strings', () => {
        const collected = api.collectLocalStorage(store);
        expect(Object.keys(collected).sort()).toEqual([
            'ms365-dashboard-expert-open-v1',
            'ms365-dashboard-favorites-v1',
            'ms365-schooltool-data-v2',
            'webuntis-teams-creator-state-v1'
        ]);
        expect(collected['ms365-dashboard-expert-open-v1']).toBe('1');
        expect(collected['ms365-schooltool-data-v2']).toEqual({ version: 3, core: { domain: 'schule.at' } });
        expect(collected['msal.token.keys.abc']).toBeUndefined();
        expect(collected['unrelated-other-app']).toBeUndefined();
    });

    it('baut ein erkennbares Backup-Paket', () => {
        const now = new Date(2026, 7, 17, 15, 0, 0);
        const storeWithSchool = createMemoryStorage({
            'ms365-schooltool-data-v2': JSON.stringify({ version: 4, core: { schoolName: 'BRG Muster', domain: 'brg-muster.at' }, structure: {}, match: {} }),
            'ms365-dashboard-favorites-v1': JSON.stringify(['kursteams']),
            'webuntis-teams-creator-state-v1': JSON.stringify({ step: 2 }),
            'ms365-dashboard-expert-open-v1': '1'
        });
        const apiWithSchool = loadBackup(storeWithSchool, sessionStore).ms365BrowserBackup;
        const payload = apiWithSchool.buildBackup(storeWithSchool, now);
        expect(apiWithSchool.isBackupPayload(payload)).toBe(true);
        expect(payload.kind).toBe('ms365-browser-backup-v1');
        expect(payload.version).toBe(7);
        expect(payload.exportedAt).toBe(now.toISOString());
        expect(payload.schoolName).toBe('BRG Muster');
        expect(payload.domain).toBe('brg-muster.at');
        expect(payload.keyCount).toBe(6);
        expect(payload.contentFingerprint).toMatch(/^[0-9a-f]+:\d+$/);
        expect(payload.includesNote).toMatch(/Power-Automate/);
        expect(payload.legacyMirror).toBeTruthy();
        expect(payload.legacyMirror.writeEnabled).toBe(false);
        expect(payload.legacyMirror.canonicalKey).toBe('ms365-schooltool-data-v2');
        expect(payload.legacyMirror.canonicalPresent).toBe(true);
        expect(Object.keys(payload.sessionStorage || {}).sort()).toEqual([
            'ms365-gaeste-verwalten-active-tab-v1',
            'ms365-gast-zugaenge-snapshot-v1'
        ]);
        expect(apiWithSchool.backupFilename(now, storeWithSchool)).toBe(
            'ms365-browser-backup-2026-08-17_1500-BRG_Muster.json'
        );
    });

    it('syncStorageBeforeBackup behält Verwaltungs-Zielgruppen in v2', () => {
        const v2 = {
            version: 4,
            core: {
                schoolName: 'Test',
                domain: 't.at',
                verwaltungAudienceGroups: [{ id: 'schulleitung', label: 'Schulleitung', builtin: true }],
                adminAudienceMemberships: [{ groupId: 'schulleitung', email: 'dir@t.at' }]
            },
            years: { current: '2025/26', byLabel: { '2025/26': { students: [], classes: [] } } },
            setup: {},
            structure: {},
            match: {},
            tenant: {}
        };
        const stStore = createMemoryStorage({
            'ms365-schooltool-data-v2': JSON.stringify(v2),
            'ms365-tenant-settings-v1': JSON.stringify({ version: 3, schoolName: 'Test', domain: 't.at' })
        });
        const stSandbox = loadBackup(stStore, sessionStore);
        const wipeSave = (partial) => {
            const raw = stStore.getItem('ms365-tenant-settings-v1');
            const cur = raw ? JSON.parse(raw) : {};
            stStore.setItem('ms365-tenant-settings-v1', JSON.stringify({ ...cur, ...partial }));
            if (partial.verwaltungAudienceGroups !== undefined || partial.adminAudienceMemberships !== undefined) {
                const v2raw = JSON.parse(stStore.getItem('ms365-schooltool-data-v2'));
                if (partial.verwaltungAudienceGroups !== undefined) {
                    v2raw.core.verwaltungAudienceGroups = partial.verwaltungAudienceGroups;
                }
                if (partial.adminAudienceMemberships !== undefined) {
                    v2raw.core.adminAudienceMemberships = partial.adminAudienceMemberships;
                }
                stStore.setItem('ms365-schooltool-data-v2', JSON.stringify(v2raw));
            }
        };
        mountAppDataV2Stub(stSandbox, stStore, v2, wipeSave);
        const stApi = stSandbox.ms365BrowserBackup;
        stApi.syncStorageBeforeBackup();
        const after = JSON.parse(stStore.getItem('ms365-schooltool-data-v2'));
        expect(after.core.verwaltungAudienceGroups).toHaveLength(1);
        expect(after.core.adminAudienceMemberships).toHaveLength(1);
    });

    it('nimmt Power-Automate-Konfiguration in den Export auf', () => {
        const paStore = createMemoryStorage({
            'ms365-schooltool-data-v2': JSON.stringify({ version: 4, core: { schoolName: 'Test', domain: 't.at' } }),
            'ms365-freistellung-setup-v1': JSON.stringify({ siteUrl: 'https://schule.sharepoint.com/sites/Intranet', emailMailbox: 'automate@t.at' }),
            'ms365-pa-termine-sync-v1': JSON.stringify({ mailboxSchool: 'kalender@t.at' }),
            'ms365-pa-done-freistellung': '1'
        });
        const paApi = loadBackup(paStore, sessionStore).ms365BrowserBackup;
        const payload = paApi.buildBackup(paStore);
        expect(payload.localStorage['ms365-freistellung-setup-v1'].emailMailbox).toBe('automate@t.at');
        expect(payload.localStorage['ms365-pa-termine-sync-v1'].mailboxSchool).toBe('kalender@t.at');
        expect(payload.localStorage['ms365-pa-done-freistellung']).toBe('1');
        const inv = (payload.inventory && payload.inventory.categories) || [];
        expect(inv.some((c) => /Power Automate/i.test(c.label))).toBe(true);
        expect(typeof paApi.notifyLocalDataChanged).toBe('function');
    });

    it('stellt ein Backup 1:1 wieder her und räumt Ziel-Reste weg', () => {
        const payload = api.buildBackup(store);
        const target = createMemoryStorage({
            'ms365-jahrgang-state-v1': '{"old":true}',
            'ms365-dashboard-tab-v1': 'sharepoint',
            'msal.cache': 'keep-me'
        });
        const targetSession = createMemoryStorage({
            'ms365-gaeste-verwalten-active-tab-v1': 'teams',
            'ms365-post-login-url': '/keep-login',
            'ms365-gast-zugaenge-snapshot-v1': '{"old":true}'
        });
        const result = api.applyBackup(payload, target, { sessionStorage: targetSession });
        expect(target.getItem('ms365-jahrgang-state-v1')).toBeNull();
        expect(target.getItem('ms365-dashboard-tab-v1')).toBeNull();
        expect(target.getItem('msal.cache')).toBe('keep-me');
        expect(JSON.parse(target.getItem('ms365-schooltool-data-v2')).core.domain).toBe('schule.at');
        expect(target.getItem('ms365-dashboard-expert-open-v1')).toBe('1');
        expect(target.getItem('webuntis-teams-creator-state-v1')).toContain('"step":2');
        expect(targetSession.getItem('ms365-gaeste-verwalten-active-tab-v1')).toBe('audit');
        expect(targetSession.getItem('ms365-gast-zugaenge-snapshot-v1')).toContain('"savedAt":"2026-08-17T13:00:00.000Z"');
        expect(targetSession.getItem('ms365-post-login-url')).toBe('/keep-login');
        expect(result.written).toContain('ms365-schooltool-data-v2');
        expect(result.removed).toContain('ms365-jahrgang-state-v1');
        expect(result.sessionWritten).toContain('ms365-gaeste-verwalten-active-tab-v1');
    });

    it('schreibt MSAL-Keys aus einer Datei nicht zurück', () => {
        const target = createMemoryStorage();
        const result = api.applyBackup(
            {
                kind: api.KIND,
                version: 3,
                localStorage: {
                    'ms365-school-email-domain-v1': 'schule.at',
                    'msal.token.keys.abc': 'stolen'
                }
            },
            target
        );
        expect(target.getItem('ms365-school-email-domain-v1')).toBe('schule.at');
        expect(target.getItem('msal.token.keys.abc')).toBeNull();
        expect(result.skipped).toContain('msal.token.keys.abc');
    });

    it('erkennt Legacy-Schuldaten und lehnt unbekanntes JSON ab', async () => {
        expect(api.isLegacyAppDataPayload({ version: 3, core: {}, structure: {}, match: {} })).toBe(true);
        expect(api.isBackupPayload({ version: 3, core: {}, structure: {}, match: {} })).toBe(false);
        await expect(api.importPayload({ foo: 1 }, store)).rejects.toThrow(/Unbekanntes JSON-Format/);
        expect(() => api.applyBackup({ kind: 'nope', localStorage: {} }, store)).toThrow(/Kein gültiges Browser-Backup/);
    });

    it('importPayload spielt ein Backup auf den übergebenen Store', async () => {
        const payload = api.buildBackup(store);
        const target = createMemoryStorage({ 'ms365-arge-state-v2': '{}' });
        const imported = await api.importPayload(payload, target);
        expect(imported.mode).toBe('backup');
        expect(target.getItem('ms365-arge-state-v2')).toBeNull();
        expect(target.getItem('ms365-dashboard-favorites-v1')).toContain('kursteams');
    });

    it('deckt Stammdaten, E-Mail-Muster und Schul-Apps in criticalCoverage ab', () => {
        const demoStore = createMemoryStorage({
            'ms365-schooltool-data-v2': JSON.stringify({
                version: 4,
                core: {
                    schoolName: 'Demo',
                    domain: 'demo.at',
                    teachers: [{ name: 'A' }],
                    subjects: [{ code: 'D' }],
                    arges: [{ code: 'AG' }]
                },
                years: {
                    current: '2025/26',
                    byLabel: { '2025/26': { students: [{ name: 'S' }], classes: [{ name: '1A' }] } }
                },
                setup: {
                    schoolProfile: {
                        schoolCode: '302123',
                        logoDataUrl: 'data:image/png;base64,iVBORw0KGgo=',
                        logoFileName: 'logo.png',
                        street: 'Schulweg 1',
                        postalCode: '1010',
                        city: 'Wien',
                        country: 'Österreich',
                        phone: '+43 1 234',
                        email: 'office@demo.at',
                        website: 'https://demo.at'
                    },
                    intranetSiteUrl: 'https://schule.sharepoint.com/sites/Intranet',
                    intranetListTitles: { lehrer: 'Lehrer' }
                }
            }),
            'ms365-tenant-school-email-pattern-v1': JSON.stringify({
                patternId: 'vorname.nachname',
                firstNameMode: 'all'
            }),
            'ms365-sa-site-url': 'https://schule.sharepoint.com/sites/Schularbeiten',
            'ms365-sa-settings-v1': JSON.stringify({ foo: 1 }),
            'ms365-pw-site-url': 'https://schule.sharepoint.com/sites/Projektwochen',
            'ms365-intranet-list-links-v1': JSON.stringify({ lehrer: 'https://...' }),
            'ms365-dashboard-tool-access-v1': JSON.stringify({ tools: {} }),
            'ms365-einsammler-settings-v1': JSON.stringify({ siteUrl: 'https://schule.sharepoint.com/sites/Intranet' }),
            'ms365-elternbrief-sp-settings-v1': JSON.stringify({ listTitle: 'Elternbriefe' }),
            'ms365-schueler-stammdaten-portal-v1': JSON.stringify({ tenants: {} })
        });
        const demoApi = loadBackup(demoStore, sessionStore).ms365BrowserBackup;
        const payload = demoApi.buildBackup(demoStore);
        expect(payload.schooltoolSummary.schoolCode).toBe('302123');
        expect(payload.schooltoolSummary.hasSchoolProfileLogo).toBe(true);
        expect(payload.schooltoolSummary.schoolProfileCity).toBe('Wien');
        expect(payload.schooltoolSummary.teacherCount).toBe(1);
        expect(payload.schooltoolSummary.intranetListTitleCount).toBe(1);
        expect(payload.criticalCoverage.readyForDemoTransfer).toBe(true);
        expect(payload.criticalCoverage.missingIds).toEqual([]);
        const apps = payload.criticalCoverage.schoolApps.filter((a) => a.present).map((a) => a.id);
        expect(apps).toContain('schularbeiten');
        expect(apps).toContain('projektwochen');
        expect(apps).toContain('einsammler');
        expect(apps).toContain('elternbrief');
        expect(apps).toContain('stammdaten-portal');
        expect(payload.indexedDB).toEqual({});
        expect(Array.isArray(payload.includesIndexedDb)).toBe(true);
        // Volles Schulprofil inkl. Logo muss im kanonischen v2-Block stecken.
        const v2Export = payload.localStorage['ms365-schooltool-data-v2'];
        expect(v2Export.setup.schoolProfile.logoDataUrl).toContain('data:image/png;base64,');
        expect(v2Export.setup.schoolProfile.street).toBe('Schulweg 1');
        expect(v2Export.setup.schoolProfile.email).toBe('office@demo.at');
        const target = createMemoryStorage();
        demoApi.applyBackup(payload, target);
        expect(JSON.parse(target.getItem('ms365-tenant-school-email-pattern-v1')).firstNameMode).toBe('all');
        expect(target.getItem('ms365-sa-site-url')).toContain('Schularbeiten');
        expect(target.getItem('ms365-pw-site-url')).toContain('Projektwochen');
        expect(target.getItem('ms365-einsammler-settings-v1')).toContain('Intranet');
        expect(target.getItem('ms365-schueler-stammdaten-portal-v1')).toContain('tenants');
        const restored = JSON.parse(target.getItem('ms365-schooltool-data-v2'));
        expect(restored.setup.schoolProfile.logoDataUrl).toContain('data:image/png;base64,');
        expect(restored.setup.schoolProfile.logoFileName).toBe('logo.png');
        expect(restored.setup.schoolProfile.city).toBe('Wien');
        expect(restored.setup.schoolProfile.website).toBe('https://demo.at');
    });

    it('nimmt alle Schul-App-Keys per Präfix mit und listet sie im Katalog', () => {
        const schoolAppKeys = [
            'ms365-einsammler-v1',
            'ms365-elternbrief-settings-v1',
            'ms365-krankmeldung-sp-settings-v1',
            'ms365-kollegium-umfragen-v1',
            'ms365-ausgaben-settings-v1',
            'ms365-wahlfach-zuteilung-v1',
            'ms365-schulsprecherwahl-v1',
            'ms365-teams-antrag-settings-v1',
            'ms365-diplomarbeiten-state-v1',
            'ms365-raumplan-v1',
            'ms365-aufgaben-sp-settings-v1',
            'ms365-technik-meldungen-v1',
            'ms365-termine-hub-settings-v1',
            'ms365-kalender-postfaecher-v1',
            'ms365-akt-ics-feed-v1',
            'ms365-class-chats-mirror-v1'
        ];
        const init = { 'ms365-schooltool-data-v2': JSON.stringify({ version: 4, core: { domain: 't.at' } }) };
        schoolAppKeys.forEach((k) => {
            init[k] = JSON.stringify({ ok: k });
        });
        const fullStore = createMemoryStorage(init);
        const fullApi = loadBackup(fullStore, sessionStore).ms365BrowserBackup;
        const payload = fullApi.buildBackup(fullStore);
        schoolAppKeys.forEach((k) => {
            expect(payload.localStorage[k]).toEqual({ ok: k });
        });
        const catalogKeys = new Set();
        fullApi.STORAGE_CATALOG.forEach((cat) => cat.keys.forEach((k) => catalogKeys.add(k)));
        schoolAppKeys.forEach((k) => {
            expect(catalogKeys.has(k)).toBe(true);
        });
        expect(fullApi.INDEXED_DB_CATALOG.map((e) => e.name)).toEqual(
            expect.arrayContaining([
                'ms365-einsammler-attachments',
                'ms365-elternbrief-attachments',
                'ms365-krankmeldung-attachments',
                'ms365-kollegium-umfragen-attachments'
            ])
        );
        const target = createMemoryStorage();
        fullApi.applyBackup(payload, target);
        schoolAppKeys.forEach((k) => {
            expect(JSON.parse(target.getItem(k)).ok).toBe(k);
        });
    });

    it('buildBackupAsync liefert indexedDB-Feld auch ohne IndexedDB-Runtime', async () => {
        const payload = await api.buildBackupAsync(store);
        expect(payload.version).toBe(7);
        expect(payload.indexedDB).toEqual({});
        expect(payload.indexedDbCount).toBe(0);
    });

    it('lässt SharePoint-Session-Pull-Marker nicht ins Backup importieren', () => {
        const payload = api.buildBackup(store);
        payload.sessionStorage['ms365-spo-auto-pull-done-v1:tenant-abc'] = '1';
        const targetSession = createMemoryStorage();
        api.applyBackup(payload, store, { sessionStorage: targetSession });
        expect(targetSession.getItem('ms365-spo-auto-pull-done-v1:tenant-abc')).toBeNull();
    });

    it('ignoriert SharePoint-Sync-Meta im Inhalts-Fingerprint', () => {
        const base = createMemoryStorage({
            'ms365-schooltool-data-v2': JSON.stringify({ version: 4, core: { domain: 't.at' } })
        });
        const apiBase = loadBackup(base, sessionStore).ms365BrowserBackup;
        const fp1 = apiBase.contentFingerprint(base);
        base.setItem(
            'ms365-stammdaten-spo-sync-by-tenant-v2',
            JSON.stringify({ 'tid-1': { dirty: true, at: '2099-01-01T00:00:00.000Z' } })
        );
        expect(apiBase.contentFingerprint(base)).toBe(fp1);
    });

    it('invalidiert app-data-v2-Cache nach Backup-Import', () => {
        let invalidated = false;
        const target = createMemoryStorage();
        const sandbox = loadBackup(target, sessionStore);
        sandbox.ms365AppDataV2 = {
            invalidateCache() {
                invalidated = true;
            },
            getContainer() {
                return { core: {}, years: { current: '2025/26', byLabel: {} } };
            }
        };
        sandbox.ms365TenantSettingsLoad = function () {
            return { domain: 'schule.at' };
        };
        sandbox.ms365SetSchoolDomainNoAt = function () {};
        const payload = sandbox.ms365BrowserBackup.buildBackup(store);
        sandbox.ms365BrowserBackup.applyBackup(payload, target);
        expect(invalidated).toBe(true);
    });

    it('Config-Bundle mit domain löscht vorhandene Lehrkräfte nicht', async () => {
        const target = createMemoryStorage({
            'ms365-schooltool-data-v2': JSON.stringify({
                version: 4,
                core: { teachers: [{ code: 'ADA', name: 'Ada' }] },
                years: { current: '2026/27', byLabel: { '2026/27': { classes: [{ code: '1A' }], students: [] } } }
            })
        });
        const sandbox = loadBackup(target, sessionStore);
        let saved = false;
        sandbox.ms365TenantSettingsSave = function () {
            saved = true;
        };
        sandbox.ms365AppDataV2 = {
            importJson() {
                saved = true;
            }
        };
        const imported = await sandbox.ms365BrowserBackup.importPayload({
            bundleKind: 'ms365-spo-config-bundle-v1',
            domain: 'ms365.schule',
            localStorage: {
                'ms365-dashboard-favorites-v1': JSON.stringify(['kursteams'])
            }
        }, target);
        expect(imported.mode).toBe('config-bundle');
        expect(saved).toBe(false);
        const v2 = JSON.parse(target.getItem('ms365-schooltool-data-v2'));
        expect(v2.core.teachers).toEqual([{ code: 'ADA', name: 'Ada' }]);
        expect(v2.years.byLabel['2026/27'].classes).toEqual([{ code: '1A' }]);
        expect(target.getItem('ms365-dashboard-favorites-v1')).toContain('kursteams');
    });

    it('Config-Manifest ohne Listen weist auf die Backup-Datei hin', async () => {
        await expect(
            api.importPayload({
                bundleKind: 'ms365-spo-config-bundle-v1',
                domain: 'ms365.schule',
                files: [{ id: 'stammdaten-core', path: 'config/stammdaten-core.json' }]
            }, store)
        ).rejects.toThrow(/ms365-browser-backup/);
    });

    it('IndexedDB-Fehler lassen das geschriebene Backup stehen', async () => {
        const target = createMemoryStorage();
        const sandbox = loadBackup(target, sessionStore);
        sandbox.indexedDB = {
            open() {
                throw new Error('idb kaputt');
            },
            deleteDatabase() {
                throw new Error('idb kaputt');
            }
        };
        const payload = sandbox.ms365BrowserBackup.buildBackup(store);
        payload.indexedDB = {
            'ms365-einsammler-attachments': {
                stores: { files: [{ key: 'a', value: { n: 1 } }] }
            }
        };
        const result = await sandbox.ms365BrowserBackup.applyBackupAsync(payload, target);
        expect(result.warnings.join(' ')).toMatch(/idb kaputt|IndexedDB/);
        expect(target.getItem('ms365-schooltool-data-v2')).toContain('schule.at');
    });
});

describe('browser-backup Seiten', () => {
    const pages = [
        { file: 'index.html', scriptPrefix: 'src/shared/browser-backup.js' },
        { file: 'tenant.html', scriptPrefix: 'src/shared/browser-backup.js' }
    ];

    it.each(pages)('$file bindet Backup-Skript und Steuerelemente ein', ({ file, scriptPrefix }) => {
        const html = readFileSync(join(projectRoot, file), 'utf8');
        expect(html).toContain(scriptPrefix);
        if (file === 'tenant.html') {
            expect(html).toContain('data-ms365-backup="export"');
        } else {
            expect(html).toContain('id="browserBackupExport"');
        }
        expect(html).toContain('id="browserBackupImportFile"');
    });
});
