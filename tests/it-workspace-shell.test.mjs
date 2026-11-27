import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
    canonicalizeWorkspaceHref,
    clearItWorkspaceNav,
    deactivateItWorkspace,
    extractToolsFromTask,
    IT_WORKSPACE_FLAG_KEY,
    IT_WORKSPACE_NAV_KEY,
    IT_WORKSPACE_STAMMDATEN_TASK_ID,
    isDashboardItShellActive,
    normalizeWorkspaceHref,
    persistItWorkspaceNav,
    resolveWorkspaceHrefForPage,
    sanitizeItWorkspaceTree,
    workspaceHrefsMatch
} from '../src/shared/it-workspace-nav.js';
import { mergeStammdatenWorkspaceTools } from '../src/shared/stammdaten-nav-catalog.js';
import {
    mountItWorkspaceShell,
    resolveActiveWorkspaceTaskId,
    shouldMountItWorkspaceShell,
    trySoftItWorkspaceNavigate,
    unmountItWorkspaceShell,
    isItDashboardHomePath,
    isItContentSwapHref,
    itWorkspaceFrameNeedsSrc,
    trySoftItWorkspaceFrameNavigate
} from '../src/shared/it-workspace-shell.js';

function makeStorage() {
    const map = new Map();
    return {
        getItem: (k) => (map.has(k) ? map.get(k) : null),
        setItem: (k, v) => map.set(k, String(v)),
        removeItem: (k) => map.delete(k),
        clear: () => map.clear()
    };
}

describe('it-workspace-shell & nav gating', () => {
    beforeEach(() => {
        const local = makeStorage();
        const session = makeStorage();
        vi.stubGlobal('localStorage', local);
        vi.stubGlobal('sessionStorage', session);
        vi.stubGlobal('location', {
            href: 'https://app.ms365.schule/tools/elternbrief.html',
            pathname: '/tools/elternbrief.html',
            search: '',
            hash: ''
        });
        const attrs = new Set();
        const docEl = {
            hasAttribute: (k) => attrs.has(k),
            setAttribute: (k) => attrs.add(k),
            removeAttribute: (k) => attrs.delete(k),
            toggleAttribute: (k, force) => (force ? attrs.add(k) : attrs.delete(k))
        };
        const bodyClasses = new Set();
        const body = {
            classList: {
                contains: (c) => bodyClasses.has(c),
                add: (...cs) => cs.forEach((c) => bodyClasses.add(c)),
                remove: (...cs) => cs.forEach((c) => bodyClasses.delete(c))
            }
        };
        vi.stubGlobal('document', {
            documentElement: docEl,
            body,
            getElementById: () => null,
            querySelectorAll: () => []
        });
        vi.stubGlobal('window', {
            location: globalThis.location,
            addEventListener: () => {},
            __ms365DashboardPersonasLast: null
        });
    });

    it('lässt Portal vorbereiten als Root-Seite (nicht unter tools/)', () => {
        expect(normalizeWorkspaceHref('dashboard-werkzeug-zugriff.html')).toBe(
            'dashboard-werkzeug-zugriff.html'
        );
        expect(normalizeWorkspaceHref('dashboard-werkzeug-zugriff.html#freigeben')).toBe(
            'dashboard-werkzeug-zugriff.html#freigeben'
        );
        expect(resolveWorkspaceHrefForPage('dashboard-werkzeug-zugriff.html', true)).toBe(
            '../dashboard-werkzeug-zugriff.html'
        );
        expect(resolveWorkspaceHrefForPage('dashboard-werkzeug-zugriff.html', false)).toBe(
            'dashboard-werkzeug-zugriff.html'
        );
    });

    it('erkennt Dashboard-IT-Shell nur bei gesetztem data-dash-it-shell UND full-Modus', () => {
        expect(isDashboardItShellActive()).toBe(false);
        document.documentElement.setAttribute('data-dash-it-shell', '');
        expect(isDashboardItShellActive()).toBe(false);
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'full');
        expect(isDashboardItShellActive()).toBe(true);
    });

    it('blockiert IT-Linke-Leiste im Portal-Modus selbst dann, wenn Nav-Baum im sessionStorage liegt', () => {
        persistItWorkspaceNav(
            [{ id: 'dashTaskGruppen', label: 'Mitgliedschaften', tools: [] }],
            'dashTaskGruppen'
        );
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'portal');

        expect(shouldMountItWorkspaceShell()).toBe(false);
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBeNull();
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).toBeNull();
    });

    it('hält die IT-Leiste, wenn die Shell bewusst auf IT steht – auch bei Lehrkraft-Persona', () => {
        persistItWorkspaceNav(
            [{ id: 'dashTaskGruppen', label: 'Mitgliedschaften', tools: [] }],
            'dashTaskGruppen'
        );
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'full');
        sessionStorage.setItem(
            'ms365-dashboard-persona-cache-v2',
            JSON.stringify({
                key: 'teacher-1',
                at: Date.now(),
                ttl: 60000,
                personas: { loggedIn: true, isIt: false, isLehrer: true, unknownUser: false }
            })
        );

        expect(shouldMountItWorkspaceShell()).toBe(true);
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBe('1');
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).toBeNull();
    });

    it('blockiert IT-Workspace für Lehrkraft ohne IT-Shell-Wahl', () => {
        persistItWorkspaceNav(
            [{ id: 'dashTaskGruppen', label: 'Mitgliedschaften', tools: [] }],
            'dashTaskGruppen'
        );
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'portal');
        sessionStorage.setItem(
            'ms365-dashboard-persona-cache-v2',
            JSON.stringify({
                key: 'teacher-1',
                at: Date.now(),
                ttl: 60000,
                personas: { loggedIn: true, isIt: false, isLehrer: true, unknownUser: false }
            })
        );

        expect(shouldMountItWorkspaceShell()).toBe(false);
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBeNull();
    });

    it('hält IT-Workspace bei pending-directory / unknownUser (kein Lehrer-Flash)', () => {
        persistItWorkspaceNav(
            [{ id: 'dashTaskOrdnung', label: 'Aufräumen', tools: [] }],
            'dashTaskOrdnung'
        );
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'full');
        sessionStorage.setItem(
            'ms365-dashboard-persona-cache-v2',
            JSON.stringify({
                key: 'pending-1',
                at: Date.now(),
                ttl: 15000,
                personas: {
                    loggedIn: true,
                    isIt: false,
                    isLehrer: false,
                    unknownUser: true,
                    personaSource: 'pending-directory'
                }
            })
        );

        expect(shouldMountItWorkspaceShell()).toBe(true);
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBe('1');
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).toBeNull();
    });

    it('erlaubt IT-Linke-Leiste für echten IT-Modus (full + aktives Flag + IT-Persona)', () => {
        persistItWorkspaceNav(
            [
                {
                    id: 'dashTaskPersonen',
                    label: 'Personen',
                    tools: [{ href: 'tools/personen-verwaltung.html', label: 'Personen suchen', icon: 'bi bi-search' }]
                }
            ],
            'dashTaskPersonen'
        );
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'full');
        window.__ms365DashboardPersonasLast = { loggedIn: true, isIt: true };

        expect(shouldMountItWorkspaceShell()).toBe(true);
    });

    it('hält IT-Leiste auf Stammdaten auch bei Kopfzeile „Portal“ (IT-Persona)', () => {
        vi.stubGlobal('location', {
            href: 'https://app.ms365.schule/tenant.html',
            pathname: '/tenant.html',
            search: '',
            hash: ''
        });
        window.location = globalThis.location;
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'portal');
        window.__ms365DashboardPersonasLast = { loggedIn: true, isIt: true };
        sessionStorage.removeItem(IT_WORKSPACE_FLAG_KEY);
        sessionStorage.removeItem(IT_WORKSPACE_NAV_KEY);

        expect(shouldMountItWorkspaceShell()).toBe(true);
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).toBeNull();
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBe('1');
    });

    it('markiert auf fremden Werkzeugseiten (z. B. Elternbrief) keinen veralteten Bereich wie Mitgliedschaften', () => {
        const tree = [
            {
                id: 'dashTaskGruppen',
                label: 'Mitgliedschaften',
                tools: [
                    { href: 'tools/verwaltung.html?kind=schulleitung', label: 'Schulleitung', icon: 'bi' },
                    { href: 'tools/verwaltung.html?kind=verwaltung', label: 'Verwaltung (Personal)', icon: 'bi' }
                ]
            },
            {
                id: 'dashTaskSchulApps',
                label: 'Schul-Apps',
                tools: [{ href: 'tools/schulaktivitaeten-planer.html', label: 'Schulaktivitäten', icon: 'bi' }]
            }
        ];

        expect(resolveActiveWorkspaceTaskId(tree, 'tools/elternbrief.html')).toBe('');
        expect(resolveActiveWorkspaceTaskId(tree, 'tools/schulaktivitaeten-planer.html')).toBe('dashTaskSchulApps');
        expect(resolveActiveWorkspaceTaskId(tree, 'tools/verwaltung.html?kind=verwaltung')).toBe('dashTaskGruppen');
        expect(resolveActiveWorkspaceTaskId(tree, 'tools/verwaltung.html')).toBe('dashTaskGruppen');
    });

    it('führt tenant.html immer zur eigenen Stammdaten-Gruppe, nie zu Schuljahresstart', () => {
        const tree = [
            {
                id: IT_WORKSPACE_STAMMDATEN_TASK_ID,
                label: 'Stammdaten',
                tools: [{ href: 'tenant.html', label: 'Stammdaten öffnen', icon: 'bi' }]
            },
            {
                id: 'dashTaskSchuljahr',
                label: 'Schuljahresstart',
                tools: [
                    { href: 'tools/playbook-schuljahresstart.html', label: 'Playbook', icon: 'bi' },
                    { href: 'tenant.html#classes', label: 'Klassen-Anzeigenamen', icon: 'bi' }
                ]
            }
        ];
        expect(resolveActiveWorkspaceTaskId(tree, 'tenant.html')).toBe(IT_WORKSPACE_STAMMDATEN_TASK_ID);
        expect(resolveActiveWorkspaceTaskId(tree, 'tenant.html#klassenlehrer')).toBe(
            IT_WORKSPACE_STAMMDATEN_TASK_ID
        );
        expect(resolveActiveWorkspaceTaskId(tree, 'tenant.html#classes')).toBe(
            IT_WORKSPACE_STAMMDATEN_TASK_ID
        );

        const clean = sanitizeItWorkspaceTree(tree);
        const schuljahr = clean.find((n) => n.id === 'dashTaskSchuljahr');
        expect(schuljahr.tools.map((t) => t.label)).toEqual(['Playbook']);
        const stamm = clean.find((n) => n.id === IT_WORKSPACE_STAMMDATEN_TASK_ID);
        const primary = stamm.tools.filter((t) => !t.expert);
        expect(primary[0].label).toBe('Schulprofil');
        expect(primary.some((t) => t.label === 'Lehrer je Klasse')).toBe(true);
        expect(primary.every((t) => /^tenant\.html#/.test(t.href))).toBe(true);
        expect(stamm.tools.some((t) => t.expert && /playbook|webuntis|quelle|bildung/i.test(t.href))).toBe(
            true
        );
    });

    it('unterscheidet tenant-Tab-Hashes beim Exact-Match', () => {
        expect(workspaceHrefsMatch('tenant.html#faecher', 'tenant.html#klassen')).toBe(false);
        expect(workspaceHrefsMatch('tenant.html#klassenlehrer', 'tenant.html#klassenlehrer')).toBe(true);
        expect(
            workspaceHrefsMatch('tenant.html#faecher', 'tenant.html#klassen', { looseFile: true })
        ).toBe(true);
    });

    it('baut Stammdaten-Untermenü aus Register-Bereichen statt Import-Karten', () => {
        const tools = mergeStammdatenWorkspaceTools([
            { href: 'tenant.html', label: 'Stammdaten', icon: 'bi bi-database' },
            { href: 'tools/playbook-daten-import-verknuepfen.html', label: 'Blaupause', icon: 'bi bi-flag' }
        ]);
        expect(tools.filter((t) => !t.expert).map((t) => t.label)).toEqual([
            'Schulprofil',
            'Verwaltung',
            'Fächer',
            'ARGE',
            'Klassen',
            'Lehrer je Klasse',
            'Lehrer',
            'Schüler',
            'Schülervertretung',
            'SGA',
            'Datenlandkarte'
        ]);
        expect(tools.some((t) => t.expert && t.label === 'Blaupause')).toBe(true);
        expect(tools.some((t) => t.label === 'Stammdaten' && /tenant\.html$/i.test(t.href))).toBe(false);
    });

    it('filtert echte Duplikate (gleiche Datei + gleicher Titel) in extractToolsFromTask heraus, behält aber getrennte Varianten', () => {
        const fakeAnchor = (href, title, expert) => ({
            getAttribute: (k) => (k === 'href' ? href : ''),
            closest: (sel) => (expert && sel === 'details.dash-task-expert' ? {} : null),
            querySelector: (sel) => {
                if (sel === '.dash-task-row__title') return { textContent: title };
                return { className: 'bi bi-people' };
            }
        });
        const task = {
            querySelectorAll: () => [
                fakeAnchor('tools/verwaltung.html?kind=schulleitung', 'Schulleitung', false),
                fakeAnchor('tools/verwaltung.html?kind=verwaltung', 'Verwaltung (Personal)', false),
                fakeAnchor('tools/verwaltung.html?dup=1', 'Schulleitung', false),
                fakeAnchor('tools/weitere-teams-gruppen.html', 'Weitere Teams', true)
            ]
        };

        const extracted = extractToolsFromTask(task);
        expect(extracted.map((t) => t.label)).toEqual([
            'Schulleitung',
            'Verwaltung (Personal)',
            'Weitere Teams'
        ]);
        expect(extracted.map((t) => t.expert)).toEqual([false, false, true]);
    });

    it('speichert keine Menü-Kopie mehr; clearItWorkspaceNav löscht nur den Arbeitsflächen-Zustand', () => {
        persistItWorkspaceNav([{ id: 'x', label: 'X', tools: [] }], 'x');
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).toBeNull();
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBe('1');
        deactivateItWorkspace();
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBeNull();

        clearItWorkspaceNav();
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).toBeNull();
        expect(unmountItWorkspaceShell()).toBe(false);
        expect(mountItWorkspaceShell()).toBe(false);
    });

    it('legt Aufräumen aus dem Katalog an, Cleanup-Playbook inklusive', async () => {
        const { defaultItWorkspaceTree } = await import('../src/shared/it-workspace-nav.js');
        const ordnung = defaultItWorkspaceTree().find((n) => n.id === 'dashTaskOrdnung');
        expect(ordnung).toBeTruthy();
        expect(ordnung.tools[0].href).toMatch(/datenhygiene\.html/i);
        expect(ordnung.tools.some((t) => /cleanup-playbook\.html/i.test(t.href))).toBe(true);
    });

    it('zeigt Unterrichtsteams immer mit denselben Katalog-Einträgen', async () => {
        const { defaultItWorkspaceTree, ensureItWorkspaceNavReady, loadItWorkspaceNav, IT_WORKSPACE_NAV_KEY } =
            await import('../src/shared/it-workspace-nav.js');
        sessionStorage.setItem(
            IT_WORKSPACE_NAV_KEY,
            JSON.stringify({
                version: 1,
                tree: [
                    {
                        id: 'dashTaskImportVerknuepfen',
                        label: 'Stammdaten',
                        icon: 'bi bi-box-arrow-in-down',
                        tools: [{ href: 'tenant.html#profil', label: 'Schulprofil', icon: 'bi bi-window' }]
                    }
                ]
            })
        );
        const loaded = loadItWorkspaceNav();
        const profil = loaded.tree
            .find((n) => n.id === 'dashTaskImportVerknuepfen')
            .tools.find((t) => t.label === 'Schulprofil');
        expect(profil.icon).toBe('bi bi-building');
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).toBeNull();
        persistItWorkspaceNav(
            [
                {
                    id: 'dashTaskUnterricht',
                    label: 'Alt',
                    tools: [{ href: 'tools/kursteams.html', label: 'Nur eines', icon: 'bi bi-easel' }]
                }
            ],
            'dashTaskUnterricht'
        );
        const ready = ensureItWorkspaceNavReady();
        const labels = ready.tree
            .find((n) => n.id === 'dashTaskUnterricht')
            .tools.map((t) => t.label);
        expect(labels).toEqual(defaultItWorkspaceTree().find((n) => n.id === 'dashTaskUnterricht').tools.map((t) => t.label));
        expect(labels).toEqual([
            'Klassengruppen',
            'Unterrichtsteams',
            'Einzelne Unterrichtsteams',
            'Playbook Unterrichtsteams',
            'Unterrichtsteams-Katalog',
            'Klassen zusammenlegen',
            'OneNote-Inhalte verteilen'
        ]);
    });

    it('mappt Sammelgruppen-Stubs auf schueler-lehrer-gruppen?kind=', () => {
        expect(canonicalizeWorkspaceHref('tools/schueler-sammelgruppe.html')).toBe(
            'tools/schueler-lehrer-gruppen.html?kind=schueler'
        );
        expect(canonicalizeWorkspaceHref('tools/lehrer-sammelgruppe.html')).toBe(
            'tools/schueler-lehrer-gruppen.html?kind=lehrer'
        );
        expect(
            workspaceHrefsMatch(
                'tools/schueler-sammelgruppe.html',
                'tools/schueler-lehrer-gruppen.html?kind=schueler'
            )
        ).toBe(true);
        expect(
            workspaceHrefsMatch(
                'tools/lehrer-sammelgruppe.html',
                'tools/schueler-lehrer-gruppen.html?kind=lehrer'
            )
        ).toBe(true);
        expect(
            workspaceHrefsMatch(
                'tools/schueler-lehrer-gruppen.html?kind=schueler',
                'tools/schueler-lehrer-gruppen.html?kind=lehrer'
            )
        ).toBe(false);
    });

    it('soft-navigiert verwaltung?kind= ohne Full-Reload', () => {
        vi.stubGlobal('location', {
            href: 'https://app.ms365.schule/tools/verwaltung.html?kind=schulleitung',
            pathname: '/tools/verwaltung.html',
            search: '?kind=schulleitung',
            hash: ''
        });
        window.location = globalThis.location;
        const pushed = [];
        vi.stubGlobal('history', {
            pushState: (_s, _t, url) => {
                pushed.push(String(url || ''));
            }
        });
        const kinds = [];
        window.ms365VerwaltungNavigateKind = (kind) => {
            kinds.push(kind);
        };
        document.getElementById = () => null;
        document.querySelectorAll = () => [];

        expect(trySoftItWorkspaceNavigate('tools/verwaltung.html?kind=verwaltung')).toBe(true);
        expect(kinds).toEqual(['verwaltung']);
        expect(pushed.some((u) => /kind=verwaltung/.test(u))).toBe(true);
        expect(trySoftItWorkspaceNavigate('tools/klassenchats.html')).toBe(false);
    });

    it('lässt gespeichert offene Bereiche offen, auch wenn ein anderer aktiv ist', async () => {
        const {
            itNavGroupShouldOpen,
            rememberItNavGroupOpen,
            readItNavGroupOpenState
        } = await import('../src/shared/it-nav-open-state.js');
        rememberItNavGroupOpen('dashTaskImportVerknuepfen', true);
        rememberItNavGroupOpen('dashTaskUnterricht', true);
        rememberItNavGroupOpen('dashTaskOrdnung', false);

        const state = readItNavGroupOpenState();
        expect(itNavGroupShouldOpen(state, 'dashTaskImportVerknuepfen', false)).toBe(true);
        expect(itNavGroupShouldOpen(state, 'dashTaskUnterricht', true)).toBe(true);
        expect(itNavGroupShouldOpen(state, 'dashTaskOrdnung', true)).toBe(false);
        expect(itNavGroupShouldOpen({}, 'dashTaskPersonen', true)).toBe(true);
        expect(itNavGroupShouldOpen({}, 'dashTaskPersonen', false)).toBe(false);
        expect(state.dashTaskImportVerknuepfen).toBe(true);
        expect(state.dashTaskUnterricht).toBe(true);
    });

    it('zeigt die IT-Leiste auf der Startseite, wenn die Shell auf IT steht', () => {
        vi.stubGlobal('location', {
            href: 'https://app.ms365.schule/index.html#klassen',
            pathname: '/index.html',
            search: '',
            hash: '#klassen'
        });
        window.location = globalThis.location;
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'full');
        document.documentElement.setAttribute('data-dash-it-shell', '');

        expect(isItDashboardHomePath()).toBe(true);
        expect(shouldMountItWorkspaceShell()).toBe(true);
    });

    it('lässt die IT-Leiste auf der Landingpage aus', () => {
        vi.stubGlobal('location', {
            href: 'https://ms365.schule/landing/index.html',
            pathname: '/landing/index.html',
            search: '',
            hash: ''
        });
        window.location = globalThis.location;
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'full');

        expect(isItDashboardHomePath()).toBe(false);
        expect(shouldMountItWorkspaceShell()).toBe(false);
    });

    it('tauscht Startseite und Werkzeuge im Rahmen, Willkommen und Admin nicht', () => {
        expect(isItContentSwapHref('https://app.ms365.schule/index.html#klassen')).toBe(true);
        expect(isItContentSwapHref('https://app.ms365.schule/tools/kursteams.html')).toBe(true);
        expect(isItContentSwapHref('https://app.ms365.schule/')).toBe(true);
        expect(isItContentSwapHref('https://app.ms365.schule/welcome.html')).toBe(false);
        expect(isItContentSwapHref('https://app.ms365.schule/admin.html')).toBe(false);
        expect(isItContentSwapHref('https://ms365.schule/landing/index.html')).toBe(false);
    });

    it('lädt den iframe nach Soft-Nav auch wenn src noch leer ist (pushState-Falle)', () => {
        vi.stubGlobal('location', {
            href: 'https://app.ms365.schule/tools/jahrgangsgruppen.html',
            pathname: '/tools/jahrgangsgruppen.html',
            search: '',
            hash: '',
            origin: 'https://app.ms365.schule'
        });
        window.location = globalThis.location;

        const abs = new URL('https://app.ms365.schule/tools/jahrgangsgruppen.html');
        const emptyFrame = {
            getAttribute: () => null,
            contentWindow: { location: { href: 'about:blank', pathname: 'blank', search: '', hash: '' } }
        };
        expect(itWorkspaceFrameNeedsSrc(emptyFrame, abs)).toBe(true);

        const loadedFrame = {
            getAttribute: () => 'https://app.ms365.schule/tools/jahrgangsgruppen.html',
            contentWindow: {
                location: {
                    href: 'https://app.ms365.schule/tools/jahrgangsgruppen.html',
                    pathname: '/tools/jahrgangsgruppen.html',
                    search: '',
                    hash: ''
                }
            }
        };
        expect(itWorkspaceFrameNeedsSrc(loadedFrame, abs)).toBe(false);

        const otherAbs = new URL('https://app.ms365.schule/tools/kursteams.html');
        expect(itWorkspaceFrameNeedsSrc(loadedFrame, otherAbs)).toBe(true);
    });

    it('setzt iframe.src beim Soft-Nav von der Leiste, statt nur die URL zu tauschen', () => {
        const loc = {
            href: 'https://app.ms365.schule/tools/personen-verwaltung.html',
            pathname: '/tools/personen-verwaltung.html',
            search: '',
            hash: '',
            origin: 'https://app.ms365.schule'
        };
        vi.stubGlobal('location', loc);
        window.location = loc;

        const frameState = { src: '', attr: null };
        const frame = {
            id: 'itWorkspaceFrame',
            className: '',
            title: '',
            contentWindow: { location: { href: 'about:blank', pathname: 'blank', search: '', hash: '' } },
            contentDocument: null,
            getAttribute: (name) => (name === 'src' ? frameState.attr : null),
            setAttribute: () => {},
            addEventListener: () => {},
            set src(v) {
                frameState.src = String(v || '');
                frameState.attr = frameState.src;
            },
            get src() {
                return frameState.src;
            }
        };
        const mainChildren = [];
        const main = {
            id: 'itWorkspaceMain',
            get firstChild() {
                return mainChildren[0] || null;
            },
            removeChild: (n) => {
                const i = mainChildren.indexOf(n);
                if (i >= 0) mainChildren.splice(i, 1);
                return n;
            },
            appendChild: (n) => {
                mainChildren.push(n);
                return n;
            }
        };
        const nodes = {
            itWorkspaceMain: main,
            itWorkspaceShell: { id: 'itWorkspaceShell' },
            itWorkspaceNav: {
                id: 'itWorkspaceNav',
                querySelectorAll: () => [],
                querySelector: () => null
            },
            itWorkspaceFrame: null
        };
        document.getElementById = (id) => nodes[id] || null;
        document.createElement = (tag) => {
            if (tag === 'iframe') {
                nodes.itWorkspaceFrame = frame;
                return frame;
            }
            return {
                id: '',
                className: '',
                hidden: false,
                setAttribute: () => {},
                appendChild: () => {}
            };
        };
        vi.stubGlobal('history', {
            pushState: (_s, _t, url) => {
                const next = String(url || '');
                loc.href = 'https://app.ms365.schule' + next;
                loc.pathname = next.split('?')[0].split('#')[0] || loc.pathname;
                loc.search = next.includes('?') ? '?' + next.split('?')[1].split('#')[0] : '';
                loc.hash = next.includes('#') ? '#' + next.split('#').pop() : '';
            }
        });

        expect(trySoftItWorkspaceFrameNavigate('/tools/jahrgangsgruppen.html')).toBe(true);
        expect(frameState.src).toContain('/tools/jahrgangsgruppen.html');

        frame.contentWindow.location = {
            href: 'https://app.ms365.schule/tools/jahrgangsgruppen.html',
            pathname: '/tools/jahrgangsgruppen.html',
            search: '',
            hash: ''
        };
        expect(trySoftItWorkspaceFrameNavigate('/tools/kursteams.html')).toBe(true);
        expect(frameState.src).toContain('/tools/kursteams.html');
    });
});
