import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
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
    unmountItWorkspaceShell
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
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).not.toBeNull();
    });

    it('blockiert und löscht IT-Workspace für eingeloggte Nicht-IT-Konten (Lehrkraft)', () => {
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
                personas: { loggedIn: true, isIt: false, isTeacher: true }
            })
        );

        expect(shouldMountItWorkspaceShell()).toBe(false);
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBeNull();
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
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).not.toBeNull();
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

    it('räumt über clearItWorkspaceNav und deactivateItWorkspace sauber auf', () => {
        persistItWorkspaceNav([{ id: 'x', label: 'X', tools: [] }], 'x');
        deactivateItWorkspace();
        expect(sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY)).toBeNull();
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).not.toBeNull();

        clearItWorkspaceNav();
        expect(sessionStorage.getItem(IT_WORKSPACE_NAV_KEY)).toBeNull();
        expect(unmountItWorkspaceShell()).toBe(false);
        expect(mountItWorkspaceShell()).toBe(false);
    });
});
