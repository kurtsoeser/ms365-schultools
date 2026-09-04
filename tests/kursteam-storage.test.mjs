import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

function loadStorageSandbox() {
    const store = new Map();
    const els = new Map();
    const put = (id, props = {}) => {
        els.set(id, {
            id,
            value: '',
            checked: false,
            style: { display: '' },
            textContent: '',
            ...props
        });
    };
    put('yearPrefix', { value: 'SJ26' });
    put('teamSeparator', { value: ' | ' });
    put('excludeSubjects', { value: 'ORD,DIR,KV' });
    put('removeDuplicates', { checked: true });
    put('webuntisPasteInput', { value: '' });
    put('studentRosterPreferGroup', { checked: true });
    put('studentRosterSkipCombinedClasses', { checked: true });
    put('studentRosterHideNoMatch', { checked: true });
    put('totalRecords');
    put('uniqueSubjects');
    put('uniqueTeachers');
    put('importStats');
    put('filteredRecords');
    put('filterStats');
    put('teacherCount');
    put('teacherMappingInfo');

    const sandbox = {
        console,
        localStorage: {
            getItem: (k) => (store.has(k) ? store.get(k) : null),
            setItem: (k, v) => store.set(k, String(v)),
            removeItem: (k) => store.delete(k)
        },
        document: {
            readyState: 'complete',
            getElementById: (id) => els.get(id) || null,
            addEventListener: () => {}
        },
        setInterval: () => 0,
        Blob: class {
            constructor() {}
        },
        URL: { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} }
    };
    sandbox.window = sandbox;
    sandbox.addEventListener = () => {};
    createContext(sandbox);
    // STORAGE_KEY kommt aus utils – hier manuell setzen, nur storage laden
    sandbox.ms365Kursteam = { STORAGE_KEY: 'webuntis-teams-creator-state-v1' };
    runInContext(readFileSync(join(root, 'src/tools/kursteams/kursteam-storage.js'), 'utf8'), sandbox, {
        filename: 'kursteam-storage.js'
    });
    return { ns: sandbox.ms365Kursteam, store };
}

describe('kursteam-storage snapshot', () => {
    it('save → load roundtrip und JSON-Import', () => {
        const { ns, store } = loadStorageSandbox();
        const toasts = [];
        ns.showToast = (m) => toasts.push(m);
        ns.goToStep = () => {};
        ns.renderTeamNameBuilder = () => {};
        ns.refreshSubjectFilterUI = () => {};
        ns.displayFilteredData = () => {};
        ns.displayTeamsData = () => {};
        ns.displayTeacherMappingTable = () => {};
        ns.setContinueButton = () => {};
        ns.updateStep4Checklist = () => {};
        ns.updateStep5Checklist = () => {};
        ns.downloadBlob = () => {};

        ns.rawData = [{ id: 1, klasse: '1A', fach: 'D', lehrer: 'MEI', gruppe: '' }];
        ns.filteredData = [...ns.rawData];
        ns.teamsData = [{ teamName: 'SJ26 | 1A | D', gruppenmail: 'SJ26-1A-D', isValid: true }];
        ns.teamsGenerated = true;
        ns.teacherEmailMapping = { MEI: 'mei@schule.at' };
        ns.currentStep = 5;
        ns.kursteamEntryMode = 'webuntis';
        ns.teamNamePattern = [{ type: 'yearPrefix' }, { type: 'text', value: ' | ' }, { type: 'lehrer' }];

        expect(ns.saveStateToStorage()).toBe(true);
        expect(store.has(ns.STORAGE_KEY)).toBe(true);
        expect(toasts.some((t) => /gespeichert/i.test(t))).toBe(true);

        ns.rawData = [];
        ns.filteredData = [];
        ns.teamsData = [];
        ns.teacherEmailMapping = {};
        ns.teamsGenerated = false;
        ns.currentStep = 0;

        expect(ns.loadStateFromStorage()).toBe(true);
        expect(ns.rawData).toHaveLength(1);
        expect(ns.teamsData).toHaveLength(1);
        expect(ns.teacherEmailMapping.MEI).toBe('mei@schule.at');
        expect(ns.currentStep).toBe(5);
        expect(ns.teamNamePattern.some((p) => p.type === 'lehrer')).toBe(true);

        const exported = ns.buildKursteamStateSnapshot();
        expect(exported.kind).toBe('ms365-kursteams-state');
        ns.rawData = [];
        ns.importKursteamStateJsonText(JSON.stringify(exported));
        expect(ns.rawData).toHaveLength(1);
    });
});
