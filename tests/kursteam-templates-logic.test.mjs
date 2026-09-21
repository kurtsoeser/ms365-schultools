import { describe, it, expect } from 'vitest';
import {
    normalizeTemplate,
    createEmptyTemplate,
    cloneTemplate,
    filterTemplatesBySubject,
    renameTemplateChannel,
    addTemplateChannel,
    removeTemplateChannel,
    moveTemplateChannel,
    isGeneralChannelName,
    channelNumberPrefix,
    diffChannels,
    summarizeDiff,
    buildExportPayload,
    parseImportPayload,
    mergeTemplates,
    subjectOptionsFromCore,
    normalizeModule,
    normalizeSchulstufe,
    schulstufeLabel,
    moduleLabel,
    klasseToSchulstufe,
    hakLehrplanModulToMeta,
    normalizeSemester,
    semesterLabel,
    buildTemplateTree,
    filterTemplates,
    normalizeSchoolForm,
    rememberSchoolForm,
    renameSchoolFormInTemplates,
    collectSchoolForms,
    EXPORT_VERSION
} from '../src/tools/kursteam-templates/kursteam-templates-logic.js';
import { getSeedTemplates } from '../src/tools/kursteam-templates/kursteam-templates-seed.js';

describe('kursteam-templates-logic', () => {
    it('erkennt Allgemein/General und Nummer-Präfixe', () => {
        expect(isGeneralChannelName('Allgemein')).toBe(true);
        expect(isGeneralChannelName(' general ')).toBe(true);
        expect(isGeneralChannelName('00-Allgemein')).toBe(true);
        expect(isGeneralChannelName('00 - Allgemein')).toBe(true);
        expect(isGeneralChannelName('01 - Grundlagen')).toBe(false);
        expect(channelNumberPrefix('01 - Grundlagen')).toBe('01');
        expect(channelNumberPrefix('7. Terme')).toBe('7');
        expect(channelNumberPrefix('ohne Nummer')).toBe('');
    });

    it('normalisiert Vorlagen auf Schulstufe + Semester (HAKB-Modul = Halbjahr)', () => {
        const t = normalizeTemplate({
            name: 'MAM Modul 3',
            schoolForm: 'HAKB',
            subjectCode: 'mam',
            module: '3',
            channels: ['Allgemein', '01 - Grundlagen', { displayName: '02 - Prozent' }]
        });
        expect(t.schoolForm).toBe('HAKB');
        expect(t.subjectCode).toBe('MAM');
        expect(t.schulstufe).toBe('10');
        expect(t.semester).toBe('WS');
        expect(t.channels.map((c) => c.displayName)).toEqual(['01 - Grundlagen', '02 - Prozent']);

        const m4 = normalizeTemplate({ schoolForm: 'HAKB', module: 4, name: 'x', channels: [] });
        expect(m4.schulstufe).toBe('10');
        expect(m4.semester).toBe('SS');

        const m5 = normalizeTemplate({ schoolForm: 'hakb', module: 5, name: 'x', channels: [] });
        expect(m5.schoolForm).toBe('HAKB');
        expect(m5.schulstufe).toBe('11');
        expect(m5.semester).toBe('WS');

        const ahs = normalizeTemplate({
            name: 'AHS M 5. Klasse',
            schoolForm: 'AHS',
            klasse: 5,
            subjectCode: 'M',
            semester: 'schuljahr',
            channels: ['01 - Funktionen']
        });
        expect(ahs.schulstufe).toBe('9');
        expect(ahs.semester).toBe('SJ');

        const explicit = normalizeTemplate({
            name: 'Explizit',
            schoolForm: 'HAKB',
            schulstufe: '11',
            semester: 'SS',
            module: '99',
            channels: []
        });
        expect(explicit.schulstufe).toBe('11');
        expect(explicit.semester).toBe('SS');

        expect(hakLehrplanModulToMeta(3)).toEqual({ schulstufe: '10', semester: 'WS' });
        expect(hakLehrplanModulToMeta(8)).toEqual({ schulstufe: '12', semester: 'SS' });
        expect(normalizeSchulstufe('Modul 5')).toBe('5');
        expect(normalizeModule('Modul 5')).toBe('5');
        expect(schulstufeLabel('5')).toBe('Schulstufe 5');
        expect(moduleLabel('5')).toBe('Schulstufe 5');
        expect(klasseToSchulstufe('HAKB', 1)).toBe('9');
        expect(klasseToSchulstufe('AHS', 1)).toBe('5');
        expect(normalizeSemester('ws+ss')).toBe('SJ');
        expect(semesterLabel('WS')).toBe('Wintersemester');
        expect(normalizeSchoolForm('hakb')).toBe('HAKB');
        expect(normalizeSchoolForm('nms')).toBe('Mittelschule');
    });

    it('gruppiert nach Schulform, Fach, Schulstufe oder Semester und filtert kombiniert', () => {
        const seeds = getSeedTemplates();
        expect(seeds[0].schoolForm).toBe('HAKB');
        expect(seeds[0].semester).toBe('WS');

        const bySchool = buildTemplateTree(seeds, 'schoolForm', {});
        expect(bySchool).toHaveLength(1);
        expect(bySchool[0].label).toBe('HAKB');

        const bySubject = buildTemplateTree(seeds, 'subject', { schoolForm: 'HAKB' });
        expect(bySubject[0].label).toBe('MAM');
        expect(bySubject[0].children[0].badge).toContain('HAKB');

        const byStufe = buildTemplateTree(seeds, 'schulstufe', { schoolForm: 'HAKB', subjectCode: 'MAM' });
        expect(byStufe.map((g) => g.label)).toEqual([
            'Schulstufe 10',
            'Schulstufe 11',
            'Schulstufe 12'
        ]);
        expect(byStufe[0].count).toBe(2); // WS + SS

        const bySem = buildTemplateTree(seeds, 'semester', {});
        expect(bySem.map((g) => g.label)).toEqual(['Wintersemester', 'Sommersemester']);

        expect(filterTemplates(seeds, { schoolForm: 'AHS' })).toHaveLength(0);
        expect(filterTemplates(seeds, { schoolForm: 'HAKB', schulstufe: '10' })).toHaveLength(2);
        expect(filterTemplates(seeds, { schoolForm: 'HAKB', schulstufe: '10', semester: 'WS' })).toHaveLength(1);
    });

    it('merkt und benennt Schulformen um', () => {
        expect(rememberSchoolForm(['HAKB'], 'ahs')).toEqual(['AHS', 'HAKB']);
        const list = [
            normalizeTemplate({ name: 'A', schoolForm: 'HAKB', subjectCode: 'MAM', schulstufe: '11' }),
            normalizeTemplate({ name: 'B', schoolForm: 'HAKB', subjectCode: 'MAM', schulstufe: '12' })
        ];
        const renamed = renameSchoolFormInTemplates(list, 'HAKB', 'HLW');
        expect(renamed.every((t) => t.schoolForm === 'HLW')).toBe(true);
        expect(collectSchoolForms(renamed, ['HLW'])).toContain('HLW');
    });

    it('CRUD-Helfer: add/rename/remove/move/clone/filter', () => {
        let t = createEmptyTemplate('Mathe A', 'MAM');
        t = addTemplateChannel(t, '01 - A');
        t = addTemplateChannel(t, '02 - B');
        expect(() => addTemplateChannel(t, 'Allgemein')).toThrow(/Allgemein/);
        t = renameTemplateChannel(t, t.channels[0].id, '01 - Grundlagen');
        expect(t.channels[0].displayName).toBe('01 - Grundlagen');
        t = moveTemplateChannel(t, t.channels[1].id, 'up');
        expect(t.channels[0].displayName).toBe('02 - B');
        const id0 = t.channels[0].id;
        t = removeTemplateChannel(t, id0);
        expect(t.channels.length).toBe(1);

        const copy = cloneTemplate(t);
        expect(copy.id).not.toBe(t.id);
        expect(copy.name).toContain('Kopie');
        expect(copy.channels[0].id).not.toBe(t.channels[0].id);

        const list = [t, createEmptyTemplate('Deutsch', 'D')];
        expect(filterTemplatesBySubject(list, 'MAM')).toHaveLength(1);
        expect(filterTemplatesBySubject(list, '')).toHaveLength(2);
    });

    it('diffChannels: create / ok / rename / skip_general / extra', () => {
        const tpl = normalizeTemplate({
            name: 'MAM',
            channels: ['01 - Neu', '02 - Alt Name']
        });
        const team = [
            { id: 'g', displayName: 'Allgemein', membershipType: 'standard' },
            { id: 'c2', displayName: '02 - Alter Name', membershipType: 'standard' },
            { id: 'x', displayName: 'Extra Kanal', membershipType: 'standard' }
        ];
        const rows = diffChannels(tpl, team);
        const byStatus = (s) => rows.filter((r) => r.status === s);
        expect(byStatus('skip_general')).toHaveLength(1);
        expect(byStatus('create')[0].templateChannel.displayName).toBe('01 - Neu');
        expect(byStatus('rename')[0].teamChannel.displayName).toBe('02 - Alter Name');
        expect(byStatus('extra')[0].teamChannel.displayName).toBe('Extra Kanal');

        const okTeam = [
            { id: 'g', displayName: 'General' },
            { id: '1', displayName: '01 - Neu' },
            { id: '2', displayName: '02 - Alt Name' }
        ];
        const okRows = diffChannels(tpl, okTeam);
        expect(summarizeDiff(okRows).ok).toBe(2);
        expect(summarizeDiff(okRows).create).toBe(0);
    });

    it('Export/Import/Merge und Fächeroptionen', () => {
        const a = createEmptyTemplate('A', 'MAM', '', '11', 'HAKB', 'SJ');
        const payload = buildExportPayload([a]);
        expect(payload.kind).toBe('ms365-kursteam-templates');
        expect(payload.version).toBe(EXPORT_VERSION);
        expect(payload.version).toBe(3);
        expect(payload.templates).toHaveLength(1);
        expect(payload.templates[0].schulstufe).toBe('11');
        expect(payload.templates[0].semester).toBe('SJ');
        expect(payload.templates[0].module).toBeUndefined();

        const parsed = parseImportPayload(JSON.stringify(payload));
        expect(parsed.templates[0].name).toBe('A');

        const b = { ...a, name: 'A2' };
        const merged = mergeTemplates([a], [b]);
        expect(merged).toHaveLength(1);
        expect(merged[0].name).toBe('A2');

        expect(subjectOptionsFromCore([{ code: 'd', name: 'Deutsch' }, { code: 'D', name: 'Dup' }])).toEqual([
            { code: 'D', label: 'D – Deutsch' }
        ]);
    });

    it('liefert 6 MAM-Seed-Vorlagen (Modul 3–8 → Stufe 10–12 WS/SS)', () => {
        const seeds = getSeedTemplates();
        expect(seeds).toHaveLength(6);
        expect(seeds.every((t) => t.subjectCode === 'MAM')).toBe(true);
        expect(seeds.every((t) => t.schoolForm === 'HAKB')).toBe(true);
        expect(seeds.map((t) => [t.schulstufe, t.semester])).toEqual([
            ['10', 'WS'],
            ['10', 'SS'],
            ['11', 'WS'],
            ['11', 'SS'],
            ['12', 'WS'],
            ['12', 'SS']
        ]);
        expect(seeds[0].name).toContain('Modul 3');
        expect(seeds[0].channels).toHaveLength(10);
        expect(seeds[5].channels.some((c) => c.displayName === 'sRDP - Vorbereitung')).toBe(true);
        expect(seeds.every((t) => t.channels.every((c) => !isGeneralChannelName(c.displayName)))).toBe(true);
    });
});
