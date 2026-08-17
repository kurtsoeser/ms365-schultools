import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
    PLAYBOOK_REQUIRED_IDS,
    applyInferredPlaybookDone,
    applyYearActivated,
    buildCohortRows,
    cohortDisplayName,
    cohortMailNickname,
    cohortPhase,
    currentSchoolYearLabel,
    inferPlaybookFromContainer,
    isSchoolYearLabel,
    markPlaybookStep,
    nextSchoolYearLabel,
    normalizePlaybook,
    openItemsFromContainer,
    playbookProgress,
    playbookStepDefs,
    schoolYearEndYear,
    summarizeRunPreview,
    yearAlreadyExists,
    appendRunLogEntry,
    buildSchoolYearRunPreview
} from '../src/tools/organisations-assistent/organisations-assistent-logic.js';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

describe('Schuljahr-Playbook Logik', () => {
    it('nextSchoolYearLabel erhöht 2025/26 auf 2026/27', () => {
        expect(nextSchoolYearLabel('2025/26')).toBe('2026/27');
        expect(nextSchoolYearLabel('2025/2026')).toBe('2026/27');
        expect(isSchoolYearLabel('2026/27')).toBe(true);
        expect(isSchoolYearLabel('foo')).toBe(false);
    });

    it('currentSchoolYearLabel folgt dem übergebenen Datum', () => {
        expect(currentSchoolYearLabel(new Date('2026-03-01T12:00:00Z'))).toBe('2026/27');
    });

    it('normalizePlaybook setzt den Jahr-Schritt, wenn Ziel = aktuelles Jahr', () => {
        const n = normalizePlaybook({ targetYear: '2026/27', done: { names: true } }, '2026/27');
        expect(n.done.year).toBe(true);
        expect(n.done.names).toBe(true);
        expect(n.done.kursteams).toBe(false);
        expect(n.targetYear).toBe('2026/27');
    });

    it('applyYearActivated setzt year und setzt andere Schritte bei neuem Ziel zurück', () => {
        const prev = normalizePlaybook({
            targetYear: '2025/26',
            done: { year: true, names: true, students: true }
        });
        const next = applyYearActivated(prev, '2026/27');
        expect(next.targetYear).toBe('2026/27');
        expect(next.done.year).toBe(true);
        expect(next.done.names).toBe(false);
        expect(next.done.students).toBe(false);
    });

    it('applyYearActivated behält Häkchen beim erneuten Aktivieren desselben Jahres', () => {
        const prev = {
            targetYear: '2026/27',
            done: { year: true, names: true, students: false }
        };
        const next = applyYearActivated(prev, '2026/27');
        expect(next.done.names).toBe(true);
        expect(next.done.year).toBe(true);
    });

    it('markPlaybookStep ändert nur den genannten Schritt', () => {
        const n = markPlaybookStep({ targetYear: '2026/27' }, 'kursteams', true);
        expect(n.done.kursteams).toBe(true);
        expect(n.done.names).toBe(false);
    });

    it('playbookProgress zählt nur Pflichtschritte', () => {
        const pb = {
            targetYear: '2026/27',
            done: { year: true, names: true, expert: true }
        };
        const p = playbookProgress(pb, PLAYBOOK_REQUIRED_IDS);
        expect(PLAYBOOK_REQUIRED_IDS).not.toContain('parents');
        expect(p.total).toBe(6);
        expect(p.done).toBe(2);
        expect(p.pct).toBe(33);
    });

    it('yearAlreadyExists vergleicht trim-genau', () => {
        expect(yearAlreadyExists(['2025/26', '2026/27'], '2026/27')).toBe(true);
        expect(yearAlreadyExists(['2025/26'], '2026/27')).toBe(false);
    });

    it('Schritt-Definitionen verlinken Alltagswerkzeuge, nicht den Fachschafts-Tab', () => {
        const defs = playbookStepDefs();
        const ids = defs.map((d) => d.id);
        expect(ids).toEqual(['year', 'names', 'graduates', 'students', 'kursteams', 'subjects', 'expert']);
        expect(defs.find((d) => d.id === 'year').href).toBe('#year');
        expect(defs.find((d) => d.id === 'subjects').href).toBe('arge-fachgruppen.html');
        expect(defs.find((d) => d.id === 'names').href).toBe('#namen');
        expect(defs.find((d) => d.id === 'graduates').href).toBe('#abschluss');
        expect(defs.find((d) => d.id === 'parents')).toBeUndefined();
        expect(defs.find((d) => d.id === 'expert').href).toBe('schulstruktur-sync.html?mode=struktur');
        expect(defs.find((d) => d.id === 'expert').hrefLabel).toBe('Gruppenverwaltung');
        expect(defs.find((d) => d.id === 'expert').blurb).not.toMatch(/Unterbäume kopieren/);
        expect(defs.find((d) => d.id === 'kursteams').blurb).toMatch(/CSV\/CMD/);
        expect(defs.find((d) => d.id === 'students').blurb).toMatch(/automatisch abgehakt/);
    });

    it('openItemsFromContainer listet unmatched Klassen und Archiv-Hinweis', () => {
        const items = openItemsFromContainer(
            {
                years: {
                    current: '2026/27',
                    byLabel: {
                        '2026/27': {
                            classes: [{ code: '1A' }, { code: '1B' }],
                            students: []
                        }
                    }
                },
                core: { classTeams: [{ classCode: '1A', graphGroupId: 'g1' }], subjects: [], arges: [] },
                setup: { matched: {}, catalogLinks: [] }
            },
            { targetYear: '2026/27', done: { year: true, names: true, graduates: true, students: true, kursteams: true, subjects: true } }
        );
        expect(items.some((x) => x.id === 'unmatched-classes')).toBe(true);
        expect(items.some((x) => x.id === 'archive-kursteams')).toBe(true);
        expect(items.find((x) => x.id === 'unmatched-classes').text).toMatch(/1 Klasse/);
    });

    it('buildSchoolYearRunPreview plant Namen +1 ohne Graph', () => {
        const preview = buildSchoolYearRunPreview(
            {
                years: {
                    current: '2025/26',
                    byLabel: { '2025/26': { classes: [{ code: '1A', year: '2030' }], students: [] } }
                },
                core: {
                    classTeams: [{ classCode: '1A', displayName: '1HMA', graphGroupId: 'g1' }],
                    subjects: [],
                    arges: []
                },
                setup: { matched: {}, catalogLinks: [] }
            },
            { targetYear: '2025/26', done: { year: true } }
        );
        expect(preview.nextYear).toBe('2026/27');
        const names = preview.actions.find((a) => a.id === 'names');
        expect(names.detail).toMatch(/1HMA → 2HMA/);
        expect(names.status).toBe('ready');
        const grads = preview.actions.find((a) => a.id === 'graduates');
        expect(grads.title).toMatch(/Abschlussjahrg/);
        const sum = summarizeRunPreview(preview);
        expect(sum.total).toBeGreaterThan(5);
        const log = appendRunLogEntry([], { mode: 'preview', summary: 'test', at: '2026-08-17T12:00:00Z' });
        expect(log).toHaveLength(1);
        expect(log[0].summary).toBe('test');
    });

    it('inferPlaybookFromContainer erkennt Sammelgruppe und voll verknüpfte Fächer', () => {
        const empty = inferPlaybookFromContainer(null);
        expect(empty.students).toBe(false);
        expect(empty.subjects).toBe(false);

        const partial = inferPlaybookFromContainer({
            years: { current: '2026/27', byLabel: { '2026/27': { students: [{ upn: 'a@x' }] } } },
            setup: {
                matched: { schuelerGroupId: 'gid-1' },
                catalogLinks: []
            },
            core: {
                subjects: [{ code: 'D', name: 'Deutsch' }],
                arges: []
            }
        });
        expect(partial.students).toBe(true);
        expect(partial.subjects).toBe(false);
        expect(partial.hints.students).toMatch(/Sammelgruppe/);

        const full = applyInferredPlaybookDone(
            { targetYear: '2026/27', done: {} },
            inferPlaybookFromContainer({
                years: { current: '2026/27', byLabel: { '2026/27': { students: [] } } },
                setup: {
                    matched: { schuelerGroupId: 'gid-1' },
                    catalogLinks: [
                        { kind: 'subject', code: 'D', graphGroupId: 'g1' },
                        { kind: 'arge', code: 'CHOR', graphGroupId: 'g2' }
                    ]
                },
                core: {
                    subjects: [{ code: 'D', name: 'Deutsch' }],
                    arges: [{ code: 'CHOR', name: 'Chor' }]
                }
            }),
            '2026/27'
        );
        expect(full.done.students).toBe(true);
        expect(full.done.subjects).toBe(true);
    });
});

describe('Schuljahr-Playbook Seite', () => {
    it('ist eine Checkliste ohne Fachschafts-Tab', () => {
        const html = readFileSync(join(projectRoot, 'tools/organisations-assistent.html'), 'utf8');
        expect(html).toContain('Checkliste');
        expect(html).toContain('id="oaPlaybookList"');
        expect(html).toContain('data-oa-step="year"');
        expect(html).toContain('data-oa-step="names"');
        expect(html).toContain('data-oa-step="graduates"');
        expect(html).toContain('id="oaActivateYear"');
        expect(html).toContain('id="oaCopyLists"');
        expect(html).toContain('type="module"');
        expect(html).toContain('organisations-assistent.js');
        expect(html).not.toContain('oaTabfach');
        expect(html).not.toContain('Fachschafts-Gruppen');
        expect(html).not.toContain('oaFachTbody');
        expect(html).toContain('oaExpertDetails');
        expect(html).toContain('id="oaExpertMeta"');
        expect(html).not.toContain('oaDupBtn');
        expect(html).not.toContain('oaBulkYear');
        expect(html).not.toContain('oaArchiveBtn');
        expect(html).not.toContain('oaStructTbody');
        expect(html).toContain('id="oaOpenList"');
        expect(html).toContain('id="oaRunBox"');
        expect(html).toContain('id="oaRunPreview"');
        expect(html).toContain('id="oaRunSaveLog"');
        expect(html).toContain('Abschlussjahrgang');
        expect(html).toContain('id="classTeamsRolloverRoot"');
        expect(html).toContain('id="namen"');
        expect(html).toContain('id="abschluss"');
        expect(html).not.toContain('id="oaKindEltern"');
        expect(html).not.toContain('data-oa-kind="eltern"');
        expect(html).not.toContain('Eltern-Jahrgang');
        expect(html).toContain('id="oaCohortListItems"');
        expect(html).toContain('id="oaCohortDetailHost"');
        expect(html).toContain('group-detail.css');
        expect(html).toContain('group-detail.js');
        expect(html).toContain('slg-live-details.js');
        expect(html).not.toContain('oaCohortAdd');
        expect(html).not.toContain('oaGradDetails');
        expect(html).toContain('class-teams-rollover-ui.js');
        expect(html).toContain('graph-rename-preview.js');
        expect(html).toContain('msal-auth-ui.js');
        expect(html).toContain('graph-unified-groups.js');
        expect(html).toContain('ms365-graph-client-id');
        const js = readFileSync(
            join(projectRoot, 'src/tools/organisations-assistent/organisations-assistent.js'),
            'utf8'
        );
        expect(html).toContain('id="oaStepStatus-students"');
        expect(html).toContain('id="oaStepStatus-kursteams"');
        expect(html).toContain('id="oaStepStatus-subjects"');
        expect(html).toContain('automatisch abgehakt');
        expect(js).toContain('inferPlaybookFromContainer');
        expect(js).toContain('applyInferredPlaybookDone');
        expect(js).toContain('buildSchoolYearRunPreview');
        expect(js).toContain('renderRunPreview');
        expect(js).toContain('oaStepStatus-');
        expect(js).not.toContain("kind: 'subject'");
        expect(js).not.toContain('oaFachAddLinks');
        expect(js).not.toContain('oaCohortAdd');
        expect(js).toContain('openPlaybookStep');
        expect(js).toContain('initCohortPanel');
        expect(js).not.toContain('host.replaceChildren');
        expect(js).toContain('ms365-cohort-linked');
        expect(js).not.toContain('ms365-eltern-linked');
        expect(js).not.toContain('setYearGroupKind');
        expect(js).not.toContain('duplicateSubtrees');
        expect(js).not.toContain('oaDupBtn');
        const roll = readFileSync(join(projectRoot, 'src/shared/class-teams-rollover-ui.js'), 'utf8');
        expect(roll).toContain('ms365ClassTeamsRolloverRefresh');
        expect(roll).toContain('ms365-class-teams-rollover-saved');
        expect(js).toContain('ms365-class-teams-rollover-saved');
    });
});

describe('Abschlussjahrgang aus Klassendaten', () => {
    it('cohortMailNickname und Anzeigename folgen dem Schema maturajgYYYY', () => {
        expect(cohortMailNickname('2026')).toBe('maturajg2026');
        expect(cohortDisplayName('2026')).toBe('Abschlussjahrgang 2026');
        expect(cohortMailNickname('26')).toBe('');
        expect(schoolYearEndYear('2026/27')).toBe(2027);
    });

    it('cohortPhase ist relativ zum Schuljahr-Start', () => {
        expect(cohortPhase('2025', '2026/27')).toBe('vergangen');
        expect(cohortPhase('2026', '2026/27')).toBe('gerade-abgeschlossen');
        expect(cohortPhase('2027', '2026/27')).toBe('aktuell');
        expect(cohortPhase('2028', '2026/27')).toBe('kommend');
        expect(cohortPhase('', '2026/27')).toBe('offen');
    });

    it('buildCohortRows fasst Klassen je Abschlussjahr zusammen und merkt catalogLinks', () => {
        const rows = buildCohortRows(
            [
                { code: '5HMA', name: '5HMA', year: '2026' },
                { code: '5HMB', name: '5HMB', year: '2026' },
                { code: '4HMA', name: '4HMA', year: '2027' },
                { code: 'X', name: 'ohne Jahr', year: '' }
            ],
            [
                {
                    kind: 'cohort',
                    code: '2026',
                    graphGroupId: 'g-26',
                    mailNickname: 'maturajg2026',
                    mode: 'matched'
                }
            ],
            '2026/27',
            [{ kind: 'matura', graduationYear: '2029', displayName: 'Alt 2029' }]
        );
        expect(rows.map((r) => r.year)).toEqual(['2026', '2027', '2029']);
        expect(rows[0].classCount).toBe(2);
        expect(rows[0].graphGroupId).toBe('g-26');
        expect(rows[0].phase).toBe('gerade-abgeschlossen');
        expect(rows[1].phase).toBe('aktuell');
        expect(rows[1].mailNickname).toBe('maturajg2027');
        expect(rows[2].displayName).toBe('Alt 2029');
        expect(rows[2].classCount).toBe(0);
    });

    it('buildCohortRows ignoriert Eltern-Merkungen', () => {
        const plans = [
            { kind: 'eltern', graduationYear: '2026' },
            { kind: 'matura', graduationYear: '2026' },
            { kind: 'eltern', graduationYear: 'abc' }
        ];
        const rows = buildCohortRows([], [], '2026/27', plans);
        expect(rows.map((r) => r.year)).toEqual(['2026']);
        expect(rows[0].displayName).toBe('Abschlussjahrgang 2026');
        expect(buildCohortRows([], [], '2026/27', [{ kind: 'eltern', graduationYear: '2026' }])).toEqual([]);
    });
});
