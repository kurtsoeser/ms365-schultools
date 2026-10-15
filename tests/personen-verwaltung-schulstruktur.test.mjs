import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { collectPersonSchulstrukturAssignments } from '../src/tools/personen-verwaltung/personen-verwaltung-schulstruktur-logic.js';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

describe('Personen-Verwaltung Schulstruktur-Tab', () => {
    it('Seite hat Tab und Panel Schulstruktur', () => {
        const html = readFileSync(join(projectRoot, 'tools/personen-verwaltung.html'), 'utf8');
        expect(html).toContain('data-pv-tab="schulstruktur"');
        expect(html).toContain('id="pvTabSchulstruktur"');
        expect(html).toContain('id="pvPanelSchulstruktur"');
        expect(html).toContain('id="pvSchulstrukturRoot"');
        expect(html).toContain('app-data-v2.js');
    });

    it('sammelt Funktionen, Bereiche, Leitungen und Unterricht', () => {
        const settings = {
            admin: [{ name: 'Anna Admin', email: 'anna@schule.at', role: 'Direktor:in' }],
            verwaltungAudienceGroups: [
                { id: 'schulleitung', label: 'Schulleitung' },
                { id: 'sekretariat', label: 'Sekretariat' }
            ],
            adminAudienceMemberships: [
                { email: 'anna@schule.at', groupId: 'schulleitung' }
            ],
            schulOrganigramm: {
                version: 2,
                nodes: [
                    { id: 'a1', type: 'area', label: 'Verwaltung', parentId: null },
                    { id: 'r1', type: 'role', label: 'Direktor:in', parentId: 'a1', refKind: 'adminRole', refKey: 'DIR' },
                    {
                        id: 'p1',
                        type: 'person',
                        label: 'Anna Admin',
                        parentId: 'r1',
                        refKind: 'adminPerson',
                        refKey: 'anna@schule.at'
                    }
                ]
            },
            arges: [{ code: 'MAT', name: 'Mathematik', headEmail: 'anna@schule.at' }],
            classes: [{ code: '1A', name: '1A', headEmail: 'other@schule.at' }],
            teachers: [{ code: 'AA', name: 'Anna Admin', email: 'anna@schule.at' }],
            unterrichtsbelegung: {
                rows: [
                    {
                        lehrerEmail: 'anna@schule.at',
                        fach: 'M',
                        klasse: '1A',
                        teamName: 'M · 1A'
                    }
                ]
            }
        };

        const data = collectPersonSchulstrukturAssignments(settings, 'anna@schule.at', {
            name: 'Anna Admin',
            upn: 'anna@schule.at'
        });

        expect(data.empty).toBe(false);
        expect(data.total).toBeGreaterThanOrEqual(5);

        const kinds = data.categories.flatMap((c) => c.items.map((i) => i.kind));
        expect(kinds).toContain('adminRole');
        expect(kinds).toContain('audience');
        expect(kinds).toContain('orgRoleSeat');
        expect(kinds).toContain('argeHead');
        expect(kinds).toContain('teacher');
        expect(kinds).toContain('kursteam');
        expect(kinds).not.toContain('classHead');

        const fn = data.categories.find((c) => c.id === 'functions');
        expect(fn).toBeTruthy();
        expect(fn.items.some((i) => i.kind === 'orgRoleSeat' && /Direktor/i.test(i.label))).toBe(true);

        const areas = data.categories.find((c) => c.id === 'areas');
        expect(areas.items.some((i) => i.label === 'Schulleitung')).toBe(true);
    });

    it('matcht über UPN wenn mail leer', () => {
        const settings = {
            teachers: [{ email: 'max@schule.at', name: 'Max', code: 'MX' }]
        };
        const data = collectPersonSchulstrukturAssignments(settings, '', {
            upn: 'max@schule.at',
            name: 'Max'
        });
        expect(data.total).toBe(1);
        expect(data.categories[0].items[0].kind).toBe('teacher');
    });

    it('liefert leer ohne Treffer', () => {
        const data = collectPersonSchulstrukturAssignments(
            { teachers: [{ email: 'x@y.at' }] },
            'niemand@schule.at'
        );
        expect(data.empty).toBe(true);
        expect(data.total).toBe(0);
        expect(data.categories).toEqual([]);
    });
});
