import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
    afgNeedsDisplayNameCleanup,
    buildAfgBulkCreateConfirmMessage,
    buildAfgBulkMatchExistingConfirmMessage,
    buildAfgBulkRenameConfirmMessage,
    buildAfgMailNickname,
    collectAfgCreateOwners,
    displayNameForAfgCreate,
    formatAfgBulkCreatePreviewLines,
    pickAfgGroupByNickname,
    stripAfgDisplayNamePrefix
} from '../src/tools/arge-fachgruppen/arge-fachgruppen-bulk-logic.js';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

describe('ARGE & Fachgruppen Bulk-Logik', () => {
    it('baut Alias ohne Präfix aus dem Kürzel inkl. Unterstrich', () => {
        expect(buildAfgMailNickname('', 'FSGS_DAF')).toBe('fsgs_daf');
        expect(buildAfgMailNickname('', 'FSGS_E')).toBe('fsgs_e');
        expect(buildAfgMailNickname('arge-', 'FSGS_DAF')).toBe('arge-fsgs_daf');
    });

    it('formatiert Kürzel → Alias-Vorschauzeilen', () => {
        expect(
            formatAfgBulkCreatePreviewLines([
                { code: 'FSGS_DAF', nick: 'arge-fsgsdaf' },
                { code: 'FSGS_E', nick: 'arge-fsgse' }
            ])
        ).toBe('FSGS_DAF → arge-fsgsdaf\nFSGS_E → arge-fsgse');
    });

    it('kürzt lange Vorschauen mit Auslassung', () => {
        const rows = Array.from({ length: 15 }, (_, i) => ({
            code: 'C' + i,
            nick: 'nick' + i
        }));
        const text = formatAfgBulkCreatePreviewLines(rows, 3);
        expect(text.split('\n')).toEqual(['C0 → nick0', 'C1 → nick1', 'C2 → nick2', '…']);
    });

    it('sammelt ARGE-Leitung und Direktion ohne Duplikate', () => {
        expect(
            collectAfgCreateOwners(
                { headEmail: 'leitung@schule.at' },
                ['direktion@schule.at', 'Leitung@schule.at', 'invalid']
            )
        ).toEqual(['leitung@schule.at', 'direktion@schule.at']);
        expect(collectAfgCreateOwners(null, ['a@b.c'])).toEqual(['a@b.c']);
    });

    it('baut die Bestätigungsnachricht mit Alias-Liste', () => {
        const msg = buildAfgBulkCreateConfirmMessage({
            kindLabel: 'ARGEs',
            count: 2,
            previewLines: 'FSGS_DAF → arge-fsgsdaf\nFSGS_E → arge-fsgse',
            createTeam: true
        });
        expect(msg).toContain('2 ungematchte ARGEs');
        expect(msg).toContain('Kürzel → geplanter Alias:');
        expect(msg).toContain('FSGS_DAF → arge-fsgsdaf');
        expect(msg).toContain('Microsoft Team bereitgestellt');
    });

    it('entfernt Anzeigename-Präfixe Arbeitsgruppe/Fach', () => {
        expect(stripAfgDisplayNamePrefix('Arbeitsgruppe FS GS Deutsch/DaF')).toBe('FS GS Deutsch/DaF');
        expect(stripAfgDisplayNamePrefix('Fach Mathematik')).toBe('Mathematik');
        expect(stripAfgDisplayNamePrefix('FS GYM Biologie')).toBe('FS GYM Biologie');
        expect(afgNeedsDisplayNameCleanup('Arbeitsgruppe FS GS Englisch')).toBe(true);
        expect(afgNeedsDisplayNameCleanup('FS GS Englisch')).toBe(false);
        expect(displayNameForAfgCreate({ code: 'FSGS_E', name: 'FS GS Englisch' })).toBe('FS GS Englisch');
    });

    it('wählt Gruppe exakt per Alias oder SMTP-Lokalteil', () => {
        const g = pickAfgGroupByNickname(
            [
                { id: '1', mailNickname: 'other', mail: 'other@dsrom.de' },
                { id: '2', mailNickname: 'fsgs_daf', mail: 'fsgs_daf@dsrom.de', displayName: 'Arbeitsgruppe FS GS Deutsch/DaF' }
            ],
            'fsgs_daf'
        );
        expect(g && g.id).toBe('2');
        expect(pickAfgGroupByNickname([{ id: '1', mail: 'fsgs_e@dsrom.de' }], 'fsgs_e')?.id).toBe('1');
        expect(pickAfgGroupByNickname([{ id: '1', mailNickname: 'fsgs_e' }], 'fsgs_daf')).toBeNull();
    });

    it('baut Bestätigungen für Abgleich und Namensbereinigung', () => {
        const matchMsg = buildAfgBulkMatchExistingConfirmMessage({
            kindLabel: 'ARGEs',
            count: 2,
            previewLines: 'FSGS_DAF → fsgs_daf'
        });
        expect(matchMsg).toContain('bestehenden Microsoft');
        expect(matchMsg).toContain('exakt gleichem Alias');
        const renameMsg = buildAfgBulkRenameConfirmMessage({
            count: 1,
            previewLines: 'Arbeitsgruppe FS GS Englisch  →  FS GS Englisch'
        });
        expect(renameMsg).toContain('Arbeitsgruppe');
        expect(renameMsg).toContain('Anzeigenamen');
    });
});

describe('ARGE & Fachgruppen Seite – Bulk anlegen', () => {
    it('hängt Anlegen & matchen inkl. Alias-Vorschau-Hinweis ein', () => {
        const html = readFileSync(join(projectRoot, 'tools/arge-fachgruppen.html'), 'utf8');
        expect(html).toContain('id="afgBtnBulkCreate"');
        expect(html).toContain('id="afgBtnBulkMatchExisting"');
        expect(html).toContain('id="afgBtnBulkRename"');
        expect(html).toContain('id="afgBtnSelectUnmatched"');
        expect(html).toContain('id="afgBulkCreateTeam"');
        expect(html).toContain('Anlegen &amp; matchen');
        expect(html).toContain('Bestehende matchen');
        expect(html).toContain('Namen bereinigen');
        expect(html).toContain('Kürzel → Alias');
        expect(html).toContain('Leer lassen');
        expect(html).toContain('leer = nur Kürzel');

        const js = readFileSync(join(projectRoot, 'src/tools/arge-fachgruppen/arge-fachgruppen.js'), 'utf8');
        expect(js).toContain('runBulkCreateAndMatch');
        expect(js).toContain('runBulkMatchExisting');
        expect(js).toContain('runBulkStripDisplayPrefix');
        expect(js).toContain('collectSelectedUnmatched');
        expect(js).toContain('formatAfgBulkCreatePreviewLines');
        expect(js).toContain('buildAfgMailNickname');
        expect(js).toContain('pickAfgGroupByNickname');
        expect(js).toContain('stripAfgDisplayNamePrefix');
        expect(js).toContain('afgBtnBulkCreate');
        expect(js).toContain('afgBtnBulkMatchExisting');
        expect(js).toContain('afgBtnBulkRename');
        expect(js).toContain('arge-fachgruppen-bulk-logic.js');
        expect(js).toContain("stored == null ? def");
        expect(js).not.toContain("kind === 'arge' ? 'Arbeitsgruppe '");
    });
});
