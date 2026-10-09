import { describe, it, expect } from 'vitest';
import {
    formatGuidForDebug,
    formatDiagnosticsReport,
    mergePlannerPermsForDiagnostics,
    diagnosticsIsStaffContext
} from '../src/tools/freistellung-planer/freistellung-planer-access-debug.js';
import { listDescriptionBootstrapMeta } from '../src/tools/freistellung-planer/freistellung-planer-remote-config.js';
import { isLikelySharePointTenantRoot } from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('freistellung-planer-access-debug', () => {
    it('formatGuidForDebug kürzt GUIDs', () => {
        const g = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890';
        expect(formatGuidForDebug(g)).toBe('a1b2c3d4…7890');
        expect(formatGuidForDebug('')).toBe('(leer)');
    });

    it('listDescriptionBootstrapMeta liest Site und Listen-ID', () => {
        const m = listDescriptionBootstrapMeta({
            gs: 'guid',
            w: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
            lid: '72d2028f-2ee6-4ce9-a610-0c2ef70196fe'
        });
        expect(m.siteUrl).toContain('MS365-Schultools');
        expect(m.listId).toContain('72d2028f');
    });

    it('isLikelySharePointTenantRoot erkennt Stammweb', () => {
        expect(isLikelySharePointTenantRoot('https://kurtrocks.sharepoint.com')).toBe(true);
        expect(isLikelySharePointTenantRoot('https://kurtrocks.sharepoint.com/sites/X')).toBe(false);
    });

    it('mergePlannerPermsForDiagnostics übernimmt Schüler-ID von der Liste', () => {
        const m = mergePlannerPermsForDiagnostics(
            { groupKvId: 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa', groupSchuelerId: '' },
            { groupSchuelerId: 'f5b9986a-1111-2222-3333-444444444cbf' },
            null
        );
        expect(m.groupSchuelerId).toBe('f5b9986a-1111-2222-3333-444444444cbf');
    });

    it('diagnosticsIsStaffContext erkennt KV', () => {
        expect(diagnosticsIsStaffContext({ role: 'kv' }, [])).toBe(true);
        expect(diagnosticsIsStaffContext({ role: 'schueler' }, ['kv'])).toBe(true);
        expect(diagnosticsIsStaffContext({ role: 'schueler' }, [])).toBe(false);
    });

    it('formatDiagnosticsReport baut Textbericht', () => {
        const text = formatDiagnosticsReport({
            at: '2026-01-01T00:00:00.000Z',
            accountEmail: 'schueler@demo.schule',
            planerAccessDenied: true,
            role: '',
            summary: 'Test',
            steps: [{ id: 'x', title: 'Login', status: 'ok', detail: 'ok' }]
        });
        expect(text).toContain('schueler@demo.schule');
        expect(text).toContain('[ok] Login');
    });
});
