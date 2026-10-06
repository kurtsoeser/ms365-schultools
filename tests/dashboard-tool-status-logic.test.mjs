import { describe, expect, it } from 'vitest';
import {
    formatSyncSecondary,
    klassenChatsProvisionStatus,
    resolveDashboardToolStatus,
    subjectCatalogLinkStatus,
    subjectCatalogStatusHint
} from '../src/shared/dashboard-tool-status-logic.js';

describe('dashboard-tool-status-logic', () => {
    it('subjectCatalogLinkStatus erkennt vollständige Verknüpfung', () => {
        const s = { subjects: [{ code: 'D' }], arges: [] };
        const links = [{ kind: 'subject', code: 'd', graphGroupId: 'g1' }];
        expect(subjectCatalogLinkStatus(s, links)).toBe('ok');
    });

    it('subjectCatalogStatusHint zeigt Fortschritt bei Teilverknüpfung', () => {
        const s = {
            subjects: [{ code: 'D' }, { code: 'M' }],
            arges: []
        };
        const links = [{ kind: 'subject', code: 'd', graphGroupId: 'g1' }];
        expect(subjectCatalogStatusHint(s, links)).toBe('1 von 2 verknüpft');
    });

    it('klassenChatsProvisionStatus zählt angelegte Chats', () => {
        const container = {
            years: {
                current: '2026/27',
                byLabel: {
                    '2026/27': {
                        classChats: {
                            items: [{ klasse: '1A', chatId: 'c1' }]
                        }
                    }
                }
            }
        };
        const settings = { classes: [{ code: '1A' }, { code: '2B' }] };
        expect(klassenChatsProvisionStatus(container, settings)).toBe('mismatch');
    });

    it('formatSyncSecondary zeigt heute', () => {
        const now = new Date().toISOString();
        expect(formatSyncSecondary(now)).toBe('Letzter Sync: heute');
    });

    it('resolveDashboardToolStatus für Playbook', () => {
        const info = resolveDashboardToolStatus('playbook-schuljahresstart', {
            show: true,
            container: null,
            settings: {},
            hygieneById: {},
            hygieneApi: null
        });
        expect(info && info.primary).toBeTruthy();
    });

    it('liefert ohne Stammdaten keinen Katalog-Status', () => {
        expect(
            resolveDashboardToolStatus('slg-schueler', {
                show: false,
                container: null,
                settings: {},
                hygieneById: {},
                hygieneApi: null
            })
        ).toBeNull();
    });
});
