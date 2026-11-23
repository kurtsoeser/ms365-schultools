import { describe, expect, it } from 'vitest';
import {
    formatQuietSyncLabel,
    formatSyncSecondary,
    klassenChatsProvisionStatus,
    klassenChatsStatusHint,
    klassenChatsCoverage,
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
        expect(klassenChatsStatusHint(container, settings)).toBe('1 von 2 Klassen mit Chat');
    });

    it('klassenChats findet Chats in anderem Year-Bucket (Fallback)', () => {
        const container = {
            years: {
                current: '2026/27',
                byLabel: {
                    '2026/27': { classChats: null },
                    '2025/26': {
                        classChats: {
                            items: [
                                { klasse: '1A', chatId: 'c1' },
                                { klasse: '2B', chatId: 'c2' }
                            ]
                        }
                    }
                }
            }
        };
        const settings = { classes: [{ code: '1A' }, { code: '2B' }] };
        expect(klassenChatsProvisionStatus(container, settings)).toBe('ok');
        expect(klassenChatsCoverage(container, settings).fromFallback).toBe(true);
        expect(klassenChatsStatusHint(container, settings)).toBe('2 von 2 Klassen mit Chat');
    });

    it('klassenChats zeigt angelegte Chats auch bei Kürzel-Mismatch', () => {
        const container = {
            years: {
                current: '2026/27',
                byLabel: {
                    '2026/27': {
                        classChats: {
                            items: [
                                { klasse: '1AK', chatId: 'c1' },
                                { klasse: '1BK', channelId: 'ch1', teamId: 't1', kind: 'privateChannel' }
                            ]
                        }
                    }
                }
            }
        };
        const settings = { classes: [{ code: '1A' }, { code: '1B' }] };
        expect(klassenChatsProvisionStatus(container, settings)).toBe('mismatch');
        expect(klassenChatsStatusHint(container, settings)).toBe('2 Chats angelegt');
    });

    it('klassenChats erkennt Array-Snapshot und Organigramm-Meta', () => {
        const container = {
            years: {
                current: '2026/27',
                byLabel: {
                    '2026/27': {
                        classChats: [{ klasse: '1A', chatId: 'c-arr' }]
                    }
                }
            },
            core: {
                schulOrganigramm: {
                    version: 2,
                    nodes: [
                        {
                            refKind: 'klassenchatClass',
                            refKey: '2B',
                            meta: { chatId: 'c-org' }
                        }
                    ]
                }
            }
        };
        const settings = { classes: [{ code: '1A' }, { code: '2B' }, { code: '3C' }] };
        expect(klassenChatsProvisionStatus(container, settings)).toBe('mismatch');
        expect(klassenChatsCoverage(container, settings).resourceCount).toBe(2);
        expect(klassenChatsStatusHint(container, settings)).toBe('2 von 3 Klassen mit Chat');
    });

    it('klassenChats matcht Klassenkürzel robust (1.A ≈ 1A) und byKlasse-Map', () => {
        const container = {
            years: {
                current: '2026/27',
                byLabel: {
                    '2026/27': {
                        classChats: {
                            byKlasse: {
                                '1.A': { chatId: 'c1' },
                                '2B': { channelId: 'ch2', teamId: 't1', kind: 'privateChannel' }
                            }
                        }
                    }
                }
            }
        };
        const settings = { classes: [{ code: '1A' }, { code: '2B' }] };
        expect(klassenChatsProvisionStatus(container, settings)).toBe('ok');
        expect(klassenChatsStatusHint(container, settings)).toBe('2 von 2 Klassen mit Chat');
    });

    it('klassenChats liest settings.classChats auch ohne Year-Bucket', () => {
        const container = { years: { current: '2026/27', byLabel: { '2026/27': { classChats: null } } } };
        const settings = {
            classes: [{ code: '1A' }, { code: '2B' }],
            classChats: {
                items: [
                    { klasse: '1A', chatId: 'c1' },
                    { klasse: '2B', chatId: 'c2' }
                ]
            }
        };
        expect(klassenChatsProvisionStatus(container, settings)).toBe('ok');
        expect(klassenChatsStatusHint(container, settings)).toBe('2 von 2 Klassen mit Chat');
    });

    it('formatSyncSecondary zeigt heute', () => {
        const now = new Date().toISOString();
        expect(formatSyncSecondary(now)).toBe('Letzter Sync: heute');
    });

    it('formatQuietSyncLabel bleibt kurz', () => {
        const now = new Date().toISOString();
        expect(formatQuietSyncLabel(now)).toBe('heute');
        expect(formatQuietSyncLabel('')).toBe('Verbunden');
    });

    it('SharePoint-Listen-Status ist bei Intranet-URL subtil', () => {
        const info = resolveDashboardToolStatus('sharepoint-liste-lehrer', {
            show: true,
            container: { setup: { intranetSiteUrl: 'https://contoso.sharepoint.com/sites/Listen' } },
            settings: {},
            hygieneById: {},
            hygieneApi: null
        });
        expect(info && info.tone).toBe('ok');
        expect(info && info.variant).toBe('subtle');
        expect(info && info.primary).not.toMatch(/Intranet verbunden/i);
        expect(info && info.secondary).toBe('');
    });

    it('SharePoint-Listen-Status ohne URL betont offen', () => {
        const info = resolveDashboardToolStatus('sharepoint-intranet-hub', {
            show: true,
            container: { setup: {} },
            settings: {},
            hygieneById: {},
            hygieneApi: null
        });
        expect(info && info.tone).toBe('pending');
        expect(info && info.variant).toBe('emphasis');
        expect(info && info.primary).toMatch(/Noch nicht gestartet/i);
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
