import { describe, it, expect } from 'vitest';
import {
    parseNumberedFolderName,
    suggestChannelMappings,
    buildMigrationPlan,
    validateNcMigrateSelection,
    relativeInside,
    topLevelSegment
} from '../src/tools/nextcloud-teams-migrate/nextcloud-teams-migrate-logic.js';
import { parsePropfindXml } from '../src/tools/nextcloud-teams-migrate/nextcloud-teams-migrate-webdav.js';

describe('numbered folder → channel', () => {
    it('parses 01 Fachkonferenz', () => {
        const p = parseNumberedFolderName('01 Fachkonferenz');
        expect(p).toEqual({
            number: '01',
            title: 'Fachkonferenz',
            channelName: '01 Fachkonferenz'
        });
    });

    it('parses 2. Protokolle', () => {
        const p = parseNumberedFolderName('2. Protokolle');
        expect(p?.number).toBe('02');
        expect(p?.channelName).toBe('02 Protokolle');
    });

    it('rejects plain names', () => {
        expect(parseNumberedFolderName('Physik')).toBeNull();
    });

    it('suggests mappings', () => {
        const maps = suggestChannelMappings([
            { name: '01 Fachkonferenz', isFolder: true, relPath: 'Physik/01 Fachkonferenz' },
            { name: 'readme.txt', isFolder: false, relPath: 'Physik/readme.txt' },
            { name: 'Material', isFolder: true, relPath: 'Physik/Material' }
        ]);
        expect(maps).toHaveLength(1);
        expect(maps[0].channelName).toBe('01 Fachkonferenz');
    });
});

describe('migration plan', () => {
    it('routes numbered folders to channels', () => {
        const plan = buildMigrationPlan({
            sourceRootRel: '5. Gymnasium/1. Fächer/Physik',
            files: [
                {
                    relPath: '5. Gymnasium/1. Fächer/Physik/01 Fachkonferenz/notiz.pdf',
                    name: 'notiz.pdf',
                    size: 10
                },
                {
                    relPath: '5. Gymnasium/1. Fächer/Physik/Material/folie.pptx',
                    name: 'folie.pptx',
                    size: 20
                }
            ],
            channelMaps: [
                {
                    sourceFolder: '01 Fachkonferenz',
                    sourceRelPath: '5. Gymnasium/1. Fächer/Physik/01 Fachkonferenz',
                    channelName: '01 Fachkonferenz',
                    enabled: true
                }
            ],
            destRootFolder: 'Import aus Nextcloud',
            mapNumberedToChannels: true
        });
        expect(plan.summary.fileCount).toBe(2);
        expect(plan.summary.channelCount).toBe(1);
        expect(plan.summary.teamRootCount).toBe(1);
        const ch = plan.items.find((i) => i.destKind === 'channel');
        expect(ch?.channelName).toBe('01 Fachkonferenz');
        expect(ch?.destFolderRel).toBe('');
        expect(ch?.fileName).toBe('notiz.pdf');
        const root = plan.items.find((i) => i.destKind === 'teamRoot');
        expect(root?.destFolderRel).toContain('Import aus Nextcloud');
        expect(root?.destFolderRel).toContain('Material');
    });

    it('keeps nested year folders inside one channel', () => {
        const plan = buildMigrationPlan({
            sourceRootRel: '1. Fachkonferenz',
            files: [
                {
                    relPath: '1. Fachkonferenz/21_22/protokoll.pdf',
                    name: 'protokoll.pdf',
                    size: 5
                },
                {
                    relPath: '1. Fachkonferenz/22_23/agenda.docx',
                    name: 'agenda.docx',
                    size: 7
                }
            ],
            channelMaps: [
                {
                    sourceFolder: '1. Fachkonferenz',
                    sourceRelPath: '1. Fachkonferenz',
                    channelName: '01 Fachkonferenz',
                    enabled: true
                },
                {
                    sourceFolder: '21_22',
                    sourceRelPath: '1. Fachkonferenz/21_22',
                    channelName: '21 22',
                    enabled: false
                }
            ],
            destRootFolder: 'Import aus Nextcloud',
            mapNumberedToChannels: true
        });
        expect(plan.summary.channelCount).toBe(2);
        expect(plan.summary.teamRootCount).toBe(0);
        expect(plan.items.every((i) => i.channelName === '01 Fachkonferenz')).toBe(true);
        expect(plan.items.find((i) => i.fileName === 'protokoll.pdf')?.destFolderRel).toBe('21_22');
        expect(plan.items.find((i) => i.fileName === 'agenda.docx')?.destFolderRel).toBe('22_23');
    });

    it('validates selection', () => {
        const bad = validateNcMigrateSelection({});
        expect(bad.ok).toBe(false);
        const helperOk = validateNcMigrateSelection({
            sourceMode: 'helper',
            hasHelper: true,
            destGroupId: 'gid',
            selectedRelPaths: ['Physik']
        });
        expect(helperOk.ok).toBe(true);
        const zipOk = validateNcMigrateSelection({
            sourceMode: 'zip',
            hasLocalFs: true,
            destGroupId: 'gid',
            selectedRelPaths: ['Physik']
        });
        expect(zipOk.ok).toBe(true);
        const webdavOk = validateNcMigrateSelection({
            sourceMode: 'webdav',
            nextcloudUrl: 'https://x',
            nextcloudUser: 'u',
            hasPassword: true,
            destGroupId: 'gid',
            selectedRelPaths: ['Physik']
        });
        expect(webdavOk.ok).toBe(true);
    });

    it('relative helpers', () => {
        expect(relativeInside('a/b', 'a/b/c/d.txt')).toBe('c/d.txt');
        expect(topLevelSegment('01 Fachkonferenz/x')).toBe('01 Fachkonferenz');
    });
});

describe('webdav xml parse', () => {
    it('parses collection + file', () => {
        const xml = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:">
 <d:response>
  <d:href>/remote.php/dav/files/kurt/Physik/</d:href>
  <d:propstat><d:prop><d:resourcetype><d:collection/></d:resourcetype><d:displayname>Physik</d:displayname></d:prop></d:propstat>
 </d:response>
 <d:response>
  <d:href>/remote.php/dav/files/kurt/Physik/a.txt</d:href>
  <d:propstat><d:prop><d:resourcetype/><d:getcontentlength>3</d:getcontentlength><d:displayname>a.txt</d:displayname></d:prop></d:propstat>
 </d:response>
</d:multistatus>`;
        const items = parsePropfindXml(xml, 'kurt');
        expect(items[0].isFolder).toBe(true);
        expect(items[1].relPath).toBe('Physik/a.txt');
    });
});
