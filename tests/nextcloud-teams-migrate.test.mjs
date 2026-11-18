import { describe, it, expect } from 'vitest';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
    parsePropfindXml,
    formatBytes,
    davFilesRoot,
    davUrl
} from '../scripts/nextcloud-teams-migrate/lib/webdav.mjs';
import { loadMapping } from '../scripts/nextcloud-teams-migrate/lib/config.mjs';
import { stateKey, loadState, markDone } from '../scripts/nextcloud-teams-migrate/lib/state.mjs';
import fs from 'node:fs';
import os from 'node:os';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

describe('nextcloud-teams-migrate webdav', () => {
    it('builds dav urls', () => {
        const root = davFilesRoot('https://www.dsrom-intranet.de', 'kurt');
        expect(root).toBe('https://www.dsrom-intranet.de/remote.php/dav/files/kurt');
        expect(davUrl(root, '5. Gymnasium (5-12)/Physik')).toContain('5.%20Gymnasium%20(5-12)');
    });

    it('parses propfind xml', () => {
        const xml = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:">
 <d:response>
  <d:href>/remote.php/dav/files/kurt/Physik/</d:href>
  <d:propstat><d:prop><d:resourcetype><d:collection/></d:resourcetype><d:displayname>Physik</d:displayname></d:prop></d:propstat>
 </d:response>
 <d:response>
  <d:href>/remote.php/dav/files/kurt/Physik/notiz.txt</d:href>
  <d:propstat><d:prop><d:resourcetype/><d:getcontentlength>12</d:getcontentlength><d:getcontenttype>text/plain</d:getcontenttype><d:displayname>notiz.txt</d:displayname></d:prop></d:propstat>
 </d:response>
</d:multistatus>`;
        const items = parsePropfindXml(xml, '/remote.php/dav/files/kurt');
        expect(items).toHaveLength(2);
        expect(items[0].isCollection).toBe(true);
        expect(items[0].relPath).toBe('Physik');
        expect(items[1].isCollection).toBe(false);
        expect(items[1].relPath).toBe('Physik/notiz.txt');
        expect(items[1].size).toBe(12);
    });

    it('formats bytes', () => {
        expect(formatBytes(1536)).toContain('KB');
    });
});

describe('nextcloud-teams-migrate mapping + state', () => {
    it('loads physik mapping', () => {
        const entries = loadMapping(
            path.join(__dirname, '..', 'scripts/nextcloud-teams-migrate/mapping.physik.example.json')
        );
        expect(entries).toHaveLength(1);
        expect(entries[0].id).toBe('physik');
        expect(entries[0].teamName).toBe('Physik');
        expect(entries[0].nextcloudPath).toContain('Physik');
    });

    it('persists resume state', () => {
        const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'nc-mig-'));
        const key = stateKey('physik', 'a/b', 'gid');
        const state = loadState(dir, key);
        markDone(dir, key, 'ordner/datei.pdf', 100, state);
        const again = loadState(dir, key);
        expect(again.done['ordner/datei.pdf'].size).toBe(100);
    });
});
