import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';

const FLOW_ASSET_ID = 'c9e2a4b1-7f3d-4c8e-9a2b-5d6e8f0a1b2c';
const defPath = path.join(
    'assets',
    'power-automate',
    'srdp-anmeldung',
    'Microsoft.Flow',
    'flows',
    FLOW_ASSET_ID,
    'definition.json'
);

describe('srdp bestätigungsmail flow template', () => {
    it('Trigger neue Anmeldung + Shared-Mailbox-Mail mit Zusammenfassung', () => {
        const raw = fs.readFileSync(defPath, 'utf8');
        const wrap = JSON.parse(raw);
        const s = JSON.stringify(wrap.properties?.definition || wrap);
        expect(s).toContain('GetOnNewItems');
        expect(s).toContain('SharedMailboxSendEmailV2');
        expect(s).toContain('emailMessage/MailboxAddress');
        expect(s).toContain('PruefplanKurz');
        expect(s).toContain('BetreuungslehrerDA');
        expect(s).toContain('LehrerLFS');
        expect(s).toContain('LehrerBKO');
        expect(s).toContain('WahlfachMuendlich');
        expect(s).toContain('Author/Email');
        expect(s).not.toContain('StartAndWaitForAnApproval');
    });

    it('flows manifest verweist auf sRDP asset', () => {
        const manifest = JSON.parse(
            fs.readFileSync(
                path.join(
                    'assets',
                    'power-automate',
                    'srdp-anmeldung',
                    'Microsoft.Flow',
                    'flows',
                    'manifest.json'
                ),
                'utf8'
            )
        );
        expect(manifest.flowAssets.assetPaths).toContain(FLOW_ASSET_ID);
    });
});
