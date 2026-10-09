import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';

const FLOW_ASSET_ID = '6c60dd7e-ab68-4cc8-949e-d689badc0993';
const defPath = path.join(
    'assets',
    'power-automate',
    'freistellung',
    'Microsoft.Flow',
    'flows',
    FLOW_ASSET_ID,
    'definition.json'
);

describe('freistellung flow template (v3)', () => {
    it('enthält v3-Genehmigungs- und Patch-Aktionen', () => {
        const raw = fs.readFileSync(defPath, 'utf8');
        const wrap = JSON.parse(raw);
        const s = JSON.stringify(wrap.properties?.definition || wrap);
        expect(s).toContain('StartAndWaitForAnApproval');
        expect(s).toContain('PatchItem');
        expect(s).toContain('GetOnNewItems');
        expect(s).toContain('GenehmigtVonKV');
        expect(s).toContain('Abgelehnt');
        expect(s).toContain('AlleKommentareGenehmigung');
        expect(s).toContain("toLower(trim(outputs('KvEmail')))");
    });

    it('flows manifest verweist auf v3 asset', () => {
        const manifest = JSON.parse(
            fs.readFileSync(
                path.join('assets', 'power-automate', 'freistellung', 'Microsoft.Flow', 'flows', 'manifest.json'),
                'utf8'
            )
        );
        expect(manifest.flowAssets.assetPaths).toContain(FLOW_ASSET_ID);
    });
});
