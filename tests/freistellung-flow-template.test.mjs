import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';

const FLOW_ASSET_ID = 'eff47cd0-dd67-468d-a48a-9e146aab57a7';
const defPath = path.join(
    'assets',
    'power-automate',
    'freistellung',
    'Microsoft.Flow',
    'flows',
    FLOW_ASSET_ID,
    'definition.json'
);

describe('freistellung flow template (v2)', () => {
    it('enthält v2-Genehmigungs- und Patch-Aktionen', () => {
        const raw = fs.readFileSync(defPath, 'utf8');
        const wrap = JSON.parse(raw);
        const s = JSON.stringify(wrap.properties?.definition || wrap);
        expect(s).toContain('StartAndWaitForAnApproval');
        expect(s).toContain('PatchItem');
        expect(s).toContain('GetOnNewItems');
    });

    it('flows manifest verweist auf v2 asset', () => {
        const manifest = JSON.parse(
            fs.readFileSync(
                path.join('assets', 'power-automate', 'freistellung', 'Microsoft.Flow', 'flows', 'manifest.json'),
                'utf8'
            )
        );
        expect(manifest.flowAssets.assetPaths).toContain(FLOW_ASSET_ID);
    });
});
