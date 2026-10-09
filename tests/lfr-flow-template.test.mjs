import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';

const FLOW_ASSET_ID = 'b8d4e2f1-6a3c-4d5e-8f9a-1b2c3d4e5f6a';
const defPath = path.join(
    'assets',
    'power-automate',
    'lehrer-freistellung',
    'Microsoft.Flow',
    'flows',
    FLOW_ASSET_ID,
    'definition.json'
);

describe('lehrer-freistellung flow template', () => {
    it('enthält Direktion-Genehmigung und Audit-Patches ohne KV', () => {
        const raw = fs.readFileSync(defPath, 'utf8');
        const wrap = JSON.parse(raw);
        const s = JSON.stringify(wrap.properties?.definition || wrap);
        expect(s).toContain('StartAndWaitForAnApproval');
        expect(s).toContain('PatchItem');
        expect(s).toContain('GetOnNewItems');
        expect(s).toContain('GenehmigtVonDirektion');
        expect(s).toContain('BemerkungDirektion');
        expect(s).toContain('LehrerEmail');
        expect(s).not.toContain('GenehmigtVonKV');
        expect(s).not.toContain('KvEmail');
    });

    it('flows manifest verweist auf LFR asset', () => {
        const manifest = JSON.parse(
            fs.readFileSync(
                path.join(
                    'assets',
                    'power-automate',
                    'lehrer-freistellung',
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
