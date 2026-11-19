import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';

const FLOW_ASSET_ID = 'a7f3c91e-2b4d-4e8a-9c1f-6d5e8b0a4f27';
const defPath = path.join(
    'assets',
    'power-automate',
    'schulaktivitaeten',
    'Microsoft.Flow',
    'flows',
    FLOW_ASSET_ID,
    'definition.json'
);

describe('schulaktivitaeten flow template', () => {
    it('enthält Direktion-Approval und Status-Patches ohne KV', () => {
        const raw = fs.readFileSync(defPath, 'utf8');
        const wrap = JSON.parse(raw);
        const s = JSON.stringify(wrap.properties?.definition || wrap);
        expect(s).toContain('StartAndWaitForAnApproval');
        expect(s).toContain('PatchItem');
        expect(s).toContain('GetOnNewItems');
        expect(s).toContain('LehrerEmail');
        expect(s).toContain('genehmigt');
        expect(s).toContain('abgelehnt');
        expect(s).toContain('beantragt');
        expect(s).toContain('AblehnungsGrund');
        expect(s).toContain('GenehmigtVon');
        // PatchItem verlangt Title – sonst scheitert der Legacy-Import.
        expect(s).toContain("item/Title");
        expect(s).toMatch(/Patch_Genehmigt[\s\S]*item\/Title/);
        expect(s).toMatch(/Patch_Abgelehnt[\s\S]*item\/Title/);
        expect(s).not.toContain('GenehmigtVonKV');
        expect(s).not.toContain('KvEmail');
    });

    it('InitializeVariable liegt Top-Level (nicht in If/Scope)', () => {
        const wrap = JSON.parse(fs.readFileSync(defPath, 'utf8'));
        const def = wrap.properties?.definition || wrap;
        const top = def.actions || {};
        expect(top.Variable_Kommentare?.type).toBe('InitializeVariable');
        expect(top.Nur_bei_Beantragt?.type).toBe('If');
        expect(top.Nur_bei_Beantragt?.runAfter).toEqual({ Variable_Kommentare: ['Succeeded'] });
        const nested = JSON.stringify(top.Nur_bei_Beantragt?.actions || {});
        expect(nested).not.toContain('"InitializeVariable"');
    });

    it('flows manifest verweist auf AKT asset', () => {
        const manifest = JSON.parse(
            fs.readFileSync(
                path.join(
                    'assets',
                    'power-automate',
                    'schulaktivitaeten',
                    'Microsoft.Flow',
                    'flows',
                    'manifest.json'
                ),
                'utf8'
            )
        );
        expect(manifest.flowAssets.assetPaths).toContain(FLOW_ASSET_ID);
    });

    it('copy-static liefert Flow-Vorlagen nach dist (sonst 404 im Setup)', () => {
        const copyStatic = fs.readFileSync(path.join('scripts', 'copy-static.mjs'), 'utf8');
        expect(copyStatic).toContain("copyDirToDist('assets/power-automate')");
        expect(fs.existsSync(defPath)).toBe(true);
    });
});
