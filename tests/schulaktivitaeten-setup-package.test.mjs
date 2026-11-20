import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import {
    AKT_FLOW_ASSET_ID,
    AKT_FLOW_SOURCE,
    applyAktFlowParamsToDefinition,
    assertAktFlowDefinitionMatchesSchool,
    normalizeSharePointListGuid
} from '../src/tools/power-automate/schulaktivitaeten-setup-package.js';

const defPath = path.join(
    'assets',
    'power-automate',
    'schulaktivitaeten',
    'Microsoft.Flow',
    'flows',
    AKT_FLOW_ASSET_ID,
    'definition.json'
);

describe('schulaktivitaeten setup package (Multitenant)', () => {
    it('Vorlage enthält Platzhalter-Site und Platzhalter-Listen-GUID', () => {
        const wrap = JSON.parse(fs.readFileSync(defPath, 'utf8'));
        const params = wrap.properties.definition.triggers.Item_Is_Created.inputs.parameters;
        expect(params.dataset).toBe(AKT_FLOW_SOURCE.siteUrl);
        expect(normalizeSharePointListGuid(params.table)).toBe(
            normalizeSharePointListGuid(AKT_FLOW_SOURCE.listId)
        );
    });

    it('ersetzt Site und Listen-GUID für eine beliebige Schule', () => {
        const raw = fs.readFileSync(defPath, 'utf8');
        const other = {
            siteUrl: 'https://andere-schule.sharepoint.com/sites/MS365-Schultools',
            listId: 'f1e2d3c4-b5a6-4789-8abc-def012345678',
            listName: 'Schulaktivitaeten',
            emailDirektion: 'direktion@andere-schule.at',
            emailMailbox: 'flow@andere-schule.at',
            flowServiceAccount: 'technik@andere-schule.at',
            flowDisplayName: 'Schulaktivitäten - Genehmigung Direktion'
        };
        const def = applyAktFlowParamsToDefinition(raw, other, {
            flowAccount: other.flowServiceAccount,
            newFlowId: '11111111-2222-4333-8444-555555555555'
        });
        assertAktFlowDefinitionMatchesSchool(def, other);
        const s = JSON.stringify(def);
        expect(s).toContain(other.siteUrl);
        expect(s).not.toContain(AKT_FLOW_SOURCE.siteUrl);
        expect(s).not.toContain('kurtrocks.sharepoint.com');
        expect(s.toLowerCase()).toContain(other.listId.toLowerCase());
        expect(s).not.toContain(AKT_FLOW_SOURCE.listId);
        expect(s).toContain(other.emailDirektion);
    });

    it('lehnt Paket ohne echte Listen-GUID ab', () => {
        const raw = fs.readFileSync(defPath, 'utf8');
        expect(() =>
            applyAktFlowParamsToDefinition(raw, {
                siteUrl: 'https://schule.sharepoint.com/sites/X',
                listName: 'Schulaktivitaeten'
            })
        ).toThrow(/Listen-ID/);
        expect(() =>
            applyAktFlowParamsToDefinition(raw, {
                siteUrl: 'https://schule.sharepoint.com/sites/X',
                listId: AKT_FLOW_SOURCE.listId
            })
        ).toThrow(/Platzhalter/);
    });

    it('erlaubt Ziel-Site = Platzhalter-Site (Demo-/Kurtrocks-Mandant)', () => {
        const raw = fs.readFileSync(defPath, 'utf8');
        const same = {
            siteUrl: AKT_FLOW_SOURCE.siteUrl,
            listId: '74f87ee3-2aa1-44c3-b3d8-bfc84ba31208',
            emailDirektion: 'direktor@kurtrocks.com',
            flowServiceAccount: 'technik@kurtrocks.com'
        };
        const def = applyAktFlowParamsToDefinition(raw, same, {
            flowAccount: same.flowServiceAccount,
            newFlowId: '22222222-3333-4444-8555-666666666666'
        });
        assertAktFlowDefinitionMatchesSchool(def, same);
        expect(
            def.properties.definition.triggers.Item_Is_Created.inputs.parameters.dataset
        ).toBe(AKT_FLOW_SOURCE.siteUrl);
        expect(
            normalizeSharePointListGuid(
                def.properties.definition.triggers.Item_Is_Created.inputs.parameters.table
            )
        ).toBe(same.listId);
    });
});
