/**
 * Optionaler lokaler Check. Produktiv: Setup-UI → Flow-Paket (Site + Listen-GUID aus Formular).
 *
 * AKT_SITE_URL=… AKT_LIST_ID=… npm run pack:akt-pa
 */
import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';
import JSZip from 'jszip';
import {
    AKT_FLOW_ASSET_ID,
    AKT_FLOW_SOURCE,
    applyAktFlowParamsToDefinition,
    assertAktFlowDefinitionMatchesSchool,
    normalizeSharePointListGuid,
    replaceAllLiteral
} from '../src/tools/power-automate/schulaktivitaeten-setup-package.js';

const FLOW_ID = AKT_FLOW_ASSET_ID;
const BASE = 'assets/power-automate/schulaktivitaeten';

const TARGET = {
    siteUrl: String(process.env.AKT_SITE_URL || '')
        .trim()
        .replace(/\/+$/, ''),
    listId: normalizeSharePointListGuid(process.env.AKT_LIST_ID || ''),
    listName: String(process.env.AKT_LIST_NAME || AKT_FLOW_SOURCE.listTitle).trim(),
    emailDirektion: String(process.env.AKT_EMAIL_DIREKTION || AKT_FLOW_SOURCE.emailDirektion)
        .trim()
        .toLowerCase(),
    emailMailbox: String(process.env.AKT_EMAIL_MAILBOX || AKT_FLOW_SOURCE.emailMailbox)
        .trim()
        .toLowerCase(),
    connectionOwner: String(process.env.AKT_FLOW_ACCOUNT || AKT_FLOW_SOURCE.connectionOwner)
        .trim()
        .toLowerCase(),
    flowDisplayName: 'Schulaktivitäten - Genehmigung Direktion'
};

if (!TARGET.siteUrl || !TARGET.listId) {
    console.error(
        'Abbruch: AKT_SITE_URL und AKT_LIST_ID (GUID der Liste Schulaktivitaeten) sind Pflicht.\n' +
            'Produktiv: Setup-Assistent → Listen anlegen/prüfen → Flow-Paket herunterladen.'
    );
    process.exit(1);
}

execSync('node scripts/build-akt-pa-template.mjs', { stdio: 'inherit' });

const defRaw = fs.readFileSync(
    path.join(BASE, 'Microsoft.Flow/flows', FLOW_ID, 'definition.json'),
    'utf8'
);
const rootManifest = JSON.parse(fs.readFileSync(path.join(BASE, 'manifest.json'), 'utf8'));
const flowsManifestRaw = fs.readFileSync(
    path.join(BASE, 'Microsoft.Flow/flows/manifest.json'),
    'utf8'
);
const apisMapRaw = fs.readFileSync(
    path.join(BASE, 'Microsoft.Flow/flows', FLOW_ID, 'apisMap.json'),
    'utf8'
);
const connectionsMapRaw = fs.readFileSync(
    path.join(BASE, 'Microsoft.Flow/flows', FLOW_ID, 'connectionsMap.json'),
    'utf8'
);

const definition = applyAktFlowParamsToDefinition(
    defRaw,
    {
        siteUrl: TARGET.siteUrl,
        listId: TARGET.listId,
        listName: TARGET.listName,
        emailDirektion: TARGET.emailDirektion,
        emailMailbox: TARGET.emailMailbox,
        flowServiceAccount: TARGET.connectionOwner,
        flowDisplayName: TARGET.flowDisplayName
    },
    { flowAccount: TARGET.connectionOwner, newFlowId: FLOW_ID }
);
assertAktFlowDefinitionMatchesSchool(definition, {
    siteUrl: TARGET.siteUrl,
    listId: TARGET.listId
});

rootManifest.details = rootManifest.details || {};
rootManifest.details.displayName = TARGET.flowDisplayName;
rootManifest.details.createdTime = new Date().toISOString();

Object.keys(rootManifest.resources || {}).forEach((key) => {
    const r = rootManifest.resources[key];
    if (!r || !r.details) return;
    if (r.type === 'Microsoft.PowerApps/apis/connections') {
        if (String(r.details.displayName || '').indexOf('@') !== -1) {
            r.details.displayName = TARGET.connectionOwner;
        }
    }
    if (r.type === 'Microsoft.Flow/flows') {
        r.details.displayName = TARGET.flowDisplayName;
        r.suggestedCreationType = 'New';
    }
});

let rootStr = JSON.stringify(rootManifest);
rootStr = replaceAllLiteral(rootStr, AKT_FLOW_SOURCE.connectionOwner, TARGET.connectionOwner);
const rootOut = JSON.parse(rootStr);

const zip = new JSZip();
zip.file('manifest.json', JSON.stringify(rootOut, null, 2));
zip.file('Microsoft.Flow/flows/manifest.json', flowsManifestRaw);
const folder = zip.folder('Microsoft.Flow/flows/' + FLOW_ID);
folder.file('definition.json', JSON.stringify(definition, null, 2));
folder.file('apisMap.json', apisMapRaw);
folder.file('connectionsMap.json', connectionsMapRaw);

const safeHost = TARGET.siteUrl.replace(/^https?:\/\//i, '').replace(/[^\w.-]+/g, '_');
const outPath = path.join(BASE, '_local', 'Schulaktivitaeten_' + safeHost + '.zip');
fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));

console.log('OK lokales Test-ZIP:', outPath);
console.log('  site:', TARGET.siteUrl);
console.log('  listId:', TARGET.listId);
