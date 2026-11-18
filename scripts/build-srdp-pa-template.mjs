/**
 * Power-Automate-Legacy-Paket: sRDP-Anmelde-Bestätigung aus freigegebenem Postfach.
 * node scripts/build-srdp-pa-template.mjs
 */
import fs from 'fs';
import path from 'path';

const FLOW_ID = 'c9e2a4b1-7f3d-4c8e-9a2b-5d6e8f0a1b2c';
const BASE = 'assets/power-automate/srdp-anmeldung';

const SOURCE = {
    siteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
    listId: 'bbbbbbbb-cccc-dddd-eeee-ffffffffffff',
    emailMailbox: 'matura@ms365.schule',
    emailBcc: 'verwaltung@ms365.schule',
    connectionOwner: 'kurt@kurtsoeser.at'
};

function choice(name) {
    return `@{coalesce(triggerBody()?['${name}/Value'], triggerBody()?['${name}'])}`;
}

function field(name) {
    return `@{coalesce(triggerBody()?['${name}'], '')}`;
}

function mailBody() {
    return (
        '<p>Liebe/r ' +
        field('Vorname') +
        ' ' +
        field('Nachname') +
        '!</p>' +
        '<p>deine <b>Anmeldung zur sRDP / Matura</b> ist eingegangen. Hier die Zusammenfassung:</p>' +
        '<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:14px;">' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Name</td><td style="padding:4px 0;"><b>@{coalesce(triggerBody()?[\'Title\'], triggerBody()?[\'Nachname\'])}</b></td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Klasse</td><td style="padding:4px 0;">' +
        choice('Klasse') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Variante</td><td style="padding:4px 0;">' +
        choice('Variante') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Prüfplan</td><td style="padding:4px 0;">' +
        field('PruefplanKurz') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Diplomarbeit</td><td style="padding:4px 0;">' +
        field('TitelDiplomarbeit') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Betreuung DA</td><td style="padding:4px 0;">' +
        choice('BetreuungslehrerDA') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">LFS</td><td style="padding:4px 0;">' +
        choice('LFS') +
        ' · ' +
        choice('LehrerLFS') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">BKO mündlich</td><td style="padding:4px 0;">' +
        choice('LehrerBKO') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Wahlfach mündlich</td><td style="padding:4px 0;">' +
        choice('WahlfachMuendlich') +
        ' · ' +
        choice('LehrerWahlfach') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Seminar</td><td style="padding:4px 0;">' +
        choice('Seminar') +
        '</td></tr>' +
        '<tr><td style="padding:4px 12px 4px 0;color:#666;">Terminjahr</td><td style="padding:4px 0;">' +
        field('TerminJahr') +
        '</td></tr>' +
        '</table>' +
        '<p style="margin-top:16px;">Bei Unklarheiten bitte das Sekretariat / die Prüfungsorganisation kontaktieren.<br>' +
        '<i>Bitte nicht auf diese E-Mail antworten – sie wird nicht überwacht.</i></p>'
    );
}

function buildDefinitionInner() {
    const dataset = SOURCE.siteUrl;
    const table = SOURCE.listId;
    const mailbox = SOURCE.emailMailbox;
    const bcc = SOURCE.emailBcc;

    return {
        $schema:
            'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
        contentVersion: '1.0.0.0',
        parameters: {
            $authentication: { defaultValue: {}, type: 'SecureObject' },
            $connections: { defaultValue: {}, type: 'Object' }
        },
        triggers: {
            Item_Is_Created: {
                recurrence: { frequency: 'Minute', interval: 1 },
                evaluatedRecurrence: { frequency: 'Minute', interval: 1 },
                splitOn: "@triggerBody()?['value']",
                type: 'OpenApiConnection',
                inputs: {
                    parameters: { dataset, table },
                    host: {
                        apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
                        connectionName: 'shared_sharepointonline',
                        operationId: 'GetOnNewItems'
                    },
                    authentication: "@parameters('$authentication')"
                },
                description: 'Neue sRDP-/Matura-Anmeldung in der Jahresliste'
            }
        },
        actions: {
            Empfaenger: {
                type: 'Compose',
                inputs: "@coalesce(triggerBody()?['Author/Email'], '')",
                runAfter: {}
            },
            Bestaetigungsmail: {
                type: 'OpenApiConnection',
                runAfter: { Empfaenger: ['Succeeded'] },
                inputs: {
                    parameters: {
                        'emailMessage/MailboxAddress': mailbox,
                        'emailMessage/To': "@outputs('Empfaenger')",
                        'emailMessage/Cc': '',
                        'emailMessage/Bcc': bcc,
                        'emailMessage/Subject':
                            "Bestätigung sRDP-Anmeldung: @{coalesce(triggerBody()?['Title'], triggerBody()?['Nachname'])}",
                        'emailMessage/Body': mailBody(),
                        'emailMessage/Importance': 'Normal'
                    },
                    host: {
                        apiId: '/providers/Microsoft.PowerApps/apis/shared_office365',
                        connectionName: 'shared_office365',
                        operationId: 'SharedMailboxSendEmailV2'
                    },
                    authentication: "@parameters('$authentication')"
                }
            }
        },
        outputs: {}
    };
}

function buildFlowWrapper() {
    return {
        name: FLOW_ID,
        id: '/providers/Microsoft.Flow/flows/' + FLOW_ID,
        type: 'Microsoft.Flow/flows',
        properties: {
            apiId: '/providers/Microsoft.PowerApps/apis/shared_logicflows',
            displayName: 'sRDP-Anmeldung - Bestätigungsmail',
            definition: buildDefinitionInner(),
            connectionReferences: {
                shared_sharepointonline: {
                    api: { name: 'shared_sharepointonline' },
                    connection: { id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline' },
                    source: 'Embedded',
                    connectionName: 'shared_sharepointonline',
                    id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline'
                },
                shared_office365: {
                    api: { name: 'shared_office365' },
                    connection: { id: '/providers/Microsoft.PowerApps/apis/shared_office365' },
                    source: 'Embedded',
                    connectionName: 'shared_office365',
                    id: '/providers/Microsoft.PowerApps/apis/shared_office365'
                }
            },
            flowFailureAlertSubscribed: false,
            isManaged: false
        }
    };
}

function buildRootManifest() {
    const student = JSON.parse(
        fs.readFileSync('assets/power-automate/freistellung/manifest.json', 'utf8')
    );
    const manifest = JSON.parse(JSON.stringify(student));
    manifest.details.displayName = 'sRDP-Anmeldung - Bestätigungsmail';
    manifest.details.description =
        'Bei neuer Matura-/sRDP-Anmeldung: Bestätigungsmail mit Zusammenfassung aus freigegebenem Postfach.';
    manifest.details.createdTime = new Date().toISOString();

    // Approvals entfernen – dieser Flow braucht nur SharePoint + Outlook
    const drop = [
        '6c60dd7e-ab68-4cc8-949e-d689badc0993',
        'a78c7b1c-b62c-42a7-9519-c47c021ff177',
        '455a544e-4d8b-412a-8521-953e66c97aa4'
    ];
    drop.forEach((k) => delete manifest.resources[k]);

    manifest.resources[FLOW_ID] = {
        type: 'Microsoft.Flow/flows',
        suggestedCreationType: 'New',
        creationType: 'Existing, New, Update',
        details: { displayName: 'sRDP-Anmeldung - Bestätigungsmail' },
        configurableBy: 'User',
        hierarchy: 'Root',
        dependsOn: [
            'c04304d8-8809-418f-962b-db94c03db5dc',
            '33431c00-56f9-4901-8914-f9ac63d3525c',
            '78e2cb2d-0683-49c9-9c71-62ddcbf3930d',
            '0d335309-e936-4762-90f4-6a5ce122fbab'
        ]
    };
    return manifest;
}

function ensureDir(p) {
    fs.mkdirSync(p, { recursive: true });
}

const flowDir = path.join(BASE, 'Microsoft.Flow/flows', FLOW_ID);
ensureDir(flowDir);

fs.writeFileSync(path.join(flowDir, 'definition.json'), JSON.stringify(buildFlowWrapper()));
fs.writeFileSync(
    path.join(flowDir, 'apisMap.json'),
    JSON.stringify({
        shared_sharepointonline: 'c04304d8-8809-418f-962b-db94c03db5dc',
        shared_office365: '78e2cb2d-0683-49c9-9c71-62ddcbf3930d'
    })
);
fs.writeFileSync(
    path.join(flowDir, 'connectionsMap.json'),
    JSON.stringify({
        shared_sharepointonline: '33431c00-56f9-4901-8914-f9ac63d3525c',
        shared_office365: '0d335309-e936-4762-90f4-6a5ce122fbab'
    })
);
fs.writeFileSync(
    path.join(BASE, 'Microsoft.Flow/flows/manifest.json'),
    JSON.stringify({ packageSchemaVersion: '1.0', flowAssets: { assetPaths: [FLOW_ID] } }, null, 2) + '\n'
);
fs.writeFileSync(path.join(BASE, 'manifest.json'), JSON.stringify(buildRootManifest(), null, 2) + '\n');

console.log('OK:', BASE, 'flow', FLOW_ID);
