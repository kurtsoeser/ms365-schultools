/**
 * Erzeugt das Power-Automate-Legacy-Paket für Schulaktivitäten (nur Direktion / Approvals).
 * node scripts/build-akt-pa-template.mjs
 */
import fs from 'fs';
import path from 'path';

const FLOW_ID = 'a7f3c91e-2b4d-4e8a-9c1f-6d5e8b0a4f27';
const BASE = 'assets/power-automate/schulaktivitaeten';

const SOURCE = {
    siteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
    listId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    emailDirektion: 'direktor@ms365.schule',
    emailMailbox: 'automate@ms365.schule',
    connectionOwner: 'kurt@kurtsoeser.at'
};

function actorExpr() {
    return (
        "@concat(first(body('Nur_Direktion')?['responses'])?['responder/displayName'], ' <', first(body('Nur_Direktion')?['responses'])?['responder/email'], '>')"
    );
}

function approvalDetails() {
    return (
        "Neuer Antrag auf Schulaktivität / Exkursion.\n\n" +
        "Titel: @{triggerBody()?['Title']}\n" +
        "Typ: @{triggerBody()?['Typ/Value']}\n" +
        "Klasse(n): @{coalesce(triggerBody()?['KlasseCodes'], triggerBody()?['KlasseCode'])}\n" +
        "Gruppe: @{triggerBody()?['GruppenText']}\n" +
        "Nicht teilnehmend: @{triggerBody()?['NichtTeilnehmende']}\n" +
        "Lehrkraft: @{triggerBody()?['LehrerCode']} (@{triggerBody()?['LehrerEmail']})\n" +
        "Begleitung: @{coalesce(triggerBody()?['BegleitungCodes'], triggerBody()?['Begleitung'])}\n" +
        "Ort: @{triggerBody()?['Ort']}\n" +
        "Zeitraum: @{triggerBody()?['Startdatum']} – @{triggerBody()?['Enddatum']}\n" +
        "Unterrichtsbezug: @{triggerBody()?['Unterrichtsbezug']}\n" +
        "Ablauf: @{triggerBody()?['Ablauf']}\n" +
        "Notiz: @{triggerBody()?['Notiz']}\n"
    );
}

function buildDefinitionInner() {
    const dataset = SOURCE.siteUrl;
    const table = SOURCE.listId;
    const assignedTo = SOURCE.emailDirektion;

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
                description: 'Neuer Antrag auf Schulaktivitäten-Liste'
            }
        },
        actions: {
            Variable_Kommentare: {
                type: 'InitializeVariable',
                inputs: { variables: [{ name: 'AlleKommentareGenehmigung', type: 'Array', value: [] }] },
                runAfter: {}
            },
            Nur_bei_Beantragt: {
                actions: {
                    Nur_Direktion: {
                        type: 'OpenApiConnectionWebhook',
                        inputs: {
                            parameters: {
                                approvalType: 'Basic',
                                'WebhookApprovalCreationInput/title':
                                    "Schulaktivität: @{triggerBody()?['Title']}",
                                'WebhookApprovalCreationInput/assignedTo': assignedTo,
                                'WebhookApprovalCreationInput/details': approvalDetails(),
                                'WebhookApprovalCreationInput/requestor':
                                    "@coalesce(triggerBody()?['LehrerEmail'], triggerBody()?['BeantragtVon'])",
                                'WebhookApprovalCreationInput/enableNotifications': true,
                                'WebhookApprovalCreationInput/enableReassignment': true
                            },
                            host: {
                                apiId: '/providers/Microsoft.PowerApps/apis/shared_approvals',
                                connectionName: 'shared_approvals',
                                operationId: 'StartAndWaitForAnApproval'
                            },
                            authentication: "@parameters('$authentication')"
                        }
                    },
                    For_each_Kommentar: {
                        runAfter: { Nur_Direktion: ['Succeeded'] },
                        type: 'Foreach',
                        foreach: "@outputs('Nur_Direktion')?['body/responses']",
                        actions: {
                            Append_Kommentar: {
                                type: 'AppendToArrayVariable',
                                inputs: {
                                    name: 'AlleKommentareGenehmigung',
                                    value:
                                        "Genehmiger: @{items('For_each_Kommentar')?['responder/displayName']}\nKommentar: @{items('For_each_Kommentar')?['comments']}\nDatum: @{outputs('Nur_Direktion')?['body/completionDate']}\n---"
                                }
                            }
                        }
                    },
                    Kommentare_Text: {
                        runAfter: { For_each_Kommentar: ['Succeeded'] },
                        type: 'Compose',
                        inputs:
                            "@join(variables('AlleKommentareGenehmigung'), decodeUriComponent('%0A%0A'))"
                    },
                    Direktion_Entscheidung: {
                        runAfter: { Kommentare_Text: ['Succeeded'] },
                        type: 'If',
                        expression: {
                            equals: ["@outputs('Nur_Direktion')?['body/outcome']", 'Approve']
                        },
                        actions: {
                            Patch_Genehmigt: {
                                type: 'OpenApiConnection',
                                inputs: {
                                    parameters: {
                                        dataset,
                                        table,
                                        id: "@triggerBody()?['ID']",
                                        // SharePoint PatchItem verlangt Title (required), auch wenn unverändert.
                                        'item/Title': "@triggerBody()?['Title']",
                                        'item/Status/Value': 'genehmigt',
                                        'item/GenehmigtVon': actorExpr(),
                                        'item/GenehmigtAm':
                                            "@outputs('Nur_Direktion')?['body/completionDate']",
                                        'item/AblehnungsGrund': '',
                                        'item/AenderungsLog':
                                            "@concat(coalesce(triggerBody()?['AenderungsLog'], ''), ' | Approval genehmigt · ', outputs('Kommentare_Text'))"
                                    },
                                    host: {
                                        apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
                                        connectionName: 'shared_sharepointonline',
                                        operationId: 'PatchItem'
                                    },
                                    authentication: "@parameters('$authentication')"
                                }
                            },
                            Mail_Genehmigt: {
                                runAfter: { Patch_Genehmigt: ['Succeeded'] },
                                type: 'OpenApiConnection',
                                inputs: {
                                    parameters: {
                                        'emailMessage/To':
                                            "@coalesce(triggerBody()?['LehrerEmail'], triggerBody()?['BeantragtVon'])",
                                        'emailMessage/Subject':
                                            "GENEHMIGT: @{triggerBody()?['Title']}",
                                        'emailMessage/Body':
                                            "<p>Guten Tag,<br><br>Ihr Antrag <b>@{triggerBody()?['Title']}</b> (@{triggerBody()?['Startdatum']} – @{triggerBody()?['Enddatum']}) wurde <b>genehmigt</b>.<br><br>Bemerkung Direktion:<br>@{outputs('Kommentare_Text')}</p>",
                                        'emailMessage/Importance': 'Normal'
                                    },
                                    host: {
                                        apiId: '/providers/Microsoft.PowerApps/apis/shared_office365',
                                        connectionName: 'shared_office365',
                                        operationId: 'SendEmailV2'
                                    },
                                    authentication: "@parameters('$authentication')"
                                }
                            }
                        },
                        else: {
                            actions: {
                                Patch_Abgelehnt: {
                                    type: 'OpenApiConnection',
                                    inputs: {
                                        parameters: {
                                            dataset,
                                            table,
                                            id: "@triggerBody()?['ID']",
                                            'item/Title': "@triggerBody()?['Title']",
                                            'item/Status/Value': 'abgelehnt',
                                            'item/AblehnungsGrund': "@outputs('Kommentare_Text')",
                                            'item/GenehmigtVon': actorExpr(),
                                            'item/GenehmigtAm':
                                                "@outputs('Nur_Direktion')?['body/completionDate']",
                                            'item/AenderungsLog':
                                                "@concat(coalesce(triggerBody()?['AenderungsLog'], ''), ' | Approval abgelehnt · ', outputs('Kommentare_Text'))"
                                        },
                                        host: {
                                            apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
                                            connectionName: 'shared_sharepointonline',
                                            operationId: 'PatchItem'
                                        },
                                        authentication: "@parameters('$authentication')"
                                    }
                                },
                                Mail_Abgelehnt: {
                                    runAfter: { Patch_Abgelehnt: ['Succeeded'] },
                                    type: 'OpenApiConnection',
                                    inputs: {
                                        parameters: {
                                            'emailMessage/To':
                                                "@coalesce(triggerBody()?['LehrerEmail'], triggerBody()?['BeantragtVon'])",
                                            'emailMessage/Subject':
                                                "ABGELEHNT: @{triggerBody()?['Title']}",
                                            'emailMessage/Body':
                                                "<p>Guten Tag,<br><br>Ihr Antrag <b>@{triggerBody()?['Title']}</b> wurde <b>abgelehnt</b>.<br><br>Bemerkung Direktion:<br>@{outputs('Kommentare_Text')}</p>",
                                            'emailMessage/Importance': 'Normal'
                                        },
                                        host: {
                                            apiId: '/providers/Microsoft.PowerApps/apis/shared_office365',
                                            connectionName: 'shared_office365',
                                            operationId: 'SendEmailV2'
                                        },
                                        authentication: "@parameters('$authentication')"
                                    }
                                }
                            }
                        }
                    }
                },
                else: { actions: {} },
                runAfter: { Variable_Kommentare: ['Succeeded'] },
                type: 'If',
                expression: {
                    equals: ["@toLower(coalesce(triggerBody()?['Status/Value'], 'beantragt'))", 'beantragt']
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
            displayName: 'Schulaktivitäten - Genehmigung Direktion',
            definition: buildDefinitionInner(),
            connectionReferences: {
                shared_sharepointonline: {
                    api: { name: 'shared_sharepointonline' },
                    connection: { id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline' },
                    source: 'Embedded',
                    connectionName: 'shared_sharepointonline',
                    id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline'
                },
                shared_approvals: {
                    api: { name: 'shared_approvals' },
                    connection: { id: '/providers/Microsoft.PowerApps/apis/shared_approvals' },
                    source: 'Embedded',
                    connectionName: 'shared_approvals',
                    id: '/providers/Microsoft.PowerApps/apis/shared_approvals'
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
    manifest.details.displayName = 'Schulaktivitäten - Workflow-Paket';
    manifest.details.description =
        'Schulaktivitäten: Trigger Liste, Microsoft Approvals an die Direktion, Status/Mail zurück an die Lehrkraft.';
    manifest.details.createdTime = new Date().toISOString();

    const oldFlowKey = '6c60dd7e-ab68-4cc8-949e-d689badc0993';
    delete manifest.resources[oldFlowKey];
    manifest.resources[FLOW_ID] = {
        type: 'Microsoft.Flow/flows',
        suggestedCreationType: 'New',
        creationType: 'Existing, New, Update',
        details: { displayName: 'Schulaktivitäten - Genehmigung Direktion' },
        configurableBy: 'User',
        hierarchy: 'Root',
        dependsOn: [
            'c04304d8-8809-418f-962b-db94c03db5dc',
            '33431c00-56f9-4901-8914-f9ac63d3525c',
            'a78c7b1c-b62c-42a7-9519-c47c021ff177',
            '455a544e-4d8b-412a-8521-953e66c97aa4',
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
    fs.readFileSync(
        'assets/power-automate/freistellung/Microsoft.Flow/flows/6c60dd7e-ab68-4cc8-949e-d689badc0993/apisMap.json'
    )
);
fs.writeFileSync(
    path.join(flowDir, 'connectionsMap.json'),
    fs.readFileSync(
        'assets/power-automate/freistellung/Microsoft.Flow/flows/6c60dd7e-ab68-4cc8-949e-d689badc0993/connectionsMap.json'
    )
);
fs.writeFileSync(
    path.join(BASE, 'Microsoft.Flow/flows/manifest.json'),
    JSON.stringify({ packageSchemaVersion: '1.0', flowAssets: { assetPaths: [FLOW_ID] } }, null, 2) + '\n'
);
fs.writeFileSync(path.join(BASE, 'manifest.json'), JSON.stringify(buildRootManifest(), null, 2) + '\n');

console.log('OK:', BASE, 'flow', FLOW_ID);
