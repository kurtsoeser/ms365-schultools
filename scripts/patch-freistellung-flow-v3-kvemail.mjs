/**
 * Stabilisiert KV=Direktion-Vergleich in der Freistellung-Flow-Vorlage v3:
 * Compose KvEmail vor der Bedingung; Vergleich per toLower(outputs('KvEmail')).
 */
import fs from 'fs';

const FLOW_ID = '6c60dd7e-ab68-4cc8-949e-d689badc0993';
const defPath = `assets/power-automate/freistellung/Microsoft.Flow/flows/${FLOW_ID}/definition.json`;

const j = JSON.parse(fs.readFileSync(defPath, 'utf8'));
const actions = j.properties.definition.actions;
if (!actions.KvEmail) {
    throw new Error('KvEmail Compose fehlt in der Vorlage.');
}
if (!actions.Condition) {
    throw new Error('Condition fehlt in der Vorlage.');
}

actions.KvEmail.inputs = "@triggerBody()?['Klassenvorstand/Email']";
actions.KvEmail.runAfter = { Variable_initialisieren: ['Succeeded'] };

actions.Condition.runAfter = { KvEmail: ['Succeeded'] };
actions.Condition.expression = {
    and: [
        {
            equals: ["@toLower(trim(outputs('KvEmail')))", 'direktor@ms365.schule']
        }
    ]
};

fs.writeFileSync(defPath, JSON.stringify(j));
console.log('Patched', defPath);
console.log('Condition:', JSON.stringify(actions.Condition.expression));
