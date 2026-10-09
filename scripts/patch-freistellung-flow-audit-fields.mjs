/**
 * Ergänzt Audit-Spalten in allen SharePoint-Patch-Aktionen des Freistellungs-Flows v2.
 * Einmal ausführen nach Flow-Export: node scripts/patch-freistellung-flow-audit-fields.mjs
 */
import fs from 'fs';

const FLOW_ID = '6c60dd7e-ab68-4cc8-949e-d689badc0993';
const defPath = `assets/power-automate/freistellung/Microsoft.Flow/flows/${FLOW_ID}/definition.json`;

function actorExpr(approvalStep) {
    return (
        "@concat(first(body('" +
        approvalStep +
        "')?['responses'])?['responder/displayName'], ' <', first(body('" +
        approvalStep +
        "')?['responses'])?['responder/email'], '>')"
    );
}

function actorExprNth(approvalStep, index) {
    const i = Number(index) || 0;
    return (
        "@concat(body('" +
        approvalStep +
        "')?['responses'][" +
        i +
        "]['responder/displayName'], ' <', body('" +
        approvalStep +
        "')?['responses'][" +
        i +
        "]['responder/email'], '>')"
    );
}

const DATE_EXPR = "@formatDateTime(utcNow(), 'yyyy-MM-dd')";

/** @param {Record<string, string>} params @param {string} actionName */
function applyAuditToPatch(params, actionName) {
    if (!params || typeof params !== 'object') return;
    const n = String(actionName || '');
    if (n.includes('Genehmigt_von_KV+Direktion')) {
        params['item/GenehmigtVonKV'] = actorExprNth('Genehmigung', 0);
        params['item/GenehmigtAmKV'] = DATE_EXPR;
        params['item/GenehmigtVonDirektion'] = actorExprNth('Genehmigung', 1);
        params['item/GenehmigtAmDirektion'] = DATE_EXPR;
        return;
    }
    if (n.includes('Genehmigt_von_KV')) {
        params['item/GenehmigtVonKV'] = actorExpr('Nur_Klassenvorstand');
        params['item/GenehmigtAmKV'] = DATE_EXPR;
        return;
    }
    if (n.includes('Abgelehnt_von_KV')) {
        params['item/AbgelehntVon'] = actorExpr('Nur_Klassenvorstand');
        params['item/AbgelehntAm'] = DATE_EXPR;
        return;
    }
    if (n.includes('Genehmigt_von_Direktion')) {
        params['item/GenehmigtVonDirektion'] = actorExpr('Nur_Direktion');
        params['item/GenehmigtAmDirektion'] = DATE_EXPR;
        return;
    }
    if (n.includes('Abgelehnt_von_Direktion')) {
        params['item/AbgelehntVon'] = actorExpr('Nur_Direktion');
        params['item/AbgelehntAm'] = DATE_EXPR;
        return;
    }
    if (n.includes('Abgelehnt')) {
        params['item/AbgelehntVon'] = actorExpr('Genehmigung');
        params['item/AbgelehntAm'] = DATE_EXPR;
    }
}

/** @param {object} o @param {string} path */
function walk(o, path) {
    if (!o || typeof o !== 'object') return;
    if (o.type === 'OpenApiConnection') {
        const params = o.inputs?.parameters;
        if (params && /PatchItem/i.test(JSON.stringify(o.inputs))) {
            const name = path.split('/').pop() || '';
            applyAuditToPatch(params, name);
        }
    }
    if (o.actions && typeof o.actions === 'object') {
        for (const [name, act] of Object.entries(o.actions)) {
            walk(act, path + '/' + name);
        }
    }
    if (o.else && typeof o.else === 'object') walk(o.else, path + '/else');
    if (Array.isArray(o.cases)) {
        o.cases.forEach((c, i) => {
            if (c?.actions) walk({ actions: c.actions }, path + '/case' + i);
        });
    }
}

const wrap = JSON.parse(fs.readFileSync(defPath, 'utf8'));
walk(wrap.properties?.definition || wrap, 'def');
fs.writeFileSync(defPath, JSON.stringify(wrap));
console.log('Audit-Felder in', defPath, 'ergänzt.');
