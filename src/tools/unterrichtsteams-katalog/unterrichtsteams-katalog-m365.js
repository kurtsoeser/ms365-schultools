/**
 * Microsoft-365-Teams für Unterrichtsteams-Katalog laden und verknüpfen.
 */
import { getGraphToken } from '../kursteams/kursteam-graph.js';
import { buildBelegungRowsFromGraphGroups } from '../../shared/kursteam-graph-import-logic.js';
import { buildKursteamAbgleichReport } from '../../shared/kursteam-belegung-abgleich-logic.js';
import {
    buildTeacherLookupFromCodeMap,
    enrichBelegungRows
} from '../../shared/unterrichtsbelegung-teacher-enrich-logic.js';
import { applyAbgleichMatchesToRows } from './unterrichtsteams-katalog-logic.js';

/** @type {object[]} */
let cachedM365Rows = [];

export function getCachedM365Rows() {
    return cachedM365Rows.slice();
}

function buildTeacherByCodeMap() {
    const map = new Map();
    try {
        const load =
            typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad : null;
        const data = load ? load() : null;
        const teachers = data && Array.isArray(data.teachers) ? data.teachers : [];
        teachers.forEach((t) => {
            const code = String((t && t.code) || '')
                .trim()
                .toUpperCase();
            if (!code) return;
            map.set(code, {
                email: String((t && t.email) || '')
                    .trim()
                    .toLowerCase(),
                name: String((t && t.name) || '').trim()
            });
        });
    } catch {
        /* ignore */
    }
    return map;
}

async function graphJson(method, path, token, extraHeaders) {
    const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
    const headers = Object.assign({ Authorization: 'Bearer ' + token }, extraHeaders || {});
    const res = await fetch(url, { method, headers });
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg =
            typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
        throw new Error(method + ' ' + path + ': ' + msg);
    }
    return data || {};
}

async function listTeamGroupsFromGraph(token) {
    const collected = [];
    let nextPath =
        "/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,mailNickname,resourceProvisioningOptions&$top=999";
    try {
        while (nextPath) {
            const useEv = nextPath.indexOf('http') !== 0;
            const data = await graphJson(
                'GET',
                nextPath,
                token,
                useEv ? { ConsistencyLevel: 'eventual' } : undefined
            );
            (data.value || []).forEach((g) => {
                if (g && g.mailNickname) collected.push(g);
            });
            nextPath = data['@odata.nextLink'] || null;
        }
        return collected;
    } catch {
        return listTeamGroupsFromGraphFallback(token);
    }
}

async function listTeamGroupsFromGraphFallback(token) {
    const collected = [];
    let nextPath = '/groups?$select=id,displayName,mailNickname,resourceProvisioningOptions&$top=999';
    while (nextPath) {
        const data = await graphJson('GET', nextPath, token);
        (data.value || []).forEach((g) => {
            const opts = g && g.resourceProvisioningOptions ? g.resourceProvisioningOptions : [];
            if (opts.indexOf('Team') !== -1 && g.mailNickname) collected.push(g);
        });
        nextPath = data['@odata.nextLink'] || null;
    }
    return collected;
}

/**
 * @param {string} yearPrefix
 * @returns {Promise<object[]>}
 */
export async function fetchM365KursteamRows(yearPrefix) {
    const yp = String(yearPrefix || '').trim();
    const token = await getGraphToken();
    const groups = await listTeamGroupsFromGraph(token);
    const teacherByCode = buildTeacherByCodeMap();
    const rawRows = buildBelegungRowsFromGraphGroups(groups, { yearPrefix: yp, teacherByCode });
    const lookup = buildTeacherLookupFromCodeMap(teacherByCode);
    cachedM365Rows = enrichBelegungRows(rawRows, lookup, new Map());
    return cachedM365Rows;
}

/**
 * @param {object[]} plannedRows
 * @param {object[]} [m365Rows]
 */
export function buildAbgleichForRows(plannedRows, m365Rows) {
    const m365 = Array.isArray(m365Rows) ? m365Rows : cachedM365Rows;
    return buildKursteamAbgleichReport(plannedRows, m365);
}

/**
 * @param {object[]} plannedRows
 * @param {object[]} [m365Rows]
 */
export function autoLinkRows(plannedRows, m365Rows) {
    const report = buildAbgleichForRows(plannedRows, m365Rows);
    const applied = applyAbgleichMatchesToRows(plannedRows, report.matched);
    return { rows: applied.rows, linked: applied.linked, report };
}

/**
 * @param {object} plannedRow
 * @param {object[]} [m365Rows]
 * @returns {object|null} Patch für graphGroupId / teamName / gruppenmail
 */
export function linkPatchForSingleRow(plannedRow, m365Rows) {
    const report = buildKursteamAbgleichReport([plannedRow], Array.isArray(m365Rows) ? m365Rows : cachedM365Rows);
    const m = report.matched && report.matched[0];
    if (!m || !String(m.graphGroupId || '').trim()) return null;
    return {
        graphGroupId: String(m.graphGroupId).trim(),
        gruppenmail: plannedRow.gruppenmail || m.gruppenmail || '',
        teamName: plannedRow.teamName || m.teamName || ''
    };
}

export function entraGroupUrl(groupId) {
    const id = String(groupId || '').trim();
    if (!id) return '';
    return (
        'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/' +
        encodeURIComponent(id)
    );
}
