/**
 * tenant-settings-ui-directory-users.js
 * @param {Record<string, any>} ctx
 */
import { directoryStatusView, paintStatusCell } from './tenant-workbench-chrome.js';

export function installTenantDirectoryUsers(ctx) {
    ctx.graphApi = function graphApi() {
        const api = window.ms365GraphUnifiedGroups;
        if (!api) throw new Error('Microsoft-Graph-Modul fehlt.');
        return api;
    }

    ctx.randomTempPassword = function randomTempPassword() {
        const u = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
        const l = 'abcdefghijkmnopqrstuvwxyz';
        const d = '23456789';
        const s = '!@#$%&*';
        function pick(set) {
            return set.charAt(Math.floor(Math.random() * set.length));
        }
        let pwd = pick(u) + pick(l) + pick(d) + pick(s);
        for (let i = 0; i < 12; i++) pwd += pick(u + l + d);
        return pwd;
    }

    ctx.mailNicknameFromUpn = function mailNicknameFromUpn(upn) {
        const local = String(upn || '').split('@')[0] || 'user';
        return ctx.graphApi().sanitizeMailNickname(local.replace(/[^a-zA-Z0-9]/g, '') || 'user');
    }

    ctx.patchDirectoryMatchKeys = function patchDirectoryMatchKeys(emailKeys, payload) {
        const patch = {};
        const seen = new Set();
        (emailKeys || []).forEach(function (k) {
            const em = ctx.normStr(k).toLowerCase();
            if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
            seen.add(em);
            patch[em] = payload;
        });
        if (!Object.keys(patch).length) return;
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: patch });
        }
    }

    ctx.directoryMatchUserPayload = function directoryMatchUserPayload(user, iso) {
        return {
            graphUserId: String(user && user.id ? user.id : '').trim(),
            displayName: String((user && (user.displayName || user.displayNameHint)) || '').trim(),
            userPrincipalName: String((user && (user.userPrincipalName || user.upn)) || '').trim(),
            notFound: false,
            checkedAt: iso || new Date().toISOString()
        };
    }

    ctx.createEntraUserViaGraph = async function createEntraUserViaGraph(token, upn, displayName) {
        const pwd = ctx.randomTempPassword();
        const mailNick = ctx.mailNicknameFromUpn(upn);
        const body = {
            accountEnabled: true,
            displayName: String(displayName).trim(),
            mailNickname: mailNick,
            userPrincipalName: String(upn).trim(),
            passwordProfile: {
                forceChangePasswordNextSignIn: true,
                password: pwd
            }
        };
        const created = await ctx.graphApi().graphJson('POST', '/users', token, body, undefined);
        const uid = String(created.id || '').trim();
        if (!uid) throw new Error('Keine Benutzer-ID von Graph erhalten.');
        return {
            id: uid,
            password: pwd,
            upn: String(upn).trim(),
            displayName: String(displayName).trim(),
            mailNickname: mailNick
        };
    }

    ctx.getDirectoryMatchByEmail = function getDirectoryMatchByEmail(emailRaw) {
        const em = ctx.normStr(emailRaw).toLowerCase();
        const api = window.ms365AppDataV2;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const map = setup && setup.directoryMatchByEmail ? setup.directoryMatchByEmail : {};
        return em ? map[em] || null : null;
    }

    ctx.verifyAdminDirectoryEmail = async function verifyAdminDirectoryEmail(emailRaw) {
        const em = ctx.normStr(emailRaw).toLowerCase();
        if (!em || em.indexOf('@') === -1) {
            ctx.setSummary('Bitte zuerst eine gültige E-Mail eintragen.', 'warn');
            return false;
        }
        try {
            const token = await ctx.graphApi().getGraphToken();
            const u = await ctx.graphApi().resolveUserByEmail(token, em);
            const iso = new Date().toISOString();
            if (u && u.id) {
                ctx.patchDirectoryMatchKeys([em], ctx.directoryMatchUserPayload(u, iso));
                ctx.setSummary('Microsoft 365 gefunden: ' + (u.displayName || em), 'ok');
            } else {
                ctx.patchDirectoryMatchKeys([em], { notFound: true, checkedAt: iso });
                ctx.setSummary('Kein Entra-Benutzer für: ' + em, 'warn');
            }
            ctx.refreshAdminAudienceBoardUi();
            return !!(u && u.id);
        } catch (e) {
            ctx.setSummary('Abgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
            return false;
        }
    }

    ctx.runVerifyVerwaltungGraphBulk = async function runVerifyVerwaltungGraphBulk() {
        if (!ctx.btnVerifyVerwaltungGraph) return;
        await ctx.runVerifyGraphDirectoryRows(
            ctx.groupsToDisplayRows(ctx.getAdministrationGroups()),
            function (r) {
                return r && r.email;
            },
            'Verwaltung',
            ctx.btnVerifyVerwaltungGraph,
            function () {
                ctx.refreshAdminAudienceBoardUi();
            }
        );
    }

    ctx.buildDirectoryMatchCell = function buildDirectoryMatchCell(tr, emailRaw, dirMatchMapOpt) {
        const tdMs = document.createElement('td');
        tdMs.style.fontSize = '0.88em';
        tdMs.style.lineHeight = '1.35';
        const em = ctx.normStr(emailRaw).toLowerCase();
        let m = null;
        if (em && em.indexOf('@') !== -1) {
            if (dirMatchMapOpt && typeof dirMatchMapOpt === 'object') {
                m = dirMatchMapOpt[em] || null;
            } else {
                m = ctx.getDirectoryMatchByEmail(em);
            }
        }
        paintStatusCell(tdMs, directoryStatusView(em, m), tr);
        return tdMs;
    }

    ctx.runVerifyGraphDirectoryRows = async function runVerifyGraphDirectoryRows(rows, getEmail, label, btn, onDone) {
        const updates = {};
        let found = 0;
        let missed = 0;
        let skipped = 0;
        const seen = new Set();
        const unique = [];
        const iso = new Date().toISOString();
        (rows || []).forEach(function (row) {
            const em = ctx.normStr(getEmail(row) || '').toLowerCase();
            if (!em || em.indexOf('@') === -1) {
                skipped++;
                return;
            }
            if (seen.has(em)) return;
            seen.add(em);
            unique.push(em);
        });
        try {
            if (btn) {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
            }
            if (!unique.length) {
                ctx.setSummary(label + ': keine gültigen E‑Mails zum Abgleich.', 'warn');
                return;
            }
            ctx.setSummary(
                label +
                    ': Microsoft-365-Abgleich läuft … (' +
                    unique.length +
                    ' Adresse' +
                    (unique.length === 1 ? '' : 'n') +
                    ')',
                'warn'
            );
            const api = ctx.graphApi();
            const token = await api.getGraphToken();
            let byEmail = null;
            if (typeof api.resolveUsersByEmailsBulk === 'function') {
                const bulk = await api.resolveUsersByEmailsBulk(token, unique, function (p) {
                    if (!p) return;
                    if (p.phase === 'directory') {
                        ctx.setSummary(
                            label +
                                ': Verzeichnis laden … ' +
                                (p.loaded || 0) +
                                ' Benutzer · Abgleich ' +
                                unique.length +
                                ' E‑Mails',
                            'warn'
                        );
                    } else if (p.phase === 'lookup') {
                        ctx.setSummary(
                            label + ': Abgleich ' + ((p.done || 0) + 1) + '/' + (p.total || unique.length) + ' …',
                            'warn'
                        );
                    }
                });
                byEmail = bulk && bulk.byEmail ? bulk.byEmail : new Map();
            } else {
                byEmail = new Map();
                for (let i = 0; i < unique.length; i++) {
                    const em = unique[i];
                    ctx.setSummary(label + ': Abgleich ' + (i + 1) + '/' + unique.length + ' …', 'warn');
                    try {
                        const u = await api.resolveUserByEmail(token, em);
                        byEmail.set(em, u && u.id ? u : null);
                    } catch {
                        byEmail.set(em, null);
                    }
                }
            }
            unique.forEach(function (em) {
                const u = byEmail.get(em);
                if (u && u.id) {
                    updates[em] = ctx.directoryMatchUserPayload(u, iso);
                    found++;
                } else {
                    updates[em] = { notFound: true, checkedAt: iso };
                    missed++;
                }
            });
            if (Object.keys(updates).length && window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: updates });
            }
            if (typeof onDone === 'function') onDone();
            ctx.renderStatusOverview();
            ctx.setSummary(
                label + ': ' + found + ' gefunden, ' + missed + ' nicht gefunden, ' + skipped + ' ohne gültige E‑Mail',
                missed ? 'warn' : 'ok'
            );
        } catch (e) {
            ctx.setSummary('Microsoft-365-Abgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        }
    }

}
