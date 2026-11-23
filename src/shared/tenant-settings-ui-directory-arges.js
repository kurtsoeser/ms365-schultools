/**
 * tenant-settings-ui-directory-arges.js
 * @param {Record<string, any>} ctx
 */
export function installTenantDirectoryArges(ctx) {
    ctx.argeGroupMailPrefixForDerive = function argeGroupMailPrefixForDerive() {
        const api = window.ms365AppDataV2;
        const su = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        // Explizit leer = kein Präfix (Alias = Kürzel). Sonst Default bzw. mit Bindestrich-Ende.
        if (su && su.argeGroupMailPrefix != null && !ctx.normStr(su.argeGroupMailPrefix)) {
            return '';
        }
        const p = ctx.normStr(su && su.argeGroupMailPrefix) || 'arge-';
        return p.endsWith('-') || p.endsWith('.') ? p : p + '-';
    }

    ctx.argeRowMailNickname = function argeRowMailNickname(row) {
        const prefix = ctx.argeGroupMailPrefixForDerive();
        const code = ctx.normCode(row && row.code) || 'arge';
        let raw = String(prefix + code).toLowerCase();
        try {
            const api = window.ms365GraphUnifiedGroups;
            if (api && typeof api.sanitizeUnifiedGroupMailNickname === 'function') {
                return api.sanitizeUnifiedGroupMailNickname(raw).slice(0, 60);
            }
            if (api && typeof api.sanitizeMailNickname === 'function') {
                return api.sanitizeMailNickname(raw);
            }
        } catch {
            /* ignore */
        }
        return raw.replace(/[@()[\]\\";:<>,\s]/g, '').slice(0, 60) || 'arge';
    }

    ctx.pickArgeGroupMatchFromSearch = function pickArgeGroupMatchFromSearch(groups, row) {
        const code = ctx.normCode(row && row.code);
        const name = ctx.normStr(row && row.name);
        const nickWanted = ctx.argeRowMailNickname(row);
        const sanitize =
            ctx.graphApi() && typeof ctx.graphApi().sanitizeMailNickname === 'function'
                ? ctx.graphApi().sanitizeMailNickname
                : function (v) {
                      return String(v || '')
                          .replace(/[^0-9a-zA-Z]/g, '')
                          .toLowerCase();
                  };
        const codeNick = sanitize(code);
        const list = Array.isArray(groups) ? groups : [];
        for (let i = 0; i < list.length; i++) {
            const g = list[i] || {};
            const dn = ctx.normStr(g.displayName).toLowerCase();
            const nick = ctx.normStr(g.mailNickname).toLowerCase();
            if (nick && nickWanted && nick === nickWanted.toLowerCase()) return g;
            if (nick && codeNick && (nick === codeNick || nick.endsWith(codeNick))) return g;
            if (name && dn === name.toLowerCase()) return g;
            if (code && dn.indexOf(code.toLowerCase()) !== -1) return g;
        }
        return null;
    }

    ctx.verifyArgeGroupForRow = async function verifyArgeGroupForRow(row, opts) {
        const code = ctx.normCode(row && row.code);
        if (!code) return { skipped: true };
        const token = opts && opts.token ? opts.token : await ctx.graphApi().getGraphToken();
        const prefix = ctx.argeGroupMailPrefixForDerive();
        const queries = [ctx.argeRowMailNickname(row), prefix + code.toLowerCase(), code, ctx.normStr(row.name)].filter(Boolean);
        const seen = new Set();
        const groups = [];
        for (let i = 0; i < queries.length; i++) {
            const q = queries[i];
            let hits = [];
            if (typeof ctx.graphApi().searchUnifiedGroups === 'function') {
                hits = await ctx.graphApi().searchUnifiedGroups(token, q);
            } else {
                const data = await ctx.graphApi().graphJson(
                    'GET',
                    '/groups?$filter=' +
                        encodeURIComponent("groupTypes/any(c:c eq 'Unified') and startswith(displayName,'" + String(q).replace(/'/g, "''") + "')") +
                        '&$select=' +
                        encodeURIComponent('id,displayName,mail,mailNickname,groupTypes') +
                        '&$top=25',
                    token,
                    undefined
                );
                hits = Array.isArray(data && data.value) ? data.value : [];
            }
            hits.forEach(function (g) {
                const id = ctx.normStr(g && g.id);
                if (!id || seen.has(id)) return;
                seen.add(id);
                groups.push(g);
            });
        }
        const match = ctx.pickArgeGroupMatchFromSearch(groups, row);
        if (match && match.id) {
            ctx.upsertArgeCatalogLink(code, match, '');
            return { found: true, skipped: false };
        }
        ctx.patchArgeCatalogNotFound(code);
        return { found: false, skipped: false };
    }

    ctx.createArgeGroupForRow = async function createArgeGroupForRow(row) {
        const code = ctx.normCode(row && row.code);
        if (!code) return { skipped: true, reason: 'no-code' };
        const link = ctx.getArgeCatalogLink(code);
        if (link && ctx.normStr(link.graphGroupId)) {
            return { skipped: true, reason: 'already-linked' };
        }
        const nick = ctx.argeRowMailNickname(row);
        if (!nick) {
            return { error: 'Kein Mail-Nickname ableitbar – Kürzel prüfen (ARGE-Präfix in Stammdaten).' };
        }
        const displayName = ctx.normStr(row.name) || 'ARGE ' + code;
        const desc = 'ARGE ' + displayName + ' (MS365-Schul-Tools)';
        const token = await ctx.graphApi().getGraphToken();
        const created = await ctx.graphApi().createUnifiedGroup(token, displayName, nick, desc);
        if (!created || !created.id) {
            return { error: 'Microsoft 365 hat keine Gruppe zurückgegeben.' };
        }
        const headEm = ctx.normStr(row.headEmail).toLowerCase();
        if (headEm.indexOf('@') !== -1 && typeof ctx.graphApi().ensureOwners === 'function') {
            try {
                await ctx.graphApi().ensureOwners(token, created.id, [headEm]);
            } catch {
                /* optional */
            }
        }
        ctx.upsertArgeCatalogLink(code, created, '');
        return { created: true, group: created, mailNickname: nick, displayName };
    }

    ctx.createAdminEntraUserInteractive = async function createAdminEntraUserInteractive(emailRaw, nameHint) {
        const em = ctx.normStr(emailRaw).toLowerCase();
        if (!em || em.indexOf('@') === -1) {
            ctx.setSummary('Bitte zuerst eine gültige E-Mail eintragen.', 'warn');
            return;
        }
        const dom = em.split('@')[1] || '';
        const upn = await ctx.dlgPrompt('Benutzer-Principalname (UPN), z. B. vorname.nachname@' + dom + ':', em, {
            title: 'Entra-Benutzer',
            inputLabel: 'UPN'
        });
        if (upn == null || !ctx.normStr(upn)) return;
        const displayName = await ctx.dlgPrompt(
            'Anzeigename in Microsoft 365:',
            ctx.normStr(nameHint) || ctx.normStr(String(upn).split('@')[0]),
            { title: 'Entra-Benutzer', inputLabel: 'Anzeigename' }
        );
        if (displayName == null || !ctx.normStr(displayName)) return;
        try {
            const tokenProbe = await ctx.graphApi().getGraphToken();
            const existing = await ctx.graphApi().resolveUserByEmail(tokenProbe, em);
            if (existing && existing.id) {
                ctx.patchDirectoryMatchKeys([em], ctx.directoryMatchUserPayload(existing));
                ctx.refreshAdminAudienceBoardUi();
                ctx.setSummary('Unter dieser E-Mail existiert bereits ein Entra-Benutzer.', 'ok');
                return;
            }
        } catch {
            // weiter mit Anlage
        }
        const mailNick = ctx.mailNicknameFromUpn(upn);
        const ok = await ctx.dlgConfirm(
            'Benutzer in Entra ID anlegen?\n\nAnzeigename: ' +
                displayName +
                '\nUPN: ' +
                upn +
                '\nMail-Nickname: ' +
                mailNick +
                '\n\nEs wird ein temporäres Kennwort gesetzt (Wechsel beim ersten Anmelden).',
            { title: 'Entra-Benutzer anlegen' }
        );
        if (!ok) return;
        try {
            const token = await ctx.graphApi().getGraphToken();
            const created = await ctx.createEntraUserViaGraph(token, upn, displayName);
            const iso = new Date().toISOString();
            ctx.patchDirectoryMatchKeys(
                [em, created.upn],
                ctx.directoryMatchUserPayload(
                    { id: created.id, displayName: created.displayName, userPrincipalName: created.upn },
                    iso
                )
            );
            ctx.refreshAdminAudienceBoardUi();
            await ctx.dlgAlert(
                'Benutzer angelegt.\n\nEinmaliges Kennwort:\n' +
                    created.password +
                    '\n\n(Bitte sicher übergeben / in Entra ändern.)',
                { title: 'Kennwort notieren', okText: 'Verstanden' }
            );
            ctx.setSummary('Entra-Benutzer angelegt: ' + created.displayName, 'ok');
        } catch (e) {
            ctx.setSummary('Benutzer anlegen: ' + (e && e.message ? e.message : String(e)), 'warn');
        }
    }

}
