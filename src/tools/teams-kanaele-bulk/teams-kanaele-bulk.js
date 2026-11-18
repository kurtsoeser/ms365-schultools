(function () {
    'use strict';

    const GRAPH_SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.Read.All',
        'https://graph.microsoft.com/Channel.ReadBasic.All',
        'https://graph.microsoft.com/Channel.Create',
        'https://graph.microsoft.com/ChannelSettings.ReadWrite.All'
    ];

    const GROUP_SELECT = 'id,displayName,mail,mailNickname,resourceProvisioningOptions';
    const STEP_HINTS = {
        1: 'Schritt 1: Teams auswählen.',
        2: 'Schritt 2: Kanalnamen festlegen.',
        3: 'Schritt 3: Vorschau prüfen.',
        4: 'Schritt 4: Ergebnis.'
    };
    const NEXT_LABELS = {
        1: 'Weiter zu Kanälen',
        2: 'Weiter zur Vorschau',
        3: 'Zum Ergebnis',
        4: 'Fertig'
    };

    let msalMod = null;
    let pca = null;
    let wizardStep = 1;
    /** @type {Array<{ id: string, displayName: string, mailNickname: string, mail: string, selected: boolean }>} */
    let bulkResults = [];
    /** @type {Array<{ displayName: string, description: string }>} */
    let channelRows = [];
    /** @type {{ created: number, existed: number, fail: number, total: number, dryRun?: boolean }|null} */
    let lastSummary = null;

    function L() {
        return window.ms365TeamsKanaeleBulkLogic || {};
    }

    function toast(msg) {
        const el = document.getElementById('toast');
        if (el) {
            el.textContent = msg;
            el.classList.add('show');
            clearTimeout(toast._t);
            toast._t = setTimeout(() => el.classList.remove('show'), 3800);
        } else if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else {
            window.alert(msg);
        }
    }

    async function loadMsal() {
        if (msalMod) return msalMod;
        try {
            msalMod = await import('https://esm.sh/@azure/msal-browser@3.26.1');
        } catch {
            msalMod = await import('https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.26.1/+esm');
        }
        return msalMod;
    }

    function isInteractionRequired(e) {
        return (
            e &&
            (e.name === 'InteractionRequiredAuthError' ||
                e.errorCode === 'interaction_required' ||
                (typeof e.message === 'string' && e.message.indexOf('interaction_required') !== -1))
        );
    }

    function resolveMsalConfig() {
        let cfg = window.MS365_MSAL_CONFIG;
        if (!cfg) cfg = {};
        let id = String(cfg.clientId || '').trim();
        if (!id) {
            const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
            const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
            if (fromMeta) id = fromMeta;
        }
        if (!id) {
            throw new Error(
                'Keine clientId: ms365-config.js fehlt/leer oder blockiert. Seite mit Strg+F5 neu laden.'
            );
        }
        return {
            clientId: id,
            authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
            redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
        };
    }

    async function getPca() {
        const m = await loadMsal();
        const PublicClientApplication =
            m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
        if (!PublicClientApplication) {
            throw new Error('MSAL: PublicClientApplication nicht gefunden (Import).');
        }
        const cfg = resolveMsalConfig();
        if (!pca) {
            pca = new PublicClientApplication({
                auth: {
                    clientId: cfg.clientId,
                    authority: cfg.authority,
                    redirectUri: cfg.redirectUri
                },
                cache: {
                    cacheLocation: 'sessionStorage',
                    storeAuthStateInCookie: true
                }
            });
            await pca.initialize();
            await pca.handleRedirectPromise();
        }
        return pca;
    }

    async function getGraphToken() {
        const instance = await getPca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await instance.loginPopup({ scopes: GRAPH_SCOPES, prompt: 'select_account' });
            accounts = instance.getAllAccounts();
        }
        if (!accounts.length) throw new Error('Anmeldung abgebrochen.');
        const req = { scopes: GRAPH_SCOPES, account: accounts[0] };
        try {
            return (await instance.acquireTokenSilent(req)).accessToken;
        } catch (e) {
            if (isInteractionRequired(e)) {
                return (await instance.acquireTokenPopup(req)).accessToken;
            }
            throw e;
        }
    }

    function sleep(ms) {
        return new Promise(function (r) {
            setTimeout(r, ms);
        });
    }

    async function graphRequest(method, path, token, body, extraHeaders) {
        const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
        let attempt = 0;
        while (true) {
            const headers = Object.assign({ Authorization: 'Bearer ' + token }, extraHeaders || {});
            if (body !== undefined) headers['Content-Type'] = 'application/json';
            const res = await fetch(url, {
                method: method,
                headers: headers,
                body: body !== undefined ? JSON.stringify(body) : undefined
            });
            if (res.status === 429 && attempt < 8) {
                const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
                await sleep((isNaN(ra) ? 5 : ra) * 1000);
                attempt++;
                continue;
            }
            return res;
        }
    }

    async function graphJson(method, path, token, body, extraHeaders) {
        const res = await graphRequest(method, path, token, body, extraHeaders);
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
                typeof data === 'object' && data && data.error
                    ? JSON.stringify(data.error)
                    : text || String(res.status);
            throw new Error(method + ' ' + path + ': ' + msg);
        }
        return data || {};
    }

    function appendLog(msg, kind) {
        const el = document.getElementById('tkbLog');
        if (!el) return;
        const prefix =
            kind === 'err' ? '✗ ' : kind === 'ok' ? '✓ ' : kind === 'warn' ? '! ' : '';
        const line = new Date().toLocaleTimeString() + '  ' + prefix + msg;
        el.textContent = (el.textContent ? el.textContent + '\n' : '') + line;
        el.scrollTop = el.scrollHeight;
    }

    function clearLog() {
        const el = document.getElementById('tkbLog');
        if (el) el.textContent = '';
    }

    function groupHasTeamProvisioning(g) {
        const opts = g && g.resourceProvisioningOptions;
        return Array.isArray(opts) && opts.indexOf('Team') !== -1;
    }

    async function listTeamGroupsFromGraphFallback(token) {
        const collected = [];
        let nextPath = '/groups?$select=' + GROUP_SELECT + '&$top=999';
        while (nextPath) {
            const data = await graphJson('GET', nextPath, token, undefined);
            (data.value || []).forEach(function (g) {
                if (g && g.id && g.mailNickname && groupHasTeamProvisioning(g)) collected.push(g);
            });
            nextPath = data['@odata.nextLink'] || null;
        }
        return collected;
    }

    async function listTeamGroupsFromGraph(token) {
        const collected = [];
        let nextPath =
            "/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=" +
            GROUP_SELECT +
            '&$top=999';
        try {
            while (nextPath) {
                const useEv = nextPath.indexOf('http') !== 0;
                const data = await graphJson(
                    'GET',
                    nextPath,
                    token,
                    undefined,
                    useEv ? { ConsistencyLevel: 'eventual' } : undefined
                );
                (data.value || []).forEach(function (g) {
                    if (g && g.id && g.mailNickname) collected.push(g);
                });
                nextPath = data['@odata.nextLink'] || null;
            }
            return collected;
        } catch {
            return listTeamGroupsFromGraphFallback(token);
        }
    }

    async function listChannels(token, teamId) {
        const out = [];
        let next =
            '/teams/' +
            encodeURIComponent(teamId) +
            '/channels?$select=id,displayName,membershipType,isFavoriteByDefault';
        while (next) {
            const data = await graphJson('GET', next, token, undefined);
            (data.value || []).forEach(function (c) {
                if (c && c.id) {
                    out.push({
                        id: String(c.id),
                        displayName: String(c.displayName || ''),
                        membershipType: String(c.membershipType || 'standard'),
                        isFavoriteByDefault: !!c.isFavoriteByDefault
                    });
                }
            });
            next = data['@odata.nextLink'] || null;
        }
        return out;
    }

    async function createChannel(token, teamId, spec, recommend) {
        const body = {
            displayName: String(spec.displayName || '').trim().slice(0, 50),
            description: String(spec.description || '').trim(),
            membershipType: 'standard'
        };
        if (recommend) body.isFavoriteByDefault = true;
        return graphJson(
            'POST',
            '/teams/' + encodeURIComponent(teamId) + '/channels',
            token,
            body
        );
    }

    async function trySetFavorite(token, teamId, channelId) {
        try {
            await graphJson(
                'PATCH',
                '/teams/' +
                    encodeURIComponent(teamId) +
                    '/channels/' +
                    encodeURIComponent(channelId),
                token,
                { isFavoriteByDefault: true }
            );
            return true;
        } catch {
            return false;
        }
    }

    function selectedTeams() {
        return bulkResults.filter(function (r) {
            return r.selected;
        });
    }

    function channelSpecs() {
        readChannelEditorIntoState();
        const logic = L();
        return typeof logic.normalizeChannelSpecs === 'function'
            ? logic.normalizeChannelSpecs(channelRows)
            : [];
    }

    function namingStyle() {
        const el = document.getElementById('tkbNamingStyle');
        return el && el.value === 'space' ? 'space' : 'hyphen';
    }

    function setNamingStyle(style) {
        const value = style === 'space' ? 'space' : 'hyphen';
        const hidden = document.getElementById('tkbNamingStyle');
        if (hidden) hidden.value = value;
        const hyphen = document.getElementById('tkbStyleHyphen');
        const space = document.getElementById('tkbStyleSpace');
        if (hyphen) hyphen.classList.toggle('is-active', value === 'hyphen');
        if (space) space.classList.toggle('is-active', value === 'space');
    }

    function recommendEnabled() {
        const el = document.getElementById('tkbRecommend');
        return !!(el && el.checked);
    }

    function defaultDescription() {
        const el = document.getElementById('tkbDescription');
        return el ? String(el.value || '').trim() : '';
    }

    function channelBadge(displayName, idx) {
        const m = String(displayName || '').match(/^(\d{1,2})/);
        if (m) return m[1].padStart(2, '0');
        return String(idx + 1).padStart(2, '0');
    }

    function applyPreset() {
        const logic = L();
        const preset =
            typeof logic.buildFachschaftPresetChannels === 'function'
                ? logic.buildFachschaftPresetChannels(namingStyle())
                : [];
        const desc = defaultDescription();
        channelRows = preset.map(function (p) {
            return { displayName: p.displayName, description: desc };
        });
        renderChannelEditor();
        updateGlanceValues();
    }

    function readChannelEditorIntoState() {
        const body = document.getElementById('tkbChannelBody');
        if (!body) return;
        const rows = [];
        body.querySelectorAll('[data-ch-row]').forEach(function (card) {
            const nameInp = card.querySelector('input[data-ch-name]');
            const descInp = card.querySelector('input[data-ch-desc]');
            rows.push({
                displayName: nameInp ? nameInp.value : '',
                description: descInp ? descInp.value : ''
            });
        });
        channelRows = rows;
    }

    function renderChannelEditor() {
        const body = document.getElementById('tkbChannelBody');
        if (!body) return;
        body.replaceChildren();
        if (!channelRows.length) {
            channelRows = [{ displayName: '', description: defaultDescription() }];
        }
        channelRows.forEach(function (row, idx) {
            const card = document.createElement('div');
            card.className = 'tkb-channel-card';
            card.setAttribute('data-ch-row', String(idx));

            const num = document.createElement('div');
            num.className = 'tkb-channel-card__num';
            num.textContent = channelBadge(row.displayName, idx);

            const fields = document.createElement('div');
            fields.className = 'tkb-channel-card__fields';

            const nameWrap = document.createElement('div');
            const nameLab = document.createElement('label');
            nameLab.textContent = 'Kanalname';
            nameLab.setAttribute('for', 'tkbChName' + idx);
            const nameInp = document.createElement('input');
            nameInp.type = 'text';
            nameInp.id = 'tkbChName' + idx;
            nameInp.setAttribute('data-ch-name', '1');
            nameInp.value = row.displayName || '';
            nameInp.placeholder = 'z. B. 01-Fachkonferenz';
            nameInp.maxLength = 50;
            nameInp.addEventListener('input', function () {
                num.textContent = channelBadge(nameInp.value, idx);
                updateChannelPreview();
                updateGlanceValues();
            });
            nameInp.addEventListener('change', readChannelEditorIntoState);
            nameWrap.appendChild(nameLab);
            nameWrap.appendChild(nameInp);

            const descWrap = document.createElement('div');
            const descLab = document.createElement('label');
            descLab.textContent = 'Beschreibung';
            descLab.setAttribute('for', 'tkbChDesc' + idx);
            const descInp = document.createElement('input');
            descInp.type = 'text';
            descInp.id = 'tkbChDesc' + idx;
            descInp.setAttribute('data-ch-desc', '1');
            descInp.value = row.description || '';
            descInp.placeholder = 'optional';
            descInp.maxLength = 1024;
            descInp.addEventListener('change', readChannelEditorIntoState);
            descWrap.appendChild(descLab);
            descWrap.appendChild(descInp);

            fields.appendChild(nameWrap);
            fields.appendChild(descWrap);

            const btnDel = document.createElement('button');
            btnDel.type = 'button';
            btnDel.className = 'tkb-channel-card__remove';
            btnDel.title = 'Kanal entfernen';
            btnDel.innerHTML = '<i class="bi bi-trash"></i>';
            btnDel.addEventListener('click', function () {
                readChannelEditorIntoState();
                channelRows.splice(idx, 1);
                renderChannelEditor();
                updateGlanceValues();
            });

            card.appendChild(num);
            card.appendChild(fields);
            card.appendChild(btnDel);
            body.appendChild(card);
        });
        updateChannelPreview();
    }

    function updateChannelPreview() {
        const el = document.getElementById('tkbChannelPreview');
        if (!el) return;
        const specs = channelSpecs();
        if (!specs.length) {
            el.textContent = 'Noch keine gültigen Kanalnamen.';
            return;
        }
        el.textContent =
            specs.length +
            ' Kanal(e): ' +
            specs
                .map(function (s) {
                    return s.displayName;
                })
                .join(' · ');
    }

    function renderBulkResults() {
        const wrap = document.getElementById('tkbBulkResultsWrap');
        const body = document.getElementById('tkbBulkResultsBody');
        const summary = document.getElementById('tkbBulkResultsSummary');
        const empty = document.getElementById('tkbTeamsEmpty');
        if (!wrap || !body) return;

        if (!bulkResults.length) {
            wrap.hidden = true;
            body.replaceChildren();
            if (empty) {
                empty.hidden = false;
                empty.textContent = 'Keine Treffer – Suchbegriff anpassen und erneut suchen.';
            }
            updateGlanceValues();
            return;
        }

        wrap.hidden = false;
        if (empty) empty.hidden = true;
        const nSel = selectedTeams().length;
        if (summary) {
            summary.textContent = bulkResults.length + ' gefunden · ' + nSel + ' ausgewählt';
        }

        body.replaceChildren();
        bulkResults.forEach(function (r, idx) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'tkb-team-card' + (r.selected ? ' is-on' : '');
            btn.setAttribute('role', 'listitem');
            btn.setAttribute('aria-pressed', r.selected ? 'true' : 'false');

            const check = document.createElement('input');
            check.type = 'checkbox';
            check.className = 'tkb-team-card__check';
            check.checked = !!r.selected;
            check.tabIndex = -1;
            check.addEventListener('click', function (ev) {
                ev.stopPropagation();
            });
            check.addEventListener('change', function () {
                bulkResults[idx].selected = check.checked;
                renderBulkResults();
            });

            const meta = document.createElement('span');
            const name = document.createElement('span');
            name.className = 'tkb-team-card__name';
            name.textContent = r.displayName || '—';
            const mail = document.createElement('span');
            mail.className = 'tkb-team-card__mail';
            mail.textContent = r.mail || r.mailNickname || '—';
            meta.appendChild(name);
            meta.appendChild(mail);

            btn.appendChild(check);
            btn.appendChild(meta);
            btn.addEventListener('click', function () {
                bulkResults[idx].selected = !bulkResults[idx].selected;
                renderBulkResults();
            });
            body.appendChild(btn);
        });
        updateGlanceValues();
    }

    function setAllBulkSelected(value) {
        bulkResults = bulkResults.map(function (r) {
            return Object.assign({}, r, { selected: value });
        });
        renderBulkResults();
    }

    function fillChipList(el, items, icon) {
        if (!el) return;
        el.replaceChildren();
        if (!items.length) {
            const li = document.createElement('li');
            li.className = 'muted';
            li.textContent = '—';
            el.appendChild(li);
            return;
        }
        items.forEach(function (text) {
            const li = document.createElement('li');
            li.className = 'tkb-chip';
            li.innerHTML = '<i class="bi ' + icon + '" aria-hidden="true"></i>';
            const span = document.createElement('span');
            span.textContent = text;
            li.appendChild(span);
            el.appendChild(li);
        });
    }

    function renderPreview() {
        const teams = selectedTeams();
        const specs = channelSpecs();
        const teamCount = document.getElementById('tkbPreviewTeamCount');
        const channelCount = document.getElementById('tkbPreviewChannelCount');
        const teamHint = document.getElementById('tkbPreviewTeamHint');
        const recHint = document.getElementById('tkbPreviewRecommendHint');
        const teamListHint = document.getElementById('tkbPreviewTeamListHint');

        if (teamCount) teamCount.textContent = String(teams.length);
        if (channelCount) channelCount.textContent = String(specs.length);
        if (teamHint) {
            teamHint.textContent = teams.length
                ? teams.length === 1
                    ? '1 Team ausgewählt'
                    : teams.length + ' Teams ausgewählt'
                : 'noch keine Auswahl';
        }
        if (recHint) {
            recHint.textContent = recommendEnabled()
                ? 'Empfehlung in Kanalliste: an'
                : 'Empfehlung in Kanalliste: aus';
        }
        if (teamListHint) {
            teamListHint.textContent = teams.length
                ? 'Diese Teams erhalten die Kanäle (fehlende Namen).'
                : 'Bitte in Schritt 1 mindestens ein Team wählen.';
        }

        fillChipList(
            document.getElementById('tkbPreviewChannelChips'),
            specs.map(function (s) {
                return s.displayName;
            }),
            'bi-hash'
        );
        fillChipList(
            document.getElementById('tkbPreviewTeamChips'),
            teams.map(function (t) {
                return t.displayName || t.mailNickname || t.id;
            }),
            'bi-people'
        );
    }

    function renderResultSummary(summary, dryRun) {
        lastSummary = summary
            ? Object.assign({}, summary, { dryRun: !!dryRun })
            : null;
        const wrap = document.getElementById('tkbResultSummary');
        const lead = document.getElementById('tkbResultLead');
        if (!summary) {
            if (wrap) wrap.hidden = true;
            if (lead) {
                lead.textContent =
                    'Noch kein Lauf – nach der Vorschau „Jetzt anlegen“ oder Dry-Run starten.';
            }
            updateGlanceValues();
            return;
        }
        if (wrap) wrap.hidden = false;
        const c = document.getElementById('tkbStatCreated');
        const e = document.getElementById('tkbStatExisted');
        const f = document.getElementById('tkbStatFail');
        if (c) c.textContent = String(summary.created);
        if (e) e.textContent = String(summary.existed);
        if (f) f.textContent = String(summary.fail);
        if (lead) {
            lead.textContent = dryRun
                ? 'Dry-Run abgeschlossen – es wurde nichts angelegt.'
                : 'Anlage abgeschlossen.';
        }
        updateGlanceValues();
    }

    function updateGlanceValues() {
        const g1 = document.getElementById('tkbGlance1');
        const g2 = document.getElementById('tkbGlance2');
        const g3 = document.getElementById('tkbGlance3');
        const g4 = document.getElementById('tkbGlance4');
        const nTeams = selectedTeams().length;
        const nCh = channelSpecs().length;
        if (g1) g1.textContent = nTeams ? nTeams + ' gewählt' : 'Auswählen';
        if (g2) g2.textContent = nCh ? nCh + ' Kanäle' : 'Festlegen';
        if (g3) g3.textContent = nTeams && nCh ? 'Bereit' : 'Prüfen';
        if (g4) {
            if (!lastSummary) g4.textContent = 'Offen';
            else if (lastSummary.fail) g4.textContent = 'Mit Fehlern';
            else g4.textContent = lastSummary.dryRun ? 'Dry-Run ok' : 'Erledigt';
        }

        const tab1 = document.getElementById('tkbTab1');
        const tab2 = document.getElementById('tkbTab2');
        const tab3 = document.getElementById('tkbTab3');
        const tab4 = document.getElementById('tkbTab4');
        if (tab1) tab1.classList.toggle('is-ok', nTeams > 0);
        if (tab2) tab2.classList.toggle('is-ok', nCh > 0);
        if (tab3) tab3.classList.toggle('is-ok', nTeams > 0 && nCh > 0);
        if (tab4) {
            tab4.classList.toggle('is-ok', !!(lastSummary && !lastSummary.fail));
            tab4.classList.toggle('is-warn', !!(lastSummary && lastSummary.fail));
        }
    }

    function setWizardStep(step) {
        const next = Math.max(1, Math.min(4, Number(step) || 1));
        if (next === 2) applyPresetIfEmpty();
        if (next === 3) renderPreview();
        wizardStep = next;

        for (let i = 1; i <= 4; i++) {
            const pane = document.getElementById('tkbStep' + i);
            const tab = document.getElementById('tkbTab' + i);
            const active = i === wizardStep;
            if (pane) pane.hidden = !active;
            if (tab) {
                tab.classList.toggle('is-active', active);
                tab.setAttribute('aria-selected', active ? 'true' : 'false');
            }
        }

        const hint = document.getElementById('tkbPhaseHint');
        if (hint) hint.textContent = STEP_HINTS[wizardStep] || '';

        const back = document.getElementById('tkbWizardBack');
        const fwd = document.getElementById('tkbWizardNext');
        if (back) back.disabled = wizardStep <= 1;
        if (fwd) {
            fwd.disabled = wizardStep >= 4;
            fwd.innerHTML =
                wizardStep >= 4
                    ? '<i class="bi bi-check2"></i>Fertig'
                    : '<i class="bi bi-arrow-right"></i>' + (NEXT_LABELS[wizardStep] || 'Weiter');
        }
        updateGlanceValues();
    }

    function applyPresetIfEmpty() {
        const specs = channelSpecs();
        if (!specs.length) applyPreset();
    }

    function canLeaveStep(from, to) {
        if (to < from) return true;
        if (from === 1 && to > 1) {
            if (!selectedTeams().length) {
                toast('Bitte mindestens ein Team auswählen.');
                return false;
            }
        }
        if (from === 2 && to > 2) {
            if (!channelSpecs().length) {
                toast('Bitte mindestens einen Kanalnamen eintragen.');
                return false;
            }
        }
        return true;
    }

    async function onBulkSearch() {
        const inp = document.getElementById('tkbBulkPrefix');
        const term = inp && inp.value ? inp.value.trim() : '';
        if (!term) {
            toast('Bitte Suchbegriff eintragen (z. B. Fachschaft, ARGE, Physik).');
            return;
        }
        const btn = document.getElementById('tkbBulkSearch');
        if (btn) btn.disabled = true;
        bulkResults = [];
        renderBulkResults();
        try {
            const token = await getGraphToken();
            const groups = await listTeamGroupsFromGraph(token);
            const logic = L();
            const matched =
                typeof logic.filterAndSortGroupsByTerm === 'function'
                    ? logic.filterAndSortGroupsByTerm(groups, term)
                    : [];
            bulkResults = matched.map(function (g) {
                return {
                    id: g.id,
                    displayName: g.displayName,
                    mailNickname: g.mailNickname,
                    mail: g.mail,
                    selected: true
                };
            });
            renderBulkResults();
            toast(bulkResults.length + ' Team(s) gefunden.');
        } catch (e) {
            toast(String((e && e.message) || e));
            const empty = document.getElementById('tkbTeamsEmpty');
            if (empty) {
                empty.hidden = false;
                empty.textContent = 'Fehler: ' + ((e && e.message) || e);
            }
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function runBulkCreate(dryRun) {
        const logic = L();
        const specs = channelSpecs();
        if (!specs.length) {
            toast('Bitte mindestens einen Kanalnamen eintragen.');
            setWizardStep(2);
            return;
        }
        const selected = selectedTeams();
        if (!selected.length) {
            toast('Bitte zuerst Teams suchen und mindestens eines auswählen.');
            setWizardStep(1);
            return;
        }

        const recommend = recommendEnabled();
        if (!dryRun) {
            const question =
                'In ' +
                selected.length +
                ' Team(s) die Kanäle anlegen?\n\n' +
                specs
                    .map(function (s) {
                        return '• ' + s.displayName;
                    })
                    .join('\n') +
                (recommend ? '\n\nEmpfehlung „in Kanalliste anzeigen“ ist aktiv.' : '');
            const confirmed =
                typeof window.ms365AppDialogConfirm === 'function'
                    ? await window.ms365AppDialogConfirm(question, {
                          title: 'Kanäle bulk anlegen',
                          okText: 'Anlegen',
                          danger: false
                      })
                    : window.confirm(question);
            if (!confirmed) return;
        }

        setWizardStep(4);
        const btnCreate = document.getElementById('tkbBulkCreate');
        const btnDry = document.getElementById('tkbBulkDryRun');
        const btnSearch = document.getElementById('tkbBulkSearch');
        if (btnCreate) btnCreate.disabled = true;
        if (btnDry) btnDry.disabled = true;
        if (btnSearch) btnSearch.disabled = true;

        clearLog();
        renderResultSummary(null);
        appendLog(
            (dryRun ? 'Dry-Run' : 'Anlage') +
                ' für ' +
                selected.length +
                ' Team(s) × ' +
                specs.length +
                ' Kanal(e) …'
        );

        let token;
        try {
            token = await getGraphToken();
        } catch (e) {
            appendLog('Anmeldung/Token: ' + ((e && e.message) || e), 'err');
            if (btnCreate) btnCreate.disabled = false;
            if (btnDry) btnDry.disabled = false;
            if (btnSearch) btnSearch.disabled = false;
            return;
        }

        /** @type {Array<{ ok: boolean, action?: string, skipped?: boolean }>} */
        const allResults = [];

        for (let i = 0; i < selected.length; i++) {
            const team = selected[i];
            const label = team.displayName || team.mailNickname || team.id;
            appendLog('[' + (i + 1) + '/' + selected.length + '] ' + label);

            try {
                const existing = await listChannels(token, team.id);
                const plan =
                    typeof logic.planChannelsForTeam === 'function'
                        ? logic.planChannelsForTeam(specs, existing)
                        : [];

                for (const step of plan) {
                    if (step.action === 'exists') {
                        appendLog('  vorhanden: ' + step.displayName, 'warn');
                        if (!dryRun && recommend && step.existingId) {
                            const okFav = await trySetFavorite(token, team.id, step.existingId);
                            if (okFav) appendLog('  → Empfehlung gesetzt: ' + step.displayName, 'ok');
                        }
                        allResults.push({ ok: true, action: 'exists', skipped: true });
                        continue;
                    }

                    if (dryRun) {
                        appendLog('  würde anlegen: ' + step.displayName, 'ok');
                        allResults.push({ ok: true, action: 'create' });
                        continue;
                    }

                    try {
                        const created = await createChannel(token, team.id, step, recommend);
                        appendLog('  angelegt: ' + step.displayName, 'ok');
                        if (recommend && created && created.id && created.isFavoriteByDefault !== true) {
                            const okFav = await trySetFavorite(token, team.id, created.id);
                            if (okFav) appendLog('  → Empfehlung gesetzt: ' + step.displayName, 'ok');
                            else {
                                appendLog(
                                    '  Hinweis: Empfehlung konnte Graph nicht setzen (UI ggf. manuell).',
                                    'warn'
                                );
                            }
                        }
                        allResults.push({ ok: true, action: 'create' });
                        await sleep(250);
                    } catch (eCreate) {
                        const msg = (eCreate && eCreate.message) || String(eCreate);
                        if (/conflict|already exists|NameAlreadyExists/i.test(msg)) {
                            appendLog('  vorhanden (Konflikt): ' + step.displayName, 'warn');
                            allResults.push({ ok: true, action: 'exists', skipped: true });
                        } else {
                            appendLog('  Fehler ' + step.displayName + ': ' + msg, 'err');
                            allResults.push({ ok: false, action: 'create' });
                        }
                    }
                }
            } catch (eTeam) {
                appendLog('  Team-Fehler: ' + ((eTeam && eTeam.message) || eTeam), 'err');
                allResults.push({ ok: false });
            }
        }

        const summary =
            typeof logic.buildBulkChannelSummary === 'function'
                ? logic.buildBulkChannelSummary(allResults)
                : { total: allResults.length, created: 0, existed: 0, fail: 0 };

        appendLog(
            'Fertig: ' +
                summary.created +
                ' angelegt, ' +
                summary.existed +
                ' bereits vorhanden, ' +
                summary.fail +
                ' Fehler (von ' +
                summary.total +
                ').',
            summary.fail ? 'warn' : 'ok'
        );
        renderResultSummary(summary, dryRun);
        toast(
            dryRun
                ? 'Dry-Run fertig (' + summary.created + ' neu, ' + summary.existed + ' vorhanden).'
                : 'Fertig: ' + summary.created + ' neu, ' + summary.existed + ' vorhanden.'
        );

        if (btnCreate) btnCreate.disabled = false;
        if (btnDry) btnDry.disabled = false;
        if (btnSearch) btnSearch.disabled = false;
    }

    async function onLogin() {
        try {
            await getGraphToken();
            toast('Angemeldet.');
        } catch (e) {
            toast(String((e && e.message) || e));
        }
    }

    function wire() {
        applyPreset();
        setWizardStep(1);

        const btnLogin = document.getElementById('tkbBtnLogin');
        if (btnLogin) btnLogin.addEventListener('click', onLogin);

        document.querySelectorAll('[data-tkb-step]').forEach(function (tab) {
            tab.addEventListener('click', function () {
                const to = Number(tab.getAttribute('data-tkb-step'));
                if (!canLeaveStep(wizardStep, to)) return;
                setWizardStep(to);
            });
        });

        const back = document.getElementById('tkbWizardBack');
        if (back) {
            back.addEventListener('click', function () {
                setWizardStep(wizardStep - 1);
            });
        }
        const fwd = document.getElementById('tkbWizardNext');
        if (fwd) {
            fwd.addEventListener('click', function () {
                if (wizardStep >= 4) return;
                if (!canLeaveStep(wizardStep, wizardStep + 1)) return;
                setWizardStep(wizardStep + 1);
            });
        }

        const btnPreset = document.getElementById('tkbBtnPreset');
        if (btnPreset) btnPreset.addEventListener('click', applyPreset);

        const btnAdd = document.getElementById('tkbBtnAddChannel');
        if (btnAdd) {
            btnAdd.addEventListener('click', function () {
                readChannelEditorIntoState();
                channelRows.push({ displayName: '', description: defaultDescription() });
                renderChannelEditor();
                updateGlanceValues();
            });
        }

        document.querySelectorAll('[data-tkb-style]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                setNamingStyle(btn.getAttribute('data-tkb-style'));
                applyPreset();
            });
        });

        const descEl = document.getElementById('tkbDescription');
        if (descEl) {
            descEl.addEventListener('change', function () {
                const d = defaultDescription();
                readChannelEditorIntoState();
                channelRows = channelRows.map(function (r) {
                    return { displayName: r.displayName, description: d || r.description };
                });
                renderChannelEditor();
            });
        }

        const rec = document.getElementById('tkbRecommend');
        if (rec) rec.addEventListener('change', updateGlanceValues);

        const btnSearch = document.getElementById('tkbBulkSearch');
        if (btnSearch) btnSearch.addEventListener('click', onBulkSearch);

        const btnAll = document.getElementById('tkbBulkSelectAll');
        if (btnAll) btnAll.addEventListener('click', function () {
            setAllBulkSelected(true);
        });
        const btnNone = document.getElementById('tkbBulkSelectNone');
        if (btnNone) btnNone.addEventListener('click', function () {
            setAllBulkSelected(false);
        });

        const btnDry = document.getElementById('tkbBulkDryRun');
        if (btnDry) btnDry.addEventListener('click', function () {
            runBulkCreate(true);
        });
        const btnCreate = document.getElementById('tkbBulkCreate');
        if (btnCreate) btnCreate.addEventListener('click', function () {
            runBulkCreate(false);
        });

        const prefix = document.getElementById('tkbBulkPrefix');
        if (prefix) {
            prefix.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    onBulkSearch();
                }
            });
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wire);
    } else {
        wire();
    }
})();
