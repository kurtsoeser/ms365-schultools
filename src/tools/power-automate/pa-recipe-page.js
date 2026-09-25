(function () {
    'use strict';

    const G = window.ms365SpoGraph;
    const catalog = window.ms365PaRecipes;
    if (!catalog) return;

    const SCOPES = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Sites.ReadWrite.All'
    ];

    function $(id) {
        return document.getElementById(id);
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    function recipeIdFromPage() {
        const body = document.body;
        return (body && body.getAttribute('data-pa-recipe')) || '';
    }

    function loadCfg(key) {
        try {
            return JSON.parse(localStorage.getItem(key) || '{}') || {};
        } catch (e) {
            return {};
        }
    }

    function saveCfg(key, cfg) {
        try {
            localStorage.setItem(key, JSON.stringify(cfg));
        } catch (e) {}
        if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.notifyLocalDataChanged === 'function') {
            window.ms365BrowserBackup.notifyLocalDataChanged('power-automate', { key: key });
        }
    }

    function doneKey(recipe) {
        return 'ms365-pa-done-' + recipe.id;
    }

    function isDone(recipe) {
        try {
            return localStorage.getItem(doneKey(recipe)) === '1';
        } catch (e) {
            return false;
        }
    }

    function setDone(recipe, on) {
        try {
            if (on) localStorage.setItem(doneKey(recipe), '1');
            else localStorage.removeItem(doneKey(recipe));
        } catch (e) {}
        // Sync with old collective recipes key
        try {
            const KEY = 'ms365-power-automate-recipes-v1';
            const s = JSON.parse(localStorage.getItem(KEY) || '{}') || {};
            const map = {
                'termine-sync': 'termine',
                antraege: 'formulare',
                seminar: 'seminar',
                'gast-erinnerung': 'gast',
                'diplom-ordner': 'dipl',
                schilf: 'schilf'
            };
            const old = map[recipe.id];
            if (old) {
                s[old] = !!on;
                localStorage.setItem(KEY, JSON.stringify(s));
            }
        } catch (e2) {}
        if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.notifyLocalDataChanged === 'function') {
            window.ms365BrowserBackup.notifyLocalDataChanged('power-automate-done', { recipeId: recipe && recipe.id });
        }
    }

    function log(msg) {
        const el = $('paLog');
        if (!el) return;
        el.textContent += (el.textContent ? '\n' : '') + msg;
        el.scrollTop = el.scrollHeight;
        const details = el.closest('details');
        if (details) details.open = true;
    }

    function escapeHtml(s) {
        return String(s || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function renderShell(recipe) {
        const nameEl = document.querySelector('.header-tool-indicator__name');
        if (nameEl) {
            nameEl.innerHTML = '<i class="bi ' + escapeHtml(recipe.icon) + '"></i>' + escapeHtml(recipe.title);
        }
        const help = document.querySelector('.header-help-link');
        if (help && recipe.helpId) help.setAttribute('href', '../hilfe.html#' + recipe.helpId);

        const heroKicker = $('paKicker');
        const heroTitle = $('paTitle');
        const heroSummary = $('paSummary');
        if (heroKicker) heroKicker.textContent = recipe.kicker || 'Power Automate';
        if (heroTitle) heroTitle.textContent = recipe.title;
        if (heroSummary) heroSummary.textContent = recipe.summary || '';

        const related = $('paRelated');
        if (related) {
            related.innerHTML = (recipe.related || [])
                .map(function (r) {
                    const ext = r.external
                        ? ' target="_blank" rel="noopener"'
                        : '';
                    return (
                        '<a class="btn" href="' +
                        escapeHtml(r.href) +
                        '"' +
                        ext +
                        '><i class="bi ' +
                        escapeHtml(r.icon || 'bi-link-45deg') +
                        '"></i>' +
                        escapeHtml(r.label) +
                        '</a>'
                    );
                })
                .join('');
        }

        const fieldsHost = $('paFields');
        if (fieldsHost) {
            fieldsHost.innerHTML = (recipe.fields || [])
                .map(function (f) {
                    const type = f.type === 'number' ? 'number' : f.type === 'email' ? 'email' : f.type === 'url' ? 'url' : 'text';
                    return (
                        '<div class="tm-field">' +
                        '<label for="paField_' +
                        escapeHtml(f.id) +
                        '">' +
                        escapeHtml(f.label) +
                        '</label>' +
                        '<input type="' +
                        type +
                        '" id="paField_' +
                        escapeHtml(f.id) +
                        '" data-pa-field="' +
                        escapeHtml(f.id) +
                        '" placeholder="' +
                        escapeHtml(f.placeholder || '') +
                        '" autocomplete="off" spellcheck="false">' +
                        '</div>'
                    );
                })
                .join('');
        }

        const hint = $('paListHint');
        if (hint) {
            if (recipe.listHint) {
                hint.hidden = false;
                hint.textContent = recipe.listHint;
            } else {
                hint.hidden = true;
            }
        }

        const btnList = $('paBtnList');
        if (btnList) {
            if (recipe.list && recipe.listActionLabel) {
                btnList.hidden = false;
                btnList.innerHTML = '<i class="bi bi-list-ul"></i>' + escapeHtml(recipe.listActionLabel);
            } else {
                btnList.hidden = true;
            }
        }

        const stepsHost = $('paSteps');
        if (stepsHost) {
            const items = recipe.checklist && recipe.checklist.length ? recipe.checklist : recipe.steps || [];
            const isCheck = !!(recipe.checklist && recipe.checklist.length);
            stepsHost.innerHTML = items
                .map(function (s, i) {
                    if (isCheck) {
                        return (
                            '<label class="checkbox-label" style="display:block;margin:6px 0;">' +
                            '<input type="checkbox" data-pa-check-step="' +
                            i +
                            '"> ' +
                            escapeHtml(s) +
                            '</label>'
                        );
                    }
                    return '<li>' + escapeHtml(s) + '</li>';
                })
                .join('');
            if (!isCheck) {
                const ol = document.createElement('ol');
                ol.style.margin = '0';
                ol.style.paddingLeft = '1.2em';
                ol.style.lineHeight = '1.55';
                ol.innerHTML = stepsHost.innerHTML;
                stepsHost.innerHTML = '';
                stepsHost.appendChild(ol);
            }
        }

        const doneLabel = $('paDoneLabel');
        if (doneLabel) doneLabel.textContent = recipe.doneLabel || 'Umgesetzt';
        const doneBox = $('paDone');
        if (doneBox) doneBox.checked = isDone(recipe);

        document.title = 'MS365-Schul-Tools – ' + recipe.title;
    }

    function readFields(recipe) {
        const out = {};
        (recipe.fields || []).forEach(function (f) {
            const el = document.querySelector('[data-pa-field="' + f.id + '"]');
            let v = el ? String(el.value || '').trim() : '';
            if (f.type === 'url') v = v.replace(/\/$/, '');
            if (f.type === 'email') v = v.toLowerCase();
            out[f.id] = v;
        });
        return out;
    }

    function writeFields(recipe, cfg) {
        (recipe.fields || []).forEach(function (f) {
            const el = document.querySelector('[data-pa-field="' + f.id + '"]');
            if (!el) return;
            if (cfg && cfg[f.id] != null && cfg[f.id] !== '') el.value = cfg[f.id];
            else if (f.defaultValue != null) el.value = f.defaultValue;
        });
    }

    async function ensureToken() {
        if (!G) throw new Error('SharePoint-Graph-Helfer nicht geladen.');
        return await G.getGraphToken(SCOPES);
    }

    async function findListByTitle(token, siteId, listTitle) {
        const title = String(listTitle || '').trim();
        const path =
            G.graphPathSite(siteId) +
            '/lists?$filter=' +
            encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
            '&$select=id,displayName,webUrl';
        const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
        return ((data && data.value) || [])[0] || null;
    }

    async function addColumns(siteId, listId, token, columns) {
        const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        for (let i = 0; i < columns.length; i++) {
            await G.graphJson('POST', base, token, columns[i], 'v1.0');
            await G.sleep(140);
        }
    }

    async function createList(recipe, cfg) {
        if (!recipe.list) throw new Error('Dieses Rezept braucht keine Liste.');
        const siteUrl = cfg.siteUrl;
        const listName = cfg.listName || recipe.list.defaultName;
        if (!siteUrl) throw new Error('Bitte die SharePoint-Website eintragen.');

        const token = await ensureToken();
        log('Löse Website auf …');
        const site = await G.resolveSiteFromWebUrl(token, siteUrl);
        const siteId = site && site.id ? String(site.id) : '';
        if (!siteId) throw new Error('Site-ID fehlt.');
        log('Site: ' + (site.displayName || siteId));

        log('Suche Liste „' + listName + '" …');
        const found = await findListByTitle(token, siteId, listName);
        if (found && found.id) {
            log('Bereits vorhanden – ID: ' + found.id + (found.webUrl ? ' · ' + found.webUrl : ''));
            return found;
        }

        log('Erstelle Liste „' + listName + '" …');
        const created = await G.graphJson(
            'POST',
            G.graphPathSite(siteId) + '/lists',
            token,
            {
                displayName: listName,
                description: recipe.list.description || '',
                list: { template: 'genericList' }
            },
            'v1.0'
        );
        const listId = created && created.id ? String(created.id) : '';
        if (!listId) throw new Error('Listen-ID fehlt.');
        log('Liste angelegt: ' + listId);
        log('Füge Spalten hinzu …');
        await addColumns(siteId, listId, token, recipe.list.columns || []);
        log('Fertig.' + (created.webUrl ? ' URL: ' + created.webUrl : ''));

        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: recipe.toolId,
                action: 'create-list',
                target: siteUrl,
                summary: listName + ' (' + listId + ')'
            });
        }
        return created;
    }

    function wire(recipe) {
        renderShell(recipe);
        const cfg = loadCfg(recipe.storageKey);
        writeFields(recipe, cfg);

        // checklist step persistence
        const stepCfg = cfg._steps || {};
        document.querySelectorAll('[data-pa-check-step]').forEach(function (el) {
            const i = el.getAttribute('data-pa-check-step');
            el.checked = !!stepCfg[i];
            el.addEventListener('change', function () {
                const c = readFields(recipe);
                c._steps = c._steps || loadCfg(recipe.storageKey)._steps || {};
                const prev = loadCfg(recipe.storageKey);
                const merged = Object.assign({}, prev, readFields(recipe));
                merged._steps = prev._steps || {};
                merged._steps[i] = !!el.checked;
                saveCfg(recipe.storageKey, merged);
            });
        });

        const btnSave = $('paBtnSave');
        if (btnSave) {
            btnSave.addEventListener('click', function () {
                const prev = loadCfg(recipe.storageKey);
                const merged = Object.assign({}, prev, readFields(recipe));
                saveCfg(recipe.storageKey, merged);
                toast('Einstellungen gespeichert.');
            });
        }

        const btnList = $('paBtnList');
        if (btnList) {
            btnList.addEventListener('click', async function () {
                try {
                    const prev = loadCfg(recipe.storageKey);
                    const merged = Object.assign({}, prev, readFields(recipe));
                    saveCfg(recipe.storageKey, merged);
                    log('— Liste —');
                    await createList(recipe, merged);
                    toast('Liste bereit.');
                } catch (e) {
                    log('Fehler: ' + (e && e.message ? e.message : e));
                    toast(e && e.message ? e.message : String(e));
                }
            });
        }

        const doneBox = $('paDone');
        if (doneBox) {
            doneBox.addEventListener('change', function () {
                setDone(recipe, !!doneBox.checked);
            });
        }
    }

    function init() {
        const id = recipeIdFromPage();
        const recipe = catalog.getById(id);
        if (!recipe) {
            const host = $('paTitle');
            if (host) host.textContent = 'Rezept nicht gefunden';
            return;
        }
        wire(recipe);
        if (window.ms365PaOnboarding && document.getElementById('paOnboardingHost')) {
            window.ms365PaOnboarding.mount('#paOnboardingHost', { compact: true });
        }
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
