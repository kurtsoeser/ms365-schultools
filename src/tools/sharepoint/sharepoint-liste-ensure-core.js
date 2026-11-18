/**
 * Gemeinsame Hilfen: SharePoint-Liste finden, Spalten nachziehen (idempotent).
 */

/**
 * @returns {typeof window.ms365SpoGraph}
 */
export function spoGraphApi() {
    const api = window.ms365SpoGraph;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

/**
 * @param {(def: object) => object} toGraphColumnBody
 */
export function createListEnsureHelpers(toGraphColumnBody) {
    const G = spoGraphApi;

    function sameChoiceSet(a, b) {
        const left = (Array.isArray(a) ? a : []).map((x) => String(x)).sort();
        const right = (Array.isArray(b) ? b : []).map((x) => String(x)).sort();
        if (left.length !== right.length) return false;
        return left.every((v, i) => v === right[i]);
    }

    /**
     * @param {string} token
     * @param {string} siteId
     * @param {string} listTitle
     */
    async function findListByTitle(token, siteId, listTitle) {
        const g = G();
        const title = String(listTitle || '').trim();
        const path =
            g.graphPathSite(siteId) +
            '/lists?$filter=' +
            encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
            '&$select=id,displayName,webUrl';
        const data = await g.graphJson('GET', path, token, undefined, 'v1.0');
        return ((data && data.value) || [])[0] || null;
    }

    /**
     * @param {string} siteId
     * @param {string} listId
     * @param {string} token
     * @param {object[]} defs
     * @param {(msg: string) => void} write
     */
    async function addMissingColumns(siteId, listId, token, defs, write) {
        const g = G();
        const colsPath =
            g.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
        const colsData = await g.graphJson('GET', colsPath, token, undefined, 'v1.0');
        const existingCols = (colsData && colsData.value) || [];
        const byName = new Map(
            existingCols
                .map((c) => [String((c && c.name) || '').trim(), c])
                .filter((pair) => pair[0])
        );
        const base = g.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
        let added = 0;
        let updated = 0;
        for (let i = 0; i < defs.length; i++) {
            const def = defs[i];
            // Title kommt mit genericList – nicht erneut anlegen
            if (String(def.name || '') === 'Title' && byName.has('Title')) continue;
            const existing = byName.get(def.name);
            if (!existing) {
                await g.graphJson('POST', base, token, toGraphColumnBody(def), 'v1.0');
                added++;
                write('  + Spalte ' + def.name);
                await g.sleep(120);
                continue;
            }
            if (def.choice && existing.id) {
                const current = (existing.choice && existing.choice.choices) || [];
                const wanted = def.choice.choices || [];
                if (!sameChoiceSet(current, wanted)) {
                    await g.graphJson(
                        'PATCH',
                        base + '/' + encodeURIComponent(existing.id),
                        token,
                        {
                            choice: {
                                allowTextEntry: !!def.choice.allowTextEntry,
                                choices: [...wanted]
                            }
                        },
                        'v1.0'
                    );
                    updated++;
                    write('  ~ Choice-Spalte ' + def.name + ' aktualisiert');
                    await g.sleep(120);
                }
            }
        }
        return { added, updated };
    }

    /**
     * @param {string} token
     * @param {string} siteId
     * @param {string} title
     * @param {object[]} columnDefs
     * @param {(msg: string) => void} write
     */
    async function ensureList(token, siteId, title, columnDefs, write) {
        const g = G();
        let list = await findListByTitle(token, siteId, title);
        if (!list || !list.id) {
            write('Erstelle Liste „' + title + '" …');
            const created = await g.graphJson(
                'POST',
                g.graphPathSite(siteId) + '/lists',
                token,
                { displayName: title, list: { template: 'genericList' } },
                'v1.0'
            );
            const listId = created && created.id ? String(created.id) : '';
            if (!listId) throw new Error('Listen-ID fehlt für „' + title + '“.');
            list = { id: listId, displayName: title, webUrl: created.webUrl || '' };
            write('Liste angelegt: ' + (list.webUrl || listId));
        } else {
            write('Liste „' + title + '" bereits vorhanden.');
        }
        write('Prüfe Spalten für „' + title + '" …');
        const result = await addMissingColumns(siteId, list.id, token, columnDefs, write);
        const parts = [];
        if (result.added) parts.push(result.added + ' Spalte(n) ergänzt');
        if (result.updated) parts.push(result.updated + ' Choice-Spalte(n) aktualisiert');
        write(parts.length ? parts.join(', ') + '.' : 'Spalten vollständig.');
        return list;
    }

    return { findListByTitle, addMissingColumns, ensureList, spoGraphApi: G };
}
