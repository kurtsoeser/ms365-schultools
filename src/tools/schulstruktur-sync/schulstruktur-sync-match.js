/**
 * Tenant-Match-Vorschläge und das zugehörige Suchfeld-Wiring für
 * „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2 Schnitt 7).
 *
 * Pure Funktionen (gut testbar):
 *  - {@link normKey}: aggressive, Unicode-bewusste Normalisierung für Match-Keys.
 *  - {@link suggestTenantGroupForUnitFromList}: schlägt eine Gruppe für eine
 *    Strukturzeile vor (exakt → Alias → enthält).
 *  - {@link suggestTenantUserForPersonFromList}: scoring-basierter
 *    Vorschlag für Person-Zeilen anhand Name/UPN/mail/Alias/Local-Part.
 *  - {@link suggestTenantMatchSelectValue}: kombiniert beide Vorschläge zum
 *    `g:<id>` / `u:<id>` Dropdown-Wert.
 *  - {@link formatEntraUserPickLabel}: einheitliches Anzeige-Label.
 *  - {@link matchTenantFilterNeedle}: Filtertext-Normalisierung.
 *  - {@link matchTenantHaystackForGroup} / {@link matchTenantHaystackForUser}:
 *    erzeugen den Vergleichs-Heuhaufen.
 *
 * DOM-Wiring (nicht testbar):
 *  - {@link rebuildMatchTenantSelectOptions}: baut die Dropdown-Liste neu auf.
 *  - {@link wireMatchTenantSearchOnce}: Event-Wiring fürs Suchfeld (idempotent).
 */

import { getEl } from '../../shared/utils/dom.js';

/**
 * Aggressive Normalisierung für Match-Keys: trim, lower, Whitespace zusammenfassen,
 * alle Zeichen außer Buchstaben/Zahlen und `- _ .` entfernen.
 *
 * @param {unknown} s
 * @returns {string}
 */
export function normKey(s) {
    return String(s || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .replace(/[^\p{L}\p{N}\-_. ]/gu, '');
}

/**
 * Abgleich-Vorschlag: passende Tenant-Gruppe für eine Strukturzeile.
 * Reihenfolge:
 *  1. exakter Match auf `bezeichnung`
 *  2. exakter Match auf `alias`
 *  3. „enthält" auf `bezeichnung` oder `alias`
 *
 * @param {{ bezeichnung?: string } | null} unit
 * @param {Array<{ id: string|number, bezeichnung?: string, alias?: string }>} list
 * @returns {string} Group-ID oder Leerstring.
 */
export function suggestTenantGroupForUnitFromList(unit, list) {
    if (!unit) return '';
    const uKey = normKey(unit.bezeichnung || '');
    if (!uKey) return '';
    const rows = Array.isArray(list) ? list : [];
    let best = rows.find((g) => normKey(g.bezeichnung) === uKey);
    if (best) return String(best.id);
    best = rows.find((g) => g.alias && normKey(g.alias) === uKey);
    if (best) return String(best.id);
    best = rows.find(
        (g) => normKey(g.bezeichnung).includes(uKey) || (g.alias && normKey(g.alias).includes(uKey))
    );
    return best ? String(best.id) : '';
}

/**
 * Local-Part einer E-Mail/UPN (vor dem @), normalisiert.
 * @param {unknown} raw
 * @returns {string}
 */
export function emailLocalPart(raw) {
    const s = String(raw || '')
        .trim()
        .toLowerCase();
    if (!s) return '';
    const at = s.indexOf('@');
    const local = at === -1 ? s : s.slice(0, at);
    return normKey(local);
}

/**
 * Alle Identitäts-Schlüssel eines Entra-Users für Matching
 * (DisplayName, UPN, mail, Aliase/otherMails, mailNickname, Vor-/Nachname).
 *
 * @param {{ displayName?: string, userPrincipalName?: string, mail?: string, mailNickname?: string, givenName?: string, surname?: string, otherMails?: string[] } | null} u
 * @returns {{ exact: string[], soft: string[] }}
 */
export function collectUserIdentityKeys(u) {
    const exact = [];
    const soft = [];
    const pushExact = (x) => {
        const k = normKey(x);
        if (k && exact.indexOf(k) === -1) exact.push(k);
    };
    const pushSoft = (x) => {
        const k = normKey(x);
        if (k && soft.indexOf(k) === -1) soft.push(k);
    };
    if (!u || typeof u !== 'object') return { exact, soft };

    pushExact(u.displayName);
    pushExact(u.userPrincipalName);
    pushExact(u.mail);
    pushExact(u.mailNickname);
    pushExact(emailLocalPart(u.userPrincipalName));
    pushExact(emailLocalPart(u.mail));

    const others = Array.isArray(u.otherMails) ? u.otherMails : [];
    for (let i = 0; i < others.length; i++) {
        pushExact(others[i]);
        pushExact(emailLocalPart(others[i]));
    }

    const given = normKey(u.givenName);
    const sur = normKey(u.surname);
    if (given && sur) {
        pushSoft(given + ' ' + sur);
        pushSoft(sur + ' ' + given);
        pushExact(given + '.' + sur);
        pushExact(sur + '.' + given);
    } else {
        pushSoft(given);
        pushSoft(sur);
    }
    return { exact, soft };
}

/**
 * Abgleich-Vorschlag: Entra-Benutzer für SOLL-Typ „Person" (Name/E-Mail/Rolle).
 * Punktet:
 *  - 100 bei exaktem Match auf DisplayName/UPN/mail/Alias/mailNickname/Local-Part
 *  - 80 bei Vorname+Nachname (auch vertauscht)
 *  - 50 bei „enthält"-Match DisplayName / zusammengesetzter Name
 *  - 45 bei „enthält"-Match UPN, mail oder Alias
 *
 * @param {{ typ?: string, personName?: string, personEmail?: string, bezeichnung?: string } | null} unit
 * @param {Array<{ id?: string|number, displayName?: string, userPrincipalName?: string, mail?: string, mailNickname?: string, givenName?: string, surname?: string, otherMails?: string[] }>} users
 * @returns {string} User-ID oder Leerstring.
 */
export function suggestTenantUserForPersonFromList(unit, users) {
    if (!unit || String(unit.typ || '') !== 'Person') return '';
    const arr = Array.isArray(users) ? users : [];
    const keys = [];
    const pushK = (x) => {
        const k = normKey(x);
        if (k && keys.indexOf(k) === -1) keys.push(k);
    };
    pushK(unit.personName);
    pushK(unit.personEmail);
    pushK(unit.bezeichnung);
    pushK(emailLocalPart(unit.personEmail));
    if (!keys.length) return '';

    function scoreUser(u) {
        const ids = collectUserIdentityKeys(u);
        let best = 0;
        for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (!k) continue;
            for (let e = 0; e < ids.exact.length; e++) {
                if (k === ids.exact[e]) return 100;
            }
            for (let s = 0; s < ids.soft.length; s++) {
                const soft = ids.soft[s];
                if (k === soft) best = Math.max(best, 80);
                else if (soft && (soft.includes(k) || k.includes(soft))) best = Math.max(best, 50);
            }
            for (let e = 0; e < ids.exact.length; e++) {
                const ex = ids.exact[e];
                if (!ex) continue;
                // kurze Nummern-UPNs nicht per Substring gegen lange Namen matchen
                if (ex.length >= 3 && k.length >= 3 && (ex.includes(k) || k.includes(ex))) {
                    best = Math.max(best, 45);
                }
            }
        }
        return best;
    }

    let bestId = '';
    let bestScore = 0;
    for (let j = 0; j < arr.length; j++) {
        const u = arr[j];
        const sc = scoreUser(u);
        if (sc > bestScore) {
            bestScore = sc;
            bestId = String(u.id || '');
        }
    }
    // Unter 45 zu unsicher (reine Substring-Zufallstreffer vermeiden)
    return bestScore >= 45 ? bestId : '';
}

/**
 * Dropdown-Wert: `g:<id>` für Gruppe/Team, `u:<id>` für Entra-Benutzer.
 * Person-Zeilen werden zuerst gegen User gematcht, sonst Gruppen.
 *
 * @param {object | null} unit
 * @param {any[]} groups
 * @param {any[]} users
 * @returns {string}
 */
export function suggestTenantMatchSelectValue(unit, groups, users) {
    if (!unit) return '';
    if (String(unit.typ || '') === 'Person') {
        const uid = suggestTenantUserForPersonFromList(unit, users);
        if (uid) return 'u:' + uid;
    }
    const gid = suggestTenantGroupForUnitFromList(unit, groups);
    return gid ? 'g:' + gid : '';
}

/**
 * Einheitliches Anzeige-Label für Entra-User im Auswahl-Dropdown.
 *
 * @param {{ id?: string|number, displayName?: string, userPrincipalName?: string, mail?: string } | null} u
 * @returns {string}
 */
export function formatEntraUserPickLabel(u) {
    if (!u || typeof u !== 'object') return '';
    const dn = u.displayName ? String(u.displayName).trim() : '';
    const upn = String(u.userPrincipalName || u.mail || '').trim();
    const alias =
        Array.isArray(u.otherMails) && u.otherMails.length
            ? String(u.otherMails[0] || '').trim()
            : '';
    if (dn && upn && dn.toLowerCase() !== upn.toLowerCase()) {
        if (alias && alias.toLowerCase() !== upn.toLowerCase() && alias.toLowerCase() !== dn.toLowerCase()) {
            return dn + ' · ' + upn + ' · Alias ' + alias + ' · Benutzer';
        }
        return dn + ' · ' + upn + ' · Benutzer';
    }
    if (alias && (!upn || alias.toLowerCase() !== upn.toLowerCase())) {
        return (dn || upn || String(u.id || '')) + ' · Alias ' + alias + ' · Benutzer';
    }
    return (dn || upn || String(u.id || '')) + ' · Benutzer';
}

/**
 * Normalisiert einen Filtertext für „contains"-Match (lower, getrimmt,
 * Whitespace zusammengefasst).
 *
 * @param {unknown} raw
 * @returns {string}
 */
export function matchTenantFilterNeedle(raw) {
    return String(raw || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ');
}

/**
 * Erzeugt den Such-„Heuhaufen" für eine Gruppe (lowercase, mit Trennzeichen).
 *
 * @param {{ bezeichnung?: string, typ?: string, alias?: string, mail?: string, description?: string, id?: string|number } | null} g
 * @returns {string}
 */
export function matchTenantHaystackForGroup(g) {
    return [
        g && g.bezeichnung,
        g && g.typ,
        g && g.alias,
        g && g.mail,
        g && g.description,
        g && g.id
    ]
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
}

/**
 * Erzeugt den Such-„Heuhaufen" für einen Entra-User – inkl. des
 * Anzeige-Labels, damit auch der Label-Text getroffen wird.
 *
 * @param {object | null} u
 * @returns {string}
 */
export function matchTenantHaystackForUser(u) {
    const others = Array.isArray(u && u.otherMails) ? u.otherMails.join(' ') : '';
    return [
        u && u.displayName,
        u && u.userPrincipalName,
        u && u.mail,
        u && u.mailNickname,
        u && u.givenName,
        u && u.surname,
        others,
        u && u.id,
        formatEntraUserPickLabel(u)
    ]
        .map((x) => String(x || '').toLowerCase())
        .join(' ');
}

/**
 * Baut die Tenant-Auswahlliste neu auf (optional gefiltert).
 * Quelldaten unter `window.__ms365MatchTenantPickSource = { groups, users }`.
 *
 * @param {HTMLSelectElement | null} selTenant
 * @param {string} filterRaw
 * @param {string} selectedValue Wert nach Auswahl (z. B. `g:…` / `u:…`).
 */
export function rebuildMatchTenantSelectOptions(selTenant, filterRaw, selectedValue) {
    const src = (typeof window !== 'undefined' && window.__ms365MatchTenantPickSource) || null;
    const cntEl = getEl('ssMatchTenantFilterCount');
    if (!selTenant || !src || typeof src !== 'object') {
        if (cntEl) cntEl.textContent = '';
        return;
    }
    const needle = matchTenantFilterNeedle(filterRaw);
    const list = Array.isArray(src.groups) ? src.groups : [];
    const users = Array.isArray(src.users) ? src.users : [];
    const prevSel = String(selectedValue != null ? selectedValue : selTenant.value || '');

    selTenant.replaceChildren();
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = '(keine Zuordnung)';
    selTenant.appendChild(opt0);

    const gFiltered = needle
        ? list.filter((g) => matchTenantHaystackForGroup(g).indexOf(needle) !== -1)
        : list.slice();
    const uFiltered = needle
        ? users.filter((u) => matchTenantHaystackForUser(u).indexOf(needle) !== -1)
        : users.slice();

    function selectionInFiltered() {
        if (!prevSel) return true;
        if (prevSel.startsWith('g:')) {
            const id = prevSel.slice(2);
            return gFiltered.some((g) => String(g.id) === id);
        }
        if (prevSel.startsWith('u:')) {
            const id = prevSel.slice(2);
            return uFiltered.some((u) => String(u.id) === id);
        }
        return false;
    }

    if (prevSel && needle && !selectionInFiltered()) {
        let label = prevSel;
        if (prevSel.startsWith('g:')) {
            const id = prevSel.slice(2);
            const g = list.find((x) => String(x.id) === id);
            if (g) label = (g.bezeichnung || id) + ' · ' + (g.typ || '') + ' · aktuell verknüpft';
        } else if (prevSel.startsWith('u:')) {
            const id = prevSel.slice(2);
            const u = users.find((x) => String(x.id) === id);
            if (u) label = formatEntraUserPickLabel(u) + ' · aktuell verknüpft';
        }
        const ox = document.createElement('option');
        ox.value = prevSel;
        ox.textContent = label;
        selTenant.appendChild(ox);
    }

    if (gFiltered.length) {
        const og = document.createElement('optgroup');
        og.label = 'Gruppen / Teams';
        for (let i = 0; i < gFiltered.length; i++) {
            const g = gFiltered[i];
            const o = document.createElement('option');
            o.value = 'g:' + String(g.id || '');
            o.textContent = (g.bezeichnung || '(ohne Name)') + ' · ' + (g.typ || '') + (g.alias ? ' · ' + g.alias : '');
            og.appendChild(o);
        }
        selTenant.appendChild(og);
    }

    if (uFiltered.length) {
        const ou = document.createElement('optgroup');
        ou.label = 'Benutzer (Entra ID)';
        for (let j = 0; j < uFiltered.length; j++) {
            const u = uFiltered[j];
            const o = document.createElement('option');
            o.value = 'u:' + String(u.id || '');
            o.textContent = formatEntraUserPickLabel(u);
            ou.appendChild(o);
        }
        selTenant.appendChild(ou);
    }

    if (prevSel) {
        const ok = Array.from(selTenant.options).some((o) => String(o.value) === prevSel);
        selTenant.value = ok ? prevSel : '';
    } else {
        selTenant.value = '';
    }

    if (cntEl) {
        const total = list.length + users.length;
        const shown = gFiltered.length + uFiltered.length;
        if (!total) cntEl.textContent = 'Noch keine Tenant-Daten – unter „Verwalten" Tenant einlesen.';
        else if (!needle) cntEl.textContent = String(total) + ' Einträge – Suchfeld nutzen, um die Liste einzugrenzen.';
        else cntEl.textContent = 'Zeige ' + shown + ' von ' + total + ' Einträgen (Filter aktiv).';
    }
}

/** Modul-privater Singleton-Guard für `wireMatchTenantSearchOnce`. */
let __matchTenantSearchWired = false;

/**
 * Verkabelt das Tenant-Suchfeld genau einmal mit `rebuildMatchTenantSelectOptions`.
 * Idempotent – wiederholte Aufrufe sind ein No-Op.
 */
export function wireMatchTenantSearchOnce() {
    if (__matchTenantSearchWired) return;
    const inp = getEl('ssMatchTenantSearch');
    if (!inp) return;
    __matchTenantSearchWired = true;
    inp.addEventListener('input', () => {
        const sel = getEl('ssMatchTenantGroup');
        if (!sel) return;
        const cur = String(sel.value || '');
        rebuildMatchTenantSelectOptions(sel, inp.value || '', cur);
    });
}
