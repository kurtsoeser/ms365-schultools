/**
 * Erzeugt MS365-/Schul-E-Mail-Kandidaten aus Vor- und Nachnamen.
 * Berücksichtigt Umlaute, Leerzeichen/Bindestriche (Doppelnamen) und Kollisionen.
 */
(function () {
    'use strict';

    var PATTERNS = [
        { id: 'vorname.nachname', label: 'vorname.nachname' },
        { id: 'nachname.vorname', label: 'nachname.vorname' },
        { id: 'v.nachname', label: 'v.nachname' },
        { id: 'vorname.n', label: 'vorname.n' },
        { id: 'vorname_nachname', label: 'vorname_nachname' },
        { id: 'nachname_vorname', label: 'nachname_vorname' }
    ];

    function normStr(v) {
        return String(v == null ? '' : v).trim();
    }

    function stripDiacritics(s) {
        return String(s || '')
            .replace(/ä/g, 'ae')
            .replace(/ö/g, 'oe')
            .replace(/ü/g, 'ue')
            .replace(/Ä/g, 'Ae')
            .replace(/Ö/g, 'Oe')
            .replace(/Ü/g, 'Ue')
            .replace(/ß/g, 'ss')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '');
    }

    function sanitizeToken(s) {
        return stripDiacritics(normStr(s))
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, '')
            .replace(/^-+|-+$/g, '');
    }

    /** Teile bei Leerzeichen und Bindestrich, leere Tokens verwerfen. */
    function nameParts(s) {
        return normStr(s)
            .split(/[\s\-]+/)
            .map(sanitizeToken)
            .filter(Boolean);
    }

    function joinParts(parts, sep) {
        return (parts || []).filter(Boolean).join(sep || '');
    }

    /**
     * Varianten für einen Namensblock (Vor- oder Nachname), z. B. „Devran Eren“ / „Al Akrad“.
     * @returns {string[]}
     */
    function blockVariants(raw) {
        const parts = nameParts(raw);
        if (!parts.length) return [''];
        if (parts.length === 1) return [parts[0]];
        const out = [];
        const seen = new Set();
        function add(v) {
            const t = sanitizeToken(v);
            if (!t || seen.has(t)) return;
            seen.add(t);
            out.push(t);
        }
        add(joinParts(parts, ''));
        add(joinParts(parts, '.'));
        add(parts[0] + joinParts(parts.slice(1), ''));
        add(parts[0].charAt(0) + joinParts(parts.slice(1), ''));
        add(parts[0].charAt(0) + '.' + joinParts(parts.slice(1), '.'));
        add(joinParts(parts.map(function (p) { return p.charAt(0); }), ''));
        return out;
    }

    function patternLocal(patternId, vorname, nachname) {
        const id = String(patternId || 'vorname.nachname').toLowerCase();
        const vParts = nameParts(vorname);
        const nParts = nameParts(nachname);
        const v0 = vParts[0] || '';
        const n0 = nParts[0] || '';
        const vAll = joinParts(vParts, '');
        const nAll = joinParts(nParts, '');
        const vDot = joinParts(vParts, '.') || vAll;
        const nDot = joinParts(nParts, '.') || nAll;

        if (id === 'nachname.vorname') return nDot && vDot ? nDot + '.' + vDot : nAll + vAll;
        if (id === 'v.nachname') return v0 && nDot ? v0.charAt(0) + '.' + nDot : (v0.charAt(0) || '') + nAll;
        if (id === 'vorname.n') return vDot && n0 ? vDot + '.' + n0.charAt(0) : vAll + (n0.charAt(0) || '');
        if (id === 'vorname_nachname') return vAll && nAll ? vAll + '_' + nAll : vAll || nAll;
        if (id === 'nachname_vorname') return nAll && vAll ? nAll + '_' + vAll : nAll || vAll;
        // default vorname.nachname
        return vDot && nDot ? vDot + '.' + nDot : vAll || nAll;
    }

    /**
     * Alle sinnvollen Local-Parts für eine Person (Primärmuster zuerst, dann Doppelname-Varianten).
     */
    function localPartCandidates(vorname, nachname, patternId) {
        const primary = sanitizeToken(patternLocal(patternId, vorname, nachname).replace(/_/g, '_'));
        // patternLocal already returns mostly clean; keep underscore patterns
        function cleanLocal(s) {
            return stripDiacritics(normStr(s))
                .toLowerCase()
                .replace(/[^a-z0-9._-]+/g, '')
                .replace(/^[._-]+|[._-]+$/g, '');
        }
        const seen = new Set();
        const out = [];
        function add(v) {
            const t = cleanLocal(v);
            if (!t || seen.has(t)) return;
            seen.add(t);
            out.push(t);
        }

        add(patternLocal(patternId, vorname, nachname));

        const vVars = blockVariants(vorname);
        const nVars = blockVariants(nachname);
        const id = String(patternId || 'vorname.nachname').toLowerCase();
        vVars.forEach(function (v) {
            nVars.forEach(function (n) {
                if (!v && !n) return;
                if (id === 'nachname.vorname') add(n && v ? n + '.' + v : n || v);
                else if (id === 'v.nachname') add(v && n ? v.charAt(0) + '.' + n : (v.charAt(0) || '') + n);
                else if (id === 'vorname.n') add(v && n ? v + '.' + n.charAt(0) : v + (n.charAt(0) || ''));
                else if (id === 'vorname_nachname') add(v && n ? v + '_' + n : v || n);
                else if (id === 'nachname_vorname') add(n && v ? n + '_' + v : n || v);
                else add(v && n ? v + '.' + n : v || n);
            });
        });

        if (primary) {
            // ensure primary is first
            const idx = out.indexOf(cleanLocal(primary));
            if (idx > 0) {
                out.splice(idx, 1);
                out.unshift(cleanLocal(primary));
            }
        }
        return out;
    }

    function buildEmail(local, domain) {
        const loc = normStr(local).toLowerCase();
        const dom = normStr(domain).replace(/^@+/, '').toLowerCase();
        if (!loc || !dom) return '';
        return loc + '@' + dom;
    }

    /**
     * @param {{ givenName?: string, surname?: string, foreName?: string, longName?: string, name?: string, email?: string }} person
     * @param {{ domain: string, pattern?: string, usedEmails?: Set<string>|string[], preferExisting?: boolean }} opts
     * @returns {{ email: string, generated: boolean, candidates: string[], pattern: string, conflict: boolean }}
     */
    function suggestEmail(person, opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const domain = normStr(o.domain).replace(/^@+/, '');
        const pattern = String(o.pattern || 'vorname.nachname');
        const preferExisting = o.preferExisting !== false;
        const used = o.usedEmails instanceof Set ? o.usedEmails : new Set(Array.isArray(o.usedEmails) ? o.usedEmails : []);

        const existing = normStr(person && person.email).toLowerCase();
        if (preferExisting && existing.indexOf('@') !== -1) {
            return {
                email: existing,
                generated: false,
                candidates: [existing],
                pattern: pattern,
                conflict: false
            };
        }

        let given = normStr(person && (person.givenName || person.foreName || person.vorname));
        let sur = normStr(person && (person.surname || person.longName || person.familienname || person.nachname));
        if ((!given || !sur) && person && person.name) {
            const bits = normStr(person.name).split(/\s+/).filter(Boolean);
            if (!given && bits.length) given = bits[0];
            if (!sur && bits.length > 1) sur = bits.slice(1).join(' ');
        }

        const locals = localPartCandidates(given, sur, pattern);
        const candidates = locals.map(function (loc) {
            return buildEmail(loc, domain);
        }).filter(Boolean);

        let chosen = '';
        let conflict = false;
        for (let i = 0; i < candidates.length; i++) {
            const em = candidates[i];
            if (!used.has(em)) {
                chosen = em;
                break;
            }
        }
        if (!chosen && candidates.length) {
            chosen = candidates[0];
            conflict = true;
        }

        return {
            email: chosen,
            generated: !!chosen,
            candidates: candidates,
            pattern: pattern,
            conflict: conflict
        };
    }

    /**
     * Weist einer Liste Mails zu; bestehende Mails bleiben, neue werden vergeben und in usedEmails eingetragen.
     * @param {array} people
     * @param {{ domain: string, pattern?: string, getNameParts?: function }} opts
     */
    function assignEmails(people, opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const used = new Set();
        (people || []).forEach(function (p) {
            const em = normStr(p && p.email).toLowerCase();
            if (em.indexOf('@') !== -1) used.add(em);
        });
        return (people || []).map(function (p) {
            const sug = suggestEmail(p, {
                domain: o.domain,
                pattern: o.pattern,
                usedEmails: used,
                preferExisting: true
            });
            if (sug.email) used.add(sug.email);
            return Object.assign({}, p, {
                email: sug.email || normStr(p && p.email).toLowerCase(),
                emailMeta: {
                    generated: sug.generated && !normStr(p && p.email),
                    candidates: sug.candidates,
                    pattern: sug.pattern,
                    conflict: sug.conflict
                }
            });
        });
    }

    window.ms365PersonEmailFromName = {
        PATTERNS: PATTERNS,
        patterns: PATTERNS,
        stripDiacritics: stripDiacritics,
        sanitizeToken: sanitizeToken,
        nameParts: nameParts,
        blockVariants: blockVariants,
        localPartCandidates: localPartCandidates,
        buildEmail: buildEmail,
        suggestEmail: suggestEmail,
        assignEmails: assignEmails
    };
})();
