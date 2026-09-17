/**
 * Reine Matching-Logik Kürzel ↔ Stammdaten-Lehrer (ohne DOM).
 */

function normalizeTeacherCode(code) {
    return String(code || '')
        .trim()
        .toUpperCase()
        .replace(/Ä/g, 'AE')
        .replace(/Ö/g, 'OE')
        .replace(/Ü/g, 'UE')
        .replace(/ß/g, 'SS');
}

function normalizeEmail(email) {
    return String(email || '')
        .trim()
        .toLowerCase();
}

function lastNameFromTeacherName(name) {
    const raw = String(name || '').trim();
    if (!raw) return '';
    // "Nachname, Vorname" oder "Vorname Nachname"
    if (raw.includes(',')) {
        return raw.split(',')[0].trim();
    }
    const parts = raw.split(/\s+/).filter(Boolean);
    if (parts.length === 1) return parts[0];
    return parts[parts.length - 1];
}

function teacherHasUsableEmail(t) {
    const email = normalizeEmail(t && t.email);
    return !!(email && email.includes('@'));
}

/**
 * Findet den besten Stammdaten-Treffer für ein Unterrichts-Kürzel.
 * @returns {{ code: string, name: string, email: string, method: string } | null}
 */
function resolveTeacherMatch(kuerzel, teachers) {
    const code = normalizeTeacherCode(kuerzel);
    if (!code) return null;
    const list = Array.isArray(teachers) ? teachers : [];

    const withCode = list
        .map((t) => ({
            raw: t,
            code: normalizeTeacherCode(t && t.code),
            name: String((t && t.name) || '').trim(),
            email: normalizeEmail(t && t.email)
        }))
        .filter((t) => t.code);

    // 1) Exakter Code (inkl. Umlaut-Normalisierung)
    const exact = withCode.filter((t) => t.code === code && t.email && t.email.includes('@'));
    if (exact.length === 1) {
        return { code: exact[0].code, name: exact[0].name, email: exact[0].email, method: 'exact' };
    }
    if (exact.length > 1) {
        // Mehrere gleiche Codes mit E-Mail → ersten nehmen (sollte selten sein)
        return { code: exact[0].code, name: exact[0].name, email: exact[0].email, method: 'exact' };
    }

    // 2) Eindeutiger Code-Präfix (min. 4 Zeichen), z. B. ENGEL ↔ ENGELM
    if (code.length >= 4) {
        const prefixHits = withCode.filter(
            (t) =>
                t.email &&
                t.email.includes('@') &&
                (t.code.startsWith(code) || code.startsWith(t.code)) &&
                Math.min(t.code.length, code.length) >= 4
        );
        if (prefixHits.length === 1) {
            return {
                code: prefixHits[0].code,
                name: prefixHits[0].name,
                email: prefixHits[0].email,
                method: 'codePrefix'
            };
        }
    }

    // 3) Eindeutiger Nachname-Präfix: Kürzel ENGEL ↔ „Anita Engelmann“
    if (code.length >= 4) {
        const nameHits = withCode.filter((t) => {
            if (!t.email || !t.email.includes('@')) return false;
            const last = normalizeTeacherCode(lastNameFromTeacherName(t.name));
            if (!last || last.length < 4) return false;
            return last.startsWith(code) || code.startsWith(last);
        });
        if (nameHits.length === 1) {
            return {
                code: nameHits[0].code,
                name: nameHits[0].name,
                email: nameHits[0].email,
                method: 'namePrefix'
            };
        }
    }

    // Treffer ohne E-Mail (für Badge/Anzeige)
    const anyExact = withCode.find((t) => t.code === code);
    if (anyExact) {
        return {
            code: anyExact.code,
            name: anyExact.name,
            email: anyExact.email || '',
            method: 'exactNoEmail'
        };
    }

    return null;
}

/**
 * Übernimmt fehlende Mappings aus Stammdaten für benötigte Kürzel.
 * Bestehende Mapping-Einträge werden nicht überschrieben.
 * Schlüssel im Mapping = Unterrichts-Kürzel (UPPERCASE, ohne Umlaut-Ersetzung).
 * @returns {{ added: number, mapping: object, details: Array }}
 */
function syncTeacherMappingFromTenant(mapping, requiredCodes, teachers) {
    const out = mapping && typeof mapping === 'object' ? { ...mapping } : {};
    const details = [];
    let added = 0;

    (requiredCodes || []).forEach((raw) => {
        const lessonCode = String(raw || '')
            .trim()
            .toUpperCase();
        if (!lessonCode) return;
        if (out[lessonCode] && String(out[lessonCode]).includes('@')) return;

        const hit = resolveTeacherMatch(lessonCode, teachers);
        if (!hit || !hit.email || !hit.email.includes('@')) return;
        if (hit.method === 'exactNoEmail') return;

        out[lessonCode] = hit.email;
        added += 1;
        details.push({ code: lessonCode, email: hit.email, method: hit.method, name: hit.name });
    });

    return { added, mapping: out, details };
}

window.ms365KursteamTeacherMatchLogic = {
    normalizeTeacherCode,
    normalizeEmail,
    lastNameFromTeacherName,
    teacherHasUsableEmail,
    resolveTeacherMatch,
    syncTeacherMappingFromTenant
};
