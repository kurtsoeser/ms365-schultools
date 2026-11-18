/**
 * Fiktive Erziehungsberechtigte für Demo-/Test-Schüler:innen ergänzen
 * (Elternbrief, Einsammler, Krankmeldung, …).
 */

/** Typische weibliche Vornamen (AT/DE) */
const PARENT_FIRST_A = [
    'Maria',
    'Anna',
    'Eva',
    'Lisa',
    'Sarah',
    'Julia',
    'Christine',
    'Petra',
    'Monika',
    'Claudia',
    'Elisabeth',
    'Susanne',
    'Birgit',
    'Katharina',
    'Barbara',
    'Sabine',
    'Andrea',
    'Martina',
    'Nicole',
    'Stefanie'
];

/** Typische männliche Vornamen (AT/DE) */
const PARENT_FIRST_B = [
    'Thomas',
    'Michael',
    'Andreas',
    'Stefan',
    'Christian',
    'Martin',
    'Peter',
    'Wolfgang',
    'Daniel',
    'Markus',
    'Josef',
    'Franz',
    'Georg',
    'Bernhard',
    'Werner',
    'Alexander',
    'David',
    'Patrick',
    'Florian',
    'Manuel'
];

/** Typische Nachnamen (AT/DE) – keine Demo/Test-Platzhalter */
const PARENT_LAST = [
    'Gruber',
    'Huber',
    'Bauer',
    'Wagner',
    'Müller',
    'Steiner',
    'Moser',
    'Meier',
    'Hofer',
    'Leitner',
    'Berger',
    'Fuchs',
    'Eder',
    'Fischer',
    'Schmid',
    'Winkler',
    'Weber',
    'Schwarz',
    'Maier',
    'Lehner'
];

/**
 * @param {string} s
 */
function norm(s) {
    return String(s || '').trim();
}

/**
 * Stabiler Index aus Schüler-Kennung (E-Mail oder Name).
 * @param {string} key
 */
function hashIndex(key) {
    const t = String(key || '').toLowerCase();
    let h = 0;
    for (let i = 0; i < t.length; i++) h = (Math.imul(h, 31) + t.charCodeAt(i)) >>> 0;
    return h >>> 0;
}

/**
 * Sicherer Index in [0, len).
 * @param {number} h
 * @param {number} len
 * @param {number} [shift]
 */
function pickIndex(h, len, shift) {
    const n = Math.max(1, Number(len) || 1);
    const shifted = ((Number(h) >>> 0) >>> (shift || 0)) >>> 0;
    return shifted % n;
}

/**
 * @param {string[]} list
 * @param {number} h
 * @param {number} [shift]
 */
function pickName(list, h, shift) {
    const arr = Array.isArray(list) && list.length ? list : ['Unbekannt'];
    return arr[pickIndex(h, arr.length, shift)] || arr[0];
}

/**
 * @param {string} studentEmail
 * @param {string} studentName
 * @param {number} slot 1|2
 */
export function demoParentEmail(studentEmail, studentName, slot) {
    const mail = norm(studentEmail).toLowerCase();
    const m = mail.match(/^([^@]+)@/);
    let local = m ? m[1] : '';
    if (!local) {
        local =
            norm(studentName)
                .toLowerCase()
                .replace(/[^a-z0-9]+/gi, '.')
                .replace(/^\.+|\.+$/g, '')
                .slice(0, 40) || 'schueler';
    }
    const suffix = slot === 2 ? 'b' : 'a';
    return 'eltern.' + local + '.' + suffix + '@example.com';
}

/**
 * @param {string} studentEmail
 * @param {string} studentName
 * @param {number} slot 1|2
 */
export function demoParentName(studentEmail, studentName, slot) {
    const h = hashIndex(studentEmail || studentName || 'x');
    const first =
        slot === 2 ? pickName(PARENT_FIRST_B, h, 0) : pickName(PARENT_FIRST_A, h, 3);
    // Leicht versetzter Nachname-Index, damit Paare nicht immer denselben haben
    const lastShift = slot === 2 ? 11 : 7;
    const last = pickName(PARENT_LAST, h, lastShift);
    return first + ' ' + last;
}

/**
 * @param {{ name?: string, email?: string, parentPairs?: Array<{ name?: string, email?: string }> }} student
 * @param {{ force?: boolean, dualShare?: number }} [opts]
 * dualShare: Anteil mit zwei Eltern (0–1), Default 0.55
 */
export function enrichStudentWithDemoParents(student, opts) {
    const o = opts || {};
    const force = !!o.force;
    const dualShare = typeof o.dualShare === 'number' ? o.dualShare : 0.55;
    const name = norm(student && student.name);
    const email = norm(student && student.email).toLowerCase();
    const existing = Array.isArray(student && student.parentPairs)
        ? student.parentPairs
              .map((p) => ({
                  name: norm(p && p.name),
                  email: norm(p && p.email).toLowerCase()
              }))
              .filter((p) => p.email.includes('@'))
        : [];
    if (existing.length && !force) {
        return { ...student, parentPairs: existing };
    }
    const h = hashIndex(email || name || 'x');
    const dual = h % 100 < Math.round(dualShare * 100);
    const pairs = [
        {
            name: demoParentName(email, name, 1),
            email: demoParentEmail(email, name, 1)
        }
    ];
    if (dual) {
        pairs.push({
            name: demoParentName(email, name, 2),
            email: demoParentEmail(email, name, 2)
        });
    }
    return { ...student, parentPairs: pairs };
}

/**
 * @param {Array<object>} students
 * @param {{ force?: boolean, dualShare?: number }} [opts]
 */
export function enrichStudentsWithDemoParents(students, opts) {
    return (Array.isArray(students) ? students : []).map((s) =>
        enrichStudentWithDemoParents(s, opts)
    );
}

/**
 * Erkennt kaputte Demo-Namen (z. B. „undefined undefined“) und Platzhalter.
 * @param {string} name
 */
export function isBrokenDemoParentName(name) {
    const n = norm(name).toLowerCase();
    if (!n) return true;
    if (/\bundefined\b/.test(n)) return true;
    if (/\bnull\b/.test(n)) return true;
    if (/^(demo|test|sample|probe|beispiel|muster)(\s|$)/i.test(n)) return true;
    if (/\s(demo|test|sample|probe|beispiel|muster)$/i.test(n)) return true;
    return false;
}

/**
 * Textarea-Zeilen: Klasse;Name;E-Mail[;Eltern…]
 * Ergänzt fehlende Eltern; bestehende Eltern bleiben (außer force).
 * Kaputte „undefined …“-Namen werden auch ohne force neu gesetzt (E-Mail bleibt).
 * @param {string} text
 * @param {{ force?: boolean, dualShare?: number }} [opts]
 * @returns {{ text: string, enriched: number, skipped: number, repaired: number, total: number }}
 */
export function enrichStudentLinesWithDemoParents(text, opts) {
    const o = opts || {};
    const force = !!o.force;
    const linesIn = String(text || '').split(/\r\n|\n|\r/);
    /** @type {string[]} */
    const out = [];
    let enriched = 0;
    let skipped = 0;
    let repaired = 0;
    let total = 0;

    for (const raw of linesIn) {
        const line = String(raw || '');
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) {
            out.push(line);
            continue;
        }
        const parts = trimmed.split(/[;\t,|]/).map((x) => norm(x));
        while (parts.length < 3) parts.push('');
        const klasse = parts[0] || '';
        const name = parts[1] || '';
        const email = (parts[2] || '').toLowerCase();
        if (!klasse && !name && !email) {
            out.push(line);
            continue;
        }
        total++;

        /** @type {Array<{ name: string, email: string }>} */
        let existing = [];
        for (let i = 3; i + 1 < parts.length; i += 2) {
            const pn = parts[i] || '';
            const pe = (parts[i + 1] || '').toLowerCase();
            if (pe.includes('@')) existing.push({ name: pn, email: pe });
            else if (pn.includes('@')) existing.push({ name: '', email: pn.toLowerCase() });
        }

        const hasBrokenNames = existing.some((p) => isBrokenDemoParentName(p.name));
        if (existing.length && !force && !hasBrokenNames) {
            skipped++;
            out.push(line);
            continue;
        }

        const student = enrichStudentWithDemoParents(
            { name, email, parentPairs: force || hasBrokenNames ? [] : existing },
            { force: true, dualShare: o.dualShare }
        );
        const pairs = student.parentPairs || [];
        const next = [klasse, name, email];
        pairs.forEach((p) => {
            next.push(p.name || '');
            next.push(p.email || '');
        });
        out.push(next.join(';'));
        if (hasBrokenNames && !force) repaired++;
        else enriched++;
    }

    return {
        text: out.join('\n'),
        enriched,
        skipped,
        repaired,
        total
    };
}
