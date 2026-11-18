/**
 * Wahltyp-Templates für Schulwahlen (Schulsprecher / SGA Lehrer).
 * Tool-ID bleibt schulsprecherwahl; electionKind steuert Verfahren & UI.
 * Schulsprecher optional: includeSgaDeputies → zusätzlich 3/2/1 für SGA-Schüler-STV.
 */

/** @typedef {'schulsprecher'|'sga_lehrer'} ElectionKind */

export const ELECTION_KIND_SCHULSPRECHER = 'schulsprecher';
export const ELECTION_KIND_SGA_LEHRER = 'sga_lehrer';

export const POINTS_SGA_SCHUELER_STV = Object.freeze([3, 2, 1]);

/** @type {readonly ElectionKind[]} */
export const ELECTION_KINDS = Object.freeze([ELECTION_KIND_SCHULSPRECHER, ELECTION_KIND_SGA_LEHRER]);

/**
 * @typedef {object} ElectionSeatDef
 * @property {string} key
 * @property {string} label
 * @property {string} badge
 */

/**
 * @typedef {object} ElectionKindDef
 * @property {ElectionKind} id
 * @property {string} label
 * @property {string} shortLabel
 * @property {string} defaultTitle
 * @property {string} legalHint
 * @property {number[]} points
 * @property {number} rankingCount
 * @property {boolean} supportsRunoff
 * @property {'students'|'teachers'} candidatePool
 * @property {'students'|'teachers'|'manual'} eligibleModeDefault
 * @property {ElectionSeatDef[]} seats
 */

/** @type {readonly ElectionSeatDef[]} */
export const SCHULSPRECHER_SEATS = Object.freeze([
    Object.freeze({ key: 'schulsprecherId', label: 'Schulsprecher:in', badge: 'Gewählt · Schulsprecher:in' }),
    Object.freeze({ key: 'stv1Id', label: '1. Stellvertretung', badge: 'Gewählt · 1. Stellvertretung' }),
    Object.freeze({ key: 'stv2Id', label: '2. Stellvertretung', badge: 'Gewählt · 2. Stellvertretung' })
]);

/** @type {readonly ElectionSeatDef[]} */
export const SGA_SCHUELER_DEPUTY_SEATS = Object.freeze([
    Object.freeze({
        key: 'sgaDeputy1Id',
        label: 'SGA-Stellvertretung 1',
        badge: 'Gewählt · SGA-Stellvertretung 1'
    }),
    Object.freeze({
        key: 'sgaDeputy2Id',
        label: 'SGA-Stellvertretung 2',
        badge: 'Gewählt · SGA-Stellvertretung 2'
    }),
    Object.freeze({
        key: 'sgaDeputy3Id',
        label: 'SGA-Stellvertretung 3',
        badge: 'Gewählt · SGA-Stellvertretung 3'
    })
]);

/** @type {Record<ElectionKind, ElectionKindDef>} */
export const ELECTION_KIND_DEFS = Object.freeze({
    [ELECTION_KIND_SCHULSPRECHER]: Object.freeze({
        id: ELECTION_KIND_SCHULSPRECHER,
        label: 'Schulsprecher:innenwahl',
        shortLabel: 'Schulsprecher:in',
        defaultTitle: 'Schulsprecher:innenwahl',
        legalHint:
            'VO über die Wahl der Schülervertreter (BGBl. Nr. 388/1993): 6 / 5 / 4 Punkte, absolute Mehrheit der Erstgereihten sonst Stichwahl. Optional gleichzeitig drei SGA-Stellvertreter:innen (3 / 2 / 1).',
        points: Object.freeze([6, 5, 4]),
        rankingCount: 3,
        supportsRunoff: true,
        candidatePool: 'students',
        eligibleModeDefault: 'students',
        seats: SCHULSPRECHER_SEATS
    }),
    [ELECTION_KIND_SGA_LEHRER]: Object.freeze({
        id: ELECTION_KIND_SGA_LEHRER,
        label: 'SGA Lehrervertretung',
        shortLabel: 'SGA Lehrer',
        defaultTitle: 'SGA-Wahl · Lehrervertretung',
        legalHint:
            'VO über die Wahl der Vertreter der Lehrer und der Erziehungsberechtigten in den SGA (BGBl. Nr. 389/1993): 6…1 Punkte, Top 3 Mitglieder + Platz 4–6 Stellvertretung, Los bei Gleichstand.',
        points: Object.freeze([6, 5, 4, 3, 2, 1]),
        rankingCount: 6,
        supportsRunoff: false,
        candidatePool: 'teachers',
        eligibleModeDefault: 'teachers',
        seats: Object.freeze([
            Object.freeze({ key: 'member1Id', label: 'Lehrervertreter:in 1', badge: 'Gewählt · Lehrervertreter:in 1' }),
            Object.freeze({ key: 'member2Id', label: 'Lehrervertreter:in 2', badge: 'Gewählt · Lehrervertreter:in 2' }),
            Object.freeze({ key: 'member3Id', label: 'Lehrervertreter:in 3', badge: 'Gewählt · Lehrervertreter:in 3' }),
            Object.freeze({ key: 'deputy1Id', label: '1. Stellvertretung', badge: 'Gewählt · 1. Stellvertretung' }),
            Object.freeze({ key: 'deputy2Id', label: '2. Stellvertretung', badge: 'Gewählt · 2. Stellvertretung' }),
            Object.freeze({ key: 'deputy3Id', label: '3. Stellvertretung', badge: 'Gewählt · 3. Stellvertretung' })
        ])
    })
});

/**
 * @param {unknown} value
 * @returns {ElectionKind}
 */
export function normalizeElectionKind(value) {
    const v = String(value || '')
        .trim()
        .toLowerCase();
    if (v === ELECTION_KIND_SGA_LEHRER || v === 'sga-lehrer' || v === 'lehrer_sga') {
        return ELECTION_KIND_SGA_LEHRER;
    }
    return ELECTION_KIND_SCHULSPRECHER;
}

/**
 * @param {unknown} kind
 * @returns {ElectionKindDef}
 */
export function getElectionKindDef(kind) {
    const id = normalizeElectionKind(kind);
    return ELECTION_KIND_DEFS[id] || ELECTION_KIND_DEFS[ELECTION_KIND_SCHULSPRECHER];
}

/**
 * Ob die Schulsprecherwahl auch SGA-Schüler-STV (3/2/1) enthält.
 * @param {unknown} electionOrFlag
 */
export function electionIncludesSgaDeputies(electionOrFlag) {
    if (electionOrFlag === true) return true;
    if (!electionOrFlag || typeof electionOrFlag !== 'object') return false;
    const kind = normalizeElectionKind(electionOrFlag.electionKind || electionOrFlag.kind);
    if (kind !== ELECTION_KIND_SCHULSPRECHER) return false;
    return !!(electionOrFlag.includeSgaDeputies || electionOrFlag.includeSgaSchuelerStv);
}

/**
 * Effektives Ballot-/Sitz-Schema für eine konkrete Wahl (inkl. optionaler SGA-STV).
 * @param {unknown} electionOrOpts
 */
export function resolveElectionBallotSchema(electionOrOpts) {
    const o =
        electionOrOpts && typeof electionOrOpts === 'object' && !Array.isArray(electionOrOpts)
            ? electionOrOpts
            : { electionKind: electionOrOpts };
    const kind = normalizeElectionKind(o.kind || o.electionKind);
    const def = getElectionKindDef(kind);
    const includeSgaDeputies = electionIncludesSgaDeputies(o);

    if (kind === ELECTION_KIND_SGA_LEHRER) {
        return {
            kind,
            rankingCount: 6,
            points: [6, 5, 4, 3, 2, 1],
            includeSgaDeputies: false,
            coreRankingCount: 6,
            corePoints: [6, 5, 4, 3, 2, 1],
            sgaPoints: [],
            seats: def.seats.slice(),
            supportsRunoff: false,
            candidatePool: def.candidatePool
        };
    }

    if (includeSgaDeputies) {
        const seats = SCHULSPRECHER_SEATS.concat(SGA_SCHUELER_DEPUTY_SEATS);
        return {
            kind,
            rankingCount: 6,
            points: [6, 5, 4, 3, 2, 1],
            includeSgaDeputies: true,
            coreRankingCount: 3,
            corePoints: [6, 5, 4],
            sgaPoints: POINTS_SGA_SCHUELER_STV.slice(),
            seats,
            supportsRunoff: true,
            candidatePool: def.candidatePool
        };
    }

    return {
        kind,
        rankingCount: 3,
        points: [6, 5, 4],
        includeSgaDeputies: false,
        coreRankingCount: 3,
        corePoints: [6, 5, 4],
        sgaPoints: [],
        seats: SCHULSPRECHER_SEATS.slice(),
        supportsRunoff: true,
        candidatePool: def.candidatePool
    };
}

/**
 * Schritt-Labels für den Stimmzettel (Reihung für Reihung).
 * @param {unknown} electionOrOpts
 * @returns {{ pts: number, label: string, key?: string }[]}
 */
export function rankingStepLabelsForElection(electionOrOpts) {
    const schema = resolveElectionBallotSchema(electionOrOpts);
    if (schema.kind === ELECTION_KIND_SGA_LEHRER) {
        return schema.seats.map((seat, i) => ({
            pts: schema.points[i],
            label: seat.label + ' (' + schema.points[i] + ' Punkte)',
            key: seat.key
        }));
    }
    if (schema.includeSgaDeputies) {
        return [
            { pts: 6, label: 'Schulsprecher:in (6 Punkte)', key: 'schulsprecherId' },
            { pts: 5, label: '1. Stellvertretung (5 Punkte)', key: 'stv1Id' },
            { pts: 4, label: '2. Stellvertretung (4 Punkte)', key: 'stv2Id' },
            { pts: 3, label: 'SGA-Stellvertretung (3 Punkte)', key: 'sgaDeputy1Id' },
            { pts: 2, label: 'SGA-Stellvertretung (2 Punkte)', key: 'sgaDeputy2Id' },
            { pts: 1, label: 'SGA-Stellvertretung (1 Punkt)', key: 'sgaDeputy3Id' }
        ];
    }
    return [
        { pts: 6, label: 'Schulsprecher:in (6 Punkte)', key: 'schulsprecherId' },
        { pts: 5, label: '1. Stellvertretung (5 Punkte)', key: 'stv1Id' },
        { pts: 4, label: '2. Stellvertretung (4 Punkte)', key: 'stv2Id' }
    ];
}

/**
 * Punkte-Labels für Stimmzettel-UI (z. B. „6 Punkte“).
 * @param {unknown} kindOrElection
 * @returns {string[]}
 */
export function rankingPointLabels(kindOrElection) {
    return resolveElectionBallotSchema(kindOrElection).points.map((p) => p + ' Punkte');
}

/**
 * Sitze für Anzeige/Auswertung.
 * @param {unknown} electionOrOpts
 */
export function seatsForElection(electionOrOpts) {
    return resolveElectionBallotSchema(electionOrOpts).seats;
}

/**
 * Leere offices-Struktur passend zum Wahltyp.
 * @param {unknown} kindOrElection
 */
export function emptyOfficesForKind(kindOrElection) {
    const schema = resolveElectionBallotSchema(kindOrElection);
    if (schema.kind === ELECTION_KIND_SGA_LEHRER) {
        return {
            member1Id: null,
            member2Id: null,
            member3Id: null,
            deputy1Id: null,
            deputy2Id: null,
            deputy3Id: null,
            decidedAfter: null
        };
    }
    /** @type {Record<string, string|null>} */
    const out = {
        schulsprecherId: null,
        stv1Id: null,
        stv2Id: null,
        decidedAfter: null
    };
    if (schema.includeSgaDeputies) {
        out.sgaDeputy1Id = null;
        out.sgaDeputy2Id = null;
        out.sgaDeputy3Id = null;
    }
    return out;
}

/**
 * Leere Los-Felder passend zum Wahltyp.
 * @param {unknown} kindOrElection
 */
export function emptyLotForKind(kindOrElection) {
    const schema = resolveElectionBallotSchema(kindOrElection);
    if (schema.kind === ELECTION_KIND_SGA_LEHRER) {
        return {
            member1Id: null,
            member2Id: null,
            member3Id: null,
            deputy1Id: null,
            deputy2Id: null,
            deputy3Id: null
        };
    }
    /** @type {Record<string, string|null>} */
    const out = {
        schulsprecherId: null,
        stv1Id: null,
        stv2Id: null
    };
    if (schema.includeSgaDeputies) {
        out.sgaDeputy1Id = null;
        out.sgaDeputy2Id = null;
        out.sgaDeputy3Id = null;
    }
    return out;
}
