/**
 * Reine Domänenlogik Wahlfach-Zuteilung (ohne DOM / localStorage).
 */

import {
    ANMELDUNG_STATUS,
    normalizeAnmeldung,
    normalizeAngebot,
    normalizeRunde,
    normalizeZuteilung,
    RUNDE_STATUS,
    toIsoDateOnly
} from './wahlfach-zuteilung-model.js';

/**
 * Anmeldefenster offen? Status muss `offen` sein und Datum in [von, bis] (inklusiv).
 * @param {{ status?: string, anmeldeVon?: string, anmeldeBis?: string }|null} runde
 * @param {Date|string} [now]
 * @returns {boolean}
 */
export function isAnmeldefensterOffen(runde, now = new Date()) {
    if (!runde) return false;
    if (String(runde.status || '').toLowerCase() !== 'offen') return false;
    const von = toIsoDateOnly(runde.anmeldeVon);
    const bis = toIsoDateOnly(runde.anmeldeBis);
    if (!von || !bis) return false;
    const today =
        now instanceof Date && !Number.isNaN(now.getTime())
            ? toIsoDateOnly(now)
            : toIsoDateOnly(now);
    if (!today) return false;
    return today >= von && today <= bis;
}

/**
 * Belegte Plätze für ein Angebot (aktive Zuteilungen mit angebotId).
 * @param {string} angebotId
 * @param {Array<{ angebotId?: string|null, anmeldungId?: string }>} zuteilungen
 * @returns {number}
 */
export function countBelegung(angebotId, zuteilungen) {
    const id = String(angebotId || '').trim();
    if (!id) return 0;
    return (zuteilungen || []).filter((z) => z && String(z.angebotId || '') === id).length;
}

/**
 * @param {{ kapazitaetMax?: number, kapazitaetMin?: number }} angebot
 * @param {number} belegt
 * @returns {{ belegt: number, max: number, min: number, frei: number, ueberbuchung: number, kommtZustande: boolean, voll: boolean, warnung: string|null }}
 */
export function kapazitaetStatus(angebot, belegt) {
    const max = Math.max(0, Number(angebot && angebot.kapazitaetMax) || 0);
    const min = Math.max(0, Number(angebot && angebot.kapazitaetMin) || 0);
    const b = Math.max(0, Number(belegt) || 0);
    const frei = Math.max(0, max - b);
    const ueberbuchung = Math.max(0, b - max);
    const kommtZustande = b >= min;
    const voll = max > 0 && b >= max;
    let warnung = null;
    if (ueberbuchung > 0) warnung = 'Überbuchung: ' + ueberbuchung + ' über Maximum';
    else if (voll) warnung = 'Kapazität ausgeschöpft';
    else if (!kommtZustande && b > 0) warnung = 'Noch unter Mindestzahl (' + min + ')';
    return { belegt: b, max, min, frei, ueberbuchung, kommtZustande, voll, warnung };
}

/**
 * Summe der Überbuchungen über alle Angebote.
 * @param {Array<{ id?: string, kapazitaetMax?: number }>} angebote
 * @param {Array<{ angebotId?: string|null }>} zuteilungen
 * @returns {number}
 */
export function countUeberbuchung(angebote, zuteilungen) {
    let sum = 0;
    for (const a of angebote || []) {
        if (!a || !a.id) continue;
        const st = kapazitaetStatus(a, countBelegung(a.id, zuteilungen));
        sum += st.ueberbuchung;
    }
    return sum;
}

/**
 * Prüft, ob eine Zuteilung die Kapazität überschreiten würde.
 * @param {object} angebot
 * @param {Array} zuteilungen
 * @param {string} [anmeldungId] – bestehende Zuteilung dieser Anmeldung wird freigegeben
 * @returns {{ ok: boolean, warnung: string|null, status: ReturnType<typeof kapazitaetStatus> }}
 */
export function kannZuteilen(angebot, zuteilungen, anmeldungId) {
    const aid = String((angebot && angebot.id) || '').trim();
    const others = (zuteilungen || []).filter(
        (z) => z && String(z.anmeldungId || '') !== String(anmeldungId || '')
    );
    const belegtNach = countBelegung(aid, others) + 1;
    const status = kapazitaetStatus(angebot, belegtNach);
    if (status.ueberbuchung > 0) {
        return {
            ok: false,
            warnung: 'Kapazität überschritten (' + status.belegt + '/' + status.max + ').',
            status
        };
    }
    return { ok: true, warnung: status.warnung, status };
}

/**
 * Manuelle Zuteilung anwenden (reine Funktion).
 * @param {{ anmeldungen?: object[], zuteilungen?: object[] }} state
 * @param {string} anmeldungId
 * @param {string|null} angebotId – null = Zuteilung aufheben
 * @param {{ begruendung?: string, zugeteiltVon?: string, zugeteiltAm?: string, force?: boolean }} [meta]
 * @param {Array} [angebote] – für Kapazitätsprüfung
 * @returns {{ ok: boolean, errors: string[], warnings: string[], anmeldungen: object[], zuteilungen: object[] }}
 */
export function applyManuelleZuteilung(state, anmeldungId, angebotId, meta = {}, angebote = []) {
    const errors = [];
    const warnings = [];
    const id = String(anmeldungId || '').trim();
    if (!id) {
        return {
            ok: false,
            errors: ['Anmeldung fehlt.'],
            warnings,
            anmeldungen: (state && state.anmeldungen) || [],
            zuteilungen: (state && state.zuteilungen) || []
        };
    }

    const anmeldungen = ((state && state.anmeldungen) || []).map((a) => ({ ...a }));
    const idx = anmeldungen.findIndex((a) => a && a.id === id);
    if (idx < 0) {
        return {
            ok: false,
            errors: ['Anmeldung nicht gefunden.'],
            warnings,
            anmeldungen: (state && state.anmeldungen) || [],
            zuteilungen: (state && state.zuteilungen) || []
        };
    }

    const anm = anmeldungen[idx];
    if (anm.status === 'storniert') {
        errors.push('Stornierte Anmeldung kann nicht zugeteilt werden.');
        return {
            ok: false,
            errors,
            warnings,
            anmeldungen: (state && state.anmeldungen) || [],
            zuteilungen: (state && state.zuteilungen) || []
        };
    }

    let zuteilungen = ((state && state.zuteilungen) || [])
        .filter((z) => z && String(z.anmeldungId || '') !== id)
        .map((z) => ({ ...z }));

    const targetId = angebotId == null || angebotId === '' ? null : String(angebotId).trim();

    if (targetId) {
        const angebot = (angebote || []).find((a) => a && a.id === targetId);
        if (!angebot) {
            errors.push('Angebot nicht gefunden.');
            return {
                ok: false,
                errors,
                warnings,
                anmeldungen: (state && state.anmeldungen) || [],
                zuteilungen: (state && state.zuteilungen) || []
            };
        }
        if (!angebot.aktiv) {
            errors.push('Angebot ist nicht aktiv.');
            return {
                ok: false,
                errors,
                warnings,
                anmeldungen: (state && state.anmeldungen) || [],
                zuteilungen: (state && state.zuteilungen) || []
            };
        }
        const check = kannZuteilen(angebot, zuteilungen, id);
        if (!check.ok && !meta.force) {
            errors.push(check.warnung || 'Kapazität überschritten.');
            return {
                ok: false,
                errors,
                warnings,
                anmeldungen: (state && state.anmeldungen) || [],
                zuteilungen: (state && state.zuteilungen) || []
            };
        }
        if (check.warnung) warnings.push(check.warnung);

        const prevZ = ((state && state.zuteilungen) || []).find(
            (z) => z && String(z.anmeldungId || '') === id
        );
        zuteilungen.push(
            normalizeZuteilung({
                anmeldungId: id,
                angebotId: targetId,
                begruendung: meta.begruendung || '',
                zugeteiltVon: meta.zugeteiltVon || 'SLG / Admin',
                zugeteiltAm: meta.zugeteiltAm || new Date().toISOString(),
                spoItemId: (prevZ && prevZ.spoItemId) || meta.spoItemId || ''
            })
        );
        anmeldungen[idx] = normalizeAnmeldung({ ...anm, status: 'zugeordnet' });
    } else {
        anmeldungen[idx] = normalizeAnmeldung({
            ...anm,
            status: anm.status === 'zugeordnet' ? 'eingereicht' : anm.status
        });
    }

    return { ok: true, errors, warnings, anmeldungen, zuteilungen };
}

/**
 * Ergebnislisten: pro Angebot die zugeteilten Anmeldungen.
 * @param {Array} angebote
 * @param {Array} anmeldungen
 * @param {Array} zuteilungen
 * @returns {Array<{ angebot: object, belegt: number, status: object, teilnehmende: Array<{ anmeldung: object, zuteilung: object }> }>}
 */
export function ergebnisListenProAngebot(angebote, anmeldungen, zuteilungen) {
    const anmById = new Map((anmeldungen || []).map((a) => [a.id, a]));
    return (angebote || []).map((angebot) => {
        const zs = (zuteilungen || []).filter((z) => z && z.angebotId === angebot.id);
        const teilnehmende = zs
            .map((z) => {
                const anm = anmById.get(z.anmeldungId);
                return anm ? { anmeldung: anm, zuteilung: z } : null;
            })
            .filter(Boolean)
            .sort((a, b) =>
                String(a.anmeldung.schuelerName).localeCompare(String(b.anmeldung.schuelerName), 'de')
            );
        const belegt = zs.length;
        return {
            angebot,
            belegt,
            status: kapazitaetStatus(angebot, belegt),
            teilnehmende
        };
    });
}

/**
 * @param {unknown} draft
 * @returns {{ ok: boolean, errors: string[] }}
 */
export function validateRunde(draft) {
    const errors = [];
    const r = normalizeRunde(draft || {});
    if (!r.schuljahr) errors.push('Schuljahr fehlt.');
    if (!r.titel || r.titel.length < 2) errors.push('Titel fehlt.');
    if (!r.anmeldeVon) errors.push('Anmeldebeginn fehlt.');
    if (!r.anmeldeBis) errors.push('Anmeldeende fehlt.');
    if (r.anmeldeVon && r.anmeldeBis && r.anmeldeBis < r.anmeldeVon) {
        errors.push('Anmeldeende liegt vor dem Beginn.');
    }
    if (!RUNDE_STATUS.includes(r.status)) errors.push('Ungültiger Runden-Status.');
    return { ok: errors.length === 0, errors };
}

/**
 * @param {unknown} draft
 * @returns {{ ok: boolean, errors: string[] }}
 */
export function validateAngebot(draft) {
    const errors = [];
    const a = normalizeAngebot(draft || {});
    if (!a.bezeichnung || a.bezeichnung.length < 2) errors.push('Bezeichnung fehlt.');
    if (a.kapazitaetMax < 1) errors.push('Kapazität (Maximum) muss mindestens 1 sein.');
    if (a.kapazitaetMin > a.kapazitaetMax) {
        errors.push('Mindestzahl darf nicht über dem Maximum liegen.');
    }
    return { ok: errors.length === 0, errors };
}

/**
 * @param {unknown} draft
 * @param {Array} [angebote]
 * @returns {{ ok: boolean, errors: string[] }}
 */
export function validateAnmeldung(draft, angebote = []) {
    const errors = [];
    const a = normalizeAnmeldung(draft || {});
    if (!a.schuelerName || a.schuelerName.length < 2) errors.push('Schülername fehlt.');
    if (!a.klasse) errors.push('Klasse fehlt.');
    if (!a.wunsch1Id) errors.push('1. Wunsch fehlt.');
    const ids = [a.wunsch1Id, a.wunsch2Id, a.wunsch3Id].filter(Boolean);
    if (new Set(ids).size !== ids.length) errors.push('Wünsche müssen unterschiedlich sein.');
    const aktivIds = new Set((angebote || []).filter((x) => x && x.aktiv).map((x) => x.id));
    for (const wid of ids) {
        if (aktivIds.size && !aktivIds.has(wid)) {
            errors.push('Unbekanntes oder inaktives Angebot in den Wünschen.');
            break;
        }
    }
    if (!ANMELDUNG_STATUS.includes(a.status)) errors.push('Ungültiger Anmeldungs-Status.');
    return { ok: errors.length === 0, errors };
}

/**
 * Graph-/Zugriff-Fehler als soft 403 erkennen (kein Cache-Wipe).
 * @param {unknown} err
 * @returns {boolean}
 */
export function isWzAccessDeniedError(err) {
    const msg = String((err && err.message) || err || '');
    return /403|Forbidden|Zugriff verweigert|access denied/i.test(msg);
}

/**
 * Phasen-Label für Admin/Schüler (wahl-portal).
 * @param {{ status?: string, anmeldeVon?: string, anmeldeBis?: string }|null} runde
 * @param {Date|string} [now]
 * @returns {{ phase: string, label: string, tone: 'ok'|'warn'|'info'|'muted', anmeldeOffen: boolean }}
 */
export function resolveWzPhase(runde, now = new Date()) {
    if (!runde) {
        return { phase: 'keine', label: 'Keine Runde', tone: 'muted', anmeldeOffen: false };
    }
    const status = String(runde.status || '').toLowerCase();
    const anmeldeOffen = isAnmeldefensterOffen(runde, now);
    if (status === 'veroeffentlicht') {
        return { phase: 'ergebnis', label: 'Ergebnis veröffentlicht', tone: 'ok', anmeldeOffen: false };
    }
    if (status === 'archiviert') {
        return { phase: 'archiv', label: 'Archiviert', tone: 'muted', anmeldeOffen: false };
    }
    if (status === 'zuteilung' || status === 'geschlossen') {
        return {
            phase: 'zuteilung',
            label: status === 'zuteilung' ? 'Zuteilung läuft' : 'Anmeldung geschlossen · Zuteilung',
            tone: 'warn',
            anmeldeOffen: false
        };
    }
    if (anmeldeOffen) {
        return { phase: 'anmeldung', label: 'Anmeldung offen', tone: 'ok', anmeldeOffen: true };
    }
    if (status === 'offen') {
        return { phase: 'fenster-zu', label: 'Anmeldefenster zu', tone: 'warn', anmeldeOffen: false };
    }
    if (status === 'entwurf') {
        return { phase: 'entwurf', label: 'Entwurf', tone: 'info', anmeldeOffen: false };
    }
    return {
        phase: status || 'unbekannt',
        label: String(runde.status || 'Unbekannt'),
        tone: 'info',
        anmeldeOffen: false
    };
}

/**
 * Ownership für Schüler-Anmeldung (E-Mail für Nachvollziehbarkeit / SPO).
 * @param {object} draft
 * @param {{ accountEmail?: string, accountName?: string, students?: object[] }} ctx
 */
export function resolveWzOwnership(draft, ctx) {
    const d = draft && typeof draft === 'object' ? draft : {};
    const students = Array.isArray(ctx && ctx.students) ? ctx.students : [];
    const account = String((ctx && ctx.accountEmail) || '')
        .trim()
        .toLowerCase();

    let schuelerEmail = String(d.schuelerEmail || d.beantragtVon || '')
        .trim()
        .toLowerCase();
    if (!schuelerEmail) schuelerEmail = account;

    const byEmail =
        schuelerEmail &&
        students.find((s) => {
            const em = String((s && s.email) || '')
                .trim()
                .toLowerCase();
            const upn = String((s && s.upn) || '')
                .trim()
                .toLowerCase();
            return em === schuelerEmail || upn === schuelerEmail;
        });

    const schuelerName = String(
        d.schuelerName ||
            (byEmail &&
                [byEmail.vorname || byEmail.firstName, byEmail.nachname || byEmail.lastName]
                    .filter(Boolean)
                    .join(' ')) ||
            (ctx && ctx.accountName) ||
            ''
    ).trim();
    const klasse = String(
        d.klasse || (byEmail && (byEmail.klasse || byEmail.klasseCode || byEmail.classCode)) || ''
    ).trim();

    return { schuelerEmail, schuelerName, klasse, matchedStudent: !!byEmail };
}

/**
 * @param {{ schuelerEmail?: string }} ownership
 * @param {{ requireEmail?: boolean }} [opts]
 */
export function assertWzOwnership(ownership, opts) {
    const requireEmail = !(opts && opts.requireEmail === false);
    if (!requireEmail) return;
    const email = String((ownership && ownership.schuelerEmail) || '')
        .trim()
        .toLowerCase();
    if (!email || email.indexOf('@') < 0) {
        throw new Error(
            'Schüler-Konto nicht erkannt. Bitte anmelden und Anmeldung erneut absenden.'
        );
    }
}

/**
 * Actor für manuelle Zuteilung (Admin/SLG).
 * @param {{ accountEmail?: string, accountName?: string }} ctx
 * @returns {string}
 */
export function resolveWzAdminActor(ctx) {
    const email = String((ctx && ctx.accountEmail) || '')
        .trim()
        .toLowerCase();
    const name = String((ctx && ctx.accountName) || '').trim();
    if (name && email) return name + ' <' + email + '>';
    if (email) return email;
    if (name) return name;
    return 'SLG / Admin';
}

/**
 * Action-Hints für Calm-Banner („Was jetzt?“).
 * @param {{ runde?: object|null, angebote?: object[], anmeldungen?: object[], zuteilungen?: object[] }} state
 * @param {Date|string} [now]
 * @returns {{ eyebrow: string, text: string, ctaView: string, ctaLabel: string }|null}
 */
export function resolveWzActionHint(state, now = new Date()) {
    const s = state || {};
    const r = s.runde;
    if (!r) {
        return {
            eyebrow: 'Was jetzt?',
            text: 'Noch keine Runde – Schuljahr und Anmeldefenster anlegen.',
            ctaView: 'uebersicht',
            ctaLabel: 'Runde anlegen'
        };
    }
    const phase = resolveWzPhase(r, now);
    const offen = (s.anmeldungen || []).filter((a) => a && a.status !== 'storniert');
    const zugeteilt = (s.zuteilungen || []).filter((z) => z && z.angebotId).length;
    const offenWunsch = offen.filter((a) => a.status === 'eingereicht' || a.status === 'offen').length;
    const ueber = countUeberbuchung(s.angebote || [], s.zuteilungen || []);

    if (phase.anmeldeOffen && offen.length === 0) {
        return {
            eyebrow: 'Was jetzt?',
            text: 'Anmeldung ist offen – noch keine Wünsche eingegangen.',
            ctaView: 'anmeldungen',
            ctaLabel: 'Anmeldungen prüfen'
        };
    }
    if (offenWunsch > 0 && (phase.phase === 'zuteilung' || phase.phase === 'fenster-zu' || !phase.anmeldeOffen)) {
        return {
            eyebrow: 'Was jetzt?',
            text:
                offenWunsch +
                ' Anmeldung(en) ohne Zuteilung – manuell zuweisen (kein Auto-Matching).',
            ctaView: 'zuteilung',
            ctaLabel: 'Jetzt zuteilen'
        };
    }
    if (ueber > 0) {
        return {
            eyebrow: 'Was jetzt?',
            text: ueber + ' Überbuchung(en) – Kapazität prüfen und Zuteilungen anpassen.',
            ctaView: 'zuteilung',
            ctaLabel: 'Kapazität prüfen'
        };
    }
    if (zugeteilt > 0 && phase.phase !== 'ergebnis') {
        return {
            eyebrow: 'Was jetzt?',
            text: 'Zuteilungen stehen – Ergebnis prüfen und CSV für Access exportieren.',
            ctaView: 'ergebnis',
            ctaLabel: 'Ergebnis / CSV'
        };
    }
    if (phase.anmeldeOffen && offenWunsch > 0) {
        return {
            eyebrow: 'Was jetzt?',
            text: offenWunsch + ' neue Wunsch(e) – Anmeldungen prüfen.',
            ctaView: 'anmeldungen',
            ctaLabel: 'Anmeldungen prüfen'
        };
    }
    return null;
}
