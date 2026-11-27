/**
 * Lädt die Portal-Briefings.
 * Sichtbare Werkzeuge, lokale Daten zuerst, SharePoint mit Timeout.
 * Eine fehlende Quelle blendet nur ihre Kachel aus.
 */
import { resolveSchoolToolSiteUrl } from './school-tool-sites.js';
import { buildAufgabenBriefing } from '../tools/aufgaben/aufgaben-briefing.js';
import { buildElternbriefBriefing } from '../tools/elternbrief/elternbrief-briefing.js';
import { buildFreistellungBriefing } from '../tools/freistellung-planer/freistellung-briefing.js';
import { buildSchularbeitenBriefing } from '../tools/schularbeiten-planer/schularbeiten-briefing.js';
import { buildSchulaktivitaetenBriefing } from '../tools/schulaktivitaeten-planer/schulaktivitaeten-briefing.js';
import { buildTechnikBriefing } from '../tools/technik-meldungen/technik-meldungen-briefing.js';
import { expandPortalWidgets, PORTAL_WIDGET_PENDING } from './portal-widgets.js';

export const PORTAL_BRIEFING_TIMEOUT_MS = 8000;
const CACHE_HIT_MS = 120000;
const CACHE_MISS_MS = 20000;

/** @type {Map<string, { at: number, ttl: number, value: object|null }>} */
const cache = new Map();

export function clearPortalBriefingCache() {
    cache.clear();
}

/**
 * @param {Promise<T>} promise
 * @param {number} ms
 * @returns {Promise<T>}
 * @template T
 */
function withTimeout(promise, ms) {
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('timeout')), ms);
        Promise.resolve(promise).then(
            (value) => {
                clearTimeout(timer);
                resolve(value);
            },
            (err) => {
                clearTimeout(timer);
                reject(err);
            }
        );
    });
}

/**
 * @param {string} key
 * @param {() => Promise<object|null>} producer
 */
async function cached(key, producer) {
    const hit = cache.get(key);
    if (hit && Date.now() - hit.at < hit.ttl) return hit.value;
    const value = await producer();
    cache.set(key, {
        at: Date.now(),
        ttl: value ? CACHE_HIT_MS : CACHE_MISS_MS,
        value: value || null
    });
    return value;
}

function readTenantSettings() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const settings = window.ms365TenantSettingsLoad();
            return settings && typeof settings === 'object' ? settings : {};
        }
    } catch {
        /* ignore */
    }
    return {};
}

/**
 * @param {object} settings
 * @param {string} email
 */
function studentKlasseCode(settings, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return '';
    const students = Array.isArray(settings && settings.students) ? settings.students : [];
    const hit = students.find((s) => String((s && s.email) || '').trim().toLowerCase() === em);
    if (!hit) return '';
    return String(hit.klasse || hit.class || hit.klasseCode || '').trim();
}

/**
 * @param {object} settings
 * @param {string} email
 */
function teacherCodeForEmail(settings, email) {
    const em = String(email || '')
        .trim()
        .toLowerCase();
    if (!em) return '';
    const teachers = Array.isArray(settings && settings.teachers) ? settings.teachers : [];
    const hit = teachers.find((t) => String((t && (t.email || t.upn)) || '').trim().toLowerCase() === em);
    return hit ? String(hit.code || '').trim() : '';
}

/**
 * @param {string} email
 * @param {object} settings
 */
async function resolveAktivitaetenRole(email, settings) {
    const { applyAktRoleFromEntra } = await import(
        '../tools/schulaktivitaeten-planer/schulaktivitaeten-planer-entra-role.js'
    );
    const { kvKlasseCodesForEmail } = await import(
        '../tools/schulaktivitaeten-planer/schulaktivitaeten-planer-state.js'
    );
    const state = {
        accountEmail: email,
        stammdaten: settings,
        role: 'lehrer',
        planerRoles: []
    };
    try {
        await applyAktRoleFromEntra(state, { preferShellDefault: true });
    } catch {
        /* Eigene Anträge bleiben über die E-Mail sichtbar. */
    }
    const roles =
        Array.isArray(state.planerRoles) && state.planerRoles.length
            ? state.planerRoles
            : [state.role || 'lehrer'];
    return {
        roles,
        kvKlassen: kvKlasseCodesForEmail(settings, email)
    };
}

async function loadTermineEvents(ctx) {
    const { loadState } = await import('../tools/termine-hub/termine-hub-store.js');
    const local = loadState().termine || [];
    try {
        const { ensureSettingsWithSuggestions, resolveEnabledSources, SOURCE_IDS } = await import(
            '../tools/termine-hub/termine-hub-settings.js'
        );
        const { fetchEventsForSources, signedInEmail } = await import('../tools/termine-hub/termine-hub-outlook.js');
        const settings = ensureSettingsWithSuggestions();
        const audience = ctx.view === 'schueler' ? 'schueler' : 'lehrer';
        let sources = resolveEnabledSources(settings, { audience }).filter((source) => source.visible !== false);
        if (!sources.some((source) => source.id === SOURCE_IDS.ME)) {
            sources.unshift({
                id: SOURCE_IDS.ME,
                label: 'Mein Kalender',
                kind: 'me',
                upn: '',
                visible: true
            });
        }
        const now = new Date();
        const from = new Date(now);
        from.setHours(0, 0, 0, 0);
        const to = new Date(from);
        to.setDate(to.getDate() + 14);
        to.setHours(23, 59, 59, 999);
        const res = await withTimeout(
            fetchEventsForSources(sources, from, to, { signedInEmail: signedInEmail() }),
            PORTAL_BRIEFING_TIMEOUT_MS
        );
        const live = res && Array.isArray(res.events) ? res.events : [];
        if (live.length) return live;
        if (res && Array.isArray(res.errors) && res.errors.length && local.length) return local;
        return live;
    } catch {
        return local;
    }
}

async function loadAufgabenCard(ctx) {
    const email = String(ctx.email || '').trim().toLowerCase();
    if (!email) return null;
    return cached('aufgaben:' + email, async () => {
        const { loadState } = await import('../tools/aufgaben/aufgaben-store.js');
        const local = buildAufgabenBriefing((loadState().aufgaben || []), { email });
        try {
            const { pullAufgabenFromSharePoint } = await import('../tools/aufgaben/aufgaben-sync.js');
            const pulled = await withTimeout(
                pullAufgabenFromSharePoint({ forceDemoIfEmpty: false }),
                PORTAL_BRIEFING_TIMEOUT_MS
            );
            if (pulled && pulled.ok) {
                return buildAufgabenBriefing((pulled.state && pulled.state.aufgaben) || [], { email });
            }
            return local;
        } catch {
            return local;
        }
    });
}

async function loadAktivitaetenCard(ctx) {
    const email = String(ctx.email || '').trim().toLowerCase();
    const view = ctx.view === 'schueler' ? 'schueler' : 'lehrer';
    if (!email) return null;
    return cached('aktivitaeten:' + view + ':' + email, async () => {
        const siteUrl = resolveSchoolToolSiteUrl('aktivitaeten');
        if (!siteUrl) return null;
        const settings = readTenantSettings();
        try {
            const [{ resolveAktContext, loadAllAktData }, role] = await Promise.all([
                import('../tools/schulaktivitaeten-planer/schulaktivitaeten-planer-graph.js'),
                resolveAktivitaetenRole(email, settings)
            ]);
            const graphCtx = await withTimeout(resolveAktContext(siteUrl), PORTAL_BRIEFING_TIMEOUT_MS);
            const data = await withTimeout(loadAllAktData(graphCtx), PORTAL_BRIEFING_TIMEOUT_MS);
            return buildSchulaktivitaetenBriefing(data.items || [], {
                view,
                email,
                roles: role.roles,
                kvKlassen: role.kvKlassen,
                klasseCode: studentKlasseCode(settings, email)
            });
        } catch {
            return null;
        }
    });
}

async function loadSchularbeitenCard(ctx) {
    const email = String(ctx.email || '').trim().toLowerCase();
    const view = ctx.view === 'schueler' ? 'schueler' : 'lehrer';
    if (!email) return null;
    return cached('schularbeiten:' + view + ':' + email, async () => {
        const siteUrl = resolveSchoolToolSiteUrl('schularbeiten');
        if (!siteUrl) return null;
        const settings = readTenantSettings();
        try {
            const { resolvePlanerContext, loadAllPlanerData } = await import(
                '../tools/schularbeiten-planer/schularbeiten-planer-graph.js'
            );
            const { loadSchuljahr } = await import('../tools/schularbeiten-planer/schularbeiten-planer-state.js');
            const graphCtx = await withTimeout(resolvePlanerContext(siteUrl), PORTAL_BRIEFING_TIMEOUT_MS);
            const schuljahr = loadSchuljahr();
            const data = await withTimeout(
                loadAllPlanerData(graphCtx, schuljahr ? { schuljahr } : {}),
                PORTAL_BRIEFING_TIMEOUT_MS
            );
            return buildSchularbeitenBriefing(data.items || [], {
                view,
                email,
                klasseCode: studentKlasseCode(settings, email),
                teacherCode: teacherCodeForEmail(settings, email),
                schuljahr
            });
        } catch {
            return null;
        }
    });
}

async function resolveFreistellungRole(email, settings) {
    const { applyPlanerRoleFromEntra } = await import(
        '../tools/freistellung-planer/freistellung-planer-entra-role.js'
    );
    const { kvKlasseCodesForEmail } = await import(
        '../tools/schulaktivitaeten-planer/schulaktivitaeten-planer-state.js'
    );
    const state = {
        accountEmail: email,
        stammdaten: settings,
        role: 'lehrer',
        planerRoles: []
    };
    try {
        await applyPlanerRoleFromEntra(state, {});
    } catch {
        /* Eigene Anträge bleiben über die E-Mail sichtbar. */
    }
    const roles =
        Array.isArray(state.planerRoles) && state.planerRoles.length
            ? state.planerRoles
            : [state.role || 'lehrer'];
    return {
        roles,
        kvKlassen: kvKlasseCodesForEmail(settings, email)
    };
}

async function loadFreistellungCard(ctx) {
    const email = String(ctx.email || '').trim().toLowerCase();
    const view = ctx.view === 'schueler' ? 'schueler' : 'lehrer';
    if (!email) return null;
    return cached('freistellung:' + view + ':' + email, async () => {
        const siteUrl = resolveSchoolToolSiteUrl('freistellung');
        if (!siteUrl) return null;
        const settings = readTenantSettings();
        try {
            const [{ resolveFrContext, loadAllFreistellungen }, role] = await Promise.all([
                import('../tools/freistellung-planer/freistellung-planer-graph.js'),
                resolveFreistellungRole(email, settings)
            ]);
            const graphCtx = await withTimeout(resolveFrContext(siteUrl), PORTAL_BRIEFING_TIMEOUT_MS);
            const items = await withTimeout(loadAllFreistellungen(graphCtx), PORTAL_BRIEFING_TIMEOUT_MS);
            return buildFreistellungBriefing(items || [], {
                view,
                email,
                roles: role.roles,
                kvKlassen: role.kvKlassen
            });
        } catch {
            return null;
        }
    });
}

async function loadTechnikCard(ctx) {
    const email = String(ctx.email || '').trim().toLowerCase();
    if (!email) return null;
    return cached('technik:' + email, async () => {
        const { loadState } = await import('../tools/technik-meldungen/technik-meldungen-store.js');
        const local = buildTechnikBriefing(loadState().meldungen || [], { email });
        const siteUrl = resolveSchoolToolSiteUrl('technik');
        if (!siteUrl) return local;
        try {
            const { pullTechnikFromSharePoint } = await import(
                '../tools/technik-meldungen/technik-meldungen-sync.js'
            );
            const pulled = await withTimeout(
                pullTechnikFromSharePoint({ forceDemoIfEmpty: false }),
                PORTAL_BRIEFING_TIMEOUT_MS
            );
            if (pulled && pulled.ok) {
                return buildTechnikBriefing((pulled.state && pulled.state.meldungen) || [], { email });
            }
            return local;
        } catch {
            return local;
        }
    });
}

async function loadElternbriefCard(ctx) {
    const email = String(ctx.email || '').trim().toLowerCase();
    if (!email) return null;
    return cached('elternbrief:' + email, async () => {
        const { loadState } = await import('../tools/elternbrief/elternbrief-store.js');
        const localState = loadState();
        const local = buildElternbriefBriefing(localState.briefe || [], localState.empfaenger || [], {
            email
        });
        const siteUrl = resolveSchoolToolSiteUrl('elternbrief');
        if (!siteUrl) return local;
        try {
            const { pullElternbriefFromSharePoint } = await import('../tools/elternbrief/elternbrief-sync.js');
            const pulled = await withTimeout(
                pullElternbriefFromSharePoint({ forceDemoIfEmpty: false }),
                PORTAL_BRIEFING_TIMEOUT_MS
            );
            if (pulled && pulled.ok) {
                const state = pulled.state || {};
                return buildElternbriefBriefing(state.briefe || [], state.empfaenger || [], { email });
            }
            return local;
        } catch {
            return local;
        }
    });
}

/**
 * @param {{ view?: string, email?: string, isVisible?: (id: string) => boolean }} ctx
 * @param {(cards: object[]) => void} [onUpdate]
 */
export async function loadPortalBriefings(ctx, onUpdate) {
    const c = ctx || {};
    const visible = typeof c.isVisible === 'function' ? c.isVisible : () => false;
    const view = c.view === 'schueler' ? 'schueler' : 'lehrer';
    /** @type {Record<string, object|object[]|string|undefined>} */
    const bag = {
        view,
        termine: undefined,
        aktivitaeten: undefined,
        freistellung: undefined,
        aufgaben: undefined,
        schularbeiten: undefined,
        technik: undefined,
        elternbrief: undefined
    };
    const emit = () => {
        if (typeof onUpdate === 'function') onUpdate(expandPortalWidgets(bag));
    };

    /** @type {Promise<void>[]} */
    const jobs = [];
    if (visible('termine-hub')) {
        bag.termine = PORTAL_WIDGET_PENDING;
        jobs.push(
            loadTermineEvents(c).then((events) => {
                bag.termine = events || [];
                emit();
            })
        );
    }
    if (visible('schulaktivitaeten-planer')) {
        bag.aktivitaeten = PORTAL_WIDGET_PENDING;
        jobs.push(
            loadAktivitaetenCard(c).then((card) => {
                bag.aktivitaeten = card;
                emit();
            })
        );
    }
    if (visible('freistellung-planer')) {
        bag.freistellung = PORTAL_WIDGET_PENDING;
        jobs.push(
            loadFreistellungCard(c).then((card) => {
                bag.freistellung = card;
                emit();
            })
        );
    }
    if (visible('aufgaben')) {
        bag.aufgaben = PORTAL_WIDGET_PENDING;
        jobs.push(
            loadAufgabenCard(c).then((card) => {
                bag.aufgaben = card;
                emit();
            })
        );
    }
    if (visible('schularbeiten-planer')) {
        bag.schularbeiten = PORTAL_WIDGET_PENDING;
        jobs.push(
            loadSchularbeitenCard(c).then((card) => {
                bag.schularbeiten = card;
                emit();
            })
        );
    }
    if (visible('technik-meldungen')) {
        bag.technik = PORTAL_WIDGET_PENDING;
        jobs.push(
            loadTechnikCard(c).then((card) => {
                bag.technik = card;
                emit();
            })
        );
    }
    if (visible('elternbrief')) {
        bag.elternbrief = PORTAL_WIDGET_PENDING;
        jobs.push(
            loadElternbriefCard(c).then((card) => {
                bag.elternbrief = card;
                emit();
            })
        );
    }
    await Promise.all(jobs);
    const cards = expandPortalWidgets(bag);
    emit();
    return cards;
}
