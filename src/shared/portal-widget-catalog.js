/**
 * Kacheln, die eine Person auf die Startseite legen kann.
 * Mehrere Kacheln dürfen dasselbe Werkzeug unterschiedlich zeigen.
 */

/** 12 Spalten: 4 ist ein Drittel, 2 Zeilen sind die kurze Zahl, 3 die kurze Liste. */
const W_THIRD = 4;
const H_LIST = 3;
const H_STAT = 2;

export const PORTAL_WIDGETS = [
    { id: 'termine-heute', toolId: 'termine-hub', group: 'Kalender', title: 'Heute', blurb: 'Termine von heute aus Outlook.', w: W_THIRD, h: H_LIST },
    { id: 'termine-liste', toolId: 'termine-hub', group: 'Kalender', title: 'Nächste Termine', blurb: 'Die kommenden 14 Tage als Liste.', w: W_THIRD, h: H_LIST },
    { id: 'termine-woche', toolId: 'termine-hub', group: 'Kalender', title: 'Diese Woche', blurb: 'Montag bis Freitag nebeneinander.', w: 8, h: H_LIST, presentation: 'week' },
    { id: 'termine-zahl', toolId: 'termine-hub', group: 'Kalender', title: 'Termine', blurb: 'Anzahl in den nächsten 14 Tagen.', w: W_THIRD, h: H_STAT, presentation: 'stat' },
    { id: 'aktivitaeten-liste', toolId: 'schulaktivitaeten-planer', group: 'Schulalltag', title: 'Schulaktivitäten', blurb: 'Eigene Anträge und offene Entscheidungen.', w: W_THIRD, h: H_LIST },
    { id: 'aktivitaeten-zahl', toolId: 'schulaktivitaeten-planer', group: 'Schulalltag', title: 'Aktivitäten', blurb: 'Wie viele Aktivitäten gerade offen sind.', w: W_THIRD, h: H_STAT, presentation: 'stat' },
    { id: 'freistellung-liste', toolId: 'freistellung-planer', group: 'Schulalltag', title: 'Freistellungen', blurb: 'Offene Anträge mit Datum und Klasse.', w: W_THIRD, h: H_LIST },
    { id: 'freistellung-zahl', toolId: 'freistellung-planer', group: 'Schulalltag', title: 'Freistellungen', blurb: 'Anzahl der offenen Freistellungen.', w: W_THIRD, h: H_STAT, presentation: 'stat' },
    { id: 'schularbeiten-liste', toolId: 'schularbeiten-planer', group: 'Schulalltag', title: 'Schularbeiten', blurb: 'Nächste eigene oder klassenbezogene Arbeiten.', w: W_THIRD, h: H_LIST },
    { id: 'schularbeiten-zahl', toolId: 'schularbeiten-planer', group: 'Schulalltag', title: 'Schularbeiten', blurb: 'Anzahl der anstehenden Arbeiten.', w: W_THIRD, h: H_STAT, presentation: 'stat' },
    { id: 'aufgaben-liste', toolId: 'aufgaben', group: 'Aufgaben', title: 'Aufgaben', blurb: 'Offene Aufgaben an die eigene E-Mail.', w: W_THIRD, h: H_LIST },
    { id: 'aufgaben-zahl', toolId: 'aufgaben', group: 'Aufgaben', title: 'Aufgaben', blurb: 'Anzahl der offenen Aufgaben.', w: W_THIRD, h: H_STAT, presentation: 'stat' },
    { id: 'technik-liste', toolId: 'technik-meldungen', group: 'Betrieb', title: 'Technik', blurb: 'Eigene offene Meldungen.', w: W_THIRD, h: H_LIST },
    { id: 'technik-zahl', toolId: 'technik-meldungen', group: 'Betrieb', title: 'Technik', blurb: 'Anzahl der eigenen offenen Meldungen.', w: W_THIRD, h: H_STAT, presentation: 'stat' },
    { id: 'elternbrief-liste', toolId: 'elternbrief', group: 'Eltern', title: 'Elternbriefe', blurb: 'Eigene Entwürfe und ungelesene Briefe.', w: W_THIRD, h: H_LIST },
    { id: 'elternbrief-zahl', toolId: 'elternbrief', group: 'Eltern', title: 'Elternbriefe', blurb: 'Anzahl der offenen Briefe.', w: W_THIRD, h: H_STAT, presentation: 'stat' }
];

export const PORTAL_WIDGET_DEFAULT_IDS = [
    'termine-heute',
    'termine-liste',
    'aktivitaeten-liste',
    'freistellung-liste',
    'elternbrief-liste',
    'aufgaben-liste',
    'schularbeiten-liste',
    'technik-liste'
];

/** Alte Kachel-IDs aus der ersten Raster-Version. */
export const PORTAL_WIDGET_LEGACY_IDS = {
    'termine-hub': 'termine-liste',
    'schulaktivitaeten-planer': 'aktivitaeten-liste',
    'freistellung-planer': 'freistellung-liste',
    aufgaben: 'aufgaben-liste',
    'schularbeiten-planer': 'schularbeiten-liste',
    'technik-meldungen': 'technik-liste',
    elternbrief: 'elternbrief-liste'
};

/**
 * @param {string} id
 */
export function canonicalizeWidgetId(id) {
    const key = String(id || '').trim();
    return PORTAL_WIDGET_LEGACY_IDS[key] || key;
}

/**
 * @param {string} id
 */
export function portalWidgetById(id) {
    const key = canonicalizeWidgetId(id);
    return PORTAL_WIDGETS.find((widget) => widget.id === key) || null;
}

/**
 * @returns {Record<string, { w: number, h: number }>}
 */
export function portalWidgetSizes() {
    /** @type {Record<string, { w: number, h: number }>} */
    const sizes = {};
    PORTAL_WIDGETS.forEach((widget) => {
        sizes[widget.id] = { w: widget.w, h: widget.h };
    });
    return sizes;
}
