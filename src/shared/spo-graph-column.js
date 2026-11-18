/**
 * Graph columnDefinition-Body für SharePoint-Listen.
 * Beachtet Limits: Single-Line maxLength ≤ 255; Mehrzeiler mit textType, ohne maxLength.
 *
 * @typedef {{
 *   name: string,
 *   displayName?: string,
 *   text?: { allowMultipleLines?: boolean, maxLength?: number, textType?: string },
 *   number?: { decimalPlaces?: string },
 *   boolean?: object,
 *   dateTime?: { displayAs?: string, format?: string },
 *   choice?: { allowTextEntry?: boolean, choices?: string[] }
 * }} GraphColumnDef
 */

/**
 * @param {GraphColumnDef} def
 * @returns {Record<string, unknown>}
 */
export function toGraphColumnBody(def) {
    const body = {
        name: def.name,
        displayName: def.displayName || def.name
    };
    if (def.text) {
        const multi = !!def.text.allowMultipleLines;
        body.text = {
            allowMultipleLines: multi,
            textType: def.text.textType || 'plain'
        };
        if (!multi) {
            const max = Number(def.text.maxLength);
            body.text.maxLength = Number.isFinite(max) ? Math.min(255, Math.max(1, max)) : 255;
        }
    } else if (def.number) {
        body.number = {
            decimalPlaces: def.number.decimalPlaces || 'automatic'
        };
    } else if (def.boolean) {
        body.boolean = {};
    } else if (def.dateTime) {
        body.dateTime = {
            displayAs: def.dateTime.displayAs || 'default',
            format: def.dateTime.format || 'dateOnly'
        };
    } else if (def.choice) {
        body.choice = {
            allowTextEntry: !!def.choice.allowTextEntry,
            choices: Array.isArray(def.choice.choices) ? def.choice.choices.slice() : []
        };
    }
    return body;
}
