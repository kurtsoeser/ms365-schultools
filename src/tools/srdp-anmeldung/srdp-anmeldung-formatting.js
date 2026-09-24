/**
 * SharePoint-Formatierung & Formular-Logik für sRDP/sRP (reine Funktionen).
 */
import { VARIANT_LABELS, getProfile } from './srdp-anmeldung-schema.js';

/** Feste Farben für Variante 1–3. */
export const VARIANT_COLORS = {
    [VARIANT_LABELS[1]]: { bg: '#f4a261', fg: '#1a1a1a' },
    [VARIANT_LABELS[2]]: { bg: '#e76f51', fg: '#ffffff' },
    [VARIANT_LABELS[3]]: { bg: '#2a9d8f', fg: '#ffffff' }
};

/** Palette für Choice-Pills (Klasse, Wahlfach, …). */
export const PILL_PALETTE = [
    { bg: '#dbeafe', fg: '#1e3a5f' },
    { bg: '#dcfce7', fg: '#14532d' },
    { bg: '#fef3c7', fg: '#78350f' },
    { bg: '#fce7f3', fg: '#831843' },
    { bg: '#e0e7ff', fg: '#312e81' },
    { bg: '#ffedd5', fg: '#7c2d12' },
    { bg: '#ccfbf1', fg: '#134e4a' },
    { bg: '#f3e8ff', fg: '#581c87' }
];

export function buildSeminarShowFormula() {
    return "=if(indexOf([$WahlfachMuendlich],'Seminar')==0,'true','false')";
}

export function buildTitleHideFormula() {
    return '=false';
}

export function buildAlwaysHideFormula() {
    return '=false';
}

export function buildTitleColumnFormatJson() {
    return {
        $schema: 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
        elmType: 'div',
        style: { 'font-weight': '600' },
        txtContent:
            "=if([$Nachname]!='' && [$Vorname]!='', [$Nachname]+', '+[$Vorname], if([$Nachname]!='',[$Nachname], if([$Vorname]!='',[$Vorname], [$Title])))"
    };
}

/**
 * @param {string} [profileId]
 */
export function buildPruefplanHeaderExpression(profileId) {
    const profile = getProfile(profileId);
    const parts = profile.variants.map((v) => {
        const text =
            'schriftlich: ' + v.schriftlich.join(', ') + ' · mündlich: ' + v.muendlich.join(', ');
        return "if([$Variante]=='" + v.label + "','" + text.replace(/'/g, "''") + "',";
    });
    return '=' + parts.join('') + "'Variante wählen …'" + ')'.repeat(profile.variants.length);
}

/**
 * @param {'variante'|'choice'} mode
 * @param {string} [profileId]
 */
export function buildPillColumnFormatJson(mode, profileId) {
    if (mode !== 'variante') {
        return buildKlasseColumnFormatJson();
    }
    const profile = getProfile(profileId);
    const children = profile.variants.map((v) => {
        const c = VARIANT_COLORS[v.label] || { bg: '#e2e8f0', fg: '#0f172a' };
        return {
            elmType: 'div',
            style: {
                display: "=if([$Variante]=='" + v.label + "','flex','none')",
                'align-items': 'center',
                'justify-content': 'center',
                'min-height': '24px',
                padding: '2px 10px',
                'border-radius': '12px',
                'background-color': c.bg,
                color: c.fg,
                'font-size': '12px',
                'font-weight': '600',
                'white-space': 'nowrap'
            },
            txtContent: '[$Variante]'
        };
    });
    return {
        $schema: 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
        elmType: 'div',
        children
    };
}

export function buildWahlfachColumnFormatJson() {
    return {
        $schema: 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
        elmType: 'div',
        style: {
            display: 'inline-flex',
            'align-items': 'center',
            'min-height': '24px',
            padding: '2px 10px',
            'border-radius': '12px',
            'font-size': '12px',
            'font-weight': '600',
            'white-space': 'nowrap',
            'background-color':
                "=if(indexOf(@currentField,'Seminar')==0,'#ffedd5', if(indexOf(@currentField,'Recht')==0,'#dbeafe', if(indexOf(@currentField,'Religion')==0,'#e0e7ff', if(indexOf(@currentField,'Geschichte')==0,'#ffedd5', if(indexOf(@currentField,'Sport')==0,'#dcfce7','#e2e8f0')))))",
            color:
                "=if(indexOf(@currentField,'Seminar')==0,'#9a3412', if(indexOf(@currentField,'Recht')==0,'#1e3a5f', if(indexOf(@currentField,'Religion')==0,'#312e81', if(indexOf(@currentField,'Geschichte')==0,'#9a3412', if(indexOf(@currentField,'Sport')==0,'#14532d','#0f172a')))))"
        },
        txtContent: '@currentField'
    };
}

export function buildKlasseColumnFormatJson() {
    return {
        $schema: 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
        elmType: 'div',
        style: {
            display: 'inline-flex',
            'align-items': 'center',
            'min-height': '24px',
            padding: '2px 10px',
            'border-radius': '12px',
            'background-color': '#dcfce7',
            color: '#14532d',
            'font-size': '12px',
            'font-weight': '600',
            'white-space': 'nowrap'
        },
        txtContent: '@currentField'
    };
}

/**
 * @param {string} [profileId]
 */
export function buildClientFormCustomFormatterObject(profileId) {
    const profile = getProfile(profileId);
    const pruefExpr = buildPruefplanHeaderExpression(profile.id);
    const titleFallback = 'Anmeldung ' + profile.examShort + ' ' + profile.label;
    return {
        headerJSONFormatter: {
            elmType: 'div',
            style: {
                width: '100%',
                padding: '12px 4px 16px',
                'border-bottom': '1px solid #e2e8f0',
                'margin-bottom': '8px'
            },
            children: [
                {
                    elmType: 'div',
                    style: {
                        'font-size': '20px',
                        'font-weight': '700',
                        'margin-bottom': '6px',
                        color: '#0f172a'
                    },
                    txtContent:
                        "=if([$Nachname]!='' && [$Vorname]!='', [$Nachname]+', '+[$Vorname], '" +
                        titleFallback.replace(/'/g, "''") +
                        "')"
                },
                {
                    elmType: 'div',
                    style: {
                        'font-size': '13px',
                        color: '#475569',
                        'line-height': '1.4'
                    },
                    txtContent: pruefExpr
                }
            ]
        },
        footerJSONFormatter: {
            elmType: 'div',
            style: {
                'font-size': '12px',
                color: '#64748b',
                padding: '8px 4px 0'
            },
            txtContent:
                'Mit der Bestätigung meldest du dich zum Haupttermin an. Seminar nur ausfüllen, wenn Wahlfach = Seminar.'
        },
        bodyJSONFormatter: {
            sections: profile.formSections.slice()
        }
    };
}

/**
 * @param {string} [profileId]
 */
export function buildClientFormCustomFormatterString(profileId) {
    return JSON.stringify(buildClientFormCustomFormatterObject(profileId));
}

/**
 * @param {string|number} [terminJahr]
 * @param {string} [profileId]
 */
export function buildFieldFormatSpecs(terminJahr, profileId) {
    const jahr = String(terminJahr || '').trim();
    const profile = getProfile(profileId);
    const titleFmt = buildTitleColumnFormatJson();
    /** @type {Array<{ internalName: string, displayName?: string, conditionalShowFormula?: string, customFormatter?: object, required?: boolean, defaultValue?: string }>} */
    const specs = [
        {
            internalName: 'Title',
            conditionalShowFormula: buildTitleHideFormula(),
            customFormatter: titleFmt,
            required: false,
            displayName: 'Name (aus Nachname, Vorname)'
        },
        {
            internalName: 'Seminar',
            conditionalShowFormula: buildSeminarShowFormula()
        },
        {
            internalName: 'Schulform',
            conditionalShowFormula: buildAlwaysHideFormula(),
            defaultValue: profile.label
        },
        {
            internalName: 'TerminJahr',
            conditionalShowFormula: buildAlwaysHideFormula(),
            defaultValue: /^\d{4}$/.test(jahr) ? jahr : undefined
        },
        {
            internalName: 'PruefplanKurz',
            conditionalShowFormula: buildAlwaysHideFormula()
        },
        {
            internalName: 'Variante',
            customFormatter: buildPillColumnFormatJson('variante', profile.id)
        },
        {
            internalName: 'Klasse',
            customFormatter: buildKlasseColumnFormatJson()
        },
        {
            internalName: 'WahlfachMuendlich',
            customFormatter: buildWahlfachColumnFormatJson()
        }
    ];

    if (profile.id === 'ahs') {
        specs.push({
            internalName: 'TitelABA',
            conditionalShowFormula: "=if([$HatABA]==true,'true','false')"
        });
        specs.push({
            internalName: 'BetreuungslehrerABA',
            conditionalShowFormula: "=if([$HatABA]==true,'true','false')"
        });
        specs.push({
            internalName: 'Muendlich3',
            conditionalShowFormula: "=if([$Variante]=='Variante 2','false','true')"
        });
        specs.push({
            internalName: 'LehrerMuendlich3',
            conditionalShowFormula: "=if([$Variante]=='Variante 2','false','true')"
        });
    }

    if (profile.extraColumnNames.indexOf('KlausurKombi') !== -1) {
        specs.push({
            internalName: 'KlausurKombi',
            conditionalShowFormula: "=if([$Variante]=='Variante 1','true','false')"
        });
    }

    return specs;
}
