/**
 * Kursteam-Vorlagen Phase 1: Kanäle (ohne Graph/DOM).
 */

/** @typedef {{ id: string, displayName: string }} TemplateChannel */
/**
 * @typedef {{
 *   id: string,
 *   name: string,
 *   schoolForm: string,
 *   subjectCode: string,
 *   schulstufe: string,
 *   semester: string,
 *   description: string,
 *   channels: TemplateChannel[],
 *   updatedAt: string
 * }} ChannelTemplate
 */
/** @typedef {{ id: string, displayName: string, membershipType?: string }} TeamChannel */
/** @typedef {'create'|'ok'|'rename'|'skip_general'|'extra'} DiffStatus */
/** @typedef {{ status: DiffStatus, templateChannel: TemplateChannel|null, teamChannel: TeamChannel|null, message: string }} DiffRow */
/**
 * @typedef {{
 *   type: 'group'|'template',
 *   key: string,
 *   label: string,
 *   count: number,
 *   template?: ChannelTemplate,
 *   badge?: string,
 *   children?: TreeNode[]
 * }} TreeNode
 */

export const STORAGE_KIND = 'ms365-kursteam-templates';
export const EXPORT_KIND = 'ms365-kursteam-templates';
export const EXPORT_VERSION = 3;

/** Bekannte Schulformen (AT) – Freitext bleibt erlaubt. */
export const KNOWN_SCHOOL_FORMS = [
    'AHS',
    'Mittelschule',
    'HAK',
    'HAKB',
    'HAS',
    'HTL',
    'HLW',
    'BAfEP',
    'PTS',
    'Berufsschule',
    'Sonderpädagogik'
];

/** Semester-Kennungen: WS / SS / SJ (ganzes Schuljahr). */
export const KNOWN_SEMESTERS = ['SJ', 'WS', 'SS'];

const GENERAL_NAMES = new Set(['allgemein', 'general']);

export function normStr(v) {
    return String(v == null ? '' : v).trim();
}

export function normName(v) {
    return normStr(v).toLowerCase().replace(/\s+/g, ' ');
}

export function isGeneralChannelName(name) {
    const n = normName(name);
    if (GENERAL_NAMES.has(n)) return true;
    // Umbenannter Standardkanal, z. B. „00-Allgemein“
    if (/^00\s*[-–.]?\s*allgemein$/.test(n)) return true;
    if (/^00\s*[-–.]?\s*general$/.test(n)) return true;
    return false;
}

export function newId(prefix) {
    const p = normStr(prefix) || 'id';
    const rand =
        typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
            ? crypto.randomUUID().replace(/-/g, '').slice(0, 12)
            : String(Date.now()) + String(Math.floor(Math.random() * 1e6));
    return p + '-' + rand;
}

/**
 * Führende Nummer aus Kanalnamen, z. B. "01 - Grundlagen" → "01".
 * @param {string} name
 * @returns {string}
 */
export function channelNumberPrefix(name) {
    const m = normStr(name).match(/^(\d+)\s*[-–.:)]/);
    return m ? m[1] : '';
}

/**
 * @param {Partial<TemplateChannel>|string} raw
 * @returns {TemplateChannel}
 */
export function normalizeChannel(raw) {
    if (typeof raw === 'string') {
        return { id: newId('ch'), displayName: normStr(raw) };
    }
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        id: normStr(o.id) || newId('ch'),
        displayName: normStr(o.displayName || o.name || '')
    };
}

/**
 * Österreichische Schulstufe (1–13) normalisieren.
 * Akzeptiert auch Altlasten: „Modul 3“, „Jahrgang 5“, „3. Klasse“.
 * @param {unknown} raw
 * @param {string} [nameHint]
 * @returns {string} z. B. "9"
 */
export function normalizeSchulstufe(raw, nameHint) {
    let m = normStr(raw);
    if (m) {
        const prefixed =
            m.match(/^schulstufe\s*(\d{1,2})\b/i) ||
            m.match(/^stufe\s*(\d{1,2})\b/i) ||
            m.match(/^modul\s*(\d{1,2})\b/i) ||
            m.match(/^jahrgang\s*(\d{1,2})\b/i) ||
            m.match(/^(\d{1,2})\.\s*klasse\b/i) ||
            m.match(/^(\d{1,2})$/);
        if (prefixed) return String(Number(prefixed[1]));
        const digits = m.match(/(\d{1,2})/);
        if (digits) return String(Number(digits[1]));
        return m;
    }
    const name = normStr(nameHint);
    const fromName =
        name.match(/\bschulstufe\s*(\d{1,2})\b/i) ||
        name.match(/\bmodul\s*(\d{1,2})\b/i) ||
        name.match(/\bjahrgang\s*(\d{1,2})\b/i) ||
        name.match(/\b(\d{1,2})\.\s*klasse\b/i);
    return fromName ? String(Number(fromName[1])) : '';
}

/** @deprecated Bitte normalizeSchulstufe verwenden. */
export function normalizeModule(raw, nameHint) {
    return normalizeSchulstufe(raw, nameHint);
}

/**
 * @param {string} schulstufe
 * @returns {string}
 */
export function schulstufeLabel(schulstufe) {
    const m = normStr(schulstufe);
    if (!m) return 'Ohne Schulstufe';
    if (/^\d+$/.test(m)) return 'Schulstufe ' + m;
    return m;
}

/** @deprecated Bitte schulstufeLabel verwenden. */
export function moduleLabel(module) {
    return schulstufeLabel(module);
}

/**
 * Klasse innerhalb der Schulform → österr. Schulstufe.
 * AHS/Mittelschule: 1. Klasse = Stufe 5 … 8. Klasse = Stufe 12
 * BHS (HAK/HTL/…): 1. Klasse = Stufe 9 … 5. Klasse = Stufe 13
 * @param {string} schoolForm
 * @param {number|string} klasseNr
 * @returns {string}
 */
export function klasseToSchulstufe(schoolForm, klasseNr) {
    const k = Number(klasseNr);
    if (!Number.isFinite(k) || k < 1) return '';
    const sf = normalizeSchoolForm(schoolForm);
    if (sf === 'AHS' || sf === 'Mittelschule') return String(4 + k);
    if (sf === 'HAK' || sf === 'HAKB' || sf === 'HAS' || sf === 'HTL' || sf === 'HLW' || sf === 'BAfEP') {
        return String(8 + k);
    }
    if (sf === 'PTS') return String(8 + k);
    return String(k);
}

/**
 * HAK/HAKB/HAS-Lehrplan-„Modul“ (Halbjahr) → Schulstufe + Semester.
 * Modul 1 = Stufe 9 WS, 2 = 9 SS, 3 = 10 WS, 4 = 10 SS, …
 * @param {number|string} modulNr
 * @returns {{ schulstufe: string, semester: 'WS'|'SS' }|null}
 */
export function hakLehrplanModulToMeta(modulNr) {
    const n = Number(normalizeSchulstufe(modulNr, ''));
    if (!Number.isFinite(n) || n < 1) return null;
    return {
        schulstufe: String(8 + Math.ceil(n / 2)),
        semester: n % 2 === 1 ? 'WS' : 'SS'
    };
}

/**
 * @param {unknown} raw
 * @returns {string} '' | 'WS' | 'SS' | 'SJ'
 */
export function normalizeSemester(raw) {
    const s = normStr(raw).toLowerCase();
    if (!s) return '';
    if (
        s === 'sj' ||
        s === 'schuljahr' ||
        s === 'ganzjahr' ||
        s === 'ganzjährig' ||
        s === 'ganzjaehrig' ||
        s === 'ws+ss' ||
        s === 'ws/ss'
    ) {
        return 'SJ';
    }
    if (s === 'ws' || s === 'wintersemester' || s === 'winter') return 'WS';
    if (s === 'ss' || s === 'sommersemester' || s === 'sommer') return 'SS';
    return normStr(raw).toUpperCase();
}

/**
 * @param {string} semester
 * @returns {string}
 */
export function semesterLabel(semester) {
    const s = normalizeSemester(semester);
    if (s === 'WS') return 'Wintersemester';
    if (s === 'SS') return 'Sommersemester';
    if (s === 'SJ') return 'Schuljahr (WS+SS)';
    return s || 'Ohne Semester';
}

/**
 * Schulform normalisieren (bekannte Kürzel vereinheitlichen).
 * @param {unknown} raw
 * @returns {string}
 */
export function normalizeSchoolForm(raw) {
    const s = normStr(raw);
    if (!s) return '';
    const map = {
        ahs: 'AHS',
        nms: 'Mittelschule',
        ms: 'Mittelschule',
        mittelschule: 'Mittelschule',
        'neue mittelschule': 'Mittelschule',
        hak: 'HAK',
        hakb: 'HAKB',
        'hak b': 'HAKB',
        has: 'HAS',
        htl: 'HTL',
        hlw: 'HLW',
        bafep: 'BAfEP',
        bakip: 'BAfEP',
        pts: 'PTS',
        polytechnisch: 'PTS',
        'polytechnische schule': 'PTS',
        berufsschule: 'Berufsschule',
        bs: 'Berufsschule'
    };
    const hit = map[s.toLowerCase()];
    if (hit) return hit;
    // bereits bekanntes Label
    const known = KNOWN_SCHOOL_FORMS.find((k) => k.toLowerCase() === s.toLowerCase());
    return known || s;
}

/**
 * @param {string} schoolForm
 * @returns {string}
 */
export function schoolFormLabel(schoolForm) {
    const s = normalizeSchoolForm(schoolForm);
    return s || 'Ohne Schulform';
}

/**
 * @param {Partial<ChannelTemplate>} raw
 * @returns {ChannelTemplate}
 */
export function normalizeTemplate(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const name = normStr(o.name) || 'Unbenannte Vorlage';
    const channels = (Array.isArray(o.channels) ? o.channels : [])
        .map(normalizeChannel)
        .filter((c) => c.displayName && !isGeneralChannelName(c.displayName));
    const schoolForm = normalizeSchoolForm(o.schoolForm || o.schulform || o.schoolType || o.schultyp);
    // Explizite Schulstufe hat Vorrang.
    // klasse/classYear = Klassennummer der Schulform.
    // module/modul (HAK/HAKB/HAS) = Lehrplan-Halbjahr → Stufe + WS/SS.
    const explicitStufe = o.schulstufe ?? o.gradeLevel ?? o.stufe ?? o.yearLevel;
    let schulstufe = '';
    let semester = normalizeSemester(o.semester ?? o.sem);
    if (explicitStufe != null && normStr(explicitStufe) !== '') {
        schulstufe = normalizeSchulstufe(explicitStufe, name);
    } else if (o.klasse != null && normStr(o.klasse) !== '') {
        schulstufe = klasseToSchulstufe(schoolForm, o.klasse) || normalizeSchulstufe(o.klasse, name);
    } else if (o.classYear != null && normStr(o.classYear) !== '') {
        schulstufe = klasseToSchulstufe(schoolForm, o.classYear) || normalizeSchulstufe(o.classYear, name);
    } else if (o.module != null || o.modul != null) {
        const modulRaw = o.module ?? o.modul;
        const sf = schoolForm;
        if (sf === 'HAK' || sf === 'HAKB' || sf === 'HAS') {
            const meta = hakLehrplanModulToMeta(modulRaw);
            if (meta) {
                schulstufe = meta.schulstufe;
                if (!semester) semester = meta.semester;
            } else {
                schulstufe = normalizeSchulstufe(modulRaw, name);
            }
        } else {
            const nr = normalizeSchulstufe(modulRaw, name);
            schulstufe = klasseToSchulstufe(schoolForm, nr) || nr;
        }
    } else {
        schulstufe = normalizeSchulstufe('', name);
    }
    return {
        id: normStr(o.id) || newId('tpl'),
        name,
        schoolForm,
        subjectCode: normStr(o.subjectCode || o.subjectHint || o.fach || '').toUpperCase(),
        schulstufe,
        semester,
        description: normStr(o.description),
        channels,
        updatedAt: normStr(o.updatedAt) || new Date().toISOString()
    };
}

/**
 * @param {unknown} list
 * @returns {ChannelTemplate[]}
 */
export function normalizeTemplateList(list) {
    if (!Array.isArray(list)) return [];
    return list.map(normalizeTemplate);
}

/**
 * @param {string} name
 * @param {string} [subjectCode]
 * @param {string} [description]
 * @param {string} [schulstufe]
 * @param {string} [schoolForm]
 * @param {string} [semester]
 * @returns {ChannelTemplate}
 */
export function createEmptyTemplate(name, subjectCode, description, schulstufe, schoolForm, semester) {
    return normalizeTemplate({
        id: newId('tpl'),
        name: name || 'Neue Vorlage',
        schoolForm: schoolForm || '',
        subjectCode: subjectCode || '',
        schulstufe: schulstufe || '',
        semester: semester || '',
        description: description || '',
        channels: [],
        updatedAt: new Date().toISOString()
    });
}

/**
 * @param {ChannelTemplate} tpl
 * @returns {ChannelTemplate}
 */
export function cloneTemplate(tpl) {
    const src = normalizeTemplate(tpl);
    return normalizeTemplate({
        ...src,
        id: newId('tpl'),
        name: src.name + ' (Kopie)',
        channels: src.channels.map((c) => ({ id: newId('ch'), displayName: c.displayName })),
        updatedAt: new Date().toISOString()
    });
}

/**
 * @param {ChannelTemplate[]} templates
 * @param {{ schoolForm?: string, subjectCode?: string, schulstufe?: string, semester?: string, module?: string }} [filters]
 * @returns {ChannelTemplate[]}
 */
export function filterTemplates(templates, filters) {
    const f = filters || {};
    const school = normalizeSchoolForm(f.schoolForm);
    const code = normStr(f.subjectCode).toUpperCase();
    const stufe = normalizeSchulstufe(f.schulstufe ?? f.module, '');
    const sem = normalizeSemester(f.semester);
    let list = normalizeTemplateList(templates);
    if (school) {
        list = list.filter((t) => !t.schoolForm || normalizeSchoolForm(t.schoolForm) === school);
    }
    if (code) {
        list = list.filter((t) => !t.subjectCode || t.subjectCode === code);
    }
    if (stufe) {
        list = list.filter((t) => !t.schulstufe || normalizeSchulstufe(t.schulstufe, '') === stufe);
    }
    if (sem) {
        list = list.filter((t) => !t.semester || normalizeSemester(t.semester) === sem);
    }
    return list.slice().sort(compareTemplates);
}

/**
 * @param {ChannelTemplate[]} templates
 * @param {string} subjectFilter leer = alle
 * @returns {ChannelTemplate[]}
 */
export function filterTemplatesBySubject(templates, subjectFilter) {
    return filterTemplates(templates, { subjectCode: subjectFilter });
}

/**
 * Einzigartige Schulformen aus Vorlagen + Katalog (+ bekannte Liste).
 * @param {ChannelTemplate[]} templates
 * @param {string[]} [catalog]
 * @returns {string[]}
 */
export function collectSchoolForms(templates, catalog) {
    const set = new Set(KNOWN_SCHOOL_FORMS);
    for (const s of Array.isArray(catalog) ? catalog : []) {
        const n = normalizeSchoolForm(s);
        if (n) set.add(n);
    }
    for (const t of normalizeTemplateList(templates)) {
        if (t.schoolForm) set.add(normalizeSchoolForm(t.schoolForm));
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'de'));
}

/**
 * @param {string[]} catalog
 * @param {string} schoolForm
 * @returns {string[]}
 */
export function rememberSchoolForm(catalog, schoolForm) {
    const n = normalizeSchoolForm(schoolForm);
    if (!n) return Array.isArray(catalog) ? catalog.slice() : [];
    const set = new Set(
        (Array.isArray(catalog) ? catalog : []).map(normalizeSchoolForm).filter(Boolean)
    );
    set.add(n);
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'de'));
}

/**
 * Schulform in allen Vorlagen umbenennen.
 * @param {ChannelTemplate[]} templates
 * @param {string} from
 * @param {string} to
 * @returns {ChannelTemplate[]}
 */
export function renameSchoolFormInTemplates(templates, from, to) {
    const src = normalizeSchoolForm(from);
    const dst = normalizeSchoolForm(to);
    if (!src || !dst || src === dst) return normalizeTemplateList(templates);
    return normalizeTemplateList(templates).map((t) => {
        if (normalizeSchoolForm(t.schoolForm) !== src) return t;
        return normalizeTemplate({ ...t, schoolForm: dst, updatedAt: new Date().toISOString() });
    });
}

function compareSchulstufen(a, b) {
    const na = /^\d+$/.test(a) ? Number(a) : NaN;
    const nb = /^\d+$/.test(b) ? Number(b) : NaN;
    if (!Number.isNaN(na) && !Number.isNaN(nb) && na !== nb) return na - nb;
    if (!Number.isNaN(na) && Number.isNaN(nb)) return -1;
    if (Number.isNaN(na) && !Number.isNaN(nb)) return 1;
    return String(a).localeCompare(String(b), 'de', { numeric: true });
}

function compareTemplates(a, b) {
    const sf = (a.schoolForm || '').localeCompare(b.schoolForm || '', 'de');
    if (sf !== 0) return sf;
    const sa = a.subjectCode.localeCompare(b.subjectCode, 'de');
    if (sa !== 0) return sa;
    const st = compareSchulstufen(a.schulstufe || '', b.schulstufe || '');
    if (st !== 0) return st;
    const sem = (a.semester || '').localeCompare(b.semester || '', 'de');
    if (sem !== 0) return sem;
    return a.name.localeCompare(b.name, 'de');
}

/**
 * Flache Gruppenansicht: Schulform | Fach | Schulstufe | Semester.
 * @param {ChannelTemplate[]} templates
 * @param {'schoolForm'|'subject'|'schulstufe'|'semester'|'module'} [groupBy]
 * @param {{ schoolForm?: string, subjectCode?: string, schulstufe?: string, semester?: string }} [filters]
 * @returns {TreeNode[]}
 */
export function buildTemplateTree(templates, groupBy, filters) {
    const mode =
        groupBy === 'module' || groupBy === 'schulstufe'
            ? 'schulstufe'
            : groupBy === 'semester'
              ? 'semester'
              : groupBy === 'schoolForm'
                ? 'schoolForm'
                : 'subject';
    const f =
        typeof filters === 'string'
            ? { subjectCode: filters }
            : filters && typeof filters === 'object'
              ? filters
              : {};
    const list = filterTemplates(templates, f);
    /** @type {Map<string, ChannelTemplate[]>} */
    const buckets = new Map();

    for (const t of list) {
        let key;
        if (mode === 'schulstufe') {
            key = t.schulstufe ? 'st:' + t.schulstufe : 'st:_none';
        } else if (mode === 'semester') {
            key = t.semester ? 'sem:' + normalizeSemester(t.semester) : 'sem:_none';
        } else if (mode === 'schoolForm') {
            key = t.schoolForm ? 'sf:' + normalizeSchoolForm(t.schoolForm) : 'sf:_none';
        } else {
            key = t.subjectCode ? 'subj:' + t.subjectCode : 'subj:_none';
        }
        if (!buckets.has(key)) buckets.set(key, []);
        buckets.get(key).push(t);
    }

    const keys = Array.from(buckets.keys()).sort((a, b) => {
        if (a.endsWith(':_none')) return 1;
        if (b.endsWith(':_none')) return -1;
        if (mode === 'schulstufe') {
            return compareSchulstufen(a.replace(/^st:/, ''), b.replace(/^st:/, ''));
        }
        if (mode === 'semester') {
            const order = { SJ: 0, WS: 1, SS: 2 };
            const aa = a.replace(/^sem:/, '');
            const bb = b.replace(/^sem:/, '');
            return (order[aa] ?? 9) - (order[bb] ?? 9) || aa.localeCompare(bb, 'de');
        }
        if (mode === 'schoolForm') {
            return a.replace(/^sf:/, '').localeCompare(b.replace(/^sf:/, ''), 'de');
        }
        return a.replace(/^subj:/, '').localeCompare(b.replace(/^subj:/, ''), 'de');
    });

    /** @type {TreeNode[]} */
    const tree = [];
    for (const key of keys) {
        const tpls = buckets.get(key).slice().sort(compareTemplates);

        let label;
        if (mode === 'schulstufe') {
            const st = key.replace(/^st:/, '');
            label = st === '_none' ? 'Ohne Schulstufe' : schulstufeLabel(st);
        } else if (mode === 'semester') {
            const sem = key.replace(/^sem:/, '');
            label = sem === '_none' ? 'Ohne Semester' : semesterLabel(sem);
        } else if (mode === 'schoolForm') {
            const sf = key.replace(/^sf:/, '');
            label = sf === '_none' ? 'Ohne Schulform' : sf;
        } else {
            const subj = key.replace(/^subj:/, '');
            label = subj === '_none' ? 'Ohne Fach' : subj;
        }

        tree.push({
            type: 'group',
            key,
            label,
            count: tpls.length,
            children: tpls.map((t) => {
                let badge = '';
                if (mode === 'schoolForm') {
                    badge = [
                        t.subjectCode,
                        t.schulstufe ? schulstufeLabel(t.schulstufe) : '',
                        t.semester ? semesterLabel(t.semester) : ''
                    ]
                        .filter(Boolean)
                        .join(' · ');
                } else if (mode === 'subject') {
                    badge = [
                        t.schoolForm,
                        t.schulstufe ? schulstufeLabel(t.schulstufe) : '',
                        t.semester ? semesterLabel(t.semester) : ''
                    ]
                        .filter(Boolean)
                        .join(' · ');
                } else if (mode === 'semester') {
                    badge = [
                        t.schoolForm,
                        t.subjectCode,
                        t.schulstufe ? schulstufeLabel(t.schulstufe) : ''
                    ]
                        .filter(Boolean)
                        .join(' · ');
                } else {
                    badge = [t.schoolForm, t.subjectCode, t.semester ? semesterLabel(t.semester) : '']
                        .filter(Boolean)
                        .join(' · ');
                }
                return {
                    type: 'template',
                    key: t.id,
                    label: t.name,
                    count: t.channels.length,
                    badge,
                    template: t
                };
            })
        });
    }
    return tree;
}

/**
 * @param {ChannelTemplate} tpl
 * @param {string} channelId
 * @param {string} newDisplayName
 * @returns {ChannelTemplate}
 */
export function renameTemplateChannel(tpl, channelId, newDisplayName) {
    const t = normalizeTemplate(tpl);
    const name = normStr(newDisplayName);
    if (!name) throw new Error('Kanalname darf nicht leer sein.');
    if (isGeneralChannelName(name)) throw new Error('„Allgemein“ gehört nicht in die Vorlage.');
    const id = normStr(channelId);
    const channels = t.channels.map((c) => (c.id === id ? { ...c, displayName: name } : c));
    if (!channels.some((c) => c.id === id)) throw new Error('Kanal nicht gefunden.');
    return { ...t, channels, updatedAt: new Date().toISOString() };
}

/**
 * @param {ChannelTemplate} tpl
 * @param {string} displayName
 * @returns {ChannelTemplate}
 */
export function addTemplateChannel(tpl, displayName) {
    const t = normalizeTemplate(tpl);
    const name = normStr(displayName);
    if (!name) throw new Error('Kanalname fehlt.');
    if (isGeneralChannelName(name)) throw new Error('„Allgemein“ nicht hinzufügen – der bleibt im Team.');
    if (t.channels.some((c) => normName(c.displayName) === normName(name))) {
        throw new Error('Kanalname existiert bereits in der Vorlage.');
    }
    return {
        ...t,
        channels: [...t.channels, { id: newId('ch'), displayName: name }],
        updatedAt: new Date().toISOString()
    };
}

/**
 * @param {ChannelTemplate} tpl
 * @param {string} channelId
 * @returns {ChannelTemplate}
 */
export function removeTemplateChannel(tpl, channelId) {
    const t = normalizeTemplate(tpl);
    const id = normStr(channelId);
    return {
        ...t,
        channels: t.channels.filter((c) => c.id !== id),
        updatedAt: new Date().toISOString()
    };
}

/**
 * @param {ChannelTemplate} tpl
 * @param {string} channelId
 * @param {'up'|'down'} dir
 * @returns {ChannelTemplate}
 */
export function moveTemplateChannel(tpl, channelId, dir) {
    const t = normalizeTemplate(tpl);
    const id = normStr(channelId);
    const idx = t.channels.findIndex((c) => c.id === id);
    if (idx < 0) return t;
    const next = t.channels.slice();
    const j = dir === 'up' ? idx - 1 : idx + 1;
    if (j < 0 || j >= next.length) return t;
    const tmp = next[idx];
    next[idx] = next[j];
    next[j] = tmp;
    return { ...t, channels: next, updatedAt: new Date().toISOString() };
}

/**
 * Soll/Ist-Vergleich für Standard-Kanäle (ohne Allgemein).
 * @param {ChannelTemplate} template
 * @param {TeamChannel[]} teamChannels
 * @returns {DiffRow[]}
 */
export function diffChannels(template, teamChannels) {
    const tpl = normalizeTemplate(template);
    const team = (Array.isArray(teamChannels) ? teamChannels : [])
        .map((c) => ({
            id: normStr(c && c.id),
            displayName: normStr(c && c.displayName),
            membershipType: normStr(c && c.membershipType) || 'standard'
        }))
        .filter((c) => c.displayName);

    /** @type {DiffRow[]} */
    const rows = [];

    const usedTeamIds = new Set();

    for (const tc of team) {
        if (isGeneralChannelName(tc.displayName)) {
            rows.push({
                status: 'skip_general',
                templateChannel: null,
                teamChannel: tc,
                message: 'Allgemein bleibt unverändert'
            });
            usedTeamIds.add(tc.id);
        }
    }

    for (const wanted of tpl.channels) {
        const exact = team.find(
            (c) => !usedTeamIds.has(c.id) && normName(c.displayName) === normName(wanted.displayName)
        );
        if (exact) {
            usedTeamIds.add(exact.id);
            rows.push({
                status: 'ok',
                templateChannel: wanted,
                teamChannel: exact,
                message: 'Bereits vorhanden'
            });
            continue;
        }

        const prefix = channelNumberPrefix(wanted.displayName);
        const byNum =
            prefix &&
            team.find(
                (c) =>
                    !usedTeamIds.has(c.id) &&
                    !isGeneralChannelName(c.displayName) &&
                    channelNumberPrefix(c.displayName) === prefix
            );
        if (byNum) {
            usedTeamIds.add(byNum.id);
            rows.push({
                status: 'rename',
                templateChannel: wanted,
                teamChannel: byNum,
                message: 'Nummer gleich, Name weicht ab → umbenennen'
            });
            continue;
        }

        rows.push({
            status: 'create',
            templateChannel: wanted,
            teamChannel: null,
            message: 'Fehlt → anlegen'
        });
    }

    for (const tc of team) {
        if (usedTeamIds.has(tc.id)) continue;
        if (isGeneralChannelName(tc.displayName)) continue;
        rows.push({
            status: 'extra',
            templateChannel: null,
            teamChannel: tc,
            message: 'Nur im Team (wird nicht gelöscht)'
        });
    }

    return rows;
}

/**
 * @param {DiffRow[]} rows
 */
export function summarizeDiff(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const counts = { create: 0, ok: 0, rename: 0, skip_general: 0, extra: 0 };
    for (const r of list) {
        if (r && counts[r.status] != null) counts[r.status]++;
    }
    return counts;
}

/**
 * @param {ChannelTemplate[]} templates
 * @returns {object}
 */
export function buildExportPayload(templates) {
    return {
        kind: EXPORT_KIND,
        version: EXPORT_VERSION,
        exportedAt: new Date().toISOString(),
        templates: normalizeTemplateList(templates)
    };
}

/**
 * @param {unknown} raw
 * @returns {{ templates: ChannelTemplate[], warnings: string[] }}
 */
export function parseImportPayload(raw) {
    /** @type {string[]} */
    const warnings = [];
    let obj = raw;
    if (typeof raw === 'string') {
        try {
            obj = JSON.parse(raw);
        } catch {
            throw new Error('JSON ungültig.');
        }
    }
    if (!obj || typeof obj !== 'object') throw new Error('Import-Objekt fehlt.');

    let list = null;
    if (Array.isArray(obj)) {
        list = obj;
        warnings.push('Array ohne kind – als Vorlagenliste interpretiert.');
    } else if (Array.isArray(obj.templates)) {
        list = obj.templates;
        if (obj.kind && obj.kind !== EXPORT_KIND) {
            warnings.push('Unerwartetes kind: ' + String(obj.kind));
        }
    } else if (obj.template && typeof obj.template === 'object') {
        list = [obj.template];
    } else {
        throw new Error('Keine templates[] im Import gefunden.');
    }

    return { templates: normalizeTemplateList(list), warnings };
}

/**
 * Vorhandene Vorlagen mit Import mergen (gleiche id → ersetzen).
 * @param {ChannelTemplate[]} existing
 * @param {ChannelTemplate[]} incoming
 * @param {{ replaceSameId?: boolean }} [opts]
 */
export function mergeTemplates(existing, incoming, opts) {
    const replace = !opts || opts.replaceSameId !== false;
    const map = new Map();
    for (const t of normalizeTemplateList(existing)) map.set(t.id, t);
    for (const t of normalizeTemplateList(incoming)) {
        if (map.has(t.id) && !replace) {
            map.set(newId('tpl'), { ...t, id: newId('tpl') });
        } else {
            map.set(t.id, t);
        }
    }
    return Array.from(map.values()).sort(compareTemplates);
}

/**
 * @param {{ code?: string, name?: string }[]} subjects
 * @returns {{ code: string, label: string }[]}
 */
export function subjectOptionsFromCore(subjects) {
    const rows = Array.isArray(subjects) ? subjects : [];
    const out = [];
    const seen = new Set();
    for (const s of rows) {
        const code = normStr(s && s.code).toUpperCase();
        if (!code || seen.has(code)) continue;
        seen.add(code);
        const name = normStr(s && s.name);
        out.push({ code, label: name ? code + ' – ' + name : code });
    }
    out.sort((a, b) => a.code.localeCompare(b.code, 'de'));
    return out;
}
