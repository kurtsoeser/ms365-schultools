/**
 * Persistenz für Kursteam-Vorlagen (localStorage).
 */
import { loadJson, saveJson } from '../../shared/utils/storage.js';
import {
    STORAGE_KIND,
    normalizeTemplateList,
    normalizeTemplate,
    normalizeSchoolForm,
    rememberSchoolForm,
    stripTemplateOrigin
} from './kursteam-templates-logic.js';
import { getSeedTemplates } from './kursteam-templates-seed.js';

export const STORAGE_KEY = 'ms365-kursteam-templates-v1';

/**
 * @returns {{
 *   kind: string,
 *   templates: import('./kursteam-templates-logic.js').ChannelTemplate[],
 *   schoolForms: string[]
 * }}
 */
export function loadState() {
    const raw = loadJson(STORAGE_KEY, null);
    // Erster Besuch: Seeds. Bereits gespeicherter leerer Stand bleibt leer.
    if (!raw || typeof raw !== 'object' || !Array.isArray(raw.templates)) {
        const seeded = getSeedTemplates();
        return saveState(seeded, ['HAKB']);
    }
    const templates = normalizeTemplateList(raw.templates);
    const schoolForms = normalizeSchoolFormCatalog(raw.schoolForms, templates);
    return { kind: STORAGE_KIND, templates, schoolForms };
}

/**
 * @param {unknown} catalog
 * @param {import('./kursteam-templates-logic.js').ChannelTemplate[]} templates
 * @returns {string[]}
 */
function normalizeSchoolFormCatalog(catalog, templates) {
    let list = Array.isArray(catalog)
        ? catalog.map(normalizeSchoolForm).filter(Boolean)
        : [];
    for (const t of templates) {
        list = rememberSchoolForm(list, t.schoolForm);
    }
    return list;
}

/**
 * @param {import('./kursteam-templates-logic.js').ChannelTemplate[]} templates
 * @param {string[]} [schoolForms]
 */
export function saveState(templates, schoolForms) {
    const tpls = normalizeTemplateList(templates).map(stripTemplateOrigin);
    const forms = normalizeSchoolFormCatalog(schoolForms, tpls);
    const payload = {
        kind: STORAGE_KIND,
        templates: tpls,
        schoolForms: forms,
        savedAt: new Date().toISOString()
    };
    saveJson(STORAGE_KEY, payload);
    return { kind: STORAGE_KIND, templates: tpls, schoolForms: forms };
}

/**
 * @param {import('./kursteam-templates-logic.js').ChannelTemplate[]} templates
 */
export function saveTemplates(templates) {
    const state = loadState();
    return saveState(templates, state.schoolForms);
}

/**
 * @param {string[]} schoolForms
 */
export function saveSchoolForms(schoolForms) {
    const state = loadState();
    return saveState(state.templates, schoolForms);
}

/**
 * @param {string} id
 * @returns {import('./kursteam-templates-logic.js').ChannelTemplate|null}
 */
export function getTemplateById(id) {
    const want = String(id || '').trim();
    if (!want) return null;
    return loadState().templates.find((t) => t.id === want) || null;
}

/**
 * @param {import('./kursteam-templates-logic.js').ChannelTemplate} tpl
 */
export function upsertTemplate(tpl) {
    const next = normalizeTemplate(tpl);
    const state = loadState();
    const idx = state.templates.findIndex((t) => t.id === next.id);
    const list = state.templates.slice();
    if (idx >= 0) list[idx] = next;
    else list.push(next);
    const schoolForms = rememberSchoolForm(state.schoolForms, next.schoolForm);
    saveState(list, schoolForms);
    return next;
}

/**
 * @param {string} id
 */
export function deleteTemplate(id) {
    const want = String(id || '').trim();
    const state = loadState();
    const list = state.templates.filter((t) => t.id !== want);
    saveState(list, state.schoolForms);
    return list;
}

/**
 * Alle Vorlagen löschen (Bibliothek leer lassen).
 * @returns {{ kind: string, templates: import('./kursteam-templates-logic.js').ChannelTemplate[], schoolForms: string[] }}
 */
export function clearAllTemplates() {
    const state = loadState();
    return saveState([], state.schoolForms);
}

/**
 * Alle Vorlagen löschen und mitgelieferte Standard-Seeds neu laden.
 * @returns {{ kind: string, templates: import('./kursteam-templates-logic.js').ChannelTemplate[], schoolForms: string[] }}
 */
export function resetToSeedTemplates() {
    const seeded = getSeedTemplates();
    return saveState(seeded, ['HAKB']);
}
