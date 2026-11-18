/**
 * SharePoint: Aufgaben-Liste anlegen.
 */
import {
    AUFGABEN_COLUMNS,
    LIST_TITLES,
    REQUIRED_COLUMNS,
    toGraphColumnBody
} from '../aufgaben/aufgaben-schema.js';
import { createListEnsureHelpers, spoGraphApi } from './sharepoint-liste-ensure-core.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

const SETTINGS_KEY = 'ms365-aufgaben-sp-settings-v1';
const { ensureList } = createListEnsureHelpers(toGraphColumnBody);

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('spafLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function loadSettings() {
    try {
        const raw = localStorage.getItem(SETTINGS_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch {
        return {};
    }
}

function saveSettings(partial) {
    const next = { ...loadSettings(), ...partial };
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(next));
    return next;
}

async function ensureToken() {
    return await spoGraphApi().getGraphToken(SCOPES);
}

export async function createAufgabenList(webUrl, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await spoGraphApi().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    write('Site: ' + (site.displayName || siteId));

    const settings = loadSettings();
    const title = settings.listAufgaben || LIST_TITLES.aufgaben;
    const list = await ensureList(token, siteId, title, AUFGABEN_COLUMNS, write);
    saveSettings({ siteUrl: url });

    return { siteId, list: { id: list.id, webUrl: list.webUrl || '' } };
}

if (typeof window !== 'undefined') {
    window.ms365SpoAufgaben = {
        createList: createAufgabenList,
        createLists: createAufgabenList
    };
}

document.addEventListener('DOMContentLoaded', () => {
    const urlEl = $('spafSiteUrl');
    const listEl = $('spafListName');
    const settings = loadSettings();
    if (urlEl && !String(urlEl.value || '').trim()) urlEl.value = settings.siteUrl || '';
    if (listEl && !String(listEl.value || '').trim()) {
        listEl.value = settings.listAufgaben || LIST_TITLES.aufgaben;
    }

    $('spafBtnSaveSettings')?.addEventListener('click', () => {
        saveSettings({
            siteUrl: String((urlEl && urlEl.value) || '').trim(),
            listAufgaben: String((listEl && listEl.value) || '').trim()
        });
        toast('Einstellungen gespeichert.');
    });

    const run = $('spafBtnRun');
    if (run) {
        run.addEventListener('click', () => {
            saveSettings({
                siteUrl: String((urlEl && urlEl.value) || '').trim(),
                listAufgaben: String((listEl && listEl.value) || '').trim()
            });
            if ($('spafLog')) $('spafLog').textContent = '';
            run.disabled = true;
            createAufgabenList(String((urlEl && urlEl.value) || '').trim(), log)
                .then(() => toast('Aufgaben-Liste angelegt bzw. geprüft.'))
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    run.disabled = false;
                });
        });
    }

    $('spafBtnProbe')?.addEventListener('click', () => {
        ensureToken()
            .then((token) =>
                spoGraphApi().resolveSiteFromWebUrl(token, String((urlEl && urlEl.value) || '').trim())
            )
            .then((site) => toast('Site OK: ' + (site.displayName || site.id)))
            .catch((e) => toast('Fehler: ' + (e && e.message ? e.message : e)));
    });

    void REQUIRED_COLUMNS;
});
