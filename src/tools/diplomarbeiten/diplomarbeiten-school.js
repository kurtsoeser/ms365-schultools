/**
 * Die Wahl „ein Team pro Jahrgang / pro Arbeit“ in der IT-Bibliothek der Schule.
 * Nur der Schlüssel der Diplomarbeiten wird gelesen und geschrieben.
 */
import {
    CONFIG_FOLDER,
    configPartRelativePath
} from '../../shared/stammdaten-sharepoint-config-bundle.js';
import {
    getJsonAtDrivePath,
    isReady,
    putJsonOnDrive,
    requireItLibrary
} from '../../shared/stammdaten-sharepoint-sync-api.js';
import {
    DIPLOM_SETTINGS_KEY,
    embedSettingsInSchoolAppsPart,
    loadDiplomSettings,
    normalizeDiplomSettings
} from './diplomarbeiten-settings.js';

const PART_FILE = 'school-apps-settings.json';
const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function partPath() {
    return configPartRelativePath(CONFIG_FOLDER, PART_FILE);
}

async function graphToken() {
    const G = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!G || typeof G.getGraphToken !== 'function') {
        throw new Error('SharePoint-Zugang ist auf dieser Seite nicht geladen.');
    }
    return G.getGraphToken(SCOPES);
}

/**
 * Schul-Wahl übernehmen, wenn dieser Browser noch keine eigene Wahl hat.
 * @returns {Promise<{ applied: boolean, reason?: string }>}
 */
export async function pullDiplomSettingsFromSchool() {
    if (!isReady()) return { applied: false, reason: 'no-library' };
    const local = loadDiplomSettings();
    if (local.setupDone) return { applied: false, reason: 'local-already' };
    try {
        const it = requireItLibrary();
        const token = await graphToken();
        const raw = await getJsonAtDrivePath(it.driveId, partPath(), token);
        const slice = raw && raw.localStorage && typeof raw.localStorage === 'object' ? raw.localStorage : null;
        const remoteRaw = slice ? slice[DIPLOM_SETTINGS_KEY] : '';
        if (!remoteRaw) return { applied: false, reason: 'no-key' };
        const remote = normalizeDiplomSettings(
            typeof remoteRaw === 'string' ? JSON.parse(remoteRaw) : remoteRaw
        );
        if (!remote.setupDone) return { applied: false, reason: 'remote-open' };
        if (loadDiplomSettings().setupDone) return { applied: false, reason: 'local-already' };
        localStorage.setItem(DIPLOM_SETTINGS_KEY, JSON.stringify(remote));
        return { applied: true };
    } catch (e) {
        return {
            applied: false,
            reason: e && e.message ? String(e.message) : 'pull-failed'
        };
    }
}

/**
 * @returns {Promise<{ ok: boolean, reason?: string }>}
 */
export async function pushDiplomSettingsToSchool() {
    if (!isReady()) return { ok: false, reason: 'no-library' };
    const json = localStorage.getItem(DIPLOM_SETTINGS_KEY) || '';
    if (!json) return { ok: false, reason: 'empty' };
    try {
        const it = requireItLibrary();
        const token = await graphToken();
        const path = partPath();
        const current = await getJsonAtDrivePath(it.driveId, path, token);
        const next = embedSettingsInSchoolAppsPart(current, json);
        await putJsonOnDrive(it.driveId, path, JSON.stringify(next, null, 2), token);
        return { ok: true };
    } catch (e) {
        return {
            ok: false,
            reason: e && e.message ? String(e.message) : 'push-failed'
        };
    }
}
