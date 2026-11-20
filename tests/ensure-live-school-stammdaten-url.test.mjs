import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { resolveToolSiteUrlForLiveStammdaten } from '../src/shared/ensure-live-school-stammdaten.js';
import { SCHOOL_TOOL_SITE_KEYS } from '../src/shared/school-tool-sites.js';

function memoryStorage() {
    /** @type {Record<string, string>} */
    const map = {};
    return {
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(map, k) ? map[k] : null;
        },
        setItem(k, v) {
            map[k] = String(v);
        },
        removeItem(k) {
            delete map[k];
        }
    };
}

describe('resolveToolSiteUrlForLiveStammdaten', () => {
    const aktKey = SCHOOL_TOOL_SITE_KEYS.aktivitaeten;
    const prevWindow = globalThis.window;
    const prevLs = globalThis.localStorage;

    beforeEach(() => {
        globalThis.localStorage = memoryStorage();
        globalThis.window = {
            ms365AppDataV2: {
                getSetup: () => ({})
            }
        };
    });

    afterEach(() => {
        if (prevWindow === undefined) delete globalThis.window;
        else globalThis.window = prevWindow;
        if (prevLs === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prevLs;
    });

    it('bevorzugt planerDefaultSiteUrl vor altem Aktivitäten-LocalStorage', () => {
        localStorage.setItem(aktKey, 'https://kurtrocks.sharepoint.com/sites/Intranet');
        window.ms365AppDataV2.getSetup = () => ({
            planerDefaultSiteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
            intranetSiteUrl: 'https://kurtrocks.sharepoint.com/sites/Intranet'
        });
        expect(resolveToolSiteUrlForLiveStammdaten()).toBe(
            'https://kurtrocks.sharepoint.com/sites/MS365-Schultools'
        );
    });

    it('nimmt LocalStorage nur wenn IT/Lizenz leer sind', () => {
        localStorage.setItem(aktKey, 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools');
        window.ms365AppDataV2.getSetup = () => ({
            planerDefaultSiteUrl: '',
            intranetSiteUrl: 'https://kurtrocks.sharepoint.com/sites/Intranet'
        });
        expect(resolveToolSiteUrlForLiveStammdaten()).toBe(
            'https://kurtrocks.sharepoint.com/sites/MS365-Schultools'
        );
    });
});
