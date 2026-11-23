import { beforeEach, describe, expect, it, vi } from 'vitest';

function installDom(fields) {
    const values = Object.assign(
        {
            schoolName: '',
            schoolEmailDomain: '',
            tenantSchoolEmailPattern: 'vorname.nachname',
            tenantIntranetSiteUrl: '',
            tenantPlanerDefaultSiteUrl: '',
            tenantItLibrarySiteUrl: '',
            tenantSchoolIntranetSiteUrl: ''
        },
        fields || {}
    );
    globalThis.document = {
        getElementById(id) {
            if (!(id in values)) return null;
            return {
                value: values[id],
                dataset: {},
                addEventListener() {}
            };
        }
    };
    globalThis.window = globalThis;
}

describe('computeStammdatenSetupPaneStatus', () => {
    beforeEach(() => {
        vi.resetModules();
        delete globalThis.document;
        delete globalThis.window;
    });

    it('marks profile empty without name/domain', async () => {
        installDom({});
        const { computeStammdatenSetupPaneStatus } = await import('../src/shared/tenant-stammdaten-setup-nav.js');
        expect(computeStammdatenSetupPaneStatus().profile).toBe('empty');
    });

    it('marks profile warn with only school name', async () => {
        installDom({ schoolName: 'Testschule' });
        const { computeStammdatenSetupPaneStatus } = await import('../src/shared/tenant-stammdaten-setup-nav.js');
        expect(computeStammdatenSetupPaneStatus().profile).toBe('warn');
    });

    it('marks spo panes from URLs and intranet fallback', async () => {
        installDom({
            schoolName: 'Testschule',
            schoolEmailDomain: 'schule.at',
            tenantIntranetSiteUrl: 'https://contoso.sharepoint.com/sites/Listen'
        });
        const { computeStammdatenSetupPaneStatus } = await import('../src/shared/tenant-stammdaten-setup-nav.js');
        const status = computeStammdatenSetupPaneStatus();
        expect(status.profile).toBe('ok');
        expect(status['spo-lists']).toBe('ok');
        expect(status['spo-planer']).toBe('warn');
        expect(status['spo-hub']).toBe('warn');
        expect(status['spo-it']).toBe('empty');
    });
});
