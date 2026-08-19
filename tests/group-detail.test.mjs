import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function read(rel) {
    return readFileSync(join(projectRoot, rel), 'utf8');
}

describe('Zentrale Gruppen-Detailansicht', () => {
    it('liefert Markup, Match-API und CSS der Klassengruppen-Ansicht', () => {
        const js = read('src/shared/group-detail/group-detail.js');
        expect(js).toContain('window.ms365GroupDetail');
        expect(js).toContain('function mount(');
        expect(js).toContain('function buildMarkup(');
        expect(js).toContain('function setTab(');
        expect(js).toContain('persistMatch');
        expect(js).toContain('aliasEditable');
        expect(js).toContain('smtpSlot');
        expect(js).toContain('matchUi');
        expect(js).toContain('teamArchive');
        expect(js).toContain('deleteGroup');
        expect(js).toContain('visibilityUnsupported');
        const ids = [
            'slgDetailTitle',
            'slgBtnOpenEntra',
            'slgBtnUnmatch',
            'slgLiveName',
            'slgLiveAlias',
            'slgLiveCreated',
            'slgLiveTeam',
            'slgLiveTeamLink',
            'slgBtnProvisionTeam',
            'slgBtnRenewExpires',
            'slgBtnUpdateGroup',
            'slgBtnRefreshGroup',
            'jgSmtpDrop',
            'jgExpiresDrop',
            'jgBtnSmtpThis',
            'slgTabBtnGeneral',
            'slgOwnersList',
            'slgMembersList',
            'slgBtnSync',
            'slgTeamArchiveWrap',
            'slgArchiveState',
            'slgArchiveSpoReadonly',
            'slgBtnDeleteGroup',
            'slgOwnerSingleWrap'
        ];
        for (const id of ids) {
            expect(js, `fehlende ID ${id}`).toContain(`id="${id}"`);
        }
        expect(js).toContain('gdEmptyHint');
        expect(js).toContain('gdDetailWrap');
        expect(js).toContain('gd-archive-card');
        expect(js).toContain('gd-detail-actions');
        expect(js).not.toContain('shouldSetSpoSiteReadOnlyForMembers');
        const css = read('src/shared/group-detail/group-detail.css');
        expect(css).toContain('.field-editable');
        expect(css).toContain('.jg-team-status');
        expect(css).toContain('.slg-owner-member-box');
        expect(css).toContain('.gd-layout');
        expect(css).toContain('.gd-filter-row');
        expect(css).toContain('.gd-page-info');
        expect(css).toContain('.gd-archive-card');
        expect(css).toContain('.gd-detail-actions');
        expect(css).toContain('.gd-btn-danger');
    });

    it('hängt die Ansicht nur einmal ein', () => {
        const js = read('src/shared/group-detail/group-detail.js');
        expect(js).toContain('data-gd-mounted');
        expect(js).toContain("host.querySelector('#slgLiveName')");
        expect(js).toContain('data-gd-wired');
    });
});

describe('Match-Module nutzen die zentrale Ansicht', () => {
    const pages = [
        ['tools/schueler-lehrer-gruppen.html', 'src/tools/schueler-lehrer-gruppen/slg-gruppenverwaltung.js'],
        ['tools/verwaltung.html', 'src/tools/verwaltung/verwaltung-gruppenverwaltung.js'],
        ['tools/arge-fachgruppen.html', 'src/tools/arge-fachgruppen/arge-fachgruppen.js'],
        ['tools/jahrgangsgruppen.html', 'src/tools/jahrgangsgruppen/jahrgangsgruppen.js']
    ];

    it.each(pages)('%s hängt Host und group-detail ein', (htmlPath, jsPath) => {
        const html = read(htmlPath);
        expect(html).toContain('id="groupDetailHost"');
        expect(html).toContain('src/shared/group-detail/group-detail.js');
        expect(html).toContain('src/shared/group-detail/group-detail.css');
        expect(html).not.toContain('id="slgLiveName"');
        const js = read(jsPath);
        expect(js).toContain("mount('#groupDetailHost'");
        expect(js).not.toContain('async function runSearchGroups');
        expect(js).not.toContain('function runUnmatch');
        expect(js).not.toContain('function openEntraForMatched');
    });

    it.each(pages)('%s nutzt das gemeinsame Listen-Chrome ohne Login-Button', (htmlPath, jsPath) => {
        const html = read(htmlPath);
        expect(html).toContain('gd-layout');
        expect(html).toContain('id="slgBtnReloadLists"');
        expect(html).toContain('Neu einlesen');
        expect(html).toContain('gd-page-info');
        expect(html).not.toContain('id="slgBtnLogin"');
        expect(html).not.toContain('Bei Microsoft anmelden');
        const js = read(jsPath);
        expect(js).not.toContain("onClick('slgBtnLogin'");
        expect(js).not.toContain('async function onLogin');
    });
});

describe('Gruppenerstellung nutzt die zentrale Ansicht', () => {
    it('hängt Host und group-detail ein (ohne Match-UI-Flow)', () => {
        const html = read('tools/gruppenerstellung.html');
        expect(html).toContain('id="groupDetailHost"');
        expect(html).toContain('src/shared/group-detail/group-detail.js');
        expect(html).toContain('src/shared/group-detail/group-detail.css');
        expect(html).toContain('src/tools/schueler-lehrer-gruppen/slg-live-details.js');
        expect(html).toContain('src/shared/graph-unified-groups.js');
        expect(html).not.toContain('id="gpDetailsDl"');
        expect(html).not.toContain('id="gpLinkOverview"');
        expect(html).not.toContain('id="slgLiveName"');
        const js = read('src/tools/gruppenerstellung/gruppenerstellung-policy.js');
        expect(js).toContain("mount('#groupDetailHost'");
        expect(js).toContain('matchUi: false');
        expect(js).toContain('visibilityUnsupported: true');
        expect(js).toContain('teams: false');
        expect(js).toContain('getResolvedGroupId');
        expect(js).toContain('refreshGroupDetailPanel');
    });
});

describe('Tenant-Gruppenverwaltung nutzt die zentrale Ansicht', () => {
    it('hängt Host und group-detail ein, ohne kopiertes Tenant-Formular', () => {
        const html = read('tools/schulstruktur-sync.html');
        expect(html).toContain('id="ssTenantDetail"');
        expect(html).toContain('id="groupDetailHost"');
        expect(html).toContain('src/shared/group-detail/group-detail.js');
        expect(html).toContain('type="module" src="../src/shared/group-detail/group-detail.js"');
        expect(html).toContain('type="module" src="../src/tools/schueler-lehrer-gruppen/slg-live-details.js"');
        expect(html).toContain('src/shared/group-detail/group-detail.css');
        expect(html).toContain('src/tools/schueler-lehrer-gruppen/slg-live-details.js');
        expect(html).toContain('src/shared/graph-unified-groups.js');
        expect(html).toContain('src/shared/group-photo-thumb.js');
        expect(html).toContain('id="ssTenantBulkWrap"');
        expect(html).not.toContain('id="ssTenantName"');
        expect(html).not.toContain('id="ssTenantUpdateBtn"');
        expect(html).not.toContain('id="ssOwnersList"');
        expect(html).not.toContain('id="slgLiveName"');
        expect(html).toContain('id="ssTabStrukturTop"');
        expect(html).toContain('id="ssTabAbgleichenTop"');
        expect(html).toContain('id="ssTabTenantTop"');
        expect(html).toContain('id="ssStrukturBanner"');
        const js = read('src/tools/schulstruktur-sync/schulstruktur-sync.js');
        expect(js).toContain("import '../../shared/group-detail/group-detail.js'");
        expect(js).toContain("import '../schueler-lehrer-gruppen/slg-live-details.js'");
        expect(js).toContain("mount('#groupDetailHost'");
        expect(js).toContain('ensureTenantGroupDetailMounted');
        expect(js).toContain('matchUi: false');
        expect(js).toContain('teamArchive: true');
        expect(js).toContain('deleteGroup: true');
        expect(js).toContain('ownerExtra: \'ssTenantBulkWrap\'');
        expect(js).not.toContain('ssTenantUpdateBtn');
        expect(js).not.toContain("window.location.href = '../tenant.html'");
    });
});
