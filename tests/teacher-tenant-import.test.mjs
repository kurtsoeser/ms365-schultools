import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { rowMatchesTextFilter } from '../src/shared/teacher-tenant-import-ui.js';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function read(rel) {
    return readFileSync(join(projectRoot, rel), 'utf8');
}

describe('Lehrkräfte aus dem Tenant einlesen', () => {
    it('Einrichtung Schritt 4 hat den Microsoft-365-Import', () => {
        const html = read('einrichtung.html');
        expect(html).toContain('id="swBtnImportTeachersFromTenant"');
        expect(html).toContain('id="swTeacherTenantImportPanel"');
        expect(html).toContain('id="swBtnTeacherTenantApply"');
        expect(html).toContain('src/shared/teacher-tenant-import-ui.js');
        expect(html).toContain('Aus Microsoft 365 einlesen');
    });

    it('Schul-Einstellungen Tab Lehrer hat denselben Import', () => {
        const html = read('tenant.html');
        expect(html).toContain('id="tenantBtnImportTeachersFromTenant"');
        expect(html).toContain('id="tenantTeacherTenantImportPanel"');
        expect(html).toContain('src/shared/teacher-tenant-import-ui.js');
    });

    it('Personen-Modul filtert nach Lizenz inkl. Schüler A1/A3', () => {
        const html = read('tools/personen-verwaltung.html');
        expect(html).toContain('id="pvFilterLicense"');
        expect(html).toContain('src/shared/graph-licenses.js');
        const js = read('src/tools/personen-verwaltung/personen-verwaltung.js');
        expect(js).toContain('assignedLicenses');
        expect(js).toContain('pvFilterLicense');
        expect(js).toContain('userMatchesLicenseFilter');
        const lic = read('src/shared/graph-licenses.js');
        expect(lic).toContain('student-a1');
        expect(lic).toContain('student-a3');
        expect(lic).toContain('A1 für Schüler:innen');
    });

    it('Personen-Modul erklärt den Ablauf für neue Konten kontrastreich', () => {
        const html = read('tools/personen-verwaltung.html');
        expect(html).toContain('class="pv-onboard"');
        expect(html).toContain('Neue Person anlegen');
        expect(html).toContain('Nutzungsort');
        expect(html).toContain('Microsoft weist keine Lizenz zu');
        expect(html).not.toContain('id="pvOnboardHint"');
        expect(html).not.toContain('1. Anlegen · 2. Nutzungsort · 3. Lizenz');
    });

    it('Personen-Modul kann Lizenzen zuweisen und entziehen', () => {
        const html = read('tools/personen-verwaltung.html');
        expect(html).toContain('id="pvTabLizenzen"');
        expect(html).toContain('id="pvPanelLizenzen"');
        expect(html).toContain('id="pvLicAssignBtn"');
        expect(html).toContain('id="pvLicUsageLocation"');
        const js = read('src/tools/personen-verwaltung/personen-verwaltung.js');
        expect(js).toContain('/assignLicense');
        expect(js).toContain('Organization.Read.All');
        expect(js).toContain('removeLicenses');
        expect(js).toContain('usageLocation');
        expect(js).toContain('mailNickname');
        expect(js).toContain('streetAddress');
        expect(js).toContain('field-editable');
        expect(html).toContain('id="pvBtnSave"');
        expect(html).toContain('id="pvProfileHint"');
        expect(html).not.toContain('id="pvBtnEdit"');
    });

    it('Personen-Modul kann Gruppenmitgliedschaften setzen', () => {
        const html = read('tools/personen-verwaltung.html');
        expect(html).toContain('id="pvGroupSearch"');
        expect(html).toContain('id="pvGroupAddBtn"');
        expect(html).toContain('Zur Gruppe hinzufügen');
        const js = read('src/tools/personen-verwaltung/personen-verwaltung.js');
        expect(js).toContain('Group.ReadWrite.All');
        expect(js).toContain('/members/$ref');
        expect(js).toContain('data-pv-group-remove');
    });

    it('Einrichtung Schritt 5 hat den Microsoft-365-Import für Schüler:innen', () => {
        const html = read('einrichtung.html');
        expect(html).toContain('id="swBtnImportStudentsFromTenant"');
        expect(html).toContain('id="swStudentTenantImportPanel"');
        expect(html).toContain('id="swBtnStudentTenantApply"');
    });

    it('Schul-Einstellungen Tab Schüler hat denselben Import', () => {
        const html = read('tenant.html');
        expect(html).toContain('id="tenantBtnImportStudentsFromTenant"');
        expect(html).toContain('id="tenantStudentTenantImportPanel"');
    });

    it('Tenant-Import hat Textfilter, Alle abwählen und Kopf-Checkbox', () => {
        const ein = read('einrichtung.html');
        const ten = read('tenant.html');
        const js = read('src/shared/teacher-tenant-import-ui.js');
        expect(ein).toContain('id="swTeacherTenantImportTextFilter"');
        expect(ein).toContain('id="swStudentTenantImportTextFilter"');
        expect(ein).toContain('id="swBtnTeacherTenantSelectNone"');
        expect(ein).toContain('id="swBtnStudentTenantSelectNone"');
        expect(ein).toContain('id="swTeacherTenantSelectAllRows"');
        expect(ein).toContain('id="swStudentTenantSelectAllRows"');
        expect(ten).toContain('id="tenantTeacherTenantImportTextFilter"');
        expect(ten).toContain('id="tenantStudentTenantImportTextFilter"');
        expect(ten).toContain('id="tenantBtnTeacherTenantSelectNone"');
        expect(ten).toContain('id="tenantBtnStudentTenantSelectNone"');
        expect(js).toContain('rowMatchesTextFilter');
        expect(js).toContain('selectNoneBtnId');
        expect(js).toContain('textFilterId');
    });

    it('Textfilter trifft Name, E-Mail und Kürzel', () => {
        const row = {
            name: 'Angus Young',
            email: 'angus.young@kurtrocks.com',
            code: 'YOU',
            licenseLabel: 'A1 Lehrpersonal'
        };
        expect(rowMatchesTextFilter(row, '', 'code')).toBe(true);
        expect(rowMatchesTextFilter(row, 'young', 'code')).toBe(true);
        expect(rowMatchesTextFilter(row, 'YOU', 'code')).toBe(true);
        expect(rowMatchesTextFilter(row, 'angus kurtrocks', 'code')).toBe(true);
        expect(rowMatchesTextFilter(row, 'scott', 'code')).toBe(false);
        expect(rowMatchesTextFilter({ name: 'Lisa', klasse: '1A', email: 'lisa@schule.at' }, '1a', 'klasse')).toBe(true);
    });
});
