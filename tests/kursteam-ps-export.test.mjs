import { describe, it, expect } from 'vitest';
import {
    buildStandaloneKursteamPs1,
    buildStandaloneKursteamPs1V2,
    buildKursteamCsvPreviewPs1,
    psEscapeForExport
} from '../src/tools/kursteams/kursteam-ps-export.js';

describe('kursteam-ps-export', () => {
    const teams = [
        { teamName: "SJ26 | 1AK | D", gruppenmail: 'sj26-1ak-d', besitzer: 'lehrer@schule.at' }
    ];

    it('psEscapeForExport escaped single quotes', () => {
        expect(psEscapeForExport("O'Brien")).toBe("O''Brien");
    });

    it('buildStandaloneKursteamPs1 enthält New-Team und Idempotenz', () => {
        const ps1 = buildStandaloneKursteamPs1(teams);
        expect(ps1).toContain('New-Team -Template "EDU_Class"');
        expect(ps1).toContain('Get-Team -MailNickName');
        expect(ps1).toContain("sj26-1ak-d");
    });

    it('buildStandaloneKursteamPs1V2 enthält Checkpoint und Retry', () => {
        const ps1 = buildStandaloneKursteamPs1V2(teams);
        expect(ps1).toContain('Kursteam-Anlage-checkpoint.json');
        expect(ps1).toContain('Invoke-KtTeamCreateWithRetry');
        expect(ps1).toContain('AadGroupCreationLimitExceeded');
        expect(ps1).toContain('ETA ca.');
    });

    it('buildKursteamCsvPreviewPs1 referenziert neueteams.csv', () => {
        const preview = buildKursteamCsvPreviewPs1();
        expect(preview).toContain('Import-Csv -Path .\\neueteams.csv');
    });
});
