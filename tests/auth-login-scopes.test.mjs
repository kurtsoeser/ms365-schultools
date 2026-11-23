import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
    GRAPH_SCOPES_GROUP_WRITE,
    GRAPH_SCOPES_USER_WRITE,
    LOGIN_SCOPE_RULES,
    PERSONA_SCOPE_SUMMARY,
    USER_READ,
    accessTokenIncludesScopes,
    clientKindForScopes,
    labelScopeDe,
    matchLoginScopeRule,
    prepareItAppScopes,
    resolveLoginScopesForPath,
    scopesForGraphCall,
    scopesForGraphKind
} from '../src/shared/auth-login-scopes.js';

function jwtWithScp(scp) {
    const payload = Buffer.from(JSON.stringify({ scp }), 'utf8').toString('base64url');
    return 'e30.' + payload + '.sig';
}

describe('auth-login-scopes (a4)', () => {
    it('Dashboard und Tenant starten nur mit User.Read', () => {
        expect(resolveLoginScopesForPath('/index.html')).toEqual([USER_READ]);
        expect(resolveLoginScopesForPath('/tenant.html')).toEqual([USER_READ]);
    });

    it('Portal-Planer fordern User.Read + Sites.ReadWrite.All', () => {
        const scopes = resolveLoginScopesForPath('/tools/lehrer-freistellung-planer.html');
        expect(scopes).toContain(USER_READ);
        expect(scopes).toContain('https://graph.microsoft.com/Sites.ReadWrite.All');
        expect(scopes).not.toContain('https://graph.microsoft.com/User.ReadWrite.All');
        expect(scopes).not.toContain('https://graph.microsoft.com/Group.ReadWrite.All');
    });

    it('Termine-Hub bekommt Kalender-Scopes ohne IT-Admin-Bundle', () => {
        const scopes = resolveLoginScopesForPath('/tools/termine-hub.html');
        expect(scopes).toContain('https://graph.microsoft.com/Calendars.ReadWrite');
        expect(scopes).not.toContain('https://graph.microsoft.com/Team.Create');
    });

    it('Unbekannte Seiten fallen auf User.Read zurück', () => {
        expect(resolveLoginScopesForPath('/tools/unknown-tool.html')).toEqual([USER_READ]);
    });

    it('liefert deutsche Scope-Labels und Persona-Zusammenfassung', () => {
        expect(labelScopeDe(USER_READ)).toMatch(/Profil/i);
        expect(LOGIN_SCOPE_RULES.length).toBeGreaterThan(5);
        expect(PERSONA_SCOPE_SUMMARY.some((p) => /Lehrkraft/i.test(p.persona))).toBe(true);
        expect(matchLoginScopeRule('/tools/srdp-anmeldung.html')?.id).toBe('srdp');
    });

    it('trennt Gruppen-Schreiben von Benutzer-Schreiben (R3)', () => {
        expect(GRAPH_SCOPES_GROUP_WRITE).toContain('https://graph.microsoft.com/Group.ReadWrite.All');
        expect(GRAPH_SCOPES_GROUP_WRITE).not.toContain('https://graph.microsoft.com/User.ReadWrite.All');
        expect(GRAPH_SCOPES_USER_WRITE).toContain('https://graph.microsoft.com/User.ReadWrite.All');
        expect(GRAPH_SCOPES_USER_WRITE).not.toContain('https://graph.microsoft.com/Group.ReadWrite.All');
        expect(scopesForGraphKind('portal')).toEqual([USER_READ]);
        for (let i = 0; i < LOGIN_SCOPE_RULES.length; i++) {
            expect(LOGIN_SCOPE_RULES[i].scopes).not.toContain('https://graph.microsoft.com/User.ReadWrite.All');
            expect(LOGIN_SCOPE_RULES[i].scopes).not.toContain('https://graph.microsoft.com/Group.ReadWrite.All');
        }
    });

    it('ordnet Graph-Aufrufe Lesen oder Schreiben zu, Portal-Pfade bleiben unangetastet', () => {
        const groupRead = scopesForGraphCall('GET', '/groups?$top=10');
        expect(groupRead).toContain('https://graph.microsoft.com/Group.Read.All');
        expect(groupRead).not.toContain('https://graph.microsoft.com/Group.ReadWrite.All');
        expect(groupRead).not.toContain('https://graph.microsoft.com/User.ReadWrite.All');

        const groupWrite = scopesForGraphCall('POST', '/groups');
        expect(groupWrite).toContain('https://graph.microsoft.com/Group.ReadWrite.All');
        expect(groupWrite).not.toContain('https://graph.microsoft.com/User.ReadWrite.All');

        const userWrite = scopesForGraphCall('POST', '/users');
        expect(userWrite).toContain('https://graph.microsoft.com/User.ReadWrite.All');
        expect(userWrite).not.toContain('https://graph.microsoft.com/Group.ReadWrite.All');

        const teamPut = scopesForGraphCall('PUT', '/groups/abc/team');
        expect(teamPut).toContain('https://graph.microsoft.com/Team.Create');
        expect(teamPut).toContain('https://graph.microsoft.com/Group.ReadWrite.All');
        expect(teamPut).not.toContain('https://graph.microsoft.com/User.ReadWrite.All');

        expect(scopesForGraphCall('GET', '/me')).toBeNull();
        expect(scopesForGraphCall('GET', '/sites/root/lists')).toBeNull();
    });

    it('erkennt Write.All im Token und leitet Read nicht daraus ab, wenn es fehlt', () => {
        const readTok = jwtWithScp('User.Read Group.Read.All');
        expect(accessTokenIncludesScopes(readTok, scopesForGraphCall('GET', '/groups'))).toBe(true);
        expect(accessTokenIncludesScopes(readTok, scopesForGraphCall('POST', '/groups'))).toBe(false);
        const writeTok = jwtWithScp('User.Read Group.ReadWrite.All');
        expect(accessTokenIncludesScopes(writeTok, scopesForGraphCall('GET', '/groups'))).toBe(true);
        expect(accessTokenIncludesScopes(writeTok, ['https://graph.microsoft.com/User.ReadWrite.All'])).toBe(
            false
        );
    });

    it('graph-unified-groups defaultet nicht mehr auf das Admin-Schreib-Bundle', () => {
        const src = readFileSync(new URL('../src/shared/graph-unified-groups.js', import.meta.url), 'utf8');
        expect(src).toContain("const GRAPH_SCOPES = ['https://graph.microsoft.com/User.Read']");
        expect(src).not.toContain('User.ReadWrite.All');
        expect(src).not.toContain('Group.ReadWrite.All');
        expect(src).toContain('scopesForGraphCall');
        expect(src).toContain('scopesRequireItApp');
        expect(src).toContain('acquireItAppToken');
        expect(src).toContain("getGraphTokenFor('team-write')");
    });

    it('ordnet IT-Scopes der zweiten App zu und Lehrkraft-Scopes der Portal-App', () => {
        expect(clientKindForScopes(['https://graph.microsoft.com/User.Read'])).toBe('portal');
        expect(clientKindForScopes(['https://graph.microsoft.com/User.ReadWrite.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/Group.ReadWrite.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/Directory.ReadWrite.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/Directory.Read.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/Group.Read.All'])).toBe('portal');
        expect(clientKindForScopes(['https://graph.microsoft.com/User.Read.All'])).toBe('portal');
        expect(clientKindForScopes(['https://graph.microsoft.com/Calendars.ReadWrite'])).toBe('portal');
        expect(clientKindForScopes(['https://graph.microsoft.com/Sites.ReadWrite.All'])).toBe('portal');
        expect(clientKindForScopes(['https://graph.microsoft.com/Chat.Create'])).toBe('portal');
        expect(clientKindForScopes(['https://graph.microsoft.com/Files.ReadWrite.All'])).toBe('portal');
        expect(clientKindForScopes(['https://graph.microsoft.com/Notes.ReadWrite.All'])).toBe('portal');
        expect(clientKindForScopes(['api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access'])).toBe(
            'portal'
        );
        expect(clientKindForScopes(['api://12e0cfe2-8337-4b35-93e8-542faf658eb3/Election.Access'])).toBe(
            'portal'
        );
        expect(clientKindForScopes(['https://tenant.sharepoint.com/AllSites.Read'])).toBe('portal');
        expect(clientKindForScopes(['https://tenant.sharepoint.com/AllSites.Write'])).toBe('portal');
        expect(clientKindForScopes(['https://tenant.sharepoint.com/AllSites.FullControl'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/Team.Create'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/TeamSettings.ReadWrite.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/Channel.Create'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/ChannelMember.ReadWrite.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/ChannelSettings.ReadWrite.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/Policy.ReadWrite.Authorization'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/RoleManagement.ReadWrite.Directory'])).toBe(
            'it'
        );
        expect(clientKindForScopes(['https://graph.microsoft.com/RoleManagement.Read.Directory'])).toBe(
            'portal'
        );
        expect(
            clientKindForScopes([
                'https://graph.microsoft.com/User.Read',
                'https://graph.microsoft.com/RoleManagement.Read.Directory'
            ])
        ).toBe('portal');
        expect(clientKindForScopes(['https://graph.microsoft.com/SharePointTenantSettings.ReadWrite.All'])).toBe(
            'it'
        );
        expect(clientKindForScopes(['https://graph.microsoft.com/Sites.Create.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/Sites.Manage.All'])).toBe('it');
        expect(clientKindForScopes(['https://graph.microsoft.com/EduRoster.ReadWrite'])).toBe('it');
        expect(
            clientKindForScopes(['api://c7e6f467-e6f3-4221-a9ee-574b35120029/Kursteams.Create'])
        ).toBe('it');
        expect(clientKindForScopes(scopesForGraphCall('POST', '/groups'))).toBe('it');
        expect(clientKindForScopes(scopesForGraphCall('GET', '/groups'))).toBe('portal');
        expect(clientKindForScopes(scopesForGraphCall('GET', '/directoryRoles'))).toBe('portal');
        expect(clientKindForScopes(resolveLoginScopesForPath('/index.html'))).toBe('portal');
        expect(clientKindForScopes(resolveLoginScopesForPath('/tools/termine-hub.html'))).toBe('portal');
    });

    it('stellt Sites.FullControl.All auf AllSites.FullControl um und lässt Write stehen', () => {
        expect(prepareItAppScopes(['https://schule.sharepoint.com/Sites.FullControl.All'])).toEqual([
            'https://schule.sharepoint.com/AllSites.FullControl'
        ]);
        expect(prepareItAppScopes(['https://schule.sharepoint.com/AllSites.Write'])).toEqual([
            'https://schule.sharepoint.com/AllSites.Write'
        ]);
        expect(clientKindForScopes(['https://schule.sharepoint.com/Sites.FullControl.All'])).toBe('it');
    });
});
