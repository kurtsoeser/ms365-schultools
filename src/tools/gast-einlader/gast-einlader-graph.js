/**
 * Graph-Transport für Gast-Einlader (Analyse 02 Phase B).
 * Nutzt shared/graph-client.js.
 */
import { getGraphToken, graphRequest, graphJson, sleep } from '../../shared/graph-client.js';

export const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/RoleManagement.ReadWrite.Directory',
    'https://graph.microsoft.com/Policy.ReadWrite.Authorization'
];

export async function giGetToken() {
    return getGraphToken(GRAPH_SCOPES);
}

export { graphRequest, graphJson, sleep };
