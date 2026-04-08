import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

/**
 * Lädt ein Browser-IIFE-Skript (window.*) in einer vm-Sandbox (ohne DOM).
 * @param {string} relFromProjectRoot z. B. "src/tools/kursteams/kursteam-filter-logic.js"
 */
export function loadScript(relFromProjectRoot) {
    const full = join(projectRoot, relFromProjectRoot);
    const code = readFileSync(full, 'utf8');
    const sandbox = { console };
    sandbox.window = sandbox;
    createContext(sandbox);
    runInContext(code, sandbox, { filename: full });
    return sandbox;
}
