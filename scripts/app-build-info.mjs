import fs from 'node:fs/promises';
import path from 'node:path';

export const APP_BUILD_INFO_FILENAME = 'app-build.json';

/**
 * Zeitpunkt der Veröffentlichung = Zeitpunkt des Produktions-Builds
 * (GitHub Pages deployt direkt danach).
 * @param {Date} [now]
 */
export function createAppBuildInfo(now = new Date()) {
    const publishedAt = now instanceof Date && !Number.isNaN(now.getTime())
        ? now.toISOString()
        : new Date().toISOString();
    return { publishedAt };
}

/**
 * @param {string} distRoot
 * @param {{ publishedAt: string }} [info]
 */
export async function writeAppBuildInfo(distRoot, info = createAppBuildInfo()) {
    const dest = path.join(distRoot, APP_BUILD_INFO_FILENAME);
    await fs.writeFile(dest, `${JSON.stringify(info)}\n`, 'utf8');
    return dest;
}

function isBuildInfoRequest(url) {
    const pathname = String(url || '').split('?')[0];
    return pathname === '/app-build.json' || pathname.endsWith('/app-build.json');
}

/** Vite-Plugin: liefert die Datei im Dev-Server, ohne sie ins Repo zu schreiben. */
export function appBuildInfoPlugin() {
    const info = createAppBuildInfo();
    return {
        name: 'app-build-info',
        configureServer(server) {
            server.middlewares.use((req, res, next) => {
                if (!isBuildInfoRequest(req.url)) {
                    next();
                    return;
                }
                res.setHeader('Content-Type', 'application/json; charset=utf-8');
                res.setHeader('Cache-Control', 'no-store');
                res.end(JSON.stringify(info));
            });
        }
    };
}
