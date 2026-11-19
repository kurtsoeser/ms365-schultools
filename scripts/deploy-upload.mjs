/**
 * Verschlüsselter Upload eines Build-Ordners per lftp (GitHub Actions).
 *
 * Ersetzt den FTP-Upload im Klartext: Zugangsdaten und Dateien gehen nur über
 * FTPS (FTP mit erzwungenem TLS, Zertifikat wird geprüft) oder SFTP (Host-Key
 * wird geprüft). Unverschlüsseltes FTP wird abgelehnt.
 *
 * Umgebungsvariablen:
 *   DEPLOY_PROTOCOL         ftps (Standard) | sftp
 *   DEPLOY_HOST             Servername (muss bei FTPS zum Zertifikat passen)
 *   DEPLOY_USER             Benutzername
 *   LFTP_PASSWORD           Passwort (liest lftp selbst aus der Umgebung)
 *   DEPLOY_REMOTE_DIR       Zielordner auf dem Server, z. B. /app/
 *   DEPLOY_LOCAL_DIR        lokaler Ordner, z. B. dist
 *   DEPLOY_PORT             optional
 *   DEPLOY_SSH_KNOWN_HOSTS  nur SFTP: known_hosts-Zeile(n) des Servers
 *
 * Es wird nichts auf dem Server gelöscht (kein --delete): Im Ziel können
 * Dateien liegen, die nicht aus diesem Repo stammen.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const CA_BUNDLE = '/etc/ssl/certs/ca-certificates.crt';

/** lftp-Argument in doppelten Anführungszeichen. */
function quote(value) {
    return `"${String(value).replace(/[\\"]/g, '\\$&')}"`;
}

/**
 * lftp hängt bei Zielpfaden mit Schrägstrich am Ende den Quellordnernamen an
 * (…/app/dist). Deshalb ohne Schrägstrich übergeben; "/" bleibt "/".
 */
export function normalizeRemoteDir(dir) {
    const trimmed = String(dir || '').trim();
    if (!trimmed) return '';
    return trimmed.replace(/\/+$/, '') || '/';
}

/**
 * @param {{ protocol?: string, host?: string, user?: string, remoteDir?: string,
 *           localDir?: string, port?: string, knownHostsFile?: string }} opts
 * @returns {string} lftp-Skript (enthält kein Passwort)
 */
export function buildLftpScript(opts) {
    const protocol = String(opts.protocol || 'ftps').trim().toLowerCase();
    if (protocol !== 'ftps' && protocol !== 'sftp') {
        throw new Error(`DEPLOY_PROTOCOL "${protocol}" wird nicht unterstützt – erlaubt sind ftps und sftp (kein unverschlüsseltes FTP).`);
    }
    const host = String(opts.host || '').trim();
    const user = String(opts.user || '').trim();
    const remoteDir = normalizeRemoteDir(opts.remoteDir);
    const localDir = String(opts.localDir || '').trim().replace(/[\\/]+$/, '');
    const port = String(opts.port || '').trim();

    if (!/^[A-Za-z0-9.-]+$/.test(host)) throw new Error('DEPLOY_HOST fehlt oder ist kein reiner Servername (ohne ftp:// und ohne Pfad).');
    if (!user) throw new Error('DEPLOY_USER fehlt.');
    if (!remoteDir) throw new Error('DEPLOY_REMOTE_DIR fehlt.');
    if (!localDir) throw new Error('DEPLOY_LOCAL_DIR fehlt.');
    if (port && !/^\d+$/.test(port)) throw new Error('DEPLOY_PORT muss eine Zahl sein.');
    for (const value of [user, remoteDir, localDir]) {
        if (/[\r\n]/.test(value)) throw new Error('Zeilenumbruch in einer Deploy-Einstellung.');
    }

    const lines = [
        'set cmd:fail-exit true',
        'set net:max-retries 3',
        'set net:timeout 30',
        'set net:reconnect-interval-base 5'
    ];
    if (protocol === 'ftps') {
        lines.push(
            // Ohne TLS wird weder angemeldet noch übertragen.
            'set ftp:ssl-force true',
            'set ftp:ssl-protect-data true',
            'set ftp:ssl-protect-list true',
            'set ssl:verify-certificate true',
            'set ssl:check-hostname true',
            `set ssl:ca-file ${quote(CA_BUNDLE)}`
        );
    } else {
        if (!opts.knownHostsFile) {
            throw new Error('SFTP braucht das Secret DEPLOY_SSH_KNOWN_HOSTS (Host-Key des Servers), sonst ist der Server nicht überprüfbar.');
        }
        const ssh = `ssh -a -x -o StrictHostKeyChecking=yes -o GlobalKnownHostsFile=/dev/null -o UserKnownHostsFile=${opts.knownHostsFile}`;
        lines.push('set sftp:auto-confirm no', `set sftp:connect-program ${quote(ssh)}`);
    }
    const scheme = protocol === 'ftps' ? 'ftp' : 'sftp';
    lines.push(
        `open --env-password -u ${quote(user)}${port ? ` -p ${port}` : ''} ${quote(`${scheme}://${host}`)}`,
        `mirror --reverse --verbose=1 --parallel=2 --no-perms ${quote(localDir)} ${quote(remoteDir)}`,
        'bye'
    );
    return `${lines.join('\n')}\n`;
}

function main() {
    const env = process.env;
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-deploy-'));
    try {
        if (!env.LFTP_PASSWORD) throw new Error('LFTP_PASSWORD (Secret FTP_PASS) fehlt.');
        const localDir = String(env.DEPLOY_LOCAL_DIR || '').trim();
        if (localDir && !fs.existsSync(localDir)) throw new Error(`Lokaler Ordner "${localDir}" existiert nicht.`);

        let knownHostsFile = '';
        const knownHosts = String(env.DEPLOY_SSH_KNOWN_HOSTS || '').trim();
        if (knownHosts) {
            knownHostsFile = path.join(tmpDir, 'known_hosts');
            fs.writeFileSync(knownHostsFile, `${knownHosts}\n`, { mode: 0o600 });
        }
        const script = buildLftpScript({
            protocol: env.DEPLOY_PROTOCOL,
            host: env.DEPLOY_HOST,
            user: env.DEPLOY_USER,
            remoteDir: env.DEPLOY_REMOTE_DIR,
            localDir,
            port: env.DEPLOY_PORT,
            knownHostsFile
        });
        const scriptFile = path.join(tmpDir, 'upload.lftp');
        fs.writeFileSync(scriptFile, script, { mode: 0o600 });

        const result = spawnSync('lftp', ['-f', scriptFile], { stdio: 'inherit' });
        if (result.error) throw new Error(`lftp ließ sich nicht starten: ${result.error.message}`);
        if (result.status !== 0) {
            throw new Error(
                `Upload fehlgeschlagen (lftp Exit-Code ${result.status}). Bei Zertifikatsfehlern: FTP_HOST muss der Servername ` +
                    'sein, auf den das TLS-Zertifikat des Hosters ausgestellt ist.'
            );
        }
        console.log('Upload abgeschlossen.');
    } catch (err) {
        console.error(`::error::${err.message}`);
        process.exitCode = 1;
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    main();
}
