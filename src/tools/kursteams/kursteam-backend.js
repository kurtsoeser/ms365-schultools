'use strict';

const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

let pollTimer = null;
let running = false;
const loggedEntryKeys = new Set();

function resetEntryLogKeys() {
    loggedEntryKeys.clear();
}

function appendEntryLog(entry, total) {
    const prefix = '[' + entry.index + '/' + total + '] ' + entry.teamName + ': ';
    const key = entry.index + '|' + entry.status + '|' + (entry.message || '');
    if (loggedEntryKeys.has(key)) return;
    loggedEntryKeys.add(key);
    if (entry.status === 'ok') {
        appendBackendLog(prefix + 'OK → ' + entry.gruppenmail, 'ok');
    } else if (entry.status === 'error') {
        appendBackendLog(prefix + (entry.message || 'Fehler'), 'err');
    } else if (entry.status === 'running') {
        appendBackendLog(prefix + 'wird angelegt …', 'warn');
    }
}

function renderJobProgress(job) {
    if (!job || !job.entries) return;
    const done = (job.completed || 0) + (job.failed || 0);
    setProgress(
        'Job ' +
            job.status +
            ' – ' +
            done +
            '/' +
            job.total +
            ' (OK: ' +
            (job.completed || 0) +
            ', Fehler: ' +
            (job.failed || 0) +
            ')'
    );
    job.entries.forEach((entry) => appendEntryLog(entry, job.total));
}

function getApiConfig() {
    const cfg = window.MS365_KURSTEAMS_API || {};
    return {
        baseUrl: String(cfg.baseUrl || '')
            .trim()
            .replace(/\/$/, ''),
        functionKey: String(cfg.functionKey || '').trim(),
        tenantId: String(cfg.tenantId || '').trim()
    };
}

function toast(msg) {
    if (typeof ns.showToast === 'function') ns.showToast(msg);
    else if (typeof window.ms365ShowToast === 'function') window.ms365ShowToast(msg);
    else window.alert(msg);
}

function appendBackendLog(msg, kind) {
    const el = document.getElementById('kursteamBackendLog');
    if (!el) return;
    const line = document.createElement('div');
    line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
    if (kind === 'err') line.style.color = '#b00020';
    else if (kind === 'ok') line.style.color = '#0d8050';
    else if (kind === 'warn') line.style.color = '#856404';
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
}

function clearBackendLog() {
    const el = document.getElementById('kursteamBackendLog');
    if (el) el.replaceChildren();
}

function setProgress(text) {
    const el = document.getElementById('kursteamBackendProgress');
    if (el) el.textContent = text || '';
}

function setButtonsDisabled(disabled) {
    const run = document.getElementById('kursteamBackendRun');
    const health = document.getElementById('kursteamBackendHealth');
    if (run) run.disabled = disabled;
    if (health) health.disabled = disabled;
}

async function apiRequest(path, options) {
    const cfg = getApiConfig();
    const method = (options && options.method) || 'GET';
    const body = options && options.body;
    const needsKey = !(options && options.anonymous);
    const sep = path.indexOf('?') >= 0 ? '&' : '?';
    const url =
        cfg.baseUrl +
        path +
        (needsKey ? sep + 'code=' + encodeURIComponent(cfg.functionKey) : '');
    const headers = {};
    if (body !== undefined) headers['Content-Type'] = 'application/json';
    const res = await fetch(url, {
        method,
        headers,
        body: body !== undefined ? JSON.stringify(body) : undefined
    });
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    if (!res.ok) {
        const msg =
            typeof data === 'object' && data && data.error
                ? String(data.error)
                : text || String(res.status);
        throw new Error(msg);
    }
    return data;
}

async function resolveTenantIdFromLogin() {
    if (typeof window.ms365AuthEnsureInitialized === 'function') {
        await window.ms365AuthEnsureInitialized();
    }
    if (typeof window.ms365AuthAcquireToken === 'function') {
        try {
            await window.ms365AuthAcquireToken(['User.Read']);
        } catch {
            /* Anmeldung optional, wenn tenantId in Config steht */
        }
    }

    const pca =
        window.__ms365Pca ||
        (typeof window.ms365AuthEnsureInitialized === 'function'
            ? await window.ms365AuthEnsureInitialized().then(() => window.__ms365Pca)
            : null);
    const accounts = pca && typeof pca.getAllAccounts === 'function' ? pca.getAllAccounts() : [];
    const account = accounts && accounts[0];
    const tid =
        account &&
        (account.tenantId ||
            (account.idTokenClaims && account.idTokenClaims.tid) ||
            '');
    return tid ? String(tid).trim() : '';
}

async function resolveTenantId() {
    const fromLogin = await resolveTenantIdFromLogin();
    if (fromLogin) return fromLogin;

    const cfg = getApiConfig();
    if (cfg.tenantId) return cfg.tenantId;

    throw new Error(
        'tenantId fehlt: Bei Microsoft anmelden oder MS365_KURSTEAMS_API.tenantId in ms365-config.js setzen.'
    );
}

function validateConfig() {
    const cfg = getApiConfig();
    if (!cfg.baseUrl) {
        return 'MS365_KURSTEAMS_API.baseUrl fehlt in ms365-config.js';
    }
    if (!cfg.functionKey) {
        return 'MS365_KURSTEAMS_API.functionKey fehlt in ms365-config.js (Azure → Function App → App-Schlüssel)';
    }
    return '';
}

ns.refreshKursteamBackendUi = function refreshKursteamBackendUi() {
    const hint = document.getElementById('kursteamBackendConfigHint');
    const run = document.getElementById('kursteamBackendRun');
    const validTeams = (ns.teamsData || []).filter((t) => t.isValid);
    const cfgErr = validateConfig();

    if (hint) {
        if (cfgErr) {
            hint.style.display = 'block';
            hint.innerHTML =
                '<strong>Backend noch nicht konfiguriert:</strong> ' +
                cfgErr +
                ' – siehe <code>ms365-config.example.js</code>.';
        } else {
            hint.style.display = 'none';
        }
    }

    if (run) {
        run.disabled = !!cfgErr || validTeams.length === 0 || running;
    }
};

async function checkBackendHealth() {
    const cfg = getApiConfig();
    if (!cfg.baseUrl) {
        toast('MS365_KURSTEAMS_API.baseUrl fehlt in ms365-config.js');
        return;
    }
    clearBackendLog();
    appendBackendLog('Prüfe Backend …');
    try {
        const health = await apiRequest('/health', { anonymous: true });
        if (health && health.ok) {
            appendBackendLog('Backend erreichbar, Graph-Token OK.', 'ok');
            toast('Kursteams-Backend ist bereit.');
        } else {
            appendBackendLog('Health-Check negativ: ' + JSON.stringify(health), 'warn');
        }
    } catch (e) {
        const origin =
            typeof window !== 'undefined' && window.location
                ? window.location.origin || window.location.href.split('/').slice(0, 3).join('/')
                : '';
        const hint =
            origin === 'null' || String(origin).startsWith('file:')
                ? ' Seite nicht per Doppelklick (file://) öffnen – lokal z. B. „npx serve“ oder GitHub Pages nutzen.'
                : origin
                  ? ' Origin: ' + origin + ' – ggf. Schul-Firewall blockiert azurewebsites.net?'
                  : '';
        appendBackendLog('Health-Check fehlgeschlagen: ' + (e.message || e) + hint, 'err');
        toast('Backend nicht erreichbar – CORS, file:// oder Firewall prüfen.');
    }
}

function stopPolling() {
    if (pollTimer) {
        clearInterval(pollTimer);
        pollTimer = null;
    }
}

async function pollJob(jobId) {
    const maxAttempts = 72;
    let attempts = 0;

    return new Promise((resolve, reject) => {
        pollTimer = setInterval(async () => {
            attempts++;
            try {
                const job = await apiRequest('/jobs/' + encodeURIComponent(jobId));
                renderJobProgress(job);

                if (job.status === 'completed' || job.status === 'failed') {
                    stopPolling();
                    resolve(job);
                    return;
                }
                if (attempts >= maxAttempts) {
                    stopPolling();
                    reject(new Error('Timeout: Job nach 6 Minuten noch nicht abgeschlossen.'));
                }
            } catch (e) {
                stopPolling();
                reject(e);
            }
        }, 5000);
    });
}

async function runKursteamBackend() {
    if (running) return;

    const cfgErr = validateConfig();
    if (cfgErr) {
        toast(cfgErr);
        ns.refreshKursteamBackendUi();
        return;
    }

    const snapshotFn = window.ms365GetKursteamSnapshotForGraph;
    if (typeof snapshotFn !== 'function') {
        toast('Interner Fehler: Team-Daten nicht verfügbar.');
        return;
    }
    const pack = snapshotFn();
    if (!pack || !pack.teams || !pack.teams.length) {
        toast('Keine gültigen Teams – bitte Team-Namen generieren und Besitzer prüfen.');
        return;
    }
    const missing = pack.teams.filter((t) => !t.besitzer);
    if (missing.length) {
        toast('Bitte für alle Teams einen Besitzer (E-Mail / UPN) eintragen.');
        return;
    }

    running = true;
    setButtonsDisabled(true);
    clearBackendLog();
    resetEntryLogKeys();
    setProgress('');
    appendBackendLog('Start – Kursteams-Backend (' + pack.teams.length + ' Teams) …');

    try {
        const tenantId = await resolveTenantId();
        appendBackendLog('Mandant: ' + tenantId);

        const created = await apiRequest('/jobs', {
            method: 'POST',
            body: { tenantId, teams: pack.teams }
        });
        appendBackendLog(
            'Job gestartet: ' + created.jobId + ' (' + created.total + ' Teams)',
            'ok'
        );
        setProgress('Job ' + created.status + ' …');

        const finalJob = await pollJob(created.jobId);
        renderJobProgress(finalJob);

        if (finalJob.status === 'completed' && (finalJob.failed || 0) === 0) {
            appendBackendLog('Fertig – alle Teams angelegt.', 'ok');
            toast('Kursteams erfolgreich angelegt (' + finalJob.completed + ').');
        } else if (finalJob.status === 'completed' && (finalJob.failed || 0) > 0) {
            appendBackendLog(
                'Teilweise fehlgeschlagen: ' + finalJob.failed + ' Fehler.',
                'warn'
            );
            toast('Fertig mit Fehlern – Log prüfen.');
        } else {
            appendBackendLog('Job fehlgeschlagen.', 'err');
            toast('Kursteam-Anlage fehlgeschlagen – Log prüfen.');
        }
    } catch (e) {
        appendBackendLog('Fehler: ' + (e.message || e), 'err');
        toast('Backend-Fehler: ' + (e.message || e));
    } finally {
        running = false;
        stopPolling();
        setButtonsDisabled(false);
        ns.refreshKursteamBackendUi();
    }
}

window.ms365KursteamBackendRun = runKursteamBackend;
window.ms365KursteamBackendHealth = checkBackendHealth;

document.addEventListener('DOMContentLoaded', () => {
    const run = document.getElementById('kursteamBackendRun');
    const health = document.getElementById('kursteamBackendHealth');
    if (run) run.addEventListener('click', () => runKursteamBackend());
    if (health) health.addEventListener('click', () => checkBackendHealth());
    ns.refreshKursteamBackendUi();
});
