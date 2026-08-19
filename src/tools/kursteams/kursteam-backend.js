'use strict';

const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

let pollTimer = null;
let running = false;
let jobPollStartedAt = 0;
const loggedEntryKeys = new Set();

/** Grobe Sekunden pro Team (Graph + Teams-Bereitstellung + Pause im Backend). */
const SECONDS_PER_TEAM_ESTIMATE = 50;
const POLL_INTERVAL_MS = 5000;
const MAX_LOG_LINES = 400;
const LARGE_JOB_THRESHOLD = 25;

function formatDuration(totalSeconds) {
    const s = Math.max(0, Math.round(totalSeconds));
    if (s < 60) return 'ca. ' + s + ' s';
    const h = Math.floor(s / 3600);
    const m = Math.floor((s % 3600) / 60);
    const sec = s % 60;
    if (h > 0) return 'ca. ' + h + ' h ' + m + ' min';
    if (m > 0 && sec > 0) return 'ca. ' + m + ' min ' + sec + ' s';
    return 'ca. ' + m + ' min';
}

function estimateJobDurationSeconds(total) {
    return Math.max(30, Math.ceil(Number(total) || 0) * SECONDS_PER_TEAM_ESTIMATE);
}

function maxPollAttemptsForTotal(total) {
    const perTeamPolls = Math.ceil(SECONDS_PER_TEAM_ESTIMATE / (POLL_INTERVAL_MS / 1000));
    return Math.min(8640, Math.max(24, Math.ceil(Number(total) || 1) * perTeamPolls * 1.25));
}

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
    const total = job.total || job.entries.length || 0;
    const done = (job.completed || 0) + (job.failed || 0);
    const ok = job.completed || 0;
    const failed = job.failed || 0;
    const pct = total > 0 ? Math.min(100, Math.round((done / total) * 100)) : 0;

    let etaText = '';
    if (jobPollStartedAt && done > 0 && done < total) {
        const elapsedSec = (Date.now() - jobPollStartedAt) / 1000;
        const remainingSec = ((total - done) * elapsedSec) / done;
        etaText = 'Restzeit ' + formatDuration(remainingSec);
    } else if (jobPollStartedAt && done >= total && total > 0) {
        etaText = 'Dauer ' + formatDuration((Date.now() - jobPollStartedAt) / 1000);
    }

    setProgressBar(
        pct,
        done + ' / ' + total + ' Teams' + (failed ? ' (' + ok + ' OK, ' + failed + ' Fehler)' : ''),
        etaText
    );
    setProgress(
        'Job ' +
            job.status +
            ' – ' +
            done +
            '/' +
            total +
            ' (OK: ' +
            ok +
            ', Fehler: ' +
            failed +
            ')'
    );
    job.entries.forEach((entry) => appendEntryLog(entry, total));
}

function setProgressBar(percent, label, etaText) {
    const wrap = document.getElementById('kursteamBackendProgressWrap');
    const bar = document.getElementById('kursteamBackendProgressBar');
    const labelEl = document.getElementById('kursteamBackendProgressLabel');
    const etaEl = document.getElementById('kursteamBackendProgressEta');
    if (wrap) wrap.style.display = 'block';
    if (bar) {
        bar.style.width = percent + '%';
        bar.setAttribute('aria-valuenow', String(percent));
    }
    if (labelEl) labelEl.textContent = label || '';
    if (etaEl) etaEl.textContent = etaText || '';
}

function hideProgressBar() {
    const wrap = document.getElementById('kursteamBackendProgressWrap');
    if (wrap) wrap.style.display = 'none';
    const bar = document.getElementById('kursteamBackendProgressBar');
    if (bar) {
        bar.style.width = '0%';
        bar.setAttribute('aria-valuenow', '0');
    }
    const labelEl = document.getElementById('kursteamBackendProgressLabel');
    const etaEl = document.getElementById('kursteamBackendProgressEta');
    if (labelEl) labelEl.textContent = '';
    if (etaEl) etaEl.textContent = '';
}

function updateDurationHint(teamCount) {
    const el = document.getElementById('kursteamBackendDurationHint');
    if (!el) return;
    const n = Number(teamCount) || 0;
    if (n < 1) {
        el.style.display = 'none';
        el.innerHTML = '';
        return;
    }
    const est = estimateJobDurationSeconds(n);
    el.style.display = 'block';
    if (n >= LARGE_JOB_THRESHOLD) {
        el.className = 'alert alert-warning';
        el.innerHTML =
            '<strong>Viele Teams (' +
            n +
            '):</strong> Die Online-Anlage kann ' +
            formatDuration(est) +
            ' dauern (oft länger bei Microsoft-Drosselung). ' +
            '<strong>Browser-Tab geöffnet lassen</strong> – Fortschritt aktualisiert sich alle 5 Sekunden. ' +
            'Das Protokoll scrollt; bei sehr vielen Einträgen werden ältere Zeilen aus dem sichtbaren Bereich entfernt.';
    } else {
        el.className = 'alert alert-info';
        el.innerHTML =
            'Geschätzte Laufzeit für ' +
            n +
            ' Team(s): ' +
            formatDuration(est) +
            '. Das Protokoll darunter hat feste Höhe und scrollt automatisch.';
    }
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
    while (el.childElementCount > MAX_LOG_LINES) {
        el.removeChild(el.firstChild);
    }
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

function resolveMailDomain(pack) {
    if (pack && pack.mailDomain) {
        return String(pack.mailDomain)
            .trim()
            .replace(/^@+/, '');
    }
    if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
        const d = String(window.ms365GetSchoolDomainNoAt() || '')
            .trim()
            .replace(/^@+/, '');
        if (d) return d;
    }
    if (pack && pack.teams && pack.teams.length) {
        return emailDomain(pack.teams[0].besitzer);
    }
    return '';
}

function emailDomain(addr) {
    const m = String(addr || '')
        .trim()
        .match(/@([^@]+)$/i);
    return m ? m[1].toLowerCase() : '';
}

async function resolveTenantIdFromLogin() {
    if (typeof window.ms365AuthGetTenantId === 'function') {
        return await window.ms365AuthGetTenantId();
    }
    return '';
}

function getLoginUpn() {
    if (typeof window.ms365AuthGetUserPrincipalName === 'function') {
        return window.ms365AuthGetUserPrincipalName();
    }
    return '';
}

/**
 * @param {string} tenantId
 * @param {Array<{ besitzer: string }>} teams
 * @returns {{ ok: boolean, message?: string }}
 */
function validateTenantContext(tenantId, teams) {
    const loginUpn = getLoginUpn();
    const loginDomain = emailDomain(loginUpn);

    if (!tenantId) {
        return {
            ok: false,
            message:
                'Kein Mandant erkannt. Bitte unten links bei Microsoft mit dem Schul-Konto anmelden.'
        };
    }

    if (!loginUpn) {
        return {
            ok: false,
            message:
                'Bitte bei Microsoft anmelden, bevor Kursteams online angelegt werden.'
        };
    }

    const ownerDomains = new Set(
        teams.map((t) => emailDomain(t.besitzer)).filter(Boolean)
    );
    if (loginDomain && ownerDomains.size) {
        const foreign = [...ownerDomains].filter((d) => d !== loginDomain);
        if (foreign.length) {
            return {
                ok: false,
                message:
                    'Die Besitzer-E-Mails (' +
                    [...ownerDomains].join(', ') +
                    ') passen nicht zum angemeldeten Konto (' +
                    loginUpn +
                    '). Bitte Konto wechseln oder Besitzer in Schritt 5 korrigieren.'
            };
        }
    }

    return { ok: true };
}

async function resolveTenantId() {
    const fromLogin = await resolveTenantIdFromLogin();
    if (fromLogin) return fromLogin;

    throw new Error(
        'Mandant nicht erkannt. Bitte unten links bei Microsoft mit dem Konto des Ziel-Mandanten anmelden.'
    );
}

async function updateBackendTenantHint() {
    const el = document.getElementById('kursteamBackendTenantHint');
    if (!el) return;

    const loginUpn = getLoginUpn();
    let tenantId = '';
    try {
        tenantId = await resolveTenantIdFromLogin();
    } catch {
        tenantId = '';
    }

    if (!loginUpn && !tenantId) {
        el.style.display = 'block';
        el.className = 'alert alert-warning';
        el.innerHTML =
            '<strong>Bitte bei Microsoft anmelden</strong> – mit dem Konto der Schule, in der die Kursteams angelegt werden sollen.';
        return;
    }

    const pack =
        typeof window.ms365GetKursteamSnapshotForGraph === 'function'
            ? window.ms365GetKursteamSnapshotForGraph()
            : null;
    if (pack && pack.teams && pack.teams.length) {
        const check = validateTenantContext(tenantId, pack.teams);
        if (!check.ok) {
            el.style.display = 'block';
            el.className = 'alert alert-warning';
            el.innerHTML = '<strong>Hinweis:</strong> ' + check.message;
            return;
        }
    }

    el.style.display = 'none';
    el.innerHTML = '';
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

    updateDurationHint(validTeams.length);
    updateBackendTenantHint();
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

async function pollJob(jobId, totalTeams) {
    const maxAttempts = maxPollAttemptsForTotal(totalTeams);
    let attempts = 0;
    jobPollStartedAt = Date.now();

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
                    reject(
                        new Error(
                            'Timeout: Job nach ' +
                                formatDuration(maxAttempts * (POLL_INTERVAL_MS / 1000)) +
                                ' noch nicht abgeschlossen – Backend läuft ggf. weiter; Job-ID im Protokoll.'
                        )
                    );
                }
            } catch (e) {
                stopPolling();
                reject(e);
            }
        }, POLL_INTERVAL_MS);
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
    jobPollStartedAt = 0;
    hideProgressBar();
    setProgress('');
    const teamCount = pack.teams.length;
    const estSec = estimateJobDurationSeconds(teamCount);
    appendBackendLog('Start – Kursteams-Backend (' + teamCount + ' Teams) …');
    appendBackendLog(
        'Geschätzte Laufzeit: ' +
            formatDuration(estSec) +
            (teamCount >= LARGE_JOB_THRESHOLD
                ? ' – bei vielen Teams kann es deutlich länger dauern; Tab geöffnet lassen.'
                : ''),
        teamCount >= LARGE_JOB_THRESHOLD ? 'warn' : undefined
    );

    try {
        const tenantId = await resolveTenantId();
        const tenantCheck = validateTenantContext(tenantId, pack.teams);
        if (!tenantCheck.ok) {
            throw new Error(tenantCheck.message);
        }

        const created = await apiRequest('/jobs', {
            method: 'POST',
            body: {
                tenantId,
                mailDomain: resolveMailDomain(pack),
                teams: pack.teams
            }
        });
        appendBackendLog(
            'Job gestartet: ' + created.jobId + ' (' + created.total + ' Teams)',
            'ok'
        );
        setProgressBar(0, '0 / ' + created.total + ' Teams', 'Restzeit ' + formatDuration(estSec));
        setProgress('Job ' + created.status + ' …');

        const finalJob = await pollJob(created.jobId, created.total);
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
