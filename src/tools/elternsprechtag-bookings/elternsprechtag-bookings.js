/**
 * Elternsprechtag: Microsoft Bookings-Grundanlage per Graph.
 * Ein Dienst + alle ausgewählten Lehrkräfte als Mitarbeiter.
 */
import {
    weekdayFromIsoDate,
    weekdayLabelDe,
    toBookingsTime,
    durationIsoFromMinutes,
    buildBusinessHoursForDay,
    buildSingleDayServiceSchedulingPolicy,
    calendarDaysBetween,
    maximumAdvanceForOpenDate,
    normalizeTeacherRows,
    defaultServiceNameForDate
} from './elternsprechtag-bookings-logic.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Bookings.Read.All',
    'https://graph.microsoft.com/Bookings.ReadWrite.All',
    'https://graph.microsoft.com/Bookings.Manage.All'
];

const STORAGE_KEY = 'ms365-elternsprechtag-bookings-v1';

/** @type {Array<{ code: string, name: string, email: string, selected: boolean, skipReason: string }>} */
let teachers = [];
/** @type {Array<{ id: string, displayName: string }>} */
let businesses = [];

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function log(msg) {
    const el = $('esLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
    el.scrollTop = el.scrollHeight;
}

function clearLog() {
    const el = $('esLog');
    if (el) el.textContent = '';
}

function G() {
    const api = window.ms365GraphUnifiedGroups;
    if (!api) throw new Error('Graph-Hilfen nicht geladen (graph-unified-groups.js).');
    return api;
}

async function getToken() {
    if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
        return window.ms365AuthAcquireTokenPopup(SCOPES);
    }
    if (typeof window.ms365AuthAcquireToken === 'function') {
        return window.ms365AuthAcquireToken(SCOPES);
    }
    throw new Error('MSAL-Anmeldung nicht verfügbar.');
}

async function graphJson(method, path, token, body, extraHeaders) {
    return G().graphJson(method, path, token, body, extraHeaders);
}

function sleep(ms) {
    return new Promise(function (r) {
        setTimeout(r, ms);
    });
}

function loadPrefs() {
    try {
        return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') || {};
    } catch {
        return {};
    }
}

function savePrefs(partial) {
    const next = Object.assign({}, loadPrefs(), partial || {});
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
}

function loadTeachersFromStammdaten() {
    let list = [];
    try {
        const s = window.ms365TenantSettingsLoad && window.ms365TenantSettingsLoad();
        list = Array.isArray(s && s.teachers) ? s.teachers : [];
    } catch {
        list = [];
    }
    const prevSelected = new Set(
        teachers.filter(function (t) {
            return t.selected && t.email;
        }).map(function (t) {
            return t.email;
        })
    );
    teachers = normalizeTeacherRows({
        teachers: list,
        selectedEmails: teachers.length ? prevSelected : null
    });
    renderTeachers();
    refreshSummary();
    toast(teachers.length + ' Lehrkräfte aus den Stammdaten.');
}

function renderTeachers() {
    const body = $('esTeacherBody');
    if (!body) return;
    body.replaceChildren();
    if (!teachers.length) {
        body.innerHTML =
            '<tr><td colspan="4" class="muted">Keine Lehrkräfte in den Stammdaten – unter Stammdaten / Einrichtung pflegen.</td></tr>';
        return;
    }
    teachers.forEach(function (t, idx) {
        const tr = document.createElement('tr');
        const tdC = document.createElement('td');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = !!t.selected && !!t.email;
        cb.disabled = !t.email;
        cb.addEventListener('change', function () {
            teachers[idx].selected = cb.checked;
            refreshSummary();
        });
        tdC.appendChild(cb);
        tr.appendChild(tdC);

        const tdN = document.createElement('td');
        tdN.textContent = t.name || '—';
        tr.appendChild(tdN);

        const tdK = document.createElement('td');
        tdK.textContent = t.code || '—';
        tr.appendChild(tdK);

        const tdE = document.createElement('td');
        if (t.email) {
            tdE.textContent = t.email;
        } else {
            tdE.textContent = t.skipReason || 'E-Mail fehlt';
            tdE.className = 'muted';
        }
        tr.appendChild(tdE);
        body.appendChild(tr);
    });
}

function selectedTeachers() {
    return teachers.filter(function (t) {
        return t.selected && t.email;
    });
}

function readForm() {
    const mode = ($('esBizMode') && $('esBizMode').value) || 'new';
    const date = String(($('esDate') && $('esDate').value) || '').trim();
    const bookingOpen = String(($('esBookingOpen') && $('esBookingOpen').value) || '').trim();
    const start = String(($('esStart') && $('esStart').value) || '').trim();
    const end = String(($('esEnd') && $('esEnd').value) || '').trim();
    const minutes = Number(($('esDuration') && $('esDuration').value) || 10);
    const bizName = String(($('esBizName') && $('esBizName').value) || '').trim();
    const serviceName = String(($('esServiceName') && $('esServiceName').value) || '').trim();
    const existingId = resolveExistingBusinessId();
    const publish = !!($('esPublish') && $('esPublish').checked);
    const timeZone = String(($('esTimeZone') && $('esTimeZone').value) || 'Europe/Vienna').trim();

    return {
        mode: mode,
        date: date,
        bookingOpen: bookingOpen || date,
        start: start,
        end: end,
        minutes: minutes,
        bizName: bizName,
        serviceName: serviceName || defaultServiceNameForDate(date),
        existingId: existingId,
        publish: publish,
        timeZone: timeZone || 'Europe/Vienna'
    };
}

function refreshSummary() {
    const el = $('esSummary');
    if (!el) return;
    const f = readForm();
    const weekday = weekdayFromIsoDate(f.date);
    const sel = selectedTeachers();
    const withoutMail = teachers.filter(function (t) {
        return !t.email;
    }).length;
    const advance = maximumAdvanceForOpenDate(f.bookingOpen, f.date);
    const openDelta = calendarDaysBetween(f.bookingOpen, f.date);
    const parts = [];
    parts.push('<p><strong>' + sel.length + '</strong> Lehrkräfte als Mitarbeiter</p>');
    if (withoutMail) {
        parts.push('<p class="muted">' + withoutMail + ' ohne E-Mail werden übersprungen.</p>');
    }
    if (f.date && weekday) {
        parts.push(
            '<p>Nur am <strong>' +
                escapeHtml(f.date) +
                '</strong> (' +
                escapeHtml(weekdayLabelDe(weekday)) +
                ') ' +
                escapeHtml(f.start) +
                '–' +
                escapeHtml(f.end) +
                '</p>'
        );
    }
    if (f.bookingOpen && openDelta != null) {
        if (openDelta < 0) {
            parts.push('<p style="color:#b91c1c;">Freischalt-Datum liegt nach dem Sprechtag – bitte korrigieren.</p>');
        } else {
            parts.push(
                '<p>Buchbar ab <strong>' +
                    escapeHtml(f.bookingOpen) +
                    '</strong>' +
                    (advance ? ' (maximumAdvance ' + escapeHtml(advance) + ')' : '') +
                    '</p>'
            );
        }
    }
    parts.push(
        '<p>' +
            (f.mode === 'existing'
                ? 'Neuer Dienst auf bestehender Buchungsseite: '
                : 'Dienst: ') +
            '<strong>' +
            escapeHtml(f.serviceName) +
            '</strong>, Dauer ' +
            escapeHtml(String(f.minutes)) +
            ' Min, Lehrerwahl aktiv.</p>'
    );
    el.innerHTML = parts.join('');
}

function showBookingResult(publicUrl, meta) {
    const card = $('esResultCard');
    const input = $('esPublicUrl');
    const openBtn = $('esBtnOpenUrl');
    const hint = $('esResultHint');
    if (!card || !input) return;
    const url = String(publicUrl || '').trim();
    if (!url) {
        card.hidden = true;
        return;
    }
    input.value = url;
    if (openBtn) openBtn.href = url;
    if (hint) {
        const open = (meta && meta.bookingOpen) || '';
        const event = (meta && meta.eventDate) || '';
        hint.textContent =
            (open && event
                ? 'Termine erscheinen für Eltern ab ' + open + ' (Sprechtag ' + event + '). '
                : '') + 'Link kopieren und an Eltern weitergeben.';
    }
    card.hidden = false;
    savePrefs({ publicUrl: url });
}

async function copyPublicUrl() {
    const input = $('esPublicUrl');
    const url = input && input.value ? String(input.value).trim() : '';
    if (!url) {
        toast('Kein Link vorhanden.');
        return;
    }
    try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(url);
        } else {
            input.select();
            document.execCommand('copy');
        }
        toast('Link kopiert.');
    } catch (e) {
        toast('Kopieren fehlgeschlagen – bitte manuell markieren.');
    }
}

function syncBizModeUi() {
    const mode = ($('esBizMode') && $('esBizMode').value) || 'new';
    const newRow = $('esBizNewRow');
    const existRow = $('esBizExistRow');
    if (newRow) newRow.hidden = mode !== 'new';
    if (existRow) existRow.hidden = mode !== 'existing';
}

async function listAllPages(token, firstPath) {
    const items = [];
    let next = firstPath;
    while (next) {
        const data = await graphJson('GET', next, token);
        const arr = Array.isArray(data.value) ? data.value : [];
        for (let i = 0; i < arr.length; i++) items.push(arr[i]);
        next = data['@odata.nextLink'] || null;
        if (next && next.indexOf('https://graph.microsoft.com/v1.0') === 0) {
            next = next.slice('https://graph.microsoft.com/v1.0'.length);
        }
    }
    return items;
}

function bizPath(id) {
    return '/solutions/bookingBusinesses/' + encodeURIComponent(id);
}

function populateBusinessSelect(prevId) {
    const sel = $('esBizExisting');
    if (!sel) return;
    const keep = String(prevId || sel.value || '').trim();
    sel.replaceChildren();
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = businesses.length ? '— bitte wählen —' : '— keine Treffer —';
    sel.appendChild(opt0);
    businesses.forEach(function (b) {
        const o = document.createElement('option');
        o.value = b.id;
        o.textContent = b.displayName + ' (' + b.id + ')';
        if (b.id === keep) o.selected = true;
        sel.appendChild(o);
    });
}

function resolveExistingBusinessId() {
    const manual = String(($('esBizIdManual') && $('esBizIdManual').value) || '').trim();
    if (manual) return manual;
    return String(($('esBizExisting') && $('esBizExisting').value) || '').trim();
}

function explainBookingsListError(err) {
    const msg = String((err && err.message) || err || '');
    const lower = msg.toLowerCase();
    if (
        lower.indexOf('unknownerror') !== -1 ||
        lower.indexOf('errorexceededfindcountlimit') !== -1 ||
        lower.indexOf('too many results') !== -1
    ) {
        return (
            'Graph konnte die Bookings-Liste nicht liefern (häufig bei vielen Kalendern oder ohne Admin-Rolle). ' +
            'Tipp: Suchbegriff eingeben oder die Business-ID / Bookings-Mail manuell eintragen und „Prüfen“.'
        );
    }
    if (lower.indexOf('consent') !== -1 || lower.indexOf('forbidden') !== -1 || lower.indexOf('401') !== -1 || lower.indexOf('403') !== -1) {
        return (
            'Berechtigung fehlt oder Consent ausstehend. In Entra brauchen Sie ' +
            'Bookings.Read.All / Bookings.ReadWrite.All (Admin-Zustimmung) und ggf. Bookings.Manage.All.'
        );
    }
    return msg;
}

async function loadBusinesses() {
    clearLog();
    const query = String(($('esBizQuery') && $('esBizQuery').value) || '').trim();
    log(
        query
            ? 'Suche Bookings-Umgebungen mit query=\"' + query + '\" …'
            : 'Lade Bookings-Umgebungen (ohne Filter – kann bei Graph scheitern) …'
    );
    const token = await getToken();
    const path = query
        ? '/solutions/bookingBusinesses?query=' + encodeURIComponent(query)
        : '/solutions/bookingBusinesses';
    let list;
    try {
        list = await listAllPages(token, path);
    } catch (e) {
        const hint = explainBookingsListError(e);
        log('FEHLER beim Listen: ' + ((e && e.message) || e));
        log(hint);
        // Fallback: manuelle ID oder gespeicherte ID direkt abrufen
        const tryId = resolveExistingBusinessId() || String(loadPrefs().businessId || '').trim();
        if (tryId) {
            log('Versuche Einzelabruf: ' + tryId);
            try {
                const one = await graphJson('GET', bizPath(tryId), token);
                list = one && one.id ? [one] : [];
                log('Einzelabruf ok.');
            } catch (e2) {
                log('Einzelabruf fehlgeschlagen: ' + ((e2 && e2.message) || e2));
                toast(hint);
                throw e;
            }
        } else {
            toast(hint);
            throw e;
        }
    }
    businesses = list
        .map(function (b) {
            return {
                id: String((b && b.id) || ''),
                displayName: String((b && b.displayName) || b.id || '')
            };
        })
        .filter(function (b) {
            return b.id;
        })
        .sort(function (a, b) {
            return a.displayName.localeCompare(b.displayName, 'de');
        });
    const prefs = loadPrefs();
    populateBusinessSelect(prefs.businessId || '');
    if (businesses.length === 1 && $('esBizExisting')) {
        $('esBizExisting').value = businesses[0].id;
        if ($('esBizIdManual') && !$('esBizIdManual').value) {
            $('esBizIdManual').value = businesses[0].id;
        }
    }
    log(businesses.length + ' Treffer.');
    toast(businesses.length ? businesses.length + ' Umgebung(en) gefunden.' : 'Keine Treffer – ID manuell eintragen.');
}

async function verifyBusinessById() {
    clearLog();
    const id = resolveExistingBusinessId();
    if (!id) {
        toast('Bitte Business-ID / Bookings-Mail eintragen oder einen Treffer wählen.');
        return;
    }
    log('Prüfe Umgebung: ' + id);
    try {
        const token = await getToken();
        const one = await graphJson('GET', bizPath(id), token);
        const bid = String((one && one.id) || id);
        const name = String((one && one.displayName) || bid);
        businesses = [{ id: bid, displayName: name }];
        populateBusinessSelect(bid);
        if ($('esBizExisting')) $('esBizExisting').value = bid;
        if ($('esBizIdManual')) $('esBizIdManual').value = bid;
        savePrefs({ businessId: bid, mode: 'existing' });
        log('OK: ' + name + ' (' + bid + ')');
        if (one && one.isPublished != null) log('Veröffentlicht: ' + (one.isPublished ? 'ja' : 'nein'));
        if (one && one.publicUrl) {
            log('publicUrl: ' + one.publicUrl);
            showBookingResult(one.publicUrl, {});
        }
        toast('Umgebung gefunden: ' + name);
    } catch (e) {
        log('FEHLER: ' + ((e && e.message) || e));
        toast(
            'Umgebung nicht gefunden. ID prüfen (oft die Bookings-Mail wie name@tenant.onmicrosoft.com). ' +
                ((e && e.message) || '')
        );
    }
}

function isMailboxNotFoundError(err) {
    const msg = String((err && err.message) || err || '');
    return (
        /Bookings mailbox was not found/i.test(msg) ||
        /Mailbox does not exist/i.test(msg) ||
        /"code"\s*:\s*"NotFound"/i.test(msg) ||
        /NotFound/i.test(msg) && /mailbox/i.test(msg)
    );
}

/**
 * Nach Create braucht Exchange oft einige Sekunden, bis die Bookings-Mailbox steht.
 */
async function waitForBookingMailbox(token, businessId) {
    const maxAttempts = 15;
    for (let i = 0; i < maxAttempts; i++) {
        try {
            await graphJson('GET', bizPath(businessId), token);
            await graphJson('GET', bizPath(businessId) + '/staffMembers?$top=1', token);
            if (i > 0) log('Bookings-Mailbox ist bereit.');
            return true;
        } catch (e) {
            if (!isMailboxNotFoundError(e) && i >= 2) {
                throw e;
            }
            log('Warte auf Bookings-Mailbox-Provisionierung … (' + (i + 1) + '/' + maxAttempts + ')');
            await sleep(4000);
        }
    }
    return false;
}

async function graphJsonWithMailboxRetry(method, path, token, body, extraHeaders) {
    const maxAttempts = 8;
    let lastErr = null;
    for (let i = 0; i < maxAttempts; i++) {
        try {
            return await graphJson(method, path, token, body, extraHeaders);
        } catch (e) {
            lastErr = e;
            if (!isMailboxNotFoundError(e)) throw e;
            log(
                'Mailbox noch nicht bereit für ' +
                    method +
                    ' – warte und versuche erneut (' +
                    (i + 1) +
                    '/' +
                    maxAttempts +
                    ') …'
            );
            await sleep(4000);
        }
    }
    throw lastErr;
}

function buildCreateBusinessBody(form) {
    const body = {
        displayName: form.bizName,
        defaultCurrencyIso: 'EUR'
    };
    const weekday = weekdayFromIsoDate(form.date);
    const startTime = toBookingsTime(form.start);
    const endTime = toBookingsTime(form.end);
    const policy = buildSingleDayServiceSchedulingPolicy({
        eventDate: form.date,
        startHhmm: form.start,
        endHhmm: form.end,
        durationMin: form.minutes,
        bookingOpenDate: form.bookingOpen
    });
    if (weekday && startTime && endTime) {
        const hours = buildBusinessHoursForDay({
            weekday: weekday,
            startTime: startTime,
            endTime: endTime
        });
        if (hours) body.businessHours = hours;
    }
    if (policy) {
        body.schedulingPolicy = {
            timeSlotInterval: policy.timeSlotInterval,
            minimumLeadTime: policy.minimumLeadTime,
            maximumAdvance: policy.maximumAdvance,
            sendConfirmationsToOwner: true,
            allowStaffSelection: true
        };
    }
    return body;
}

/**
 * @returns {Promise<{ id: string, reused: boolean, freshlyCreated: boolean }>}
 */
async function ensureBusiness(token, form) {
    if (form.mode === 'existing') {
        const existingId = resolveExistingBusinessId() || form.existingId;
        if (!existingId) {
            throw new Error(
                'Bitte eine bestehende Umgebung wählen oder die Business-ID / Bookings-Mail eintragen.'
            );
        }
        log('Prüfe bestehende Umgebung: ' + existingId);
        try {
            const one = await graphJsonWithMailboxRetry('GET', bizPath(existingId), token);
            const id = String((one && one.id) || existingId);
            log('Nutze: ' + String((one && one.displayName) || id) + ' (' + id + ')');
            return { id: id, reused: true, freshlyCreated: false };
        } catch (e) {
            if (isMailboxNotFoundError(e)) {
                throw new Error(
                    'Bookings-Mailbox nicht gefunden für „' +
                        existingId +
                        '“. ID im Bookings-Portal prüfen (Geschäftsinformationen) oder die Seite dort neu anlegen.'
                );
            }
            throw e;
        }
    }
    if (!form.bizName) throw new Error('Bitte einen Namen für die neue Bookings-Umgebung angeben.');
    if (!businesses.length) {
        try {
            const list = await listAllPages(token, '/solutions/bookingBusinesses?query=' + encodeURIComponent(form.bizName));
            businesses = list
                .map(function (b) {
                    return {
                        id: String((b && b.id) || ''),
                        displayName: String((b && b.displayName) || b.id || '')
                    };
                })
                .filter(function (b) {
                    return b.id;
                });
        } catch (e) {
            log('Hinweis: Suche nach bestehender Umgebung übersprungen. ' + ((e && e.message) || e));
        }
    }
    const hit = businesses.find(function (b) {
        return String(b.displayName).toLowerCase() === form.bizName.toLowerCase();
    });
    if (hit) {
        log('Name bereits vorhanden – prüfe Mailbox: ' + hit.id);
        try {
            await graphJsonWithMailboxRetry('GET', bizPath(hit.id), token);
            log('Vorhandene Umgebung wird wiederverwendet (neuer Dienst, Öffnungszeiten der Seite bleiben).');
            return { id: hit.id, reused: true, freshlyCreated: false };
        } catch (e) {
            if (isMailboxNotFoundError(e)) {
                throw new Error(
                    'Es gibt einen Bookings-Eintrag „' +
                        form.bizName +
                        '“ (' +
                        hit.id +
                        '), aber keine Mailbox. Bitte in Bookings prüfen/löschen oder einen anderen Umgebungsnamen wählen.'
                );
            }
            throw e;
        }
    }
    log('Lege Bookings-Umgebung an: ' + form.bizName);
    const created = await graphJson('POST', '/solutions/bookingBusinesses', token, buildCreateBusinessBody(form));
    const id = String((created && created.id) || '');
    if (!id) throw new Error('Bookings-Umgebung angelegt, aber keine ID zurückgegeben.');
    businesses.push({ id: id, displayName: form.bizName });
    log('  → ID: ' + id);
    log('Warte, bis die Bookings-Mailbox provisioniert ist …');
    const ready = await waitForBookingMailbox(token, id);
    if (!ready) {
        throw new Error(
            'Umgebung angelegt (' +
                id +
                '), Mailbox aber noch nicht bereit. Bitte 1–2 Minuten warten, dann Modus „Bestehende Umgebung“ mit dieser ID erneut ausführen.'
        );
    }
    return { id: id, reused: false, freshlyCreated: true };
}

async function patchBusinessHours(token, businessId, form) {
    const weekday = weekdayFromIsoDate(form.date);
    const startTime = toBookingsTime(form.start);
    const endTime = toBookingsTime(form.end);
    if (!weekday || !startTime || !endTime) {
        throw new Error('Datum oder Uhrzeiten ungültig.');
    }
    if (form.start >= form.end) {
        throw new Error('Ende muss nach dem Start liegen.');
    }
    const hours = buildBusinessHoursForDay({
        weekday: weekday,
        startTime: startTime,
        endTime: endTime
    });
    const policy = buildSingleDayServiceSchedulingPolicy({
        eventDate: form.date,
        startHhmm: form.start,
        endHhmm: form.end,
        durationMin: form.minutes,
        bookingOpenDate: form.bookingOpen
    });
    log(
        'Setze Business-Zeiten für ' +
            form.date +
            ' (' +
            weekdayLabelDe(weekday) +
            ' ' +
            form.start +
            '–' +
            form.end +
            ') …'
    );
    await graphJsonWithMailboxRetry('PATCH', bizPath(businessId), token, {
        businessHours: hours,
        schedulingPolicy: policy
            ? {
                  timeSlotInterval: policy.timeSlotInterval,
                  minimumLeadTime: policy.minimumLeadTime,
                  maximumAdvance: policy.maximumAdvance,
                  sendConfirmationsToOwner: true,
                  allowStaffSelection: true
              }
            : undefined
    });
}

function isConflictError(err) {
    const msg = String((err && err.message) || err || '');
    return /Conflict/i.test(msg) || /already exists/i.test(msg);
}

function isUnknownBookingsError(err) {
    const msg = String((err && err.message) || err || '');
    return /UnknownError/i.test(msg);
}

function normEmail(v) {
    return String(v || '')
        .trim()
        .toLowerCase();
}

function normName(v) {
    return String(v || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ');
}

function indexStaffMembers(list) {
    const byEmail = new Map();
    const byName = new Map();
    (list || []).forEach(function (s) {
        if (!s || !s.id) return;
        const em = normEmail(s.emailAddress);
        if (em) byEmail.set(em, s);
        const nm = normName(s.displayName);
        if (nm && !byName.has(nm)) byName.set(nm, s);
    });
    return { byEmail: byEmail, byName: byName, list: list || [] };
}

async function loadStaffMembers(token, businessId) {
    let raw;
    try {
        raw = await listAllPages(token, bizPath(businessId) + '/staffMembers');
    } catch (e) {
        if (isMailboxNotFoundError(e)) {
            const ready = await waitForBookingMailbox(token, businessId);
            if (!ready) throw e;
            raw = await listAllPages(token, bizPath(businessId) + '/staffMembers');
        } else {
            throw e;
        }
    }
    // Manche Tenant-Antworten liefern in der Liste keine E-Mail → Einzelabruf.
    const full = [];
    for (let i = 0; i < raw.length; i++) {
        const s = raw[i];
        if (s && s.emailAddress) {
            full.push(s);
            continue;
        }
        if (s && s.id) {
            try {
                const one = await graphJson(
                    'GET',
                    bizPath(businessId) + '/staffMembers/' + encodeURIComponent(s.id),
                    token
                );
                full.push(one && one.id ? one : s);
            } catch {
                full.push(s);
            }
            await sleep(80);
        }
    }
    return indexStaffMembers(full);
}

function findStaffInIndex(index, teacher) {
    if (!index) return null;
    const em = normEmail(teacher && teacher.email);
    if (em && index.byEmail.has(em)) return index.byEmail.get(em);
    const nm = normName(teacher && teacher.name);
    if (nm && index.byName.has(nm)) return index.byName.get(nm);
    return null;
}

async function createStaffMember(token, businessId, teacher, timeZone, role) {
    const body = {
        displayName: teacher.name,
        emailAddress: teacher.email,
        role: role || 'guest',
        timeZone: timeZone,
        useBusinessHours: true,
        availabilityIsAffectedByPersonalCalendar: true,
        isEmailNotificationEnabled: true
    };
    return graphJson('POST', bizPath(businessId) + '/staffMembers', token, body);
}

async function syncStaff(token, businessId, selected, timeZone) {
    let index = await loadStaffMembers(token, businessId);
    log('Vorhandene Mitarbeiter in Bookings: ' + index.list.length);

    const staffIds = [];
    let created = 0;
    let reused = 0;

    for (let i = 0; i < selected.length; i++) {
        const t = selected[i];
        const em = t.email;
        let found = findStaffInIndex(index, t);
        if (found && found.id) {
            staffIds.push(String(found.id));
            reused++;
            continue;
        }
        log('  + Mitarbeiter: ' + t.name + ' <' + em + '>');
        try {
            let createdStaff = null;
            try {
                createdStaff = await createStaffMember(token, businessId, t, timeZone, 'guest');
            } catch (e1) {
                if (isConflictError(e1)) {
                    log('  → existiert bereits – lade Mitarbeiterliste neu …');
                    index = await loadStaffMembers(token, businessId);
                    found = findStaffInIndex(index, t);
                    if (found && found.id) {
                        staffIds.push(String(found.id));
                        reused++;
                        await sleep(120);
                        continue;
                    }
                    throw e1;
                }
                if (isUnknownBookingsError(e1)) {
                    log('  → UnknownError, zweiter Versuch als externalGuest …');
                    try {
                        createdStaff = await createStaffMember(
                            token,
                            businessId,
                            t,
                            timeZone,
                            'externalGuest'
                        );
                    } catch (e2) {
                        if (isConflictError(e2)) {
                            index = await loadStaffMembers(token, businessId);
                            found = findStaffInIndex(index, t);
                            if (found && found.id) {
                                staffIds.push(String(found.id));
                                reused++;
                                await sleep(120);
                                continue;
                            }
                        }
                        throw e2;
                    }
                } else {
                    throw e1;
                }
            }
            const sid = String((createdStaff && createdStaff.id) || '');
            if (!sid) throw new Error('Keine Staff-ID');
            staffIds.push(sid);
            index.byEmail.set(normEmail(em), createdStaff);
            if (normName(t.name)) index.byName.set(normName(t.name), createdStaff);
            created++;
        } catch (e) {
            log('  ! Fehler bei ' + em + ': ' + ((e && e.message) || e));
        }
        await sleep(180);
    }

    // Letzter Fallback: Liste nochmals laden und alle ausgewählten zuordnen
    if (staffIds.length < selected.length) {
        index = await loadStaffMembers(token, businessId);
        selected.forEach(function (t) {
            const f = findStaffInIndex(index, t);
            if (!f || !f.id) return;
            const id = String(f.id);
            if (staffIds.indexOf(id) === -1) {
                staffIds.push(id);
                reused++;
                log('  → nachgeladen: ' + t.name + ' <' + t.email + '>');
            }
        });
    }

    log('Mitarbeiter: ' + created + ' neu, ' + reused + ' bereits vorhanden, ' + staffIds.length + ' gesamt.');
    return staffIds;
}

async function ensureService(token, businessId, form, staffIds) {
    const duration = durationIsoFromMinutes(form.minutes);
    if (!duration) throw new Error('Slotdauer ungültig (5–120 Min).');

    const schedulingPolicy = buildSingleDayServiceSchedulingPolicy({
        eventDate: form.date,
        startHhmm: form.start,
        endHhmm: form.end,
        durationMin: form.minutes,
        bookingOpenDate: form.bookingOpen
    });
    if (!schedulingPolicy) {
        throw new Error('Scheduling-Policy konnte nicht gebaut werden (Datum/Zeiten prüfen).');
    }

    // Bestehende / wiederverwendete Buchungsseite: immer neuer Dienst.
    const forceNew = form.mode === 'existing' || !!form.forceNewService;

    let existing = null;
    if (!forceNew) {
        const services = await listAllPages(token, bizPath(businessId) + '/services');
        existing = services.find(function (s) {
            return String((s && s.displayName) || '').toLowerCase() === form.serviceName.toLowerCase();
        });
    } else {
        log('Bestehende Buchungsseite → lege neuen Dienst an (kein Update bestehender Dienste).');
        try {
            const services = await listAllPages(token, bizPath(businessId) + '/services');
            const sameName = services.filter(function (s) {
                return String((s && s.displayName) || '').toLowerCase() === form.serviceName.toLowerCase();
            });
            if (sameName.length) {
                log(
                    'Hinweis: Es gibt bereits ' +
                        sameName.length +
                        ' Dienst(e) mit diesem Namen – es wird trotzdem ein neuer angelegt.'
                );
            }
        } catch (e) {
            log('Hinweis: bestehende Dienste konnten nicht gelesen werden: ' + ((e && e.message) || e));
        }
    }

    log(
        'Dienst nur am ' +
            form.date +
            ' buchbar; Freischaltung ab ' +
            form.bookingOpen +
            ' (maximumAdvance ' +
            schedulingPolicy.maximumAdvance +
            ').'
    );

    const payload = {
        displayName: form.serviceName,
        description: 'Persönliches Gespräch am Elternsprechtag (' + form.date + ').',
        defaultDuration: duration,
        defaultPrice: 0,
        defaultPriceType: 'free',
        isHiddenFromCustomers: false,
        maximumAttendeesCount: 1,
        staffMemberIds: staffIds,
        schedulingPolicy: schedulingPolicy
    };

    if (!forceNew && existing && existing.id) {
        log('Aktualisiere bestehenden Dienst: ' + form.serviceName);
        const prevIds = Array.isArray(existing.staffMemberIds) ? existing.staffMemberIds.map(String) : [];
        const merged = Array.from(new Set(prevIds.concat(staffIds.map(String))));
        payload.staffMemberIds = merged;
        const updated = await graphJson(
            'PATCH',
            bizPath(businessId) + '/services/' + encodeURIComponent(existing.id),
            token,
            payload
        );
        return {
            id: String(existing.id),
            webUrl: (updated && updated.webUrl) || existing.webUrl || '',
            created: false
        };
    }

    log('Lege neuen Dienst an: ' + form.serviceName);
    const created = await graphJson('POST', bizPath(businessId) + '/services', token, payload);
    return {
        id: String((created && created.id) || ''),
        webUrl: (created && created.webUrl) || '',
        created: true
    };
}

async function publishBusiness(token, businessId) {
    log('Veröffentliche Buchungsseite …');
    await graphJson('POST', bizPath(businessId) + '/publish', token, {});
}

async function fetchPublicUrl(token, businessId) {
    const biz = await graphJson('GET', bizPath(businessId), token);
    return {
        publicUrl: String((biz && biz.publicUrl) || '').trim(),
        isPublished: !!(biz && biz.isPublished),
        displayName: String((biz && biz.displayName) || '')
    };
}

async function runSetup() {
    const form = readForm();
    const selected = selectedTeachers();
    clearLog();

    if (!form.date) {
        toast('Bitte das Datum des Elternsprechtags wählen.');
        return;
    }
    if (!form.bookingOpen) {
        toast('Bitte angeben, ab wann die Buchung freigeschaltet wird.');
        return;
    }
    const openDelta = calendarDaysBetween(form.bookingOpen, form.date);
    if (openDelta == null || openDelta < 0) {
        toast('Freischalt-Datum muss am Sprechtag oder davor liegen.');
        return;
    }
    if (!toBookingsTime(form.start) || !toBookingsTime(form.end)) {
        toast('Bitte Start- und Endzeit angeben.');
        return;
    }
    if (form.start >= form.end) {
        toast('Ende muss nach dem Start liegen.');
        return;
    }
    if (!selected.length) {
        toast('Bitte mindestens eine Lehrkraft mit E-Mail auswählen.');
        return;
    }
    if (!durationIsoFromMinutes(form.minutes)) {
        toast('Slotdauer muss zwischen 5 und 120 Minuten liegen.');
        return;
    }

    const btn = $('esBtnRun');
    if (btn) btn.disabled = true;
    try {
        const token = await getToken();
        const biz = await ensureBusiness(token, form);
        const businessId = biz.id;
        form.forceNewService = !!biz.reused;
        savePrefs({
            businessId: businessId,
            bizName: form.bizName,
            serviceName: form.serviceName,
            mode: biz.reused ? 'existing' : form.mode,
            bookingOpen: form.bookingOpen,
            eventDate: form.date
        });

        if (biz.reused) {
            log('Bestehende Buchungsseite: Geschäfts-Öffnungszeiten bleiben unverändert (nur neuer Dienst + Mitarbeiter).');
        } else if (biz.freshlyCreated) {
            log('Neue Umgebung: Zeiten wurden beim Anlegen gesetzt; ggf. Nachziehen …');
            try {
                await patchBusinessHours(token, businessId, form);
            } catch (e) {
                if (isMailboxNotFoundError(e)) {
                    log(
                        'Hinweis: Business-Zeiten konnten noch nicht gesetzt werden (Mailbox). ' +
                            'Der Dienst mit Tages-Verfügbarkeit wird trotzdem angelegt. ' +
                            ((e && e.message) || e)
                    );
                } else {
                    throw e;
                }
            }
        } else {
            await patchBusinessHours(token, businessId, form);
        }
        const staffIds = await syncStaff(token, businessId, selected, form.timeZone);
        if (!staffIds.length) {
            throw new Error('Kein Mitarbeiter konnte angelegt/gefunden werden.');
        }
        const service = await ensureService(token, businessId, form, staffIds);

        let publicUrl = '';
        if (form.publish) {
            try {
                await publishBusiness(token, businessId);
            } catch (e) {
                log('Hinweis: Publish fehlgeschlagen. ' + ((e && e.message) || e));
            }
        }
        try {
            const info = await fetchPublicUrl(token, businessId);
            publicUrl = info.publicUrl;
            log('Veröffentlicht: ' + (info.isPublished ? 'ja' : 'nein'));
            if (publicUrl) log('publicUrl: ' + publicUrl);
            else log('Noch keine publicUrl – Haken „veröffentlichen“ setzen und erneut ausführen.');
        } catch (e) {
            log('publicUrl konnte nicht gelesen werden: ' + ((e && e.message) || e));
        }

        if (!publicUrl && service.webUrl) {
            publicUrl = service.webUrl;
            log('Fallback Service-webUrl: ' + publicUrl);
        }

        log('');
        log('Fertig. Business-ID: ' + businessId);
        if (service.id) {
            log(
                (service.created ? 'Neuer Dienst angelegt' : 'Dienst aktualisiert') +
                    ': ' +
                    form.serviceName +
                    ' (' +
                    service.id +
                    ')'
            );
        }

        showBookingResult(publicUrl, {
            bookingOpen: form.bookingOpen,
            eventDate: form.date
        });
        if (publicUrl) toast('Grundanlage fertig – Buchungslink unten.');
        else toast('Grundanlage fertig – bitte veröffentlichen, um den Link zu erhalten.');
    } catch (e) {
        log('FEHLER: ' + ((e && e.message) || e));
        toast('Fehler: ' + ((e && e.message) || e));
    } finally {
        if (btn) btn.disabled = false;
        refreshSummary();
    }
}

function applyDefaultsFromPrefs() {
    const p = loadPrefs();
    if (p.bizName && $('esBizName')) $('esBizName').value = p.bizName;
    if (p.serviceName && $('esServiceName')) $('esServiceName').value = p.serviceName;
    if (p.mode && $('esBizMode')) {
        $('esBizMode').value = p.mode === 'existing' ? 'existing' : 'new';
        syncBizModeUi();
    }
    if ($('esBizName') && !$('esBizName').value) {
        const y = new Date().getFullYear();
        $('esBizName').value = 'Elternsprechtag ' + y;
    }
    if ($('esServiceName') && !$('esServiceName').value) {
        const dateVal = ($('esDate') && $('esDate').value) || '';
        $('esServiceName').value = defaultServiceNameForDate(dateVal);
    }
    if ($('esDate') && !$('esDate').value) {
        const d = new Date();
        const day = d.getDay();
        const add = (5 - day + 7) % 7 || 7;
        d.setDate(d.getDate() + add);
        $('esDate').value =
            d.getFullYear() +
            '-' +
            String(d.getMonth() + 1).padStart(2, '0') +
            '-' +
            String(d.getDate()).padStart(2, '0');
    }
    if ($('esBookingOpen') && !$('esBookingOpen').value) {
        if (p.bookingOpen) {
            $('esBookingOpen').value = p.bookingOpen;
        } else if ($('esDate') && $('esDate').value) {
            // Vorschlag: 14 Tage vor dem Sprechtag (nicht vor heute)
            const m = $('esDate').value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
            if (m) {
                const ev = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
                ev.setDate(ev.getDate() - 14);
                const today = new Date();
                const startToday = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 12, 0, 0);
                const open = ev.getTime() < startToday.getTime() ? startToday : ev;
                $('esBookingOpen').value =
                    open.getFullYear() +
                    '-' +
                    String(open.getMonth() + 1).padStart(2, '0') +
                    '-' +
                    String(open.getDate()).padStart(2, '0');
            }
        }
    }
    if ($('esStart') && !$('esStart').value) $('esStart').value = '15:00';
    if ($('esEnd') && !$('esEnd').value) $('esEnd').value = '19:00';
    if (p.businessId && $('esBizIdManual') && !$('esBizIdManual').value) {
        $('esBizIdManual').value = p.businessId;
    }
    if (p.publicUrl) {
        showBookingResult(p.publicUrl, {
            bookingOpen: p.bookingOpen || '',
            eventDate: p.eventDate || ($('esDate') && $('esDate').value) || ''
        });
    }
}

function wire() {
    applyDefaultsFromPrefs();
    syncBizModeUi();
    loadTeachersFromStammdaten();

    const mode = $('esBizMode');
    if (mode) {
        mode.addEventListener('change', function () {
            syncBizModeUi();
            refreshSummary();
        });
    }

    ['esDate', 'esBookingOpen', 'esStart', 'esEnd', 'esDuration', 'esServiceName', 'esBizName'].forEach(
        function (id) {
            const el = $(id);
            if (el) el.addEventListener('input', refreshSummary);
            if (el) el.addEventListener('change', refreshSummary);
        }
    );

    const dateEl = $('esDate');
    if (dateEl) {
        dateEl.addEventListener('change', function () {
            const openEl = $('esBookingOpen');
            if (openEl && !openEl.value) {
                const m = dateEl.value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
                if (m) {
                    const ev = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
                    ev.setDate(ev.getDate() - 14);
                    const today = new Date();
                    const startToday = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 12, 0, 0);
                    const open = ev.getTime() < startToday.getTime() ? startToday : ev;
                    openEl.value =
                        open.getFullYear() +
                        '-' +
                        String(open.getMonth() + 1).padStart(2, '0') +
                        '-' +
                        String(open.getDate()).padStart(2, '0');
                }
            }
            // Dienstname im Schulstil vorschlagen, wenn noch leer oder alter Jahres-Vorschlag
            const nameEl = $('esServiceName');
            if (nameEl && dateEl.value) {
                const suggested = defaultServiceNameForDate(dateEl.value);
                const cur = String(nameEl.value || '').trim();
                if (!cur || /^Termin Elternsprechtag \d{4}$/i.test(cur)) {
                    nameEl.value = suggested;
                }
            }
            refreshSummary();
        });
    }

    const btnReload = $('esBtnReloadTeachers');
    if (btnReload) btnReload.addEventListener('click', loadTeachersFromStammdaten);

    const btnAll = $('esBtnSelectAll');
    if (btnAll) {
        btnAll.addEventListener('click', function () {
            teachers.forEach(function (t) {
                if (t.email) t.selected = true;
            });
            renderTeachers();
            refreshSummary();
        });
    }
    const btnNone = $('esBtnSelectNone');
    if (btnNone) {
        btnNone.addEventListener('click', function () {
            teachers.forEach(function (t) {
                t.selected = false;
            });
            renderTeachers();
            refreshSummary();
        });
    }

    const btnBiz = $('esBtnLoadBiz');
    if (btnBiz) {
        btnBiz.addEventListener('click', function () {
            loadBusinesses().catch(function (e) {
                log('FEHLER: ' + explainBookingsListError(e));
                toast(explainBookingsListError(e));
            });
        });
    }

    const btnVerify = $('esBtnVerifyBiz');
    if (btnVerify) btnVerify.addEventListener('click', verifyBusinessById);

    const selBiz = $('esBizExisting');
    if (selBiz) {
        selBiz.addEventListener('change', function () {
            if (selBiz.value && $('esBizIdManual')) $('esBizIdManual').value = selBiz.value;
        });
    }

    const qEl = $('esBizQuery');
    if (qEl) {
        qEl.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                loadBusinesses().catch(function (e) {
                    toast(explainBookingsListError(e));
                });
            }
        });
    }

    const btnCopy = $('esBtnCopyUrl');
    if (btnCopy) btnCopy.addEventListener('click', copyPublicUrl);

    const btnRun = $('esBtnRun');
    if (btnRun) btnRun.addEventListener('click', runSetup);

    refreshSummary();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wire);
} else {
    wire();
}
