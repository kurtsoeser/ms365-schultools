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
    normalizeTeacherRows
} from './elternsprechtag-bookings-logic.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Bookings.ReadWrite.All'
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
    const existingId = String(($('esBizExisting') && $('esBizExisting').value) || '').trim();
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
        serviceName: serviceName || 'Elternsprechtag (' + minutes + ' Min)',
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
        '<p>Ein Dienst: <strong>' +
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

async function loadBusinesses() {
    clearLog();
    log('Melde an und lade Bookings-Umgebungen …');
    const token = await getToken();
    const list = await listAllPages(token, '/solutions/bookingBusinesses');
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
    const sel = $('esBizExisting');
    if (sel) {
        const prefs = loadPrefs();
        const prev = prefs.businessId || '';
        sel.replaceChildren();
        const opt0 = document.createElement('option');
        opt0.value = '';
        opt0.textContent = businesses.length ? '— bitte wählen —' : 'Keine Umgebung gefunden';
        sel.appendChild(opt0);
        businesses.forEach(function (b) {
            const o = document.createElement('option');
            o.value = b.id;
            o.textContent = b.displayName + ' (' + b.id + ')';
            if (b.id === prev) o.selected = true;
            sel.appendChild(o);
        });
    }
    log(businesses.length + ' Bookings-Umgebung(en) gefunden.');
    toast(businesses.length + ' Bookings-Umgebung(en) geladen.');
}

async function ensureBusiness(token, form) {
    if (form.mode === 'existing') {
        if (!form.existingId) throw new Error('Bitte eine bestehende Bookings-Umgebung wählen.');
        log('Nutze bestehende Umgebung: ' + form.existingId);
        return form.existingId;
    }
    if (!form.bizName) throw new Error('Bitte einen Namen für die neue Bookings-Umgebung angeben.');
    // Bestehende Umgebungen nachladen (Duplikate vermeiden)
    if (!businesses.length) {
        try {
            const list = await listAllPages(token, '/solutions/bookingBusinesses');
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
            log('Hinweis: bestehende Umgebungen konnten nicht geladen werden – lege neu an. ' + ((e && e.message) || e));
        }
    }
    const hit = businesses.find(function (b) {
        return String(b.displayName).toLowerCase() === form.bizName.toLowerCase();
    });
    if (hit) {
        log('Umgebung mit diesem Namen existiert bereits – verwende sie: ' + hit.id);
        return hit.id;
    }
    log('Lege Bookings-Umgebung an: ' + form.bizName);
    const created = await graphJson('POST', '/solutions/bookingBusinesses', token, {
        displayName: form.bizName,
        defaultCurrencyIso: 'EUR'
    });
    const id = String((created && created.id) || '');
    if (!id) throw new Error('Bookings-Umgebung angelegt, aber keine ID zurückgegeben.');
    businesses.push({ id: id, displayName: form.bizName });
    log('  → ID: ' + id);
    return id;
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
    await graphJson('PATCH', bizPath(businessId), token, {
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

async function syncStaff(token, businessId, selected, timeZone) {
    const existing = await listAllPages(token, bizPath(businessId) + '/staffMembers');
    const byEmail = new Map();
    existing.forEach(function (s) {
        const em = String((s && s.emailAddress) || '')
            .trim()
            .toLowerCase();
        if (em) byEmail.set(em, s);
    });

    const staffIds = [];
    let created = 0;
    let reused = 0;

    for (let i = 0; i < selected.length; i++) {
        const t = selected[i];
        const em = t.email;
        const found = byEmail.get(em);
        if (found && found.id) {
            staffIds.push(String(found.id));
            reused++;
            continue;
        }
        log('  + Mitarbeiter: ' + t.name + ' <' + em + '>');
        try {
            const body = {
                displayName: t.name,
                emailAddress: em,
                role: 'guest',
                timeZone: timeZone,
                useBusinessHours: true,
                availabilityIsAffectedByPersonalCalendar: true,
                isEmailNotificationEnabled: true
            };
            const createdStaff = await graphJson(
                'POST',
                bizPath(businessId) + '/staffMembers',
                token,
                body
            );
            const sid = String((createdStaff && createdStaff.id) || '');
            if (!sid) throw new Error('Keine Staff-ID');
            staffIds.push(sid);
            byEmail.set(em, createdStaff);
            created++;
        } catch (e) {
            log('  ! Fehler bei ' + em + ': ' + ((e && e.message) || e));
        }
        await sleep(180);
    }

    log('Mitarbeiter: ' + created + ' neu, ' + reused + ' bereits vorhanden, ' + staffIds.length + ' gesamt.');
    return staffIds;
}

async function ensureService(token, businessId, form, staffIds) {
    const services = await listAllPages(token, bizPath(businessId) + '/services');
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

    const existing = services.find(function (s) {
        return String((s && s.displayName) || '').toLowerCase() === form.serviceName.toLowerCase();
    });

    if (existing && existing.id) {
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
            webUrl: (updated && updated.webUrl) || existing.webUrl || ''
        };
    }

    log('Lege Dienst an: ' + form.serviceName);
    const created = await graphJson('POST', bizPath(businessId) + '/services', token, payload);
    return {
        id: String((created && created.id) || ''),
        webUrl: (created && created.webUrl) || ''
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
        const businessId = await ensureBusiness(token, form);
        savePrefs({
            businessId: businessId,
            bizName: form.bizName,
            serviceName: form.serviceName,
            mode: form.mode,
            bookingOpen: form.bookingOpen,
            eventDate: form.date
        });

        await patchBusinessHours(token, businessId, form);
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
        if (service.id) log('Service-ID: ' + service.id);

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
        $('esServiceName').value = 'Elternsprechtag (10 Min)';
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
    if (mode) mode.addEventListener('change', syncBizModeUi);

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
            if (!openEl || openEl.value) return;
            const m = dateEl.value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
            if (!m) return;
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
                log('FEHLER: ' + ((e && e.message) || e));
                toast((e && e.message) || String(e));
            });
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
