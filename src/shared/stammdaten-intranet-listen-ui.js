/**
 * Einbettung: Stammdaten → SharePoint-Intranet-Listen anlegen/abgleichen.
 * Mount: <div data-ms365-intranet-listen-mount data-mount-prefix="tsdSpo"></div>
 */
import { applyStammdatenPackagePermissions } from '../tools/sharepoint/stammdaten-liste-permissions.js';
import {
    buildStammdatenGroupFields,
    initEmbeddedPermissionsUi,
    readPermissionsFromPickers,
    persistPickersToStorage
} from '../tools/sharepoint/stammdaten-permissions-ui.js';

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function confirmAsync(msg) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, { title: 'SharePoint-Listen', confirmLabel: 'Fortfahren', cancelLabel: 'Abbrechen' });
    }
    return Promise.resolve(window.confirm(msg));
}

function api() {
    const a = window.ms365SpoStammdatenListen;
    if (!a || typeof a.syncSelectedLists !== 'function') {
        throw new Error('Listen-Synchronisation nicht geladen (sharepoint-liste-stammdaten.js).');
    }
    return a;
}

function el(id) {
    return document.getElementById(id);
}

function readRunOpts(p) {
    return {
        syncMode: !el(p + 'AlwaysNew') || !el(p + 'AlwaysNew').checked,
        removeOrphans: !el(p + 'RemoveOrphans') || el(p + 'RemoveOrphans').checked,
        klassenPersonen: !el(p + 'KlassenPersonen') || el(p + 'KlassenPersonen').checked
    };
}

function collectListOpts(p) {
    return {
        schueler: el(p + 'WantSchueler') && el(p + 'WantSchueler').checked,
        faecher: el(p + 'WantFaecher') && el(p + 'WantFaecher').checked,
        fachgruppen: el(p + 'WantFachgruppen') && el(p + 'WantFachgruppen').checked,
        arges: el(p + 'WantArge') && el(p + 'WantArge').checked,
        klassen: el(p + 'WantKlassen') && el(p + 'WantKlassen').checked,
        schuelerTitle: String(el(p + 'SchuelerName') && el(p + 'SchuelerName').value || '').trim(),
        faecherTitle: String(el(p + 'FaecherName') && el(p + 'FaecherName').value || '').trim(),
        fachgruppenTitle: String(el(p + 'FachgruppenName') && el(p + 'FachgruppenName').value || '').trim(),
        argesTitle: String(el(p + 'ArgeName') && el(p + 'ArgeName').value || '').trim(),
        klassenTitle: String(el(p + 'KlassenName') && el(p + 'KlassenName').value || '').trim()
    };
}

function logTo(p, msg) {
    const logEl = el(p + 'Log');
    if (!logEl) return;
    logEl.textContent += (logEl.textContent ? '\n' : '') + msg;
    logEl.scrollTop = logEl.scrollHeight;
}

function fillSiteUrl(p) {
    const urlEl = el(p + 'SiteUrl');
    if (!urlEl || String(urlEl.value || '').trim()) return;
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if (saved) urlEl.value = saved;
    } catch {
        /* ignore */
    }
}

function renderPanel(mount) {
    const p = mount.getAttribute('data-mount-prefix') || 'tsdSpo';
    const compact = mount.getAttribute('data-compact') === 'true';
    mount.dataset.intranetPrefix = p;
    mount.innerHTML =
        '<div class="stammdaten-intranet-listen" style="border:1px solid var(--border);border-radius:12px;padding:14px 16px;background:var(--card);">' +
        '<div style="display:flex;flex-wrap:wrap;gap:10px;align-items:flex-start;justify-content:space-between;margin-bottom:10px;">' +
        '<div><h3 style="margin:0 0 4px;font-size:1.05em;"><i class="bi bi-database"></i> Intranet-Listen aus Stammdaten</h3>' +
        '<p style="margin:0;color:var(--muted);font-size:0.9em;line-height:1.4;max-width:52ch;">Listen auf der Schulwebsite anlegen oder abgleichen – ohne separates Werkzeug. ' +
        '<a href="tools/sharepoint-liste-stammdaten.html">Erweitertes Tool</a></p></div></div>' +
        '<label style="display:block;margin-bottom:8px;"><span style="font-weight:700;font-size:0.9em;">SharePoint-Website</span>' +
        '<input type="url" id="' +
        p +
        'SiteUrl" placeholder="https://…/sites/Intranet" autocomplete="off" spellcheck="false" style="width:100%;margin-top:4px;"></label>' +
        '<div style="display:grid;gap:6px;margin:10px 0;font-size:0.88em;">' +
        checkbox(p, 'WantSchueler', 'Schülerinnen', true) +
        checkbox(p, 'WantFaecher', 'Fächer', true) +
        checkbox(p, 'WantFachgruppen', 'Fachgruppen', true) +
        checkbox(p, 'WantArge', 'ARGEs', true) +
        checkbox(p, 'WantKlassen', 'Klassen', true) +
        checkbox(p, 'KlassenPersonen', 'Klassen: Schüler als Personenfeld (M365)', true) +
        '</div>' +
        (compact
            ? ''
            : '<details style="margin:8px 0;"><summary style="cursor:pointer;font-weight:700;">Listenname je Typ</summary>' +
              nameField(p, 'SchuelerName', 'Schülerinnen') +
              nameField(p, 'FaecherName', 'Fächer') +
              nameField(p, 'FachgruppenName', 'Fachgruppen') +
              nameField(p, 'ArgeName', 'ARGEs') +
              nameField(p, 'KlassenName', 'Klassen') +
              '</details>') +
        '<details style="margin:8px 0;"><summary style="cursor:pointer;font-weight:700;">Abgleich &amp; Berechtigungen</summary>' +
        '<div style="margin-top:8px;display:grid;gap:6px;font-size:0.88em;">' +
        checkbox(p, 'AlwaysNew', 'Immer neue Listen anlegen (statt Abgleich)', false) +
        checkbox(p, 'RemoveOrphans', 'Verwaiste Zeilen entfernen', true) +
        checkbox(p, 'SkipPerms', 'Berechtigungen überspringen', false) +
        '</div>' +
        groupPickerRow(p, 'GroupAdmin', 'Verwaltung / Admin') +
        groupPickerRow(p, 'GroupLehrer', 'Lehrkräfte') +
        groupPickerRow(p, 'GroupSchueler', 'Schüler (Sammelgruppe)') +
        '<p style="margin:8px 0 0;font-size:0.82em;color:var(--muted);">Rollen wie Schularbeiten-Planer: z. B. Klassen/Fächer für Schüler nur Lesen, Schülerinnen-Stammliste ohne Schüler-Gruppe.</p>' +
        '</details>' +
        '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;">' +
        '<button type="button" class="btn btn-success" id="' +
        p +
        'BtnSync"><i class="bi bi-arrow-repeat"></i>Listen abgleichen</button>' +
        '<button type="button" class="btn" id="' +
        p +
        'BtnPerms"><i class="bi bi-shield-lock"></i>Nur Berechtigungen</button>' +
        '</div>' +
        '<pre id="' +
        p +
        'Log" class="tm-log" style="margin-top:10px;max-height:160px;overflow:auto;font-size:0.8em;" aria-live="polite"></pre>' +
        '</div>';

    if (compact) {
        ['SchuelerName', 'FaecherName', 'FachgruppenName', 'ArgeName', 'KlassenName'].forEach(function (suffix, i) {
            const defaults = ['Schülerinnen', 'Fächer', 'Fachgruppen', 'ARGEs', 'Klassen'];
            const hidden = document.createElement('input');
            hidden.type = 'hidden';
            hidden.id = p + suffix;
            hidden.value = defaults[i];
            mount.querySelector('.stammdaten-intranet-listen').appendChild(hidden);
        });
    }

    fillSiteUrl(p);
    initEmbeddedPermissionsUi(p, p + 'SkipPerms');

    const syncBtn = el(p + 'BtnSync');
    if (syncBtn) {
        syncBtn.addEventListener('click', function () {
            runSync(p).catch(function (e) {
                toast(e && e.message ? e.message : String(e));
            });
        });
    }
    const permsBtn = el(p + 'BtnPerms');
    if (permsBtn) {
        permsBtn.addEventListener('click', function () {
            runPermsOnly(p).catch(function (e) {
                toast(e && e.message ? e.message : String(e));
            });
        });
    }
}

function checkbox(p, id, label, checked) {
    return (
        '<label style="display:flex;gap:8px;align-items:center;"><input type="checkbox" id="' +
        p +
        id +
        '"' +
        (checked ? ' checked' : '') +
        '> ' +
        label +
        '</label>'
    );
}

function nameField(p, id, label) {
    const defaults = { SchuelerName: 'Schülerinnen', FaecherName: 'Fächer', FachgruppenName: 'Fachgruppen', ArgeName: 'ARGEs', KlassenName: 'Klassen' };
    return (
        '<label style="display:block;margin-top:6px;"><span style="font-size:0.85em;">' +
        label +
        '</span><input type="text" id="' +
        p +
        id +
        '" value="' +
        (defaults[id] || '') +
        '" maxlength="200" style="width:100%;margin-top:2px;"></label>'
    );
}

function groupPickerRow(p, id, label) {
    return (
        '<div style="margin-top:10px;"><label style="font-size:0.85em;font-weight:700;">' +
        label +
        '</label>' +
        '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:4px;">' +
        '<input type="text" id="' +
        p +
        id +
        '" readonly placeholder="Gruppe wählen …" style="flex:1;min-width:180px;">' +
        '<input type="hidden" id="' +
        p +
        id +
        'Id">' +
        '<button type="button" class="btn btn-sm" id="' +
        p +
        id +
        'Pick"><i class="bi bi-search"></i></button>' +
        '<button type="button" class="btn btn-sm alt" id="' +
        p +
        id +
        'Clear"><i class="bi bi-x-lg"></i></button>' +
        '</div></div>'
    );
}

async function applyPerms(webUrl, p, listOpts, write) {
    const defs = buildStammdatenGroupFields(p);
    const perms = readPermissionsFromPickers(defs, p + 'SkipPerms');
    persistPickersToStorage(defs, p + 'SkipPerms');
    return await applyStammdatenPackagePermissions(webUrl, perms, write, listOpts);
}

async function runSync(p) {
    const webUrl = String(el(p + 'SiteUrl') && el(p + 'SiteUrl').value || '').trim();
    if (!webUrl) {
        toast('Bitte die SharePoint-Website eintragen.');
        return;
    }
    const listOpts = collectListOpts(p);
    if (!listOpts.schueler && !listOpts.faecher && !listOpts.fachgruppen && !listOpts.arges && !listOpts.klassen) {
        toast('Mindestens eine Liste auswählen.');
        return;
    }
    const ok = await confirmAsync(
        'Ausgewählte Listen auf der Website abgleichen?\n\n' + webUrl + '\n\nDaten aus den lokalen Stammdaten in diesem Browser.'
    );
    if (!ok) return;

    const logEl = el(p + 'Log');
    if (logEl) logEl.textContent = '';
    const write = function (msg) {
        logTo(p, msg);
    };
    const runOpts = readRunOpts(p);

    await api().syncSelectedLists(webUrl, listOpts, write, runOpts);

    const skipPerms = el(p + 'SkipPerms') && el(p + 'SkipPerms').checked;
    if (!skipPerms) {
        try {
            await applyPerms(webUrl, p, listOpts, write);
        } catch (e) {
            write('Berechtigungen: ' + (e && e.message ? e.message : String(e)));
        }
    } else {
        write('Berechtigungen übersprungen.');
    }

    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ intranetSiteUrl: webUrl });
        }
    } catch {
        /* ignore */
    }

    toast('Intranet-Listen abgeglichen.');
}

async function runPermsOnly(p) {
    const webUrl = String(el(p + 'SiteUrl') && el(p + 'SiteUrl').value || '').trim();
    if (!webUrl) {
        toast('Bitte die SharePoint-Website eintragen.');
        return;
    }
    const logEl = el(p + 'Log');
    if (logEl) logEl.textContent = '';
    await applyPerms(webUrl, p, collectListOpts(p), function (msg) {
        logTo(p, msg);
    });
    toast('Berechtigungen angewendet.');
}

function boot() {
    document.querySelectorAll('[data-ms365-intranet-listen-mount]').forEach(function (mount) {
        if (mount.dataset.intranetMounted === '1') return;
        mount.dataset.intranetMounted = '1';
        renderPanel(mount);
    });
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();

window.ms365StammdatenIntranetListenUi = {
    runSync: runSync,
    runPermsOnly: runPermsOnly
};
