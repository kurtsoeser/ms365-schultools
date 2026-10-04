/**
 * UI-Bootstrap: Berechtigungen für sharepoint-liste-stammdaten.html
 */
import { applyStammdatenPackagePermissions } from './stammdaten-liste-permissions.js';
import {
    initStammdatenPermissionsUi,
    readPermissionsFromPickers,
    persistPickersToStorage,
    STAMMDATEN_GROUP_FIELDS
} from './stammdaten-permissions-ui.js';

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('spsLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function collectListOptsFromForm() {
    return {
        schueler: $('spsWantSchueler') && $('spsWantSchueler').checked,
        faecher: $('spsWantFaecher') && $('spsWantFaecher').checked,
        fachgruppen: $('spsWantFachgruppen') && $('spsWantFachgruppen').checked,
        arges: $('spsWantArge') && $('spsWantArge').checked,
        klassen: $('spsWantKlassen') && $('spsWantKlassen').checked,
        schuelerTitle: String($('spsSchuelerName') && $('spsSchuelerName').value || '').trim(),
        faecherTitle: String($('spsFaecherName') && $('spsFaecherName').value || '').trim(),
        fachgruppenTitle: String($('spsFachgruppenName') && $('spsFachgruppenName').value || '').trim(),
        argesTitle: String($('spsArgeName') && $('spsArgeName').value || '').trim(),
        klassenTitle: String($('spsKlassenName') && $('spsKlassenName').value || '').trim()
    };
}

window.ms365SpoStammdatenApplyPermissions = async function (webUrl, listOpts, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const perms = readPermissionsFromPickers(STAMMDATEN_GROUP_FIELDS, 'spsSkipPerms');
    persistPickersToStorage();
    return await applyStammdatenPackagePermissions(webUrl, perms, write, listOpts || collectListOptsFromForm());
};

document.addEventListener('DOMContentLoaded', function () {
    initStammdatenPermissionsUi();

    const permsBtn = $('spsmBtnPerms');
    if (permsBtn) {
        permsBtn.addEventListener('click', function () {
            const webUrl = String($('spsSiteUrl') && $('spsSiteUrl').value || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            if (
                !window.confirm(
                    'Berechtigungen auf die gewählten Stammdaten-Listen anwenden?\n\nVererbung wird gebrochen; breite Site-Gruppen entfernt; Entra-Gruppen erhalten die Rollen gemäß Profil.'
                )
            ) {
                return;
            }
            if ($('spsLog')) $('spsLog').textContent = '';
            window
                .ms365SpoStammdatenApplyPermissions(webUrl, collectListOptsFromForm(), log)
                .then(function () {
                    toast('Berechtigungen angewendet.');
                })
                .catch(function (e) {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }
});
