/**
 * Schul-IT: Entra-Gruppen + Werkzeug-Matrix für das Dashboard.
 */
import {
    loadDashboardAudienceGroups,
    saveDashboardAudienceGroups
} from './dashboard-audience-groups-store.js';
import { mountDashboardAudienceAdmin } from './dashboard-audience-admin-ui.js';
import {
    wireEntraGroupPickerFields,
    readGroupPickerField,
    fillGroupPickerField
} from './entra-group-picker.js';
import { clearDashboardPersonaCache } from './dashboard-persona-session.js';

const FIELDS = {
    lehrer: {
        labelInputId: 'dashAudLehrerLabel',
        idInputId: 'dashAudLehrerId',
        pickBtnId: 'dashAudLehrerPick',
        clearBtnId: 'dashAudLehrerClear',
        dialogTitle: 'Entra-Gruppe: alle Lehrkräfte'
    },
    schueler: {
        labelInputId: 'dashAudSchuelerLabel',
        idInputId: 'dashAudSchuelerId',
        pickBtnId: 'dashAudSchuelerPick',
        clearBtnId: 'dashAudSchuelerClear',
        dialogTitle: 'Entra-Gruppe: alle Schülerinnen/Schüler'
    }
};

function persistGroupsFromForm() {
    const lehrer = readGroupPickerField(FIELDS.lehrer);
    const schueler = readGroupPickerField(FIELDS.schueler);
    saveDashboardAudienceGroups({
        groupLehrerId: lehrer.id,
        groupLehrerName: lehrer.label,
        groupSchuelerId: schueler.id,
        groupSchuelerName: schueler.label
    });
    clearDashboardPersonaCache();
    try {
        window.dispatchEvent(new CustomEvent('ms365-dashboard-persona-ready'));
    } catch {
        /* ignore */
    }
}

function loadGroupsIntoForm() {
    const cfg = loadDashboardAudienceGroups();
    fillGroupPickerField(FIELDS.lehrer, { id: cfg.groupLehrerId, label: cfg.groupLehrerName });
    fillGroupPickerField(FIELDS.schueler, { id: cfg.groupSchuelerId, label: cfg.groupSchuelerName });
}

export function mountDashboardWerkzeugZugriffPage() {
    const root = document.getElementById('dashWerkzeugZugriffRoot');
    if (!root) return;

    root.innerHTML = `
        <section class="dash-access-groups tile" aria-labelledby="dashAudGroupsTitle">
            <div class="tile-head">
                <h2 id="dashAudGroupsTitle">Entra-Gruppen für das Dashboard</h2>
                <p class="tile-subtitle">
                    Wer in der <strong>Lehrer-Gruppe</strong> ist, sieht das eingeschränkte Lehrkraft-Dashboard;
                    wer in der <strong>Schüler-Gruppe</strong> ist, die Schüler-Ansicht.
                    Global Administratorinnen/Administratoren sehen immer alles und pflegen Stammdaten.
                    Ist jemand in keiner Gruppe, gilt die Schüler-Ansicht (Minimum).
                </p>
            </div>
            <div class="fr-setup-perm-grid dash-wz-perm-grid">
                <article class="fr-setup-perm-card">
                    <div class="fr-setup-perm-card__head">
                        <span class="fr-setup-perm-card__icon" aria-hidden="true"><i class="bi bi-person-workspace"></i></span>
                        <div>
                            <h4>Lehrkräfte-Gruppe</h4>
                            <p>Alle Lehrkräfte der Schule (Microsoft-365- oder Sicherheitsgruppe).</p>
                        </div>
                    </div>
                    <label class="visually-hidden" for="dashAudLehrerLabel">Lehrkräfte-Gruppe</label>
                    <div class="fr-setup-egp fr-setup-egp--tight">
                        <input type="text" id="dashAudLehrerLabel" readonly placeholder="Noch nicht gewählt" />
                        <input type="hidden" id="dashAudLehrerId" />
                        <button type="button" class="btn btn-sm" id="dashAudLehrerPick"><i class="bi bi-search" aria-hidden="true"></i> Wählen</button>
                        <button type="button" class="btn btn-sm alt" id="dashAudLehrerClear" title="Zuordnung entfernen" aria-label="Lehrkräfte-Gruppe entfernen"><i class="bi bi-x-lg" aria-hidden="true"></i></button>
                    </div>
                </article>
                <article class="fr-setup-perm-card">
                    <div class="fr-setup-perm-card__head">
                        <span class="fr-setup-perm-card__icon" aria-hidden="true"><i class="bi bi-mortarboard"></i></span>
                        <div>
                            <h4>Schüler-Gruppe</h4>
                            <p>Alle Schülerinnen und Schüler (Microsoft-365- oder Sicherheitsgruppe).</p>
                        </div>
                    </div>
                    <label class="visually-hidden" for="dashAudSchuelerLabel">Schüler-Gruppe</label>
                    <div class="fr-setup-egp fr-setup-egp--tight">
                        <input type="text" id="dashAudSchuelerLabel" readonly placeholder="Noch nicht gewählt" />
                        <input type="hidden" id="dashAudSchuelerId" />
                        <button type="button" class="btn btn-sm" id="dashAudSchuelerPick"><i class="bi bi-search" aria-hidden="true"></i> Wählen</button>
                        <button type="button" class="btn btn-sm alt" id="dashAudSchuelerClear" title="Zuordnung entfernen" aria-label="Schüler-Gruppe entfernen"><i class="bi bi-x-lg" aria-hidden="true"></i></button>
                    </div>
                </article>
            </div>
            <p class="dash-access-groups__note">Speicherung lokal (<code>ms365-dashboard-audience-groups-v1</code>). Mit <strong>In SharePoint sichern</strong> oben werden Gruppen und Werkzeug-Matrix in die IT-Bibliothek übernommen.</p>
        </section>
        <div data-ms365-dashboard-access-mount></div>`;

    loadGroupsIntoForm();
    wireEntraGroupPickerFields({
        fields: [FIELDS.lehrer, FIELDS.schueler],
        onChange: persistGroupsFromForm
    });

    mountDashboardAudienceAdmin('[data-ms365-dashboard-access-mount]');
}
