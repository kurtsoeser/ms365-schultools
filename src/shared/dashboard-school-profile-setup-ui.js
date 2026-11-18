/**
 * IT-Schritt auf „Portal vorbereiten“: Logo (Upload oder HTTPS-URL) + Adresse → setup.schoolProfile.
 */
import {
    emptySchoolProfile,
    hasSchoolLogo,
    normalizeSchoolLogoUrl,
    normalizeSchoolProfile,
    resolveSchoolLogoSrc,
    SCHOOL_LOGO_MAX_BYTES
} from './school-profile-logic.js';
import { notifyAppLocalDataChanged } from './app-local-data-notify.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else if (msg) window.alert(msg);
}

function readSchoolProfile() {
    try {
        const setup =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        return normalizeSchoolProfile(setup && setup.schoolProfile);
    } catch {
        return emptySchoolProfile();
    }
}

function readSchoolName() {
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            return String((s && s.schoolName) || '').trim();
        }
    } catch {
        /* ignore */
    }
    try {
        const api = window.ms365AppDataV2;
        if (api && typeof api.getContainer === 'function') {
            const c = api.getContainer();
            return String((c && c.core && c.core.schoolName) || '').trim();
        }
    } catch {
        /* ignore */
    }
    return '';
}

function patchSchoolProfile(partial) {
    if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') {
        toast('Stammdaten-Speicher nicht verfügbar.');
        return;
    }
    const cur = readSchoolProfile();
    const next = normalizeSchoolProfile(Object.assign({}, cur, partial || {}));
    window.ms365AppDataV2.patchSetup({ schoolProfile: next });
    notifyAppLocalDataChanged('school-profile');
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-tenant-settings-changed', {
                detail: { reason: 'autosave', source: 'dashboard-school-profile-setup' }
            })
        );
    } catch {
        /* ignore */
    }
}

function patchSchoolName(name) {
    const schoolName = String(name || '').trim();
    try {
        const api = window.ms365AppDataV2;
        if (!api || typeof api.getContainer !== 'function' || typeof api.setContainer !== 'function') {
            return;
        }
        const c = api.getContainer();
        if (!c || !c.core || typeof c.core !== 'object') return;
        if (String(c.core.schoolName || '').trim() === schoolName) return;
        c.core.schoolName = schoolName;
        api.setContainer(c);
        notifyAppLocalDataChanged('school-name');
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-tenant-settings-changed', {
                    detail: { reason: 'autosave', source: 'dashboard-school-profile-setup' }
                })
            );
        } catch {
            /* ignore */
        }
    } catch {
        /* ignore */
    }
}

/**
 * @param {HTMLElement} root
 */
function fillForm(root) {
    const p = readSchoolProfile();
    const nameEl = root.querySelector('[data-dsp-school-name]');
    const streetEl = root.querySelector('[data-dsp-street]');
    const plzEl = root.querySelector('[data-dsp-postal]');
    const cityEl = root.querySelector('[data-dsp-city]');
    const phoneEl = root.querySelector('[data-dsp-phone]');
    const urlEl = root.querySelector('[data-dsp-logo-url]');
    if (nameEl) nameEl.value = readSchoolName();
    if (streetEl) streetEl.value = p.street;
    if (plzEl) plzEl.value = p.postalCode;
    if (cityEl) cityEl.value = p.city;
    if (phoneEl) phoneEl.value = p.phone;
    if (urlEl) urlEl.value = p.logoUrl || '';
    updatePreview(root, p);
}

/**
 * @param {HTMLElement} root
 * @param {ReturnType<typeof emptySchoolProfile>} [profile]
 */
function updatePreview(root, profile) {
    const p = profile || readSchoolProfile();
    const img = root.querySelector('[data-dsp-logo-preview]');
    const placeholder = root.querySelector('[data-dsp-logo-placeholder]');
    const meta = root.querySelector('[data-dsp-logo-meta]');
    const src = resolveSchoolLogoSrc(p);
    if (img) {
        if (src) {
            img.src = src;
            img.hidden = false;
        } else {
            img.removeAttribute('src');
            img.hidden = true;
        }
    }
    if (placeholder) placeholder.hidden = !!src;
    if (meta) {
        if (p.logoUrl) meta.textContent = 'HTTPS-URL';
        else if (p.logoFileName) meta.textContent = p.logoFileName;
        else if (p.logoDataUrl) meta.textContent = 'Hochgeladenes Logo';
        else meta.textContent = '';
    }
}

/**
 * @param {HTMLElement} root
 */
function refreshStatus(root) {
    const p = readSchoolProfile();
    const el = root.querySelector('[data-dsp-status]');
    if (!el) return;
    el.textContent = hasSchoolLogo(p)
        ? 'Logo gesetzt'
        : p.street || p.city
          ? 'Adresse gesetzt, Logo fehlt noch'
          : 'Noch nicht hinterlegt';
}

function bindForm(root) {
    if (root.dataset.dspBound === '1') return;
    root.dataset.dspBound = '1';

    const debounce = (fn, ms) => {
        let t = null;
        return () => {
            if (t) clearTimeout(t);
            t = setTimeout(fn, ms);
        };
    };

    const nameEl = root.querySelector('[data-dsp-school-name]');
    const streetEl = root.querySelector('[data-dsp-street]');
    const plzEl = root.querySelector('[data-dsp-postal]');
    const cityEl = root.querySelector('[data-dsp-city]');
    const phoneEl = root.querySelector('[data-dsp-phone]');
    const urlEl = root.querySelector('[data-dsp-logo-url]');
    const fileEl = root.querySelector('[data-dsp-logo-file]');
    const removeBtn = root.querySelector('[data-dsp-logo-remove]');

    const persistAddress = debounce(() => {
        patchSchoolProfile({
            street: streetEl ? String(streetEl.value || '').trim() : '',
            postalCode: plzEl ? String(plzEl.value || '').trim() : '',
            city: cityEl ? String(cityEl.value || '').trim() : '',
            phone: phoneEl ? String(phoneEl.value || '').trim() : ''
        });
        refreshStatus(root);
    }, 400);

    [streetEl, plzEl, cityEl, phoneEl].forEach((el) => {
        if (!el) return;
        el.addEventListener('input', persistAddress);
        el.addEventListener('change', persistAddress);
    });

    if (nameEl) {
        const persistName = debounce(() => patchSchoolName(nameEl.value), 400);
        nameEl.addEventListener('input', persistName);
        nameEl.addEventListener('change', persistName);
    }

    if (urlEl) {
        const applyUrl = () => {
            const raw = String(urlEl.value || '').trim();
            if (!raw) {
                patchSchoolProfile({ logoUrl: '' });
                updatePreview(root);
                return;
            }
            const normalized = normalizeSchoolLogoUrl(raw);
            if (!normalized) {
                toast('Bitte eine öffentliche HTTPS-URL angeben (z. B. https://schule.at/logo.png).');
                urlEl.value = readSchoolProfile().logoUrl || '';
                return;
            }
            urlEl.value = normalized;
            // URL ersetzt eingebettetes Upload – schlanker für SharePoint-Backup
            patchSchoolProfile({ logoUrl: normalized, logoDataUrl: '', logoFileName: '' });
            updatePreview(root);
            refreshStatus(root);
        };
        urlEl.addEventListener('change', applyUrl);
        urlEl.addEventListener('blur', applyUrl);
    }

    if (fileEl) {
        fileEl.addEventListener('change', () => {
            const file = fileEl.files && fileEl.files[0];
            fileEl.value = '';
            if (!file) return;
            if (!/^image\//i.test(file.type || '')) {
                toast('Bitte eine Bilddatei wählen (PNG, JPEG, GIF oder WebP).');
                return;
            }
            if (file.size > SCHOOL_LOGO_MAX_BYTES) {
                toast('Das Logo ist zu groß (max. ca. 400 KB). Besser: öffentliche HTTPS-URL nutzen.');
                return;
            }
            const reader = new FileReader();
            reader.onload = () => {
                patchSchoolProfile({
                    logoDataUrl: String(reader.result || ''),
                    logoFileName: file.name || 'logo',
                    logoUrl: ''
                });
                if (urlEl) urlEl.value = '';
                updatePreview(root);
                refreshStatus(root);
            };
            reader.onerror = () => toast('Logo konnte nicht gelesen werden.');
            reader.readAsDataURL(file);
        });
    }

    if (removeBtn) {
        removeBtn.addEventListener('click', () => {
            patchSchoolProfile({ logoDataUrl: '', logoFileName: '', logoUrl: '' });
            if (urlEl) urlEl.value = '';
            updatePreview(root, emptySchoolProfile());
            refreshStatus(root);
        });
    }
}

/**
 * @param {string} [mountSelector]
 */
export function mountDashboardSchoolProfileSetup(mountSelector) {
    const sel = mountSelector || '[data-ms365-dashboard-school-profile-mount]';
    const root = document.querySelector(sel);
    if (!root) return;

    const profile = readSchoolProfile();
    const status = hasSchoolLogo(profile)
        ? 'Logo gesetzt'
        : profile.street || profile.city
          ? 'Adresse gesetzt, Logo fehlt noch'
          : 'Noch nicht hinterlegt';

    root.innerHTML = `
        <section class="dash-pg-section dash-school-setup" aria-labelledby="dashSchoolProfileSetupTitle">
            <div class="dash-pg-section__head">
                <div>
                    <h3 id="dashSchoolProfileSetupTitle">Schule im Portal</h3>
                    <p class="dash-pg-section__lead">
                        Name, Logo und Adresse für die Karte „Meine Schule“ im Lehrer-/Schüler-Portal.
                        Speicherung im Schulprofil (Stammdaten) – landet mit der IT-Bibliothek in SharePoint.
                    </p>
                </div>
                <span class="dash-school-setup__status" data-dsp-status>${escapeHtml(status)}</span>
            </div>
            <div class="dash-school-setup__grid">
                <div class="dash-school-setup__logo">
                    <div class="dash-school-setup__logo-frame" data-dsp-logo-placeholder ${hasSchoolLogo(profile) ? 'hidden' : ''}>
                        <i class="bi bi-building" aria-hidden="true"></i>
                        <span>Kein Logo</span>
                    </div>
                    <img class="dash-school-setup__logo-img" data-dsp-logo-preview alt="Schul-Logo Vorschau" width="72" height="72" referrerpolicy="no-referrer" hidden>
                    <p class="dash-school-setup__logo-meta muted" data-dsp-logo-meta></p>
                    <div class="dash-school-setup__logo-actions">
                        <label class="btn btn-sm" for="dashSchoolLogoFile"><i class="bi bi-upload" aria-hidden="true"></i> Hochladen</label>
                        <input type="file" id="dashSchoolLogoFile" data-dsp-logo-file accept="image/png,image/jpeg,image/gif,image/webp" hidden>
                        <button type="button" class="btn btn-sm alt" data-dsp-logo-remove>Entfernen</button>
                    </div>
                    <label class="dash-school-setup__field" for="dashSchoolLogoUrl">
                        <span>Oder öffentliche HTTPS-URL</span>
                        <input type="url" id="dashSchoolLogoUrl" data-dsp-logo-url placeholder="https://schule.at/media/logo.png" inputmode="url" autocomplete="off" spellcheck="false">
                    </label>
                    <small class="muted">Empfohlen: Logo auf der Schulwebsite oder in SharePoint öffentlich verlinken – schlanker als Upload.</small>
                </div>
                <div class="dash-school-setup__fields">
                    <label class="dash-school-setup__field" for="dashSchoolName">
                        <span>Schulname</span>
                        <input type="text" id="dashSchoolName" data-dsp-school-name placeholder="z. B. MS Musterstadt" autocomplete="organization">
                    </label>
                    <label class="dash-school-setup__field" for="dashSchoolStreet">
                        <span>Straße</span>
                        <input type="text" id="dashSchoolStreet" data-dsp-street placeholder="Musterstraße 1" autocomplete="street-address">
                    </label>
                    <div class="dash-school-setup__row">
                        <label class="dash-school-setup__field" for="dashSchoolPostal">
                            <span>PLZ</span>
                            <input type="text" id="dashSchoolPostal" data-dsp-postal maxlength="16" autocomplete="postal-code" inputmode="numeric">
                        </label>
                        <label class="dash-school-setup__field dash-school-setup__field--grow" for="dashSchoolCity">
                            <span>Ort</span>
                            <input type="text" id="dashSchoolCity" data-dsp-city autocomplete="address-level2">
                        </label>
                    </div>
                    <label class="dash-school-setup__field" for="dashSchoolPhone">
                        <span>Telefon</span>
                        <input type="tel" id="dashSchoolPhone" data-dsp-phone placeholder="+43 …" autocomplete="tel">
                    </label>
                    <p class="dash-pg-section__note">
                        Weitere Felder (Website, SKZ, Domain): <a href="tenant.html#stammdaten">Stammdaten → Schulprofil</a>.
                        Danach hier oder dort „In SharePoint sichern“.
                    </p>
                </div>
            </div>
        </section>`;

    fillForm(/** @type {HTMLElement} */ (root));
    bindForm(/** @type {HTMLElement} */ (root));

    window.addEventListener('ms365-tenant-settings-changed', () => {
        fillForm(/** @type {HTMLElement} */ (root));
        refreshStatus(/** @type {HTMLElement} */ (root));
    });
    window.addEventListener('storage', (ev) => {
        if (ev && ev.key === 'ms365-schooltool-data-v2') {
            fillForm(/** @type {HTMLElement} */ (root));
            refreshStatus(/** @type {HTMLElement} */ (root));
        }
    });
}
