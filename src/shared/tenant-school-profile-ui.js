/**
 * Schulregister Stammdaten: Logo, Adresse, Kontakt (setup.schoolProfile).
 */
import {
    emptySchoolProfile,
    normalizeSchoolProfile,
    SCHOOL_LOGO_MAX_BYTES
} from './school-profile-logic.js';
import { notifyAppLocalDataChanged } from './app-local-data-notify.js';

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else if (msg) window.alert(msg);
}

function patchSchoolProfile(partial) {
    if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
    const setup =
        typeof window.ms365AppDataV2.getSetup === 'function' ? window.ms365AppDataV2.getSetup() : {};
    const cur = normalizeSchoolProfile(setup && setup.schoolProfile);
    const next = normalizeSchoolProfile(Object.assign({}, cur, partial || {}));
    window.ms365AppDataV2.patchSetup({ schoolProfile: next });
    notifyAppLocalDataChanged('school-profile');
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-tenant-settings-changed', {
                detail: { reason: 'autosave', source: 'school-profile' }
            })
        );
    } catch {
        /* ignore */
    }
}

function readCoreSchoolIdentity() {
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            return {
                schoolName: String(s.schoolName || '').trim(),
                domain: String(s.domain || '').trim()
            };
        }
    } catch {
        /* ignore */
    }
    try {
        const api = window.ms365AppDataV2;
        if (api && typeof api.getContainer === 'function') {
            const c = api.getContainer();
            const core = c && c.core ? c.core : {};
            return {
                schoolName: String(core.schoolName || '').trim(),
                domain: String(core.domain || '').trim()
            };
        }
    } catch {
        /* ignore */
    }
    return { schoolName: '', domain: '' };
}

function readSchoolProfileFromSetup() {
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

function bindTextField(id, key, debounceMs) {
    const el = document.getElementById(id);
    if (!el || el.dataset.schoolProfileBound === '1') return;
    el.dataset.schoolProfileBound = '1';
    let timer = null;
    function persist() {
        patchSchoolProfile({ [key]: String(el.value || '').trim() });
    }
    el.addEventListener('input', function () {
        if (timer) clearTimeout(timer);
        timer = setTimeout(persist, debounceMs || 400);
    });
    el.addEventListener('change', persist);
}

function updateLogoPreview(profile) {
    const img = document.getElementById('tenantSchoolLogoPreview');
    const placeholder = document.getElementById('tenantSchoolLogoPlaceholder');
    const url = profile && profile.logoDataUrl ? String(profile.logoDataUrl) : '';
    if (img) {
        if (url) {
            img.src = url;
            img.hidden = false;
        } else {
            img.removeAttribute('src');
            img.hidden = true;
        }
    }
    if (placeholder) placeholder.hidden = !!url;
    const nameEl = document.getElementById('tenantSchoolLogoFileName');
    if (nameEl) {
        nameEl.textContent = profile && profile.logoFileName ? profile.logoFileName : '';
    }
}

function loadSchoolProfileIntoForm() {
    const identity = readCoreSchoolIdentity();
    const nameEl = document.getElementById('schoolName');
    const domainEl = document.getElementById('schoolEmailDomain');
    if (nameEl) nameEl.value = identity.schoolName;
    if (domainEl) domainEl.value = identity.domain;
    const p = readSchoolProfileFromSetup();
    const map = {
        tenantSchoolCode: 'schoolCode',
        tenantSchoolStreet: 'street',
        tenantSchoolPostalCode: 'postalCode',
        tenantSchoolCity: 'city',
        tenantSchoolCountry: 'country',
        tenantSchoolPhone: 'phone',
        tenantSchoolPhoneAlt: 'phoneAlt',
        tenantSchoolOfficeEmail: 'email',
        tenantSchoolWebsite: 'website'
    };
    Object.keys(map).forEach(function (id) {
        const el = document.getElementById(id);
        if (!el) return;
        const k = map[id];
        el.value = p[k] != null ? String(p[k]) : '';
    });
    updateLogoPreview(p);
}

function bindLogoUpload() {
    const fileInput = document.getElementById('tenantSchoolLogoFile');
    const btnRemove = document.getElementById('tenantSchoolLogoRemove');
    if (!fileInput || fileInput.dataset.schoolProfileBound === '1') return;
    fileInput.dataset.schoolProfileBound = '1';

    fileInput.addEventListener('change', function () {
        const file = fileInput.files && fileInput.files[0];
        fileInput.value = '';
        if (!file) return;
        if (!/^image\//i.test(file.type || '')) {
            toast('Bitte eine Bilddatei wählen (PNG, JPEG, GIF oder WebP).');
            return;
        }
        if (file.size > SCHOOL_LOGO_MAX_BYTES) {
            toast('Das Logo ist zu groß (max. ca. 400 KB). Bitte verkleinern oder komprimieren.');
            return;
        }
        const reader = new FileReader();
        reader.onload = function () {
            const dataUrl = String(reader.result || '');
            patchSchoolProfile({
                logoDataUrl: dataUrl,
                logoFileName: file.name || 'logo'
            });
            updateLogoPreview(readSchoolProfileFromSetup());
        };
        reader.onerror = function () {
            toast('Logo konnte nicht gelesen werden.');
        };
        reader.readAsDataURL(file);
    });

    if (btnRemove && btnRemove.dataset.schoolProfileBound !== '1') {
        btnRemove.dataset.schoolProfileBound = '1';
        btnRemove.addEventListener('click', function () {
            patchSchoolProfile({ logoDataUrl: '', logoFileName: '' });
            updateLogoPreview(emptySchoolProfile());
        });
    }
}

export function mountTenantSchoolProfileUi() {
    if (!document.getElementById('tenantSchoolProfileCard')) return;
    loadSchoolProfileIntoForm();
    bindLogoUpload();
    bindTextField('tenantSchoolCode', 'schoolCode');
    bindTextField('tenantSchoolStreet', 'street');
    bindTextField('tenantSchoolPostalCode', 'postalCode');
    bindTextField('tenantSchoolCity', 'city');
    bindTextField('tenantSchoolCountry', 'country');
    bindTextField('tenantSchoolPhone', 'phone');
    bindTextField('tenantSchoolPhoneAlt', 'phoneAlt');
    bindTextField('tenantSchoolOfficeEmail', 'email');
    bindTextField('tenantSchoolWebsite', 'website');
}

export function readSchoolProfileForApp() {
    return readSchoolProfileFromSetup();
}

export function reloadTenantSchoolProfileForm() {
    if (!document.getElementById('tenantSchoolProfileCard')) return;
    loadSchoolProfileIntoForm();
}
