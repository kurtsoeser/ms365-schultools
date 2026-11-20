/**
 * Microsoft-365-Personensuche für Flow-/Mail-Konten in IT-Setup-Assistenten.
 * Einheitlich für Lehrer-Freistellung, Schulaktivitäten, Freistellung Schüler, …
 */
import { pickEntraUser } from './entra-user-picker.js';

/**
 * @typedef {{
 *   inputId: string,
 *   btnId?: string,
 *   title: string,
 *   hint?: string,
 *   onPicked?: (mail: string, user: { mail?: string, id?: string, displayName?: string }) => void
 * }} FlowEmailPickerSpec
 */

/**
 * @param {FlowEmailPickerSpec[]} specs
 * @param {{ wiredFlag?: string }} [opts]
 */
export function wireFlowSetupEmailPickers(specs, opts) {
    const flag = (opts && opts.wiredFlag) || 'flowEmailPickerWired';
    (Array.isArray(specs) ? specs : []).forEach((spec) => {
        if (!spec || !spec.inputId) return;
        const input = document.getElementById(spec.inputId);
        const btn = document.getElementById(spec.btnId || spec.inputId + 'Pick');
        if (!input || !btn || btn.dataset[flag] === '1') return;
        btn.dataset[flag] = '1';
        btn.addEventListener('click', () => {
            pickEntraUser({ title: spec.title, hint: spec.hint })
                .then((user) => {
                    if (!user) return;
                    const mail = String(user.mail || user.userPrincipalName || '')
                        .trim()
                        .toLowerCase();
                    if (!mail.includes('@')) {
                        toast('Keine E-Mail-Adresse für dieses Konto gefunden.');
                        return;
                    }
                    input.value = mail;
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    if (typeof spec.onPicked === 'function') spec.onPicked(mail, user);
                })
                .catch((e) => {
                    toast(e && e.message ? e.message : String(e));
                });
        });
    });
}

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else if (typeof window.alert === 'function') window.alert(msg);
}
