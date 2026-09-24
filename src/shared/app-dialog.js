/**
 * Zentrale App-Dialoge + Toasts (ersetzt native alert / confirm / prompt).
 * Lädt app.css (modal-overlay / modal-box / ms365-toast*) mit.
 * @file
 */
(function () {
    'use strict';

    let root = null;
    let titleEl;
    let msgEl;
    let iconEl;
    let promptWrap;
    let inputLabelTextEl;
    let inputEl;
    let okBtn;
    let cancelBtn;
    /** @type {'alert'|'confirm'|'prompt'} */
    let mode = 'alert';
    let resolver = null;

    /** @type {HTMLElement|null} */
    let toastHost = null;
    let toastSeq = 0;

    const KIND_META = {
        info: {
            title: 'Hinweis',
            icon: 'info',
            toastMs: 4200
        },
        success: {
            title: 'Erfolg',
            icon: 'success',
            toastMs: 3800
        },
        error: {
            title: 'Fehler',
            icon: 'error',
            toastMs: 6500
        },
        warning: {
            title: 'Hinweis',
            icon: 'warning',
            toastMs: 5200
        }
    };

    const ICON_SVG = {
        info:
            '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
            '<path fill="currentColor" d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/></svg>',
        success:
            '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
            '<path fill="currentColor" d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm-1.2 14.2-3.5-3.5 1.4-1.4 2.1 2.1 4.5-4.5 1.4 1.4-5.9 5.9z"/></svg>',
        error:
            '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
            '<path fill="currentColor" d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>',
        warning:
            '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
            '<path fill="currentColor" d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg>'
    };

    /**
     * @param {unknown} message
     * @param {string} [explicit]
     * @returns {'info'|'success'|'error'|'warning'}
     */
    function resolveKind(message, explicit) {
        const allowed = ['info', 'success', 'error', 'warning'];
        if (explicit && allowed.indexOf(explicit) >= 0) return explicit;
        const t = String(message ?? '');
        if (/fehler|fehlgeschlagen|error|exception|ungültig|nicht möglich|abgebrochen/i.test(t)) {
            return 'error';
        }
        if (/warn|achtung|bitte.*(prüfen|kontrollieren)|unvollständig/i.test(t)) {
            return 'warning';
        }
        if (
            /\bok\b|erfolg|erfolgreich|gespeichert|angelegt|gelöscht|fertig|abgeschlossen|synchronisiert/i.test(
                t
            )
        ) {
            return 'success';
        }
        return 'info';
    }

    function cacheRefs() {
        titleEl = root.querySelector('[data-app-dialog-title]');
        msgEl = root.querySelector('[data-app-dialog-message]');
        iconEl = root.querySelector('[data-app-dialog-icon]');
        promptWrap = root.querySelector('[data-app-dialog-prompt]');
        inputLabelTextEl = root.querySelector('[data-app-dialog-input-label-text]');
        inputEl = root.querySelector('[data-app-dialog-input]');
        okBtn = root.querySelector('[data-app-dialog-ok]');
        cancelBtn = root.querySelector('[data-app-dialog-cancel]');
    }

    function setDialogKind(kind) {
        const k = KIND_META[kind] ? kind : 'info';
        root.setAttribute('data-kind', k);
        if (iconEl) {
            iconEl.className = 'app-dialog-icon app-dialog-icon--' + k;
            iconEl.innerHTML = ICON_SVG[KIND_META[k].icon] || ICON_SVG.info;
        }
    }

    function teardown(result) {
        if (!resolver) return;
        document.removeEventListener('keydown', onDocKey, true);
        if (root) root.classList.remove('open');
        const fn = resolver;
        resolver = null;
        fn(result);
    }

    function onDocKey(ev) {
        if (!root || !root.classList.contains('open')) return;
        if (ev.key === 'Escape') {
            ev.preventDefault();
            ev.stopPropagation();
            if (mode === 'confirm') teardown(false);
            else if (mode === 'prompt') teardown(null);
            else teardown();
        } else if (ev.key === 'Enter' && mode === 'prompt' && !ev.shiftKey) {
            ev.preventDefault();
            teardown(inputEl.value);
        } else if (ev.key === 'Enter' && mode === 'confirm' && !ev.shiftKey) {
            ev.preventDefault();
            teardown(true);
        }
    }

    function ensure() {
        if (root) return root;
        root = document.createElement('div');
        root.id = 'ms365AppDialog';
        root.className = 'modal-overlay ms365-app-dialog-overlay';
        root.setAttribute('role', 'dialog');
        root.setAttribute('aria-modal', 'true');
        root.setAttribute('data-kind', 'info');
        root.innerHTML =
            '<div class="modal-box ms365-app-dialog-box" tabindex="-1">' +
            '<div class="app-dialog-head">' +
            '<div class="app-dialog-icon app-dialog-icon--info" data-app-dialog-icon aria-hidden="true"></div>' +
            '<div class="app-dialog-head-text">' +
            '<h3 class="app-dialog-title" data-app-dialog-title></h3>' +
            '<p class="app-dialog-message" data-app-dialog-message></p>' +
            '</div></div>' +
            '<div class="app-dialog-prompt" data-app-dialog-prompt style="display:none">' +
            '<label class="app-dialog-input-label">' +
            '<span data-app-dialog-input-label-text></span>' +
            '<input type="text" class="app-dialog-input" data-app-dialog-input autocomplete="off" spellcheck="false" />' +
            '</label></div>' +
            '<div class="modal-actions">' +
            '<button type="button" class="btn" data-app-dialog-cancel>Abbrechen</button>' +
            '<button type="button" class="btn btn-success" data-app-dialog-ok>OK</button>' +
            '</div></div>';
        document.body.appendChild(root);
        cacheRefs();
        setDialogKind('info');
        root.addEventListener('click', function (e) {
            if (e.target !== root) return;
            if (mode === 'confirm') teardown(false);
            else if (mode === 'prompt') teardown(null);
            else teardown();
        });
        cancelBtn.addEventListener('click', function () {
            if (mode === 'confirm') teardown(false);
            else if (mode === 'prompt') teardown(null);
            else teardown();
        });
        okBtn.addEventListener('click', function () {
            if (mode === 'prompt') teardown(inputEl.value);
            else if (mode === 'confirm') teardown(true);
            else teardown();
        });
        return root;
    }

    function ensureToastHost() {
        if (toastHost && document.body.contains(toastHost)) return toastHost;
        toastHost = document.createElement('div');
        toastHost.id = 'ms365ToastHost';
        toastHost.className = 'ms365-toast-host';
        toastHost.setAttribute('aria-live', 'polite');
        toastHost.setAttribute('aria-relevant', 'additions text');
        document.body.appendChild(toastHost);
        return toastHost;
    }

    /**
     * @param {string} message
     * @param {{ title?: string, okText?: string, kind?: string }} [options]
     * @returns {Promise<void>}
     */
    function ms365AppDialogAlert(message, options) {
        ensure();
        mode = 'alert';
        const kind = resolveKind(message, options && options.kind);
        setDialogKind(kind);
        titleEl.textContent = (options && options.title) || KIND_META[kind].title;
        msgEl.textContent = String(message ?? '');
        promptWrap.style.display = 'none';
        cancelBtn.style.display = 'none';
        okBtn.textContent = (options && options.okText) || 'OK';
        okBtn.className = kind === 'error' ? 'btn btn-danger' : kind === 'warning' ? 'btn' : 'btn btn-success';
        return new Promise(function (resolve) {
            resolver = function () {
                resolve();
            };
            root.classList.add('open');
            document.addEventListener('keydown', onDocKey, true);
            requestAnimationFrame(function () {
                okBtn.focus();
            });
        });
    }

    /**
     * @param {string} message
     * @param {{ title?: string, okText?: string, cancelText?: string, danger?: boolean, kind?: string }} [options]
     * @returns {Promise<boolean>}
     */
    function ms365AppDialogConfirm(message, options) {
        ensure();
        mode = 'confirm';
        const danger = !!(options && options.danger);
        const kind = resolveKind(message, (options && options.kind) || (danger ? 'warning' : 'info'));
        setDialogKind(kind);
        titleEl.textContent = (options && options.title) || (danger ? 'Bitte bestätigen' : 'Bestätigung');
        msgEl.textContent = String(message ?? '');
        promptWrap.style.display = 'none';
        cancelBtn.style.display = '';
        cancelBtn.textContent = (options && options.cancelText) || 'Abbrechen';
        okBtn.textContent = (options && options.okText) || 'OK';
        okBtn.className = danger || kind === 'error' ? 'btn btn-danger' : 'btn btn-success';
        return new Promise(function (resolve) {
            resolver = function (v) {
                resolve(!!v);
            };
            root.classList.add('open');
            document.addEventListener('keydown', onDocKey, true);
            requestAnimationFrame(function () {
                if (danger) cancelBtn.focus();
                else okBtn.focus();
            });
        });
    }

    /**
     * @param {string} message
     * @param {string} [defaultValue]
     * @param {{ title?: string, inputLabel?: string, okText?: string, cancelText?: string, kind?: string }} [options]
     * @returns {Promise<string|null>} null bei Abbrechen
     */
    function ms365AppDialogPrompt(message, defaultValue, options) {
        ensure();
        mode = 'prompt';
        const kind = resolveKind(message, (options && options.kind) || 'info');
        setDialogKind(kind);
        titleEl.textContent = (options && options.title) || 'Eingabe';
        msgEl.textContent = String(message ?? '');
        if (inputLabelTextEl) {
            inputLabelTextEl.textContent = (options && options.inputLabel) || 'Eingabe';
        }
        inputEl.value = defaultValue != null ? String(defaultValue) : '';
        promptWrap.style.display = 'block';
        cancelBtn.style.display = '';
        cancelBtn.textContent = (options && options.cancelText) || 'Abbrechen';
        okBtn.textContent = (options && options.okText) || 'OK';
        okBtn.className = 'btn btn-success';
        return new Promise(function (resolve) {
            resolver = function (v) {
                if (v === null || v === undefined) resolve(null);
                else resolve(String(v));
            };
            root.classList.add('open');
            document.addEventListener('keydown', onDocKey, true);
            requestAnimationFrame(function () {
                inputEl.focus();
                try {
                    inputEl.select();
                } catch {
                    // ignore
                }
            });
        });
    }

    /**
     * Zentraler Toast – kein Markup in der Seite nötig.
     * @param {string} message
     * @param {{ kind?: string, title?: string, durationMs?: number, dismissible?: boolean }} [opts]
     */
    function ms365ShowToast(message, opts) {
        opts = opts || {};
        const host = ensureToastHost();
        const kind = resolveKind(message, opts.kind);
        const meta = KIND_META[kind];
        const durationMs =
            typeof opts.durationMs === 'number' && opts.durationMs >= 0
                ? opts.durationMs
                : meta.toastMs;
        const dismissible = opts.dismissible !== false;
        const id = 'ms365-toast-' + ++toastSeq;

        const el = document.createElement('div');
        el.id = id;
        el.className = 'ms365-toast ms365-toast--' + kind;
        el.setAttribute('role', kind === 'error' ? 'alert' : 'status');

        const title = opts.title != null ? String(opts.title) : meta.title;
        const text = String(message ?? '');

        el.innerHTML =
            '<div class="ms365-toast-icon" aria-hidden="true">' +
            (ICON_SVG[meta.icon] || ICON_SVG.info) +
            '</div>' +
            '<div class="ms365-toast-body">' +
            '<div class="ms365-toast-title"></div>' +
            '<div class="ms365-toast-msg"></div>' +
            '</div>' +
            (dismissible
                ? '<button type="button" class="ms365-toast-close" aria-label="Schließen">&times;</button>'
                : '');

        el.querySelector('.ms365-toast-title').textContent = title;
        el.querySelector('.ms365-toast-msg').textContent = text;

        let hideTimer = null;
        function removeToast() {
            if (hideTimer) {
                clearTimeout(hideTimer);
                hideTimer = null;
            }
            el.classList.remove('ms365-toast--show');
            el.classList.add('ms365-toast--hide');
            setTimeout(function () {
                if (el.parentNode) el.parentNode.removeChild(el);
            }, 220);
        }

        const closeBtn = el.querySelector('.ms365-toast-close');
        if (closeBtn) closeBtn.addEventListener('click', removeToast);

        host.appendChild(el);
        requestAnimationFrame(function () {
            el.classList.add('ms365-toast--show');
        });

        if (durationMs > 0) {
            hideTimer = setTimeout(removeToast, durationMs);
        }

        return { dismiss: removeToast, id: id };
    }

    /**
     * Kurzinfo: Toast (bevorzugt), sonst modal „Hinweis“.
     * @param {string} msg
     * @param {{ title?: string, kind?: string, durationMs?: number, forceDialog?: boolean }} [opts]
     */
    function ms365ToastOrAlert(msg, opts) {
        opts = opts || {};
        if (!opts.forceDialog && typeof ms365ShowToast === 'function') {
            ms365ShowToast(msg, opts);
            return;
        }
        void ms365AppDialogAlert(msg, Object.assign({ title: 'Hinweis' }, opts));
    }

    window.ms365AppDialogAlert = ms365AppDialogAlert;
    window.ms365AppDialogConfirm = ms365AppDialogConfirm;
    window.ms365AppDialogPrompt = ms365AppDialogPrompt;
    window.ms365ShowToast = ms365ShowToast;
    window.ms365ToastOrAlert = ms365ToastOrAlert;
})();
