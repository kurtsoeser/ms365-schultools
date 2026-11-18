/**
 * Globales Feedback: Footer-Link + Dialog → License-API → Betreiber-SharePoint.
 */
(function (global) {
    'use strict';

    var MAX_FILES = 3;
    var MAX_BYTES = 1.5 * 1024 * 1024;
    var pendingFiles = [];

    function $(id) {
        return document.getElementById(id);
    }

    function escapeHtml(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function ensureStyles() {
        if (document.getElementById('ms365FeedbackStyles')) return;
        var style = document.createElement('style');
        style.id = 'ms365FeedbackStyles';
        style.textContent =
            '#ms365FeedbackDialog{border:none;padding:0;border-radius:16px;max-width:560px;width:calc(100vw - 2rem);' +
            'max-height:min(92vh,820px);overflow:auto;margin:0;position:fixed;top:50%;left:50%;' +
            'transform:translate(-50%,-50%);' +
            'box-shadow:0 24px 64px rgba(28,20,60,.34);background:var(--surface,#fff);color:inherit;}' +
            '#ms365FeedbackDialog::backdrop{background:rgba(20,16,40,.52);backdrop-filter:blur(2px);}' +
            '.ms365-fb__wrap{padding:1.15rem 1.25rem 1.25rem;}' +
            '.ms365-fb__head{display:flex;gap:.75rem;align-items:flex-start;margin-bottom:1rem;}' +
            '.ms365-fb__icon{width:2.4rem;height:2.4rem;border-radius:10px;display:grid;place-items:center;' +
            'background:color-mix(in srgb, var(--accent,#5b4bdb) 16%, transparent);color:var(--accent,#5b4bdb);flex:0 0 auto;}' +
            '.ms365-fb__head h2{margin:0;font-size:1.15rem;}' +
            '.ms365-fb__head p{margin:.2rem 0 0;opacity:.8;font-size:.9rem;}' +
            '.ms365-fb__field{margin:0 0 .85rem;}' +
            '.ms365-fb__field label{display:block;font-size:.85rem;font-weight:600;margin-bottom:.3rem;}' +
            '.ms365-fb__field input[type=text],.ms365-fb__field textarea,.ms365-fb__field select{' +
            'width:100%;box-sizing:border-box;border:1px solid color-mix(in srgb, currentColor 18%, transparent);' +
            'border-radius:8px;padding:.55rem .65rem;background:var(--surface-2,#f7f6fb);color:inherit;font:inherit;}' +
            '.ms365-fb__field textarea{min-height:7.5rem;resize:vertical;}' +
            '.ms365-fb__kinds{display:flex;flex-wrap:wrap;gap:.45rem;}' +
            '.ms365-fb__kind{display:inline-flex;align-items:center;gap:.35rem;padding:.35rem .65rem;border-radius:999px;' +
            'border:1px solid color-mix(in srgb, currentColor 16%, transparent);cursor:pointer;font-size:.85rem;}' +
            '.ms365-fb__kind:has(input:checked){border-color:var(--accent,#5b4bdb);background:color-mix(in srgb, var(--accent,#5b4bdb) 12%, transparent);}' +
            '.ms365-fb__kind input{margin:0;}' +
            '.ms365-fb__drop{border:1px dashed color-mix(in srgb, currentColor 28%, transparent);border-radius:10px;' +
            'padding:.85rem;text-align:center;font-size:.88rem;opacity:.95;}' +
            '.ms365-fb__drop.is-drag{border-color:var(--accent,#5b4bdb);background:color-mix(in srgb, var(--accent,#5b4bdb) 8%, transparent);}' +
            '.ms365-fb__thumbs{display:flex;flex-wrap:wrap;gap:.5rem;margin-top:.6rem;}' +
            '.ms365-fb__thumb{position:relative;width:72px;height:72px;border-radius:8px;overflow:hidden;border:1px solid color-mix(in srgb, currentColor 14%, transparent);}' +
            '.ms365-fb__thumb img{width:100%;height:100%;object-fit:cover;display:block;}' +
            '.ms365-fb__thumb button{position:absolute;top:2px;right:2px;border:0;border-radius:999px;width:1.35rem;height:1.35rem;' +
            'background:rgba(0,0,0,.65);color:#fff;cursor:pointer;font-size:.75rem;line-height:1;}' +
            '.ms365-fb__meta{font-size:.8rem;opacity:.75;margin:.25rem 0 .9rem;}' +
            '.ms365-fb__actions{display:flex;flex-wrap:wrap;gap:.5rem;justify-content:flex-end;margin-top:1rem;}' +
            '.ms365-fb__status{min-height:1.2rem;font-size:.85rem;margin:.35rem 0 0;}' +
            '.ms365-fb__status.is-err{color:#b42318;}' +
            '.ms365-fb__status.is-ok{color:#067647;}' +
            '.site-feedback-link{cursor:pointer;}';
        document.head.appendChild(style);
    }

    function ensureDialog() {
        var dlg = $('ms365FeedbackDialog');
        if (dlg) return dlg;
        ensureStyles();
        dlg = document.createElement('dialog');
        dlg.id = 'ms365FeedbackDialog';
        dlg.setAttribute('aria-labelledby', 'ms365FeedbackTitle');
        dlg.innerHTML =
            '<form class="ms365-fb__wrap" id="ms365FeedbackForm" method="dialog">' +
            '<div class="ms365-fb__head">' +
            '<div class="ms365-fb__icon" aria-hidden="true"><i class="bi bi-chat-heart"></i></div>' +
            '<div><h2 id="ms365FeedbackTitle">Feedback senden</h2>' +
            '<p>Bug, Idee oder Hinweis – landet direkt beim Betreiber.</p></div></div>' +
            '<div class="ms365-fb__field"><label for="ms365FeedbackTitleInput">Titel</label>' +
            '<input type="text" id="ms365FeedbackTitleInput" maxlength="200" required placeholder="Kurz: Was ist das Thema?"></div>' +
            '<div class="ms365-fb__field"><span id="ms365FeedbackKindLabel">Art</span>' +
            '<div class="ms365-fb__kinds" role="radiogroup" aria-labelledby="ms365FeedbackKindLabel">' +
            '<label class="ms365-fb__kind"><input type="radio" name="ms365FbKind" value="Feedback" checked> Feedback</label>' +
            '<label class="ms365-fb__kind"><input type="radio" name="ms365FbKind" value="Bug"> Bug</label>' +
            '<label class="ms365-fb__kind"><input type="radio" name="ms365FbKind" value="Idee"> Idee</label>' +
            '<label class="ms365-fb__kind"><input type="radio" name="ms365FbKind" value="Frage"> Frage</label>' +
            '</div></div>' +
            '<div class="ms365-fb__field"><label for="ms365FeedbackBody">Beschreibung</label>' +
            '<textarea id="ms365FeedbackBody" maxlength="8000" required placeholder="Was ist passiert? Erwartetes Verhalten? Schritte zum Nachstellen?"></textarea></div>' +
            '<div class="ms365-fb__field"><span>Screenshots (optional, max. 3)</span>' +
            '<div class="ms365-fb__drop" id="ms365FeedbackDrop" tabindex="0">' +
            '<i class="bi bi-image" aria-hidden="true"></i> Bild hierher ziehen, einfügen oder ' +
            '<button type="button" class="btn btn-sm alt" id="ms365FeedbackPickBtn">wählen</button>' +
            '<input type="file" id="ms365FeedbackFile" accept="image/*" multiple hidden></div>' +
            '<div class="ms365-fb__thumbs" id="ms365FeedbackThumbs"></div></div>' +
            '<p class="ms365-fb__meta" id="ms365FeedbackMeta"></p>' +
            '<p class="ms365-fb__status" id="ms365FeedbackStatus" role="status"></p>' +
            '<div class="ms365-fb__actions">' +
            '<button type="button" class="btn alt" id="ms365FeedbackCancel">Abbrechen</button>' +
            '<button type="submit" class="btn" id="ms365FeedbackSend"><i class="bi bi-send"></i> Absenden</button>' +
            '</div></form>';
        document.body.appendChild(dlg);

        var form = $('ms365FeedbackForm');
        form.addEventListener('submit', onSubmit);
        $('ms365FeedbackCancel').addEventListener('click', function () {
            dlg.close();
        });
        $('ms365FeedbackPickBtn').addEventListener('click', function () {
            $('ms365FeedbackFile').click();
        });
        $('ms365FeedbackFile').addEventListener('change', function (ev) {
            addFiles(ev.target.files);
            ev.target.value = '';
        });
        var drop = $('ms365FeedbackDrop');
        drop.addEventListener('dragover', function (ev) {
            ev.preventDefault();
            drop.classList.add('is-drag');
        });
        drop.addEventListener('dragleave', function () {
            drop.classList.remove('is-drag');
        });
        drop.addEventListener('drop', function (ev) {
            ev.preventDefault();
            drop.classList.remove('is-drag');
            addFiles(ev.dataTransfer && ev.dataTransfer.files);
        });
        dlg.addEventListener('paste', function (ev) {
            var items = ev.clipboardData && ev.clipboardData.items;
            if (!items) return;
            var files = [];
            for (var i = 0; i < items.length; i++) {
                if (items[i].type && items[i].type.indexOf('image/') === 0) {
                    var f = items[i].getAsFile();
                    if (f) files.push(f);
                }
            }
            if (files.length) {
                ev.preventDefault();
                addFiles(files);
            }
        });
        return dlg;
    }

    function setStatus(msg, kind) {
        var el = $('ms365FeedbackStatus');
        if (!el) return;
        el.textContent = msg || '';
        el.classList.toggle('is-err', kind === 'err');
        el.classList.toggle('is-ok', kind === 'ok');
    }

    function renderThumbs() {
        var box = $('ms365FeedbackThumbs');
        if (!box) return;
        box.innerHTML = pendingFiles
            .map(function (f, idx) {
                return (
                    '<div class="ms365-fb__thumb">' +
                    '<img src="' +
                    escapeHtml(f.preview) +
                    '" alt="">' +
                    '<button type="button" data-fb-rm="' +
                    idx +
                    '" aria-label="Entfernen">×</button></div>'
                );
            })
            .join('');
        box.querySelectorAll('[data-fb-rm]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                var i = Number(btn.getAttribute('data-fb-rm'));
                if (pendingFiles[i] && pendingFiles[i].preview) {
                    try {
                        URL.revokeObjectURL(pendingFiles[i].preview);
                    } catch (e) {
                        /* ignore */
                    }
                }
                pendingFiles.splice(i, 1);
                renderThumbs();
            });
        });
    }

    function addFiles(fileList) {
        if (!fileList || !fileList.length) return;
        for (var i = 0; i < fileList.length; i++) {
            var file = fileList[i];
            if (!file || !/^image\//i.test(file.type || '')) {
                setStatus('Nur Bilder sind als Screenshot erlaubt.', 'err');
                continue;
            }
            if (file.size > MAX_BYTES) {
                setStatus('„' + file.name + '“ ist größer als 1,5 MB.', 'err');
                continue;
            }
            if (pendingFiles.length >= MAX_FILES) {
                setStatus('Maximal ' + MAX_FILES + ' Screenshots.', 'err');
                break;
            }
            pendingFiles.push({
                file: file,
                name: file.name || 'screenshot.png',
                contentType: file.type || 'image/png',
                preview: URL.createObjectURL(file)
            });
        }
        renderThumbs();
    }

    function fileToBase64(file) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function () {
                var s = String(reader.result || '');
                var comma = s.indexOf(',');
                resolve(comma >= 0 ? s.slice(comma + 1) : s);
            };
            reader.onerror = function () {
                reject(new Error('Datei konnte nicht gelesen werden.'));
            };
            reader.readAsDataURL(file);
        });
    }

    function currentUserHint() {
        try {
            if (global.ms365AuthGetAccount && typeof global.ms365AuthGetAccount === 'function') {
                var acc = global.ms365AuthGetAccount();
                if (acc) {
                    return String(acc.username || acc.name || '').trim();
                }
            }
        } catch (e) {
            /* ignore */
        }
        return '';
    }

    function openDialog() {
        var dlg = ensureDialog();
        pendingFiles.forEach(function (f) {
            try {
                if (f.preview) URL.revokeObjectURL(f.preview);
            } catch (e) {
                /* ignore */
            }
        });
        pendingFiles = [];
        renderThumbs();
        $('ms365FeedbackForm').reset();
        var kind = dlg.querySelector('input[name="ms365FbKind"][value="Feedback"]');
        if (kind) kind.checked = true;
        var who = currentUserHint();
        $('ms365FeedbackMeta').textContent = who
            ? 'Wird gesendet als ' + who + ' · Seite: ' + location.pathname
            : 'Bitte zuerst mit Microsoft 365 anmelden. · Seite: ' + location.pathname;
        setStatus('', '');
        if (typeof dlg.showModal === 'function') dlg.showModal();
        else dlg.setAttribute('open', '');
        setTimeout(function () {
            var t = $('ms365FeedbackTitleInput');
            if (t) t.focus();
        }, 30);
    }

    async function onSubmit(ev) {
        ev.preventDefault();
        var title = ($('ms365FeedbackTitleInput') && $('ms365FeedbackTitleInput').value) || '';
        var body = ($('ms365FeedbackBody') && $('ms365FeedbackBody').value) || '';
        var kindEl = document.querySelector('input[name="ms365FbKind"]:checked');
        var category = kindEl ? kindEl.value : 'Feedback';
        var sendBtn = $('ms365FeedbackSend');
        if (!title.trim() || !body.trim()) {
            setStatus('Titel und Beschreibung sind nötig.', 'err');
            return;
        }
        if (!global.ms365LicenseApi || typeof global.ms365LicenseApi.submitFeedback !== 'function') {
            setStatus('License-API-Client fehlt – bitte Seite neu laden.', 'err');
            return;
        }
        if (sendBtn) sendBtn.disabled = true;
        setStatus('Sende Feedback …', '');
        try {
            var attachments = [];
            for (var i = 0; i < pendingFiles.length; i++) {
                var pf = pendingFiles[i];
                attachments.push({
                    name: pf.name,
                    contentType: pf.contentType,
                    base64: await fileToBase64(pf.file)
                });
            }
            var token = await global.ms365LicenseApi.acquireLicenseToken({ popup: true });
            await global.ms365LicenseApi.submitFeedback(token, {
                title: title.trim(),
                description: body.trim(),
                category: category,
                pageUrl: String(location.href || '').slice(0, 2000),
                userAgent: String(navigator.userAgent || '').slice(0, 1000),
                attachments: attachments
            });
            setStatus('Danke – Feedback ist angekommen.', 'ok');
            setTimeout(function () {
                var dlg = $('ms365FeedbackDialog');
                if (dlg) dlg.close();
            }, 900);
        } catch (e) {
            var msg = (e && e.message) || 'Senden fehlgeschlagen.';
            if (/anmeld|login|token|silent/i.test(msg)) {
                msg = 'Bitte zuerst mit Microsoft 365 anmelden, dann erneut senden.';
            }
            setStatus(msg, 'err');
        } finally {
            if (sendBtn) sendBtn.disabled = false;
        }
    }

    function ensureFooterLink() {
        var row = document.getElementById('siteCreditKurtrocks') || document.querySelector('.site-credit-row');
        if (!row) return;
        var existing = row.querySelector('.site-feedback-link');
        if (existing) return;
        var a = document.createElement('a');
        a.href = '#';
        a.className = 'site-credit-link site-feedback-link';
        a.title = 'Feedback an den Betreiber senden';
        a.setAttribute('aria-label', 'Feedback senden');
        var icon = document.createElement('i');
        icon.className = 'bi bi-chat-heart';
        icon.setAttribute('aria-hidden', 'true');
        a.appendChild(icon);
        a.appendChild(document.createTextNode('Feedback'));
        a.addEventListener('click', function (ev) {
            ev.preventDefault();
            openDialog();
        });
        var privacy = row.querySelector('.site-privacy-link');
        if (privacy && privacy.nextSibling) row.insertBefore(a, privacy.nextSibling);
        else if (privacy) privacy.after(a);
        else row.appendChild(a);
    }

    function mount() {
        ensureFooterLink();
        ensureDialog();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mount);
    } else {
        mount();
    }
    // Footer wird oft nachgeladen
    setTimeout(ensureFooterLink, 400);
    setTimeout(ensureFooterLink, 1500);

    global.ms365Feedback = {
        open: openDialog,
        mount: mount
    };
})(typeof window !== 'undefined' ? window : globalThis);
