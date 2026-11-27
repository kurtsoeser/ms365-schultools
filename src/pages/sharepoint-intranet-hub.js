/*
 * Seiten-Skript für tools/sharepoint-intranet-hub.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    function syncPreview() {
        var titleEl = document.getElementById('ihPreviewTitle');
        var descEl = document.getElementById('ihPreviewDesc');
        var chrome = document.getElementById('ihPreviewChromeUrl');
        var title = document.getElementById('fTitle');
        var desc = document.getElementById('fDesc');
        var preview = document.getElementById('fPreviewUrl');
        if (titleEl && title) titleEl.textContent = String(title.value || '').trim() || 'Intranet';
        if (descEl && desc) descEl.textContent = String(desc.value || '').trim() || 'Schulintranet';
        var url = preview ? String(preview.textContent || '').trim() : '';
        if (chrome) chrome.textContent = url && url !== '–' ? url : 'https://…/sites/intranet';
    }
    ['fTitle', 'fDesc', 'fHost', 'fSlug'].forEach(function (id) {
        var el = document.getElementById(id);
        if (!el) return;
        el.addEventListener('input', syncPreview);
        el.addEventListener('change', syncPreview);
    });
    var prev = document.getElementById('fPreviewUrl');
    if (prev && typeof MutationObserver !== 'undefined') {
        new MutationObserver(syncPreview).observe(prev, { childList: true, characterData: true, subtree: true });
    }
    document.addEventListener('DOMContentLoaded', syncPreview);
    setTimeout(syncPreview, 200);

    var IH_KEY = 'ms365-intranet-starter-v1';
    var ihIds = ['lehrer', 'stammdaten-listen', 'termine', 'formulare', 'medien', 'news', 'stammdaten'];
    var ihListenIds = ['lehrer', 'stammdaten-listen', 'termine'];
    function ihLoad() {
        try { return JSON.parse(localStorage.getItem(IH_KEY) || '{}') || {}; } catch (e) { return {}; }
    }
    function ihSave(state) {
        try { localStorage.setItem(IH_KEY, JSON.stringify(state)); } catch (e) { /* ignore */ }
    }
    function ihSetFill(id, ratio) {
        var fill = document.getElementById(id);
        if (fill) fill.style.width = Math.round(Math.max(0, Math.min(1, ratio)) * 100) + '%';
    }
    function ihRefresh() {
        var state = ihLoad();
        var done = 0;
        var listenDone = 0;
        ihIds.forEach(function (id) {
            var el = document.querySelector('[data-ih-check="' + id + '"]');
            if (!el) return;
            el.checked = !!state[id];
            if (el.checked) done++;
            var row = el.closest('.ih-check');
            if (row) row.classList.toggle('is-done', !!el.checked);
        });
        ihListenIds.forEach(function (id) {
            if (state[id]) listenDone++;
        });
        var p = document.getElementById('ihStarterProgress');
        if (p) p.textContent = done + ' / ' + ihIds.length + ' erledigt';
        ihSetFill('ihStarterProgressFill', done / ihIds.length);
        var lp = document.getElementById('ihListenProgress');
        if (lp) lp.textContent = listenDone + ' / ' + ihListenIds.length;
        ihSetFill('ihListenProgressFill', listenDone / ihListenIds.length);
    }
    document.querySelectorAll('[data-ih-check]').forEach(function (el) {
        el.addEventListener('change', function () {
            var state = ihLoad();
            state[el.getAttribute('data-ih-check')] = !!el.checked;
            ihSave(state);
            ihRefresh();
        });
    });
    ihRefresh();
})();
