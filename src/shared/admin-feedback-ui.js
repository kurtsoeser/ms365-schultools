/**
 * Admin-Tab: Feedback aus Betreiber-SharePoint – Inbox mit Vorschau & Status-Ablauf.
 */
(function (global) {
    'use strict';

    var cache = [];
    var statuses = ['Neu', 'In Bearbeitung', 'Erledigt', 'Abgelehnt'];
    var selectedId = '';
    /** @type {Record<string, string>} path -> object URL */
    var previewCache = {};
    var previewLoading = {};

    var STATUS_META = {
        Neu: {
            label: 'Neu',
            hint: 'Noch ungelesen – kurz sichten',
            next: 'In Bearbeitung',
            nextLabel: 'In Prüfung nehmen',
            tone: 'new'
        },
        'In Bearbeitung': {
            label: 'In Prüfung',
            hint: 'Du schaust es dir an / planst es',
            next: 'Erledigt',
            nextLabel: 'Als erledigt markieren',
            tone: 'progress'
        },
        Erledigt: {
            label: 'Erledigt',
            hint: 'Umgesetzt oder beantwortet',
            next: null,
            nextLabel: '',
            tone: 'done'
        },
        Abgelehnt: {
            label: 'Ablage',
            hint: 'Nicht relevant / Duplikat / später',
            next: null,
            nextLabel: '',
            tone: 'park'
        }
    };

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

    function formatWhen(iso) {
        if (!iso) return '–';
        try {
            var d = new Date(iso);
            if (Number.isNaN(d.getTime())) return String(iso);
            return d.toLocaleString('de-AT', {
                dateStyle: 'short',
                timeStyle: 'short'
            });
        } catch (e) {
            return String(iso);
        }
    }

    function setBanner(msg, kind) {
        var el = $('adminFeedbackBanner');
        if (!el) return;
        if (!msg) {
            el.hidden = true;
            el.textContent = '';
            return;
        }
        el.hidden = false;
        el.textContent = msg;
        el.dataset.state = kind || '';
    }

    function api() {
        return global.ms365LicenseApi;
    }

    async function withToken(fn) {
        var client = api();
        if (!client || typeof client.acquireLicenseToken !== 'function') {
            throw new Error('License-API-Client fehlt.');
        }
        var token = await client.acquireLicenseToken({ popup: true });
        return fn(token, client);
    }

    function countByStatus(items) {
        var c = { Neu: 0, 'In Bearbeitung': 0, Erledigt: 0, Abgelehnt: 0 };
        (items || []).forEach(function (it) {
            var s = it.status || 'Neu';
            if (c[s] == null) c[s] = 0;
            c[s] += 1;
        });
        return c;
    }

    function renderStats(items) {
        var box = $('adminFeedbackStats');
        if (!box) return;
        var c = countByStatus(items);
        var open = c.Neu + c['In Bearbeitung'];
        box.innerHTML =
            '<button type="button" class="admin-fb__stat admin-fb__stat--new" data-fb-quick="">' +
            '<strong>' +
            (items || []).length +
            '</strong><span>Gesamt</span></button>' +
            '<button type="button" class="admin-fb__stat admin-fb__stat--new" data-fb-quick="Neu">' +
            '<strong>' +
            c.Neu +
            '</strong><span>Neu</span></button>' +
            '<button type="button" class="admin-fb__stat admin-fb__stat--progress" data-fb-quick="In Bearbeitung">' +
            '<strong>' +
            c['In Bearbeitung'] +
            '</strong><span>In Prüfung</span></button>' +
            '<button type="button" class="admin-fb__stat admin-fb__stat--open" data-fb-quick="__open">' +
            '<strong>' +
            open +
            '</strong><span>Offen</span></button>' +
            '<button type="button" class="admin-fb__stat admin-fb__stat--done" data-fb-quick="Erledigt">' +
            '<strong>' +
            c.Erledigt +
            '</strong><span>Erledigt</span></button>';
        box.querySelectorAll('[data-fb-quick]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                var v = btn.getAttribute('data-fb-quick') || '';
                var filter = $('adminFeedbackStatusFilter');
                if (!filter) return;
                if (v === '__open') {
                    filter.value = '';
                    $('adminFeedbackSearch').value = '';
                    // show only open via temporary filter marker
                    filter.dataset.openOnly = '1';
                } else {
                    filter.dataset.openOnly = '';
                    filter.value = v;
                }
                renderList(cache);
            });
        });
    }

    function filteredRows() {
        var filterEl = $('adminFeedbackStatusFilter');
        var filter = (filterEl && filterEl.value) || '';
        var openOnly = filterEl && filterEl.dataset.openOnly === '1';
        var q = String(($('adminFeedbackSearch') && $('adminFeedbackSearch').value) || '')
            .trim()
            .toLowerCase();
        return cache.filter(function (it) {
            if (openOnly) {
                if (it.status !== 'Neu' && it.status !== 'In Bearbeitung') return false;
            } else if (filter && it.status !== filter) {
                return false;
            }
            if (!q) return true;
            var hay = [
                it.title,
                it.description,
                it.reporterEmail,
                it.reporterName,
                it.schoolName,
                it.domain,
                it.category
            ]
                .join(' ')
                .toLowerCase();
            return hay.indexOf(q) >= 0;
        });
    }

    function statusBadge(status) {
        var meta = STATUS_META[status] || STATUS_META.Neu;
        return (
            '<span class="admin-fb__badge admin-fb__badge--' +
            meta.tone +
            '" title="' +
            escapeHtml(meta.hint) +
            '">' +
            escapeHtml(meta.label) +
            '</span>'
        );
    }

    function categoryPill(cat) {
        var c = String(cat || 'Feedback');
        var tone =
            c === 'Bug' ? 'bug' : c === 'Idee' ? 'idea' : c === 'Frage' ? 'ask' : 'fb';
        return (
            '<span class="admin-fb__pill admin-fb__pill--' +
            tone +
            '">' +
            escapeHtml(c) +
            '</span>'
        );
    }

    function previewSrc(att) {
        if (!att) return '';
        var path = String(att.path || '').trim();
        if (path && Object.prototype.hasOwnProperty.call(previewCache, path)) {
            return previewCache[path] || String(att.url || '').trim();
        }
        return String(att.url || '').trim();
    }

    function ensurePreviews(items) {
        var need = [];
        (items || []).forEach(function (it) {
            (it.attachments || []).forEach(function (a) {
                var path = String(a.path || '').trim();
                if (
                    !path ||
                    Object.prototype.hasOwnProperty.call(previewCache, path) ||
                    previewLoading[path]
                ) {
                    return;
                }
                need.push(path);
            });
        });
        if (!need.length) return;
        void withToken(function (token, client) {
            return Promise.all(
                need.map(function (path) {
                    previewLoading[path] = true;
                    return client
                        .adminFetchFeedbackFile(token, path)
                        .then(function (file) {
                            previewCache[path] = file.objectUrl;
                        })
                        .catch(function () {
                            previewCache[path] = '';
                        })
                        .finally(function () {
                            delete previewLoading[path];
                        });
                })
            ).then(function () {
                renderList(cache);
            });
        }).catch(function () {
            /* ignore */
        });
    }

    function thumbHtml(attachments) {
        var list = attachments || [];
        if (!list.length) {
            return '<div class="admin-fb__thumb admin-fb__thumb--empty" aria-hidden="true"><i class="bi bi-chat-left-text"></i></div>';
        }
        var first = list[0];
        var src = previewSrc(first);
        var more = list.length > 1 ? '<span class="admin-fb__thumb-more">+' + (list.length - 1) + '</span>' : '';
        return (
            '<a class="admin-fb__thumb" href="' +
            escapeHtml(first.url || src) +
            '" target="_blank" rel="noopener" title="' +
            escapeHtml(first.name || 'Screenshot') +
            '">' +
            (src
                ? '<img src="' + escapeHtml(src) + '" alt="" loading="lazy">'
                : '<i class="bi bi-image"></i>') +
            more +
            '</a>'
        );
    }

    function renderList(items) {
        var box = $('adminFeedbackList');
        if (!box) return;
        if (Array.isArray(items)) cache = items;
        renderStats(cache);
        ensurePreviews(cache);
        var rows = filteredRows();

        $('adminFeedbackCount') &&
            ($('adminFeedbackCount').textContent = rows.length + ' sichtbar · ' + cache.length + ' gesamt');

        if (!rows.length) {
            selectedId = '';
            box.innerHTML =
                '<div class="admin-fb__empty">' +
                '<i class="bi bi-inbox" aria-hidden="true"></i>' +
                '<p>Keine Einträge' +
                (rows !== cache ? ' für diesen Filter' : '') +
                '.</p></div>';
            renderDetail(null);
            return;
        }

        if (!selectedId || !rows.some(function (r) {
            return String(r.id) === String(selectedId);
        })) {
            selectedId = String(rows[0].id);
        }

        box.innerHTML = rows
            .map(function (it) {
                var active = String(it.id) === String(selectedId) ? ' is-active' : '';
                var unread = it.status === 'Neu' ? ' is-unread' : '';
                return (
                    '<button type="button" class="admin-fb__row' +
                    active +
                    unread +
                    '" data-fb-open="' +
                    escapeHtml(it.id) +
                    '">' +
                    thumbHtml(it.attachments) +
                    '<div class="admin-fb__row-main">' +
                    '<div class="admin-fb__row-top">' +
                    categoryPill(it.category) +
                    statusBadge(it.status) +
                    '</div>' +
                    '<strong class="admin-fb__row-title">' +
                    escapeHtml(it.title) +
                    '</strong>' +
                    '<p class="admin-fb__row-meta">' +
                    '<span><i class="bi bi-clock"></i> ' +
                    escapeHtml(formatWhen(it.createdAt)) +
                    '</span>' +
                    '<span><i class="bi bi-person"></i> ' +
                    escapeHtml(it.reporterName || it.reporterEmail || 'unbekannt') +
                    '</span>' +
                    (it.schoolName
                        ? '<span><i class="bi bi-building"></i> ' + escapeHtml(it.schoolName) + '</span>'
                        : '') +
                    '</p></div></button>'
                );
            })
            .join('');

        box.querySelectorAll('[data-fb-open]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                selectedId = btn.getAttribute('data-fb-open') || '';
                renderList(cache);
            });
        });

        var sel = cache.find(function (it) {
            return String(it.id) === String(selectedId);
        });
        renderDetail(sel || null);
    }

    function renderDetail(it) {
        var pane = $('adminFeedbackDetail');
        if (!pane) return;
        if (!it) {
            pane.innerHTML =
                '<div class="admin-fb__detail-empty">' +
                '<i class="bi bi-layout-sidebar-inset" aria-hidden="true"></i>' +
                '<p>Eintrag links auswählen.</p></div>';
            return;
        }

        var meta = STATUS_META[it.status] || STATUS_META.Neu;
        var gallery = (it.attachments || [])
            .map(function (a) {
                var src = previewSrc(a);
                return (
                    '<a class="admin-fb__shot" href="' +
                    escapeHtml(a.url || src) +
                    '" target="_blank" rel="noopener">' +
                    (src
                        ? '<img src="' +
                          escapeHtml(src) +
                          '" alt="' +
                          escapeHtml(a.name || 'Screenshot') +
                          '" loading="lazy">'
                        : '<div class="admin-fb__shot-ph"><i class="bi bi-image"></i></div>') +
                    '<span>' +
                    escapeHtml(a.name || 'Screenshot') +
                    '</span></a>'
                );
            })
            .join('');

        var statusOpts = statuses
            .map(function (s) {
                var m = STATUS_META[s] || { label: s };
                return (
                    '<option value="' +
                    escapeHtml(s) +
                    '"' +
                    (it.status === s ? ' selected' : '') +
                    '>' +
                    escapeHtml(m.label) +
                    '</option>'
                );
            })
            .join('');

        var nextBtn = meta.next
            ? '<button type="button" class="btn" data-fb-next="' +
              escapeHtml(meta.next) +
              '"><i class="bi bi-arrow-right-circle"></i>' +
              escapeHtml(meta.nextLabel) +
              '</button>'
            : '';

        var mail = it.reporterEmail
            ? 'mailto:' +
              encodeURIComponent(it.reporterEmail) +
              '?subject=' +
              encodeURIComponent('Re: ' + (it.title || 'Feedback MS365-Schultools'))
            : '';

        pane.innerHTML =
            '<article class="admin-fb__detail">' +
            '<header class="admin-fb__detail-head">' +
            '<div class="admin-fb__detail-tags">' +
            categoryPill(it.category) +
            statusBadge(it.status) +
            '</div>' +
            '<h2>' +
            escapeHtml(it.title) +
            '</h2>' +
            '<p class="admin-fb__detail-sub">' +
            escapeHtml(meta.hint) +
            ' · ' +
            escapeHtml(formatWhen(it.createdAt)) +
            '</p></header>' +
            '<div class="admin-fb__who">' +
            '<div><span class="admin-fb__who-label">Von</span><strong>' +
            escapeHtml(it.reporterName || '–') +
            '</strong>' +
            (it.reporterEmail
                ? '<a href="' +
                  escapeHtml(mail) +
                  '">' +
                  escapeHtml(it.reporterEmail) +
                  '</a>'
                : '') +
            '</div>' +
            '<div><span class="admin-fb__who-label">Schule</span><strong>' +
            escapeHtml(it.schoolName || '–') +
            '</strong><span>' +
            escapeHtml(it.domain || '') +
            '</span></div></div>' +
            '<div class="admin-fb__detail-body">' +
            escapeHtml(it.description).replace(/\n/g, '<br>') +
            '</div>' +
            (gallery ? '<div class="admin-fb__gallery">' + gallery + '</div>' : '') +
            (it.pageUrl
                ? '<p class="admin-fb__page"><i class="bi bi-link-45deg"></i><a href="' +
                  escapeHtml(it.pageUrl) +
                  '" target="_blank" rel="noopener">' +
                  escapeHtml(it.pageUrl) +
                  '</a></p>'
                : '') +
            '<div class="admin-fb__workflow">' +
            '<p class="admin-fb__workflow-title">Weiter mit diesem Feedback</p>' +
            '<p class="admin-fb__workflow-hint">Kein Antwort-Chat in der App – Status für dich, optional Mail an die Person.</p>' +
            '<div class="admin-fb__workflow-actions">' +
            nextBtn +
            (it.status !== 'Abgelehnt'
                ? '<button type="button" class="btn alt" data-fb-next="Abgelehnt"><i class="bi bi-archive"></i>In Ablage</button>'
                : '<button type="button" class="btn alt" data-fb-next="Neu"><i class="bi bi-arrow-counterclockwise"></i>Wieder öffnen</button>') +
            (it.status === 'Erledigt'
                ? '<button type="button" class="btn alt" data-fb-next="In Bearbeitung"><i class="bi bi-arrow-counterclockwise"></i>Zurück in Prüfung</button>'
                : '') +
            (mail
                ? '<a class="btn alt" href="' +
                  escapeHtml(mail) +
                  '"><i class="bi bi-envelope"></i>E-Mail öffnen</a>'
                : '') +
            '</div>' +
            '<label class="admin-fb__status-field"><span>Status manuell</span>' +
            '<select data-fb-status="' +
            escapeHtml(it.id) +
            '">' +
            statusOpts +
            '</select></label></div>' +
            '<p class="admin-fb__foot">Tenant <code>' +
            escapeHtml(it.tenantId || '–') +
            '</code></p></article>';

        pane.querySelectorAll('[data-fb-next]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                void updateStatus(it.id, btn.getAttribute('data-fb-next'), null);
            });
        });
        var sel = pane.querySelector('[data-fb-status]');
        if (sel) {
            sel.addEventListener('change', function () {
                void updateStatus(it.id, sel.value, sel);
            });
        }
    }

    async function updateStatus(id, status, sel) {
        try {
            setBanner('Status wird gespeichert …', '');
            await withToken(function (token, client) {
                return client.adminUpdateFeedback(token, id, { status: status });
            });
            cache = cache.map(function (it) {
                if (String(it.id) === String(id)) return Object.assign({}, it, { status: status });
                return it;
            });
            selectedId = String(id);
            setBanner('Status → ' + ((STATUS_META[status] && STATUS_META[status].label) || status), 'ok');
            renderList(cache);
        } catch (e) {
            setBanner((e && e.message) || 'Status konnte nicht gespeichert werden.', 'err');
            if (sel) {
                var prev = cache.find(function (it) {
                    return String(it.id) === String(id);
                });
                if (prev) sel.value = prev.status;
            }
        }
    }

    async function reload() {
        setBanner('Lade Feedback …', '');
        try {
            var data = await withToken(function (token, client) {
                return client.adminListFeedback(token, { top: 200 });
            });
            if (data && Array.isArray(data.statuses) && data.statuses.length) {
                statuses = data.statuses.slice();
            }
            renderList(data.items || []);
            setBanner(
                (data.items || []).length
                    ? 'Aktualisiert · ' + (data.listDisplayName || 'Feedback')
                    : 'Liste leer – ggf. „Liste sicherstellen“.',
                'ok'
            );
        } catch (e) {
            setBanner((e && e.message) || 'Feedback konnte nicht geladen werden.', 'err');
            renderList([]);
        }
    }

    async function ensureSchema() {
        setBanner('Lege SharePoint-Liste / Spalten an …', '');
        try {
            var data = await withToken(function (token, client) {
                return client.adminEnsureFeedbackSchema(token);
            });
            var created =
                data && data.schema && data.schema.columns && data.schema.columns.created
                    ? data.schema.columns.created
                    : [];
            setBanner(
                created.length
                    ? 'Schema bereit. Neu: ' + created.join(', ')
                    : 'Schema war bereits vollständig.',
                'ok'
            );
            await reload();
        } catch (e) {
            setBanner((e && e.message) || 'Schema-Anlage fehlgeschlagen.', 'err');
        }
    }

    function bind() {
        var reloadBtn = $('adminFeedbackReloadBtn');
        var schemaBtn = $('adminFeedbackEnsureBtn');
        var filter = $('adminFeedbackStatusFilter');
        var search = $('adminFeedbackSearch');
        if (reloadBtn) {
            reloadBtn.addEventListener('click', function () {
                void reload();
            });
        }
        if (schemaBtn) {
            schemaBtn.addEventListener('click', function () {
                void ensureSchema();
            });
        }
        if (filter) {
            filter.addEventListener('change', function () {
                filter.dataset.openOnly = '';
                renderList(cache);
            });
        }
        if (search) {
            var t = null;
            search.addEventListener('input', function () {
                clearTimeout(t);
                t = setTimeout(function () {
                    renderList(cache);
                }, 180);
            });
        }
    }

    function onTabActivated() {
        if (!cache.length) void reload();
        else renderList(cache);
    }

    function init() {
        if (!$('adminPanelFeedback')) return;
        bind();
        global.ms365AdminFeedback = {
            reload: reload,
            ensureSchema: ensureSchema,
            onTabActivated: onTabActivated
        };
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})(typeof window !== 'undefined' ? window : globalThis);
