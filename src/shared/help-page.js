/**
 * Hilfe-Seite: Volltextsuche, Inhaltsverzeichnis, Hash-Sprung.
 * Kein Bundling nötig – wie die anderen Shared-Skripte per <script defer>.
 */
(function () {
    'use strict';

    function normalizeQuery(q) {
        return String(q || '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, ' ');
    }

    function textHaystack(el) {
        if (!el) return '';
        const extra = el.getAttribute('data-search') || '';
        return ((el.textContent || '') + ' ' + extra).toLowerCase();
    }

    function matchesQuery(el, needle) {
        if (!needle) return true;
        return textHaystack(el).indexOf(needle) !== -1;
    }

    function setHidden(el, hidden) {
        if (!el) return;
        if (hidden) el.setAttribute('hidden', 'hidden');
        else el.removeAttribute('hidden');
    }

    function applyHelpSearch(root, query) {
        const doc = root && root.ownerDocument ? root.ownerDocument : root;
        const needle = normalizeQuery(query);
        const articles = Array.from((root || doc).querySelectorAll('[data-help-article]'));
        const faqs = Array.from((root || doc).querySelectorAll('[data-help-faq]'));
        let visible = 0;

        articles.forEach(function (el) {
            const on = matchesQuery(el, needle);
            setHidden(el, !on);
            if (on) visible += 1;
        });
        faqs.forEach(function (el) {
            const on = matchesQuery(el, needle);
            setHidden(el, !on);
            if (on) visible += 1;
        });

        (root || doc).querySelectorAll('[data-help-section]').forEach(function (sec) {
            const any = sec.querySelector(
                '[data-help-article]:not([hidden]), [data-help-faq]:not([hidden])'
            );
            setHidden(sec, needle.length > 0 && !any);
        });

        (root || doc).querySelectorAll('[data-help-nav] a[href^="#"]').forEach(function (a) {
            const href = a.getAttribute('href') || '';
            const id = href.charAt(0) === '#' ? href.slice(1) : '';
            const target = id && doc.getElementById ? doc.getElementById(id) : null;
            if (!target) return;
            const wrap =
                target.closest('[data-help-article], [data-help-faq], [data-help-section]') || target;
            const hide = needle.length > 0 && !!(wrap && wrap.hidden);
            setHidden(a, hide);
            if (a.parentElement && a.parentElement.tagName === 'LI') {
                setHidden(a.parentElement, hide);
            }
        });

        (root || doc).querySelectorAll('[data-help-nav-group]').forEach(function (group) {
            const any = group.querySelector('a[href^="#"]:not([hidden])');
            setHidden(group, needle.length > 0 && !any);
        });

        return {
            query: needle,
            visible: visible,
            total: articles.length + faqs.length
        };
    }

    function bindHelpPage(doc) {
        const documentRef = doc || (typeof document !== 'undefined' ? document : null);
        if (!documentRef || !documentRef.getElementById) return null;

        const root = documentRef.getElementById('helpRoot') || documentRef.body;
        const input = documentRef.getElementById('helpSearch');
        const meta = documentRef.getElementById('helpSearchMeta');
        const empty = documentRef.getElementById('helpSearchEmpty');
        const clearBtn = documentRef.getElementById('helpSearchClear');
        if (!root || !input) return null;

        let lastQuery = '';

        function render() {
            const stats = applyHelpSearch(root, input.value);
            lastQuery = stats.query;
            if (meta) {
                if (!stats.query) {
                    meta.textContent = '';
                } else if (stats.visible === 0) {
                    meta.textContent = 'Keine Treffer';
                } else if (stats.visible === 1) {
                    meta.textContent = '1 Abschnitt gefunden';
                } else {
                    meta.textContent = stats.visible + ' Abschnitte gefunden';
                }
            }
            if (empty) setHidden(empty, !(stats.query && stats.visible === 0));
            if (clearBtn) setHidden(clearBtn, !stats.query);
            if (documentRef.body && documentRef.body.classList && documentRef.body.classList.toggle) {
                documentRef.body.classList.toggle('help-searching', !!stats.query);
            }
            return stats;
        }

        input.addEventListener('input', render);
        input.addEventListener('search', render);

        if (clearBtn) {
            clearBtn.addEventListener('click', function () {
                input.value = '';
                render();
                input.focus();
            });
        }

        documentRef.addEventListener('keydown', function (e) {
            if (e.key !== '/' && !(e.key === 'k' && (e.ctrlKey || e.metaKey))) return;
            const t = e.target;
            const tag = t && t.tagName ? t.tagName.toLowerCase() : '';
            if (tag === 'input' || tag === 'textarea' || tag === 'select' || (t && t.isContentEditable)) {
                return;
            }
            e.preventDefault();
            input.focus();
            if (typeof input.select === 'function') input.select();
        });

        const tocLinks = Array.from(documentRef.querySelectorAll('[data-help-nav] a[href^="#"]'));
        const observed = [];
        tocLinks.forEach(function (a) {
            const id = (a.getAttribute('href') || '').slice(1);
            const target = id ? documentRef.getElementById(id) : null;
            if (target) observed.push({ a: a, el: target });
        });

        function setActiveToc(id) {
            tocLinks.forEach(function (a) {
                const on = (a.getAttribute('href') || '') === '#' + id;
                if (on) a.setAttribute('aria-current', 'location');
                else a.removeAttribute('aria-current');
            });
        }

        const Observer =
            typeof window !== 'undefined' && typeof window.IntersectionObserver === 'function'
                ? window.IntersectionObserver
                : null;
        if (Observer && observed.length) {
            const io = new Observer(
                function (entries) {
                    const visible = entries
                        .filter(function (en) {
                            return en.isIntersecting;
                        })
                        .sort(function (a, b) {
                            return b.intersectionRatio - a.intersectionRatio;
                        });
                    if (!visible.length) return;
                    setActiveToc(visible[0].target.id);
                },
                { rootMargin: '-12% 0px -70% 0px', threshold: [0.1, 0.25, 0.5] }
            );
            observed.forEach(function (item) {
                io.observe(item.el);
            });
        }

        const loc = typeof window !== 'undefined' ? window.location : null;
        if (loc && loc.hash) {
            const id = decodeURIComponent(String(loc.hash).slice(1));
            const target = documentRef.getElementById(id);
            if (target) {
                setTimeout(function () {
                    try {
                        target.scrollIntoView({ block: 'start' });
                    } catch {
                        /* ignore */
                    }
                }, 40);
            }
        }

        render();
        return {
            render: render,
            getQuery: function () {
                return lastQuery;
            }
        };
    }

    const api = {
        normalizeQuery: normalizeQuery,
        matchesQuery: matchesQuery,
        applyHelpSearch: applyHelpSearch,
        bindHelpPage: bindHelpPage
    };

    if (typeof window !== 'undefined') {
        window.ms365HelpPage = api;
    }

    if (typeof document !== 'undefined') {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', function () {
                bindHelpPage(document);
            });
        } else {
            bindHelpPage(document);
        }
    }
})();
