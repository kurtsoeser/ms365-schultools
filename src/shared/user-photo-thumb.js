/**
 * Lazy MS365-Profilfotos (IntersectionObserver + Cache), analog zu group-photo-thumb.js.
 * Markup: [data-ms365-user-photo="<email|upn|id>"] mit img + .ms365-user-photo-thumb__initials
 */
(function () {
    'use strict';

    /** @type {Map<string, { state: 'loading'|'ok'|'none', url?: string, waiters?: Array<{ resolve: Function, reject: Function }> }>} */
    const cache = new Map();
    /** @type {IntersectionObserver|null} */
    let observer = null;
    /** @type {Set<Element>} */
    const pending = new Set();
    let flushTimer = null;

    function gug() {
        return window.ms365GraphUnifiedGroups;
    }

    function cacheKey(userKey) {
        return String(userKey || '')
            .trim()
            .toLowerCase();
    }

    function initials(displayName) {
        const G = gug();
        if (G && typeof G.userPhotoInitials === 'function') {
            return G.userPhotoInitials(displayName);
        }
        if (G && typeof G.groupPhotoInitials === 'function') {
            return G.groupPhotoInitials(displayName);
        }
        const s = String(displayName || '').trim();
        if (!s) return '?';
        const parts = s.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
            return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
        }
        return s.slice(0, 2).toUpperCase();
    }

    function applyThumbState(el, state, url) {
        const img = el.querySelector('img');
        const ini = el.querySelector('.ms365-user-photo-thumb__initials');
        if (state === 'ok' && url) {
            el.classList.add('has-photo');
            el.classList.remove('is-empty');
            if (img) {
                img.src = url;
                img.hidden = false;
            }
            if (ini) ini.hidden = true;
            return;
        }
        el.classList.remove('has-photo');
        if (img) {
            img.hidden = true;
            img.removeAttribute('src');
        }
        if (ini) ini.hidden = false;
    }

    function photoSizeFromEl(el) {
        const raw = String((el && el.getAttribute && el.getAttribute('data-ms365-user-photo-size')) || '')
            .trim()
            .toLowerCase();
        if (/^(48|64|96|120|240|360|432|504|648)x(48|64|96|120|240|360|432|504|648)$/.test(raw)) {
            return raw;
        }
        if (el && el.classList && el.classList.contains('ssw-avatar--xl')) return '120x120';
        if (el && el.classList && el.classList.contains('ssw-avatar--lg')) return '96x96';
        if (el && el.classList && el.classList.contains('ssw-avatar--md')) return '64x64';
        return '48x48';
    }

    function fetchPhotoForUser(token, userKey, size) {
        const dim = String(size || '48x48').trim() || '48x48';
        const key = cacheKey(userKey) + '|' + dim;
        if (!cacheKey(userKey)) return Promise.resolve(null);

        const cached = cache.get(key);
        if (cached && cached.state === 'ok') return Promise.resolve(cached.url || null);
        if (cached && cached.state === 'none') return Promise.resolve(null);
        if (cached && cached.state === 'loading' && cached.waiters) {
            return new Promise(function (resolve, reject) {
                cached.waiters.push({ resolve: resolve, reject: reject });
            });
        }

        const entry = { state: 'loading', waiters: [] };
        cache.set(key, entry);

        const G = gug();
        const p =
            G && typeof G.fetchUserPhotoBlob === 'function'
                ? G.fetchUserPhotoBlob(token, userKey, dim)
                : Promise.resolve(null);

        return p
            .then(function (blob) {
                let url = null;
                if (blob && blob.size) {
                    const prev = cache.get(key);
                    if (prev && prev.url) {
                        try {
                            URL.revokeObjectURL(prev.url);
                        } catch {
                            /* ignore */
                        }
                    }
                    url = URL.createObjectURL(blob);
                    cache.set(key, { state: 'ok', url: url });
                } else {
                    cache.set(key, { state: 'none' });
                }
                if (entry.waiters && entry.waiters.length) {
                    entry.waiters.forEach(function (w) {
                        w.resolve(url);
                    });
                }
                return url;
            })
            .catch(function (err) {
                cache.set(key, { state: 'none' });
                if (entry.waiters && entry.waiters.length) {
                    entry.waiters.forEach(function (w) {
                        w.reject(err);
                    });
                }
                throw err;
            });
    }

    async function loadOne(el, token) {
        const userKey = cacheKey(el.getAttribute('data-ms365-user-photo'));
        if (!userKey || !el.isConnected) return;
        const dim = photoSizeFromEl(el);
        const key = userKey + '|' + dim;

        const cached = cache.get(key);
        if (cached && cached.state === 'ok') {
            applyThumbState(el, 'ok', cached.url);
            return;
        }
        if (cached && cached.state === 'none') {
            applyThumbState(el, 'none');
            return;
        }

        try {
            const url = await fetchPhotoForUser(token, userKey, dim);
            if (!el.isConnected) return;
            if (url) applyThumbState(el, 'ok', url);
            else applyThumbState(el, 'none');
        } catch {
            if (el.isConnected) applyThumbState(el, 'none');
        }
    }

    function getObserver() {
        if (observer) return observer;
        observer = new IntersectionObserver(
            function (entries) {
                entries.forEach(function (entry) {
                    if (!entry.isIntersecting) return;
                    observer.unobserve(entry.target);
                    pending.add(entry.target);
                    scheduleFlush();
                });
            },
            { rootMargin: '80px' }
        );
        return observer;
    }

    function scheduleFlush() {
        if (flushTimer) return;
        flushTimer = setTimeout(flushPending, 40);
    }

    async function flushPending() {
        flushTimer = null;
        const batch = Array.from(pending);
        pending.clear();
        if (!batch.length) return;

        let token;
        try {
            const G = gug();
            if (!G || typeof G.getGraphToken !== 'function') return;
            token = await G.getGraphToken();
        } catch {
            return;
        }

        for (let i = 0; i < batch.length; i += 6) {
            const slice = batch.slice(i, i + 6);
            await Promise.all(
                slice.map(function (el) {
                    return loadOne(el, token);
                })
            );
            if (i + 6 < batch.length) {
                await new Promise(function (r) {
                    setTimeout(r, 50);
                });
            }
        }
    }

    function createThumb(opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const userKey = String(o.userKey || o.email || '').trim();
        const displayName = String(o.displayName || o.name || '').trim();
        const size = o.size === 'chip' || o.size === 'sm' ? 'chip' : 'md';

        const wrap = document.createElement('span');
        wrap.className = 'ms365-user-photo-thumb ms365-user-photo-thumb--' + size;
        wrap.setAttribute('data-ms365-user-photo', userKey);
        wrap.setAttribute('data-ms365-user-photo-name', displayName);
        wrap.setAttribute('aria-hidden', 'true');

        const img = document.createElement('img');
        img.alt = '';
        img.hidden = true;
        wrap.appendChild(img);

        const ini = document.createElement('span');
        ini.className = 'ms365-user-photo-thumb__initials';
        ini.textContent = initials(displayName || userKey);
        wrap.appendChild(ini);

        if (!userKey) wrap.classList.add('is-empty');
        return wrap;
    }

    function hydrate(root) {
        const host = root && root.querySelectorAll ? root : document;
        if (!host || typeof host.querySelectorAll !== 'function') return;

        const nodes = host.querySelectorAll('[data-ms365-user-photo]');
        const obs = getObserver();
        nodes.forEach(function (el) {
            const userKey = cacheKey(el.getAttribute('data-ms365-user-photo'));
            if (!userKey) return;
            const dim = photoSizeFromEl(el);
            const key = userKey + '|' + dim;
            const name = el.getAttribute('data-ms365-user-photo-name') || '';
            const ini = el.querySelector('.ms365-user-photo-thumb__initials');
            if (ini && !ini.textContent.trim()) ini.textContent = initials(name || userKey);

            const cached = cache.get(key);
            if (cached && cached.state === 'ok') applyThumbState(el, 'ok', cached.url);
            else if (cached && cached.state === 'none') applyThumbState(el, 'none');

            obs.observe(el);
        });
    }

    function invalidate(userKey) {
        const key = cacheKey(userKey);
        if (!key) return;
        const prev = cache.get(key);
        if (prev && prev.url) {
            try {
                URL.revokeObjectURL(prev.url);
            } catch {
                /* ignore */
            }
        }
        cache.delete(key);

        document.querySelectorAll('[data-ms365-user-photo]').forEach(function (el) {
            if (cacheKey(el.getAttribute('data-ms365-user-photo')) !== key) return;
            applyThumbState(el, 'none');
            getObserver().observe(el);
        });
    }

    function invalidateAll() {
        cache.forEach(function (v) {
            if (v && v.url) {
                try {
                    URL.revokeObjectURL(v.url);
                } catch {
                    /* ignore */
                }
            }
        });
        cache.clear();
    }

    window.ms365UserPhotoThumb = {
        createThumb: createThumb,
        hydrate: hydrate,
        scan: hydrate,
        invalidate: invalidate,
        invalidateAll: invalidateAll
    };
})();
