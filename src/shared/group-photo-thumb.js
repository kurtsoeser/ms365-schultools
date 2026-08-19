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

    function initials(displayName) {
        const G = gug();
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
        const ini = el.querySelector('.gd-group-photo-thumb__initials');
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

    function fetchPhotoForGroup(token, groupId) {
        const gid = String(groupId || '').trim();
        if (!gid) return Promise.resolve(null);

        const cached = cache.get(gid);
        if (cached && cached.state === 'ok') return Promise.resolve(cached.url || null);
        if (cached && cached.state === 'none') return Promise.resolve(null);
        if (cached && cached.state === 'loading' && cached.waiters) {
            return new Promise(function (resolve, reject) {
                cached.waiters.push({ resolve: resolve, reject: reject });
            });
        }

        const entry = { state: 'loading', waiters: [] };
        cache.set(gid, entry);

        const G = gug();
        const p =
            G && typeof G.fetchGroupPhotoBlob === 'function'
                ? G.fetchGroupPhotoBlob(token, gid)
                : Promise.resolve(null);

        return p
            .then(function (blob) {
                let url = null;
                if (blob && blob.size) {
                    const prev = cache.get(gid);
                    if (prev && prev.url) {
                        try {
                            URL.revokeObjectURL(prev.url);
                        } catch {
                            /* ignore */
                        }
                    }
                    url = URL.createObjectURL(blob);
                    cache.set(gid, { state: 'ok', url: url });
                } else {
                    cache.set(gid, { state: 'none' });
                }
                if (entry.waiters && entry.waiters.length) {
                    entry.waiters.forEach(function (w) {
                        w.resolve(url);
                    });
                }
                return url;
            })
            .catch(function (err) {
                cache.set(gid, { state: 'none' });
                if (entry.waiters && entry.waiters.length) {
                    entry.waiters.forEach(function (w) {
                        w.reject(err);
                    });
                }
                throw err;
            });
    }

    async function loadOne(el, token) {
        const gid = String(el.getAttribute('data-gd-group-photo') || '').trim();
        if (!gid || !el.isConnected) return;

        const cached = cache.get(gid);
        if (cached && cached.state === 'ok') {
            applyThumbState(el, 'ok', cached.url);
            return;
        }
        if (cached && cached.state === 'none') {
            applyThumbState(el, 'none');
            return;
        }

        try {
            const url = await fetchPhotoForGroup(token, gid);
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
            { rootMargin: '100px' }
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

        for (let i = 0; i < batch.length; i += 4) {
            const slice = batch.slice(i, i + 4);
            await Promise.all(
                slice.map(function (el) {
                    return loadOne(el, token);
                })
            );
            if (i + 4 < batch.length) {
                await new Promise(function (r) {
                    setTimeout(r, 60);
                });
            }
        }
    }

    function createThumb(opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const groupId = String(o.groupId || '').trim();
        const displayName = String(o.displayName || '').trim();
        const size = o.size === 'search' ? 'search' : 'list';

        const wrap = document.createElement('span');
        wrap.className = 'gd-group-photo-thumb gd-group-photo-thumb--' + size;
        wrap.setAttribute('data-gd-group-photo', groupId);
        wrap.setAttribute('data-gd-group-photo-name', displayName);
        wrap.setAttribute('aria-hidden', 'true');

        const img = document.createElement('img');
        img.alt = '';
        img.hidden = true;
        wrap.appendChild(img);

        const ini = document.createElement('span');
        ini.className = 'gd-group-photo-thumb__initials';
        ini.textContent = initials(displayName);
        wrap.appendChild(ini);

        if (!groupId) wrap.classList.add('is-empty');
        return wrap;
    }

    function hydrate(root) {
        const host = root && root.querySelectorAll ? root : document;
        if (!host || typeof host.querySelectorAll !== 'function') return;

        const nodes = host.querySelectorAll('[data-gd-group-photo]');
        const obs = getObserver();
        nodes.forEach(function (el) {
            const gid = String(el.getAttribute('data-gd-group-photo') || '').trim();
            if (!gid) return;
            const name = el.getAttribute('data-gd-group-photo-name') || '';
            const ini = el.querySelector('.gd-group-photo-thumb__initials');
            if (ini && !ini.textContent.trim()) ini.textContent = initials(name);

            const cached = cache.get(gid);
            if (cached && cached.state === 'ok') applyThumbState(el, 'ok', cached.url);
            else if (cached && cached.state === 'none') applyThumbState(el, 'none');

            obs.observe(el);
        });
    }

    function invalidate(groupId) {
        const gid = String(groupId || '').trim();
        if (!gid) return;
        const prev = cache.get(gid);
        if (prev && prev.url) {
            try {
                URL.revokeObjectURL(prev.url);
            } catch {
                /* ignore */
            }
        }
        cache.delete(gid);

        document.querySelectorAll('[data-gd-group-photo]').forEach(function (el) {
            if (String(el.getAttribute('data-gd-group-photo') || '').trim() !== gid) return;
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

    window.ms365GroupPhotoThumb = {
        createThumb: createThumb,
        hydrate: hydrate,
        invalidate: invalidate,
        invalidateAll: invalidateAll
    };
})();
