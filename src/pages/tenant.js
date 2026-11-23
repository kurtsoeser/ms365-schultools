/*
 * Seiten-Skript für tenant.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    'use strict';
    function initDraggableTabOrder(containerId, storageKey) {
        const container = document.getElementById(containerId);
        if (!container) return;
        // Admin-Schiene mit festen Rubriken: keine Drag-Reihenfolge (zerstört Gruppierung).
        if (
            container.getAttribute('data-rail-fixed-groups') === '1' ||
            container.querySelector('.ts-admin-rail__group, .ts-admin-rail__sec')
        ) {
            try {
                localStorage.removeItem(storageKey);
            } catch {
                // ignore
            }
            container.querySelectorAll('.tab-btn[draggable="true"]').forEach(function (btn) {
                btn.removeAttribute('draggable');
                if (btn.title === 'Zum Umsortieren ziehen') btn.removeAttribute('title');
            });
            return;
        }
        const buttons = Array.from(container.querySelectorAll('.tab-btn[role="tab"]')).filter(function (btn) {
            return btn.id !== 'tabMainOrganigramm' && !btn.classList.contains('tab-btn--header-route') && !btn.hidden;
        });
        if (!buttons.length) return;
        const initialIds = buttons.map((btn) => btn.id).filter(Boolean);
        try {
            const saved = JSON.parse(localStorage.getItem(storageKey) || '[]');
            if (Array.isArray(saved) && saved.length) {
                saved.forEach((id) => {
                    const btn = document.getElementById(String(id));
                    if (btn && btn.parentElement === container) container.appendChild(btn);
                });
                initialIds.forEach((id) => {
                    const btn = document.getElementById(id);
                    if (btn && btn.parentElement === container) container.appendChild(btn);
                });
            }
        } catch {
            // ignore
        }

        let draggingId = '';
        function saveOrder() {
            try {
                const ids = Array.from(container.querySelectorAll('.tab-btn[role="tab"]'))
                    .map((btn) => btn.id)
                    .filter(Boolean);
                localStorage.setItem(storageKey, JSON.stringify(ids));
            } catch {
                // ignore
            }
        }
        function clearDropMarkers() {
            container.querySelectorAll('.tab-btn.is-drop-target').forEach((btn) => btn.classList.remove('is-drop-target'));
        }

        buttons.forEach((btn) => {
            btn.setAttribute('draggable', 'true');
            btn.title = 'Zum Umsortieren ziehen';
            if (btn.dataset.dndBound) return;
            btn.dataset.dndBound = '1';
            btn.addEventListener('dragstart', (ev) => {
                draggingId = btn.id;
                btn.classList.add('is-dragging');
                if (ev.dataTransfer) {
                    ev.dataTransfer.effectAllowed = 'move';
                    ev.dataTransfer.setData('text/plain', btn.id);
                }
            });
            btn.addEventListener('dragover', (ev) => {
                ev.preventDefault();
                if (!draggingId || draggingId === btn.id) return;
                clearDropMarkers();
                btn.classList.add('is-drop-target');
                const draggingEl = document.getElementById(draggingId);
                if (!draggingEl || draggingEl === btn || draggingEl.parentElement !== container) return;
                const rect = btn.getBoundingClientRect();
                const before = ev.clientX < rect.left + rect.width / 2;
                if (before) container.insertBefore(draggingEl, btn);
                else container.insertBefore(draggingEl, btn.nextSibling);
            });
            btn.addEventListener('drop', (ev) => {
                ev.preventDefault();
                clearDropMarkers();
                saveOrder();
            });
            btn.addEventListener('dragend', () => {
                btn.classList.remove('is-dragging');
                clearDropMarkers();
                draggingId = '';
                saveOrder();
            });
        });
    }

    function wireTabGroup(tabs, defaultBtnId, storageKey) {
        const byButton = {};
        tabs.forEach(({ btn, panel }) => {
            byButton[btn] = panel;
        });
        function setActive(activeBtnId) {
            tabs.forEach(({ btn, panel }) => {
                const b = document.getElementById(btn);
                const p = document.getElementById(panel);
                const on = btn === activeBtnId;
                if (b) b.setAttribute('aria-selected', on ? 'true' : 'false');
                if (p) p.classList.toggle('active', on);
            });
            const orgOn = activeBtnId === 'tabMainOrganigramm';
            document.body.classList.toggle('tenant-mode-organigramm', orgOn);
            try {
                const nameEl = document.querySelector('.header-tool-indicator__name');
                if (nameEl) {
                    nameEl.innerHTML = orgOn
                        ? '<i class="bi bi-diagram-3"></i>Schul-Organigramm'
                        : '<i class="bi bi-database"></i>Stammdaten';
                }
                const lead = document.querySelector('.header > p:not(.ts-header-mode-hint):not(.header-help-row):not(.ts-header-info-row)');
                if (lead) {
                    lead.hidden = !!orgOn;
                }
                const about = document.querySelector('.ts-register-about');
                if (about) about.hidden = !!orgOn;
                const importBtn = document.querySelector(
                    '.ts-register-header-toolbar__nav a[href*="webuntis-stammdaten-import"]'
                );
                if (importBtn) importBtn.hidden = !!orgOn;
            } catch {
                /* ignore */
            }
            try {
                const curHash = String(location.hash || '')
                    .replace(/^#/, '')
                    .toLowerCase();
                if (orgOn && curHash !== 'organigramm') {
                    history.replaceState(null, '', '#organigramm');
                } else if (
                    !orgOn &&
                    (curHash === 'organigramm' ||
                        curHash === 'schul-organigramm' ||
                        curHash === 'schulorganigramm' ||
                        curHash === 'org')
                ) {
                    history.replaceState(null, '', location.pathname + location.search);
                }
            } catch {
                /* ignore */
            }
            try {
                if (typeof window.ms365SyncSchoolContextHeader === 'function') {
                    window.ms365SyncSchoolContextHeader();
                }
            } catch {
                /* ignore */
            }
            try {
                if (storageKey && byButton[activeBtnId]) localStorage.setItem(storageKey, activeBtnId);
            } catch {
                // ignore
            }
        }
        window.ms365TenantTabActivate = setActive;
        function resolveHashTabId() {
            if (typeof window.ms365TenantResolveHashTabId === 'function') {
                const delegated = window.ms365TenantResolveHashTabId(location.hash);
                if (delegated !== undefined && delegated !== null) return delegated;
            }
            const raw = String(location.hash || '')
                .replace(/^#/, '')
                .toLowerCase();
            if (!raw) return '';
            if (
                raw === 'sync' ||
                raw === 'synchron' ||
                raw === 'import' ||
                raw === 'einspielen' ||
                raw === 'pflegen' ||
                raw === 'pflege' ||
                raw === 'intranet'
            ) {
                return '';
            }
            const mapping = {
                stammdaten: 'tabMainStammdaten',
                profile: 'tabMainStammdaten',
                profil: 'tabMainStammdaten',
                schulprofil: 'tabMainStammdaten',
                email: 'tabMainStammdaten',
                'spo-lists': 'tabMainStammdaten',
                'spo-planer': 'tabMainStammdaten',
                'spo-it': 'tabMainStammdaten',
                'spo-hub': 'tabMainStammdaten',
                'planer-apps': 'tabMainStammdaten',
                planer: 'tabMainStammdaten',
                'tool-site': 'tabMainStammdaten',
                faecher: 'tabMainSubjects',
                subjects: 'tabMainSubjects',
                arge: 'tabMainArges',
                arges: 'tabMainArges',
                verwaltung: 'tabMainVerwaltung',
                sga: 'tabMainSga',
                lehrer: 'tabMainLehrer',
                schueler: 'tabMainSchueler',
                schuelervertretung: 'tabMainSchuelervertretung',
                klassen: 'tabMainKlassen',
                classes: 'tabMainKlassen',
                panelclasses: 'tabMainKlassen',
                klassenlehrer: 'tabMainKlassenlehrer',
                klassenvorstand: 'tabMainKlassenlehrer',
                kv: 'tabMainKlassenlehrer',
                organigramm: 'tabMainOrganigramm',
                'schul-organigramm': 'tabMainOrganigramm',
                schulorganigramm: 'tabMainOrganigramm',
                org: 'tabMainOrganigramm',
                datenlandkarte: 'tabMainDatenlandkarte',
                landkarte: 'tabMainDatenlandkarte',
                'daten-landkarte': 'tabMainDatenlandkarte',
                werkzeuge: 'tabMainStammdaten'
            };
            return mapping[raw] || '';
        }
        tabs.forEach(({ btn }) => {
            const b = document.getElementById(btn);
            if (!b) return;
            b.addEventListener('click', () => setActive(btn));
        });
        let initialBtnId = defaultBtnId;
        try {
            if (storageKey) {
                const stored = localStorage.getItem(storageKey) || '';
                // Organigramm nur über Header/#organigramm – nicht aus Tab-Speicher wiederherstellen
                if (stored && byButton[stored] && stored !== 'tabMainOrganigramm') {
                    initialBtnId = stored;
                }
            }
        } catch {
            // ignore
        }
        const fromHash = resolveHashTabId();
        if (fromHash && byButton[fromHash]) initialBtnId = fromHash;
        setActive(initialBtnId);
        window.addEventListener('hashchange', () => {
            const nextBtnId = resolveHashTabId();
            if (nextBtnId && byButton[nextBtnId]) setActive(nextBtnId);
        });
    }
    initDraggableTabOrder('tenantMainTabs', 'ms365-tenant-main-tab-order-v2');
    wireTabGroup(
        [
            { btn: 'tabMainStammdaten', panel: 'panelMainStammdaten' },
            { btn: 'tabMainOrganigramm', panel: 'panelMainOrganigramm' },
            { btn: 'tabMainVerwaltung', panel: 'panelMainVerwaltung' },
            { btn: 'tabMainSubjects', panel: 'panelMainSubjects' },
            { btn: 'tabMainArges', panel: 'panelMainArges' },
            { btn: 'tabMainSga', panel: 'panelMainSga' },
            { btn: 'tabMainLehrer', panel: 'panelMainLehrer' },
            { btn: 'tabMainSchueler', panel: 'panelStudentsData' },
            { btn: 'tabMainSchuelervertretung', panel: 'panelStudentCouncilData' },
            { btn: 'tabMainKlassen', panel: 'panelMainKlassen' },
            { btn: 'tabMainKlassenlehrer', panel: 'panelMainKlassenlehrer' },
            { btn: 'tabMainDatenlandkarte', panel: 'panelMainDatenlandkarte' }
        ],
        'tabMainStammdaten',
        'ms365-tenant-main-active-tab-v1'
    );
})();
