/**
 * Schul-Landkarte: interaktiver Zoom/Pan-Canvas mit Drill-down.
 */
import {
    childrenOf,
    isAreaNode,
    isJahrgangArea,
    isSammelgruppeArea,
    isSgaBoardArea,
    isSgaScopeArea,
    sgaScopeForArea,
    isElternJahrgangArea,
    isElternClassNode,
    isKlassenchatJahrgangArea,
    isKlassenchatClassNode,
    isArgeAreaNode,
    isArgeLeiterArea,
    isKlassenvorstandArea,
    isClassAreaNode,
    isClassNode,
    normalizeSchulOrganigramm,
    sammelgruppeKeyForArea,
    setSchulOrganigrammView,
    setSchulOrganigrammNodePosition,
    setSchulOrganigrammNodeSize,
    applyLandkartePositionsFromBoxes,
    parseSchulOrgCanvasCoord,
    clampSchulOrganigrammAreaSize,
    SCHUL_ORG_AREA_MIN_W,
    SCHUL_ORG_AREA_MIN_H,
    moveSchulOrganigrammArea,
    moveSchulOrganigrammNode,
    setSchulOrganigrammAreaCollapsed,
    NODE,
    schulOrgGroupKindLabel,
    schulOrgGroupKindIcon,
    schulOrgMailboxIcon,
    normalizeSchulOrgGroupKind,
    personHasM365UserLink,
    nodeM365LinkFlags
} from './schul-organigramm-logic.js';
import {
    areaHasLinkedM365Group,
    nodeUsesGraphFocusPeople
} from './schul-organigramm-focus-people.js';

/**
 * @param {{
 *   active?: boolean,
 *   email?: string,
 *   name?: string,
 *   hitIds?: Set<string>|string[],
 *   pathIds?: Set<string>|string[]
 * }|null|undefined} raw
 */
function normalizePersonFocusOpt(raw) {
    if (!raw || !raw.active) return null;
    const toSet = function (v) {
        if (v instanceof Set) return v;
        return new Set(Array.isArray(v) ? v.map(String) : []);
    };
    return {
        active: true,
        email: String(raw.email || '')
            .trim()
            .toLowerCase(),
        name: String(raw.name || '').trim(),
        hitIds: toSet(raw.hitIds),
        pathIds: toSet(raw.pathIds)
    };
}

/**
 * Typ-spezifische M365-Marken (Chat / Gruppe / Postfach) für die Landkarte.
 * @param {{
 *   linked?: boolean,
 *   hasGroup?: boolean,
 *   hasMailbox?: boolean,
 *   hasChat?: boolean,
 *   displayKind?: string|null,
 *   groupKind?: string
 * }|null|undefined} flags
 * @param {string} [extraClass]
 */
function m365MarkHtml(flags, extraClass) {
    if (!flags || !(flags.linked || flags.showChatMark)) return '';
    /** @type {Array<{ icon: string, title: string, kind: string }>} */
    let marks = [];
    if (Array.isArray(flags.marks) && flags.marks.length) {
        marks = flags.marks.map(function (m) {
            return {
                icon: m.icon || schulOrgGroupKindIcon(m.kind),
                title: m.title || schulOrgGroupKindLabel(m.kind),
                kind: m.kind === 'groupChat' ? 'chat' : m.kind === 'mailbox' ? 'mailbox' : 'group'
            };
        });
    } else {
        if (flags.hasChat || flags.showChatMark) {
            marks.push({
                icon: schulOrgGroupKindIcon('groupChat'),
                title: flags.hasChat
                    ? 'Teams-Gruppenchat verknüpft'
                    : 'Teams-Gruppenchat (KlassenChat)',
                kind: 'chat'
            });
        }
        if (flags.hasGroup) {
            const kind =
                flags.displayKind && flags.displayKind !== 'groupChat'
                    ? flags.displayKind
                    : flags.groupKind || 'unified';
            marks.push({
                icon: schulOrgGroupKindIcon(kind),
                title: schulOrgGroupKindLabel(kind) + ' verknüpft',
                kind: 'group'
            });
        }
        if (flags.hasMailbox) {
            marks.push({
                icon: schulOrgMailboxIcon(),
                title: 'Freigegebenes Postfach verknüpft',
                kind: 'mailbox'
            });
        }
    }
    if (!marks.length) return '';
    return marks
        .map(function (m) {
            return (
                '<span class="schul-org-map__m365-mark is-' +
                m.kind +
                (extraClass ? ' ' + extraClass : '') +
                '" title="' +
                escapeHtml(m.title) +
                '" aria-label="' +
                escapeHtml(m.title) +
                '"><i class="bi ' +
                m.icon +
                '" aria-hidden="true"></i></span>'
            );
        })
        .join('');
}

/**
 * @param {{ email?: string, name?: string }|null|undefined} person
 */
function personM365MarkHtml(person) {
    if (!personHasM365UserLink(person)) return '';
    return (
        '<span class="schul-org-map__m365-user" title="M365-Konto (E-Mail) – Foto wird nachgeladen, wenn vorhanden" aria-label="Mit Microsoft 365 verknüpft">' +
        '<i class="bi bi-microsoft" aria-hidden="true"></i></span>'
    );
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/** Standardbreite Bereichs-Container (Rollen + Personen brauchen Platz) */
const AREA_W = 420;
/** Mindestbreite beim Fenster-Resize (etwas unter Default) */
const AREA_RESIZE_MIN_W = SCHUL_ORG_AREA_MIN_W || 280;
const AREA_RESIZE_MIN_H = SCHUL_ORG_AREA_MIN_H || 110;
const AREA_PAD = 12;
/** Rollen-/ARGE-Zeile inkl. Personen-Chip (muss zur sichtbaren Kachelhöhe passen) */
const LEAF_H = 52;
/** Kopf inkl. Padding/Thumb/Zwei-Zeilen-Titel (muss zu .schul-org-map__area-head passen) */
const HEAD_H = 74;
/** Kompakter Kopf für Klasse (Unterrichtsteams erst im Klassen-Zoom) */
const CLASS_NEST_HEAD_H = 48;
/** Luft zwischen Kopf und erstem Unter-Item (Rollen/Personen/Unterbereiche) */
const CONTENT_GAP = 24;
const NEST_GAP = 12;
/** Extra-Abstand zwischen verschachtelten Bereichen (vermeidet optische Überlappung) */
const NEST_AREA_GAP = 18;
/** Ab dieser Zeigerbewegung (CSS-px) zählt Press+Move als Drag; darunter = Klick/Zoom */
const DRAG_THRESHOLD_PX = 6;

function isKursteamNode(n) {
    const t = String((n && n.type) || '').toLowerCase();
    return t === NODE.KURSTEAM || t === 'kursteam';
}

/** Rollen/ARGEs/Personen-Kacheln auf der Landkarte verschiebbar */
function isMapLeafDraggable(n) {
    const t = String((n && n.type) || '').toLowerCase();
    return (
        t === NODE.ROLE ||
        t === 'role' ||
        t === NODE.ARGE ||
        t === 'arge' ||
        t === NODE.PERSON ||
        t === 'person' ||
        t === NODE.CUSTOM ||
        t === 'custom'
    );
}

function isPersonDropTargetNode(n) {
    if (!n) return false;
    if (isAreaNode(n)) return true;
    const t = String(n.type || '').toLowerCase();
    return t === NODE.ROLE || t === 'role' || t === NODE.ARGE || t === 'arge' || t === NODE.CLASS || t === 'class';
}

/**
 * Klasse auf der Übersicht: kompakt mit Anzahl Unterrichtsteams.
 * Die Teams selbst erscheinen erst beim Reinzoomen in die Klasse (Focus-Layout).
 * @param {object} base
 * @param {object} classNode
 * @param {number} absX
 * @param {number} absY
 * @param {string|null} parentAreaId
 * @param {number} [maxW] – Innenbreite des Elternfensters (sonst Default)
 */
function layoutClassWithKursteams(base, classNode, absX, absY, parentAreaId, maxW) {
    const kursCount = childrenOf(base, classNode.id).filter(isKursteamNode).length;
    const budget =
        maxW != null && Number.isFinite(maxW)
            ? Math.max(120, maxW)
            : Math.max(160, AREA_W - AREA_PAD * 2);
    const h = CLASS_NEST_HEAD_H + 12;
    return {
        boxes: [
            {
                node: classNode,
                x: absX,
                y: absY,
                w: budget,
                h: h,
                kind: 'area',
                parentAreaId: parentAreaId,
                classNest: true,
                kursCount: kursCount,
                widthLocked: maxW != null && Number.isFinite(maxW)
            }
        ],
        w: budget,
        h: h
    };
}
/** Vorschau-Personen in Bereichen (Lehrkräfte, IT-Team, Verwaltungsgruppen …) */
const PREVIEW_CHIP_H = 30;
const PREVIEW_CHIP_GAP = 6;
const PREVIEW_CHIP_MAX = 8;

/**
 * Ob ein Bereich Personen-Chips in der Übersicht zeigen darf
 * (nicht: Schüler-/Eltern-/KlassenChats-Sammel, Jahrgang, SGA-Board, SV-Sitze, Klassen).
 * SGA Lehrer/Eltern + Eltern-/KlassenChat-Klassen-Kacheln: ja.
 * @param {object} area
 */
function areaAllowsPeoplePreview(area) {
    if (!area || !isAreaNode(area)) return false;
    if (isJahrgangArea(area) || isElternJahrgangArea(area) || isKlassenchatJahrgangArea(area)) {
        return false;
    }
    if (isSgaBoardArea(area)) return false;
    // Schüler-Klassen-Bereich: KV + Schüler
    if (isClassNode(area)) return true;
    // Eltern-Klassen-Kachel: Erziehungsberechtigte aus Stammdaten
    if (isElternClassNode(area)) return true;
    // KlassenChat-Kachel: Lehrkräfte des Klassen-Gruppenchats
    if (isKlassenchatClassNode(area)) return true;
    // ARGE-Bereich: Leitung + Fachlehrer
    if (isArgeAreaNode(area)) return true;
    if (isArgeLeiterArea(area)) return true;
    if (isKlassenvorstandArea(area)) return true;
    // Schülervertretung zeigt die drei Funktions-Karten; Lehrer/Eltern brauchen Chips
    if (isSgaScopeArea(area)) {
        const scope = sgaScopeForArea(area);
        return scope === 'teacher' || scope === 'external';
    }
    const sammel = sammelgruppeKeyForArea(area);
    if (sammel === 'schueler' || sammel === 'eltern' || sammel === 'klassenchats') return false;
    if (sammel === 'lehrer') return true;
    if (isSammelgruppeArea(area)) return false;
    return true;
}

/** Extra Höhe für Leitungs-Band + Mitglieder-Trenner in der Vorschau */
const PREVIEW_LEAD_EXTRA = 18;

/**
 * Höhe des Personen-Vorschau-Streifens inkl. optionaler „+N“-Zeile
 * und etwas Luft für Leitungs-Band / Trenner.
 * @param {number} peopleCount
 */
function peoplePreviewStripHeight(peopleCount) {
    const total = Math.max(0, Math.floor(Number(peopleCount) || 0));
    const n = Math.min(PREVIEW_CHIP_MAX, total);
    if (n <= 0) return 0;
    const moreRow = total > PREVIEW_CHIP_MAX ? 1 : 0;
    return (n + moreRow) * (PREVIEW_CHIP_H + PREVIEW_CHIP_GAP) + 16 + PREVIEW_LEAD_EXTRA;
}
const FOCUS_CHIP_W = 200;
const FOCUS_CHIP_H = 64;
const FOCUS_COLS = 3;
const EXPAND_INNER_PAD = 22;
const EXPAND_MIN_W = 680;

function contentTop(areaY) {
    return areaY + HEAD_H + CONTENT_GAP;
}

/**
 * Einfügeposition unter Geschwistern anhand Drop-Y (vor diesem Knoten; '' = ans Ende).
 * @param {object[]} boxes
 * @param {string} parentId
 * @param {string} dragId
 * @param {number} worldY
 * @param {'leaf'|'area'|''} [kind] – nur diese Box-Art; leer = leaf+area
 * @returns {string}
 */
export function findInsertBeforeSiblingId(boxes, parentId, dragId, worldY, kind) {
    const pid = parentId == null || parentId === '' ? '' : String(parentId);
    const drag = String(dragId || '');
    const want = String(kind || '').toLowerCase();
    const siblings = (Array.isArray(boxes) ? boxes : [])
        .filter(function (b) {
            if (!b || !b.node || !b.node.id) return false;
            if (String(b.node.id) === drag) return false;
            const bp =
                b.parentAreaId == null || b.parentAreaId === '' ? '' : String(b.parentAreaId);
            if (bp !== pid) return false;
            if (want === 'leaf') return b.kind === 'leaf';
            if (want === 'area') return b.kind === 'area';
            return b.kind === 'leaf' || b.kind === 'area';
        })
        .sort(function (a, b) {
            return (a.y || 0) - (b.y || 0) || (a.x || 0) - (b.x || 0);
        });
    const y = Number(worldY);
    if (!Number.isFinite(y)) return '';
    for (let i = 0; i < siblings.length; i++) {
        const s = siblings[i];
        const mid = (Number(s.y) || 0) + (Number(s.h) || 0) / 2;
        if (y < mid) return String(s.node.id);
    }
    return '';
}

/**
 * Drop-Ziel aus Hit-Stack wählen (Area direkt oder Leaf → Parent-Area).
 * @param {Array<{ dropAreaId?: string|null, parentAreaId?: string|null }>} hits
 * @param {Set<string>|string[]|null|undefined} blockedIds
 * @returns {string}
 */
export function pickLandkarteDropAreaId(hits, blockedIds) {
    /** @type {Set<string>} */
    const blocked = new Set();
    if (blockedIds instanceof Set) {
        blockedIds.forEach(function (id) {
            if (id) blocked.add(String(id));
        });
    } else if (Array.isArray(blockedIds)) {
        blockedIds.forEach(function (id) {
            if (id) blocked.add(String(id));
        });
    }
    const list = Array.isArray(hits) ? hits : [];
    for (let i = 0; i < list.length; i++) {
        const hit = list[i] || {};
        const dropId = normStr(hit.dropAreaId);
        if (dropId && !blocked.has(dropId)) return dropId;
        const parentId = normStr(hit.parentAreaId);
        if (parentId && !blocked.has(parentId)) return parentId;
    }
    return '';
}

/**
 * Drop-Ziel unter dem Zeiger: Areas und überdeckende Leaves (data-parent-area-id).
 * Die gezogene Area selbst / ihre Kinder stehen in blockedIds.
 *
 * @param {number} clientX
 * @param {number} clientY
 * @param {Set<string>|string[]|null|undefined} blockedIds
 * @param {ParentNode|null|undefined} [root] nur Treffer innerhalb dieses Hosts
 * @returns {string}
 */
export function resolveLandkarteDropAreaId(clientX, clientY, blockedIds, root) {
    const stack =
        typeof document !== 'undefined' && typeof document.elementsFromPoint === 'function'
            ? document.elementsFromPoint(clientX, clientY)
            : typeof document !== 'undefined' && typeof document.elementFromPoint === 'function'
              ? [document.elementFromPoint(clientX, clientY)].filter(Boolean)
              : [];
    /** @type {Array<{ dropAreaId?: string|null, parentAreaId?: string|null }>} */
    const hits = [];
    for (let i = 0; i < stack.length; i++) {
        const el = stack[i];
        if (!el || typeof el.closest !== 'function') continue;
        if (root && !root.contains(el)) continue;
        const drop = el.closest('[data-drop-area]');
        const leaf = el.closest('[data-parent-area-id]');
        hits.push({
            dropAreaId:
                drop && (!root || root.contains(drop))
                    ? drop.getAttribute('data-area-id') || ''
                    : '',
            parentAreaId:
                leaf && (!root || root.contains(leaf))
                    ? leaf.getAttribute('data-parent-area-id') || ''
                    : ''
        });
    }
    return pickLandkarteDropAreaId(hits, blockedIds);
}

/**
 * @param {object} box
 * @param {number} x
 * @param {number} y
 */
function pointInLandkarteBox(box, x, y) {
    if (!box) return false;
    const bx = Number(box.x) || 0;
    const by = Number(box.y) || 0;
    const bw = Number(box.w) || 0;
    const bh = Number(box.h) || 0;
    return x >= bx && y >= by && x <= bx + bw && y <= by + bh;
}

/**
 * Tiefste Area unter Welt-Koordinaten (Geschwister statt großer Parent).
 * Liegt der Zeiger im Kopfbereich einer Area, gewinnt diese Area – damit man
 * befüllte Bereiche als Nest-Ziel treffen kann, ohne in ein Kind zu rutschen.
 * @param {object[]} boxes
 * @param {number} worldX
 * @param {number} worldY
 * @param {Set<string>|string[]|null|undefined} blockedIds
 * @param {object[]} [nodes] für Tiefensortierung
 * @returns {string}
 */
export function pickDeepestAreaAtPoint(boxes, worldX, worldY, blockedIds, nodes) {
    /** @type {Set<string>} */
    const blocked = new Set();
    if (blockedIds instanceof Set) {
        blockedIds.forEach(function (id) {
            if (id) blocked.add(String(id));
        });
    } else if (Array.isArray(blockedIds)) {
        blockedIds.forEach(function (id) {
            if (id) blocked.add(String(id));
        });
    }
    /** @type {Map<string, object>} */
    const byId = new Map();
    (Array.isArray(nodes) ? nodes : []).forEach(function (n) {
        if (n && n.id) byId.set(String(n.id), n);
    });
    function depthOf(id) {
        let d = 0;
        let walk = String(id || '');
        let guard = 0;
        while (walk && guard < 48) {
            guard += 1;
            const n = byId.get(walk);
            if (!n) break;
            const pid =
                n.parentId == null || n.parentId === '' ? '' : String(n.parentId);
            if (!pid) break;
            d += 1;
            walk = pid;
        }
        return d;
    }
    function headH(box) {
        return box && box.classNest ? CLASS_NEST_HEAD_H + 4 : HEAD_H + 4;
    }
    /** @type {object[]} */
    const hits = [];
    (Array.isArray(boxes) ? boxes : []).forEach(function (b) {
        if (!b || b.kind !== 'area' || !b.node || !b.node.id) return;
        const id = String(b.node.id);
        if (blocked.has(id)) return;
        if (!pointInLandkarteBox(b, worldX, worldY)) return;
        hits.push(b);
    });
    if (!hits.length) return '';
    // Kopf treffen = bewusst diesen (auch befüllten) Bereich als Parent wählen
    const headerHits = hits.filter(function (b) {
        const top = Number(b.y) || 0;
        return worldY <= top + headH(b);
    });
    const pool = headerHits.length ? headerHits : hits;
    pool.sort(function (a, b) {
        const dd = depthOf(b.node.id) - depthOf(a.node.id);
        if (dd) return dd;
        const aa = (Number(a.w) || 0) * (Number(a.h) || 0);
        const bb = (Number(b.w) || 0) * (Number(b.h) || 0);
        return aa - bb;
    });
    return String(pool[0].node.id);
}

/**
 * Nest-Drop inkl. Herausziehen: tiefstes Ziel, sonst stufenweise zum Ahnen, der den Punkt noch enthält.
 * @param {{
 *   boxes: object[],
 *   nodes: object[],
 *   worldX: number,
 *   worldY: number,
 *   dragId: string,
 *   blockedIds?: Set<string>|string[]|null,
 *   currentParentId?: string|null
 * }} opts
 * @returns {{ parentId: string|null, mode: 'nest'|'promote'|'root'|'keep' }}
 */
export function resolveLandkarteNestParent(opts) {
    const o = opts || {};
    const dragId = normStr(o.dragId);
    const curParent =
        o.currentParentId == null || o.currentParentId === ''
            ? ''
            : String(o.currentParentId);
    /** @type {Set<string>} */
    const blocked = new Set();
    if (o.blockedIds instanceof Set) {
        o.blockedIds.forEach(function (id) {
            if (id) blocked.add(String(id));
        });
    } else if (Array.isArray(o.blockedIds)) {
        o.blockedIds.forEach(function (id) {
            if (id) blocked.add(String(id));
        });
    }
    if (dragId) blocked.add(dragId);

    /** @type {Map<string, object>} */
    const byId = new Map();
    (Array.isArray(o.nodes) ? o.nodes : []).forEach(function (n) {
        if (n && n.id) byId.set(String(n.id), n);
    });
    /** Ist ancestorId ein Vorfahre von nodeId? */
    function isAncestor(ancestorId, nodeId) {
        const anc = String(ancestorId || '');
        let walk = String(nodeId || '');
        let guard = 0;
        while (walk && guard < 48) {
            guard += 1;
            const n = byId.get(walk);
            const pid =
                n && n.parentId != null && n.parentId !== '' ? String(n.parentId) : '';
            if (!pid) return false;
            if (pid === anc) return true;
            walk = pid;
        }
        return false;
    }

    const deepest = pickDeepestAreaAtPoint(
        o.boxes,
        o.worldX,
        o.worldY,
        blocked,
        o.nodes
    );
    if (deepest && deepest !== dragId) {
        if (deepest === curParent) {
            return { parentId: curParent || null, mode: 'keep' };
        }
        // Auf Ahnen des bisherigen Parents → herausziehen (wird Geschwister)
        if (curParent && isAncestor(deepest, curParent)) {
            return { parentId: deepest, mode: 'promote' };
        }
        return { parentId: deepest, mode: 'nest' };
    }

    if (!curParent) {
        return { parentId: null, mode: 'root' };
    }

    const parentBox = (Array.isArray(o.boxes) ? o.boxes : []).find(function (b) {
        return b && b.kind === 'area' && b.node && String(b.node.id) === curParent;
    });
    const stillInParent = pointInLandkarteBox(parentBox, o.worldX, o.worldY);
    if (stillInParent) {
        return { parentId: curParent, mode: 'keep' };
    }

    // Außerhalb aller Treffer / bisherigen Parents: Ahnen-Kette hoch
    let walk = curParent;
    let guard = 0;
    while (walk && guard < 48) {
        guard += 1;
        const n = byId.get(walk);
        const gp =
            n && n.parentId != null && n.parentId !== '' ? String(n.parentId) : '';
        if (!gp) {
            return { parentId: null, mode: 'root' };
        }
        const gpBox = (Array.isArray(o.boxes) ? o.boxes : []).find(function (b) {
            return b && b.kind === 'area' && b.node && String(b.node.id) === gp;
        });
        if (pointInLandkarteBox(gpBox, o.worldX, o.worldY)) {
            return { parentId: gp, mode: 'promote' };
        }
        walk = gp;
    }
    return { parentId: null, mode: 'root' };
}

/**
 * Anzahl direkter Inhalte (Unterbereiche + Leaves, ohne Personen-Knoten).
 * @param {object} base
 * @param {string} areaId
 */
function countAreaMapContents(base, areaId) {
    return childrenOf(base, areaId).filter(function (n) {
        return n && n.type !== NODE.PERSON;
    }).length;
}

/**
 * Rekursiv: Area inkl. aller Unter-Areas/Leaves messen und platzieren.
 * Eltern-Rechteck wächst, bis alle Kinder optisch enthalten sind.
 * Eingeklappte Bereiche: nur Kopfzeile (Kinder werden nicht gezeichnet).
 *
 * @param {object} base
 * @param {object} area
 * @param {number} absX
 * @param {number} absY
 * @param {string|null} parentAreaId
 * @param {{ forceExpand?: boolean, peopleCountForArea?: (area: object) => number }} [opts]
 * @returns {{ boxes: object[], w: number, h: number }}
 */
function layoutAreaTree(base, area, absX, absY, parentAreaId, opts) {
    const forceExpand = !!(opts && opts.forceExpand);
    /** Pfad-Guard gegen parentId-Zyklen / extreme Verschachtelung */
    /** @type {Set<string>} */
    const pathGuard =
        opts && opts._layoutPath instanceof Set ? opts._layoutPath : new Set();
    const areaId = area && area.id != null ? String(area.id) : '';
    if (areaId && pathGuard.has(areaId)) {
        return { boxes: [], w: AREA_W, h: HEAD_H };
    }
    if (pathGuard.size >= 32) {
        return { boxes: [], w: AREA_W, h: HEAD_H };
    }
    /** @type {Set<string>} */
    const nextPath = new Set(pathGuard);
    if (areaId) nextPath.add(areaId);
    const treeOptsBase = Object.assign({}, opts || {}, { _layoutPath: nextPath });
    const contentCount = countAreaMapContents(base, area.id);
    const isRoot = parentAreaId == null || parentAreaId === '';
    const userW = parseSchulOrgCanvasCoord(area && area.w);
    const userH = parseSchulOrgCanvasCoord(area && area.h);
    // Feste Fensterbreite nur für Root-Areas mit gespeichertem w (exakt)
    const fixedW =
        isRoot && userW != null
            ? Math.max(AREA_RESIZE_MIN_W, Math.min(2400, userW))
            : null;
    // Verschachtelte Bereiche: geerbte Maximalbreite (nicht zwingend ausfüllen)
    const constrainW = parseSchulOrgCanvasCoord(opts && opts.constrainW);
    const maxW =
        fixedW != null
            ? fixedW
            : constrainW != null
              ? Math.max(160, constrainW)
              : null;
    const forcedW = fixedW;
    const baseW =
        forcedW != null
            ? forcedW
            : maxW != null
              ? Math.max(160, Math.min(AREA_W, maxW))
              : AREA_W;
    /** Breite darf in Reflow nicht über maxW hinauswachsen */
    const widthLocked = maxW != null;
    const childConstrain = maxW != null ? Math.max(120, maxW - AREA_PAD * 2) : null;
    const childOpts =
        childConstrain != null
            ? Object.assign({}, treeOptsBase, { constrainW: childConstrain })
            : treeOptsBase;

    if (area.collapsed === true && !forceExpand) {
        const w = forcedW != null ? forcedW : maxW != null ? Math.min(AREA_W, maxW) : AREA_W;
        const h = HEAD_H;
        return {
            boxes: [
                {
                    node: area,
                    x: absX,
                    y: absY,
                    w: w,
                    h: h,
                    kind: 'area',
                    parentAreaId: parentAreaId,
                    collapsed: true,
                    contentCount: contentCount,
                    resizable: isRoot,
                    widthLocked: widthLocked
                }
            ],
            w: w,
            h: h
        };
    }

    const leaves = childrenOf(base, area.id).filter(function (n) {
        return !isAreaNode(n) && n.type !== NODE.PERSON;
    });
    const nested = childrenOf(base, area.id).filter(isAreaNode);
    /** @type {object[]} */
    const classNests = [];
    /** @type {object[]} */
    const plainLeaves = [];
    leaves.forEach(function (leaf) {
        // Legacy-Klassen-Leaves (falls noch vorhanden) als Nest
        if (isClassNode(leaf) && !isAreaNode(leaf)) {
            classNests.push(leaf);
            return;
        }
        // Kursteams: in der Übersicht ausgeblendet; im Klassen-Fokus (forceExpand) sichtbar
        if (isKursteamNode(leaf)) {
            if (opts && opts.forceExpand && isClassNode(area)) {
                plainLeaves.push(leaf);
            }
            return;
        }
        plainLeaves.push(leaf);
    });
    /** @type {object[]} */
    const childBoxes = [];
    let cursorY = contentTop(absY);
    let maxRight = absX + baseW;
    let maxBottom = absY + HEAD_H + CONTENT_GAP + 24;
    const innerMaxRight = absX + baseW - AREA_PAD;

    const nestJahrgaengeSideBySide =
        nested.length > 0 &&
        ((normStr(area.refKind) === 'sammelgruppe' &&
            (normStr(area.refKey) === 'schueler' ||
                normStr(area.refKey) === 'eltern' ||
                normStr(area.refKey) === 'klassenchats')) ||
            nested.every(function (ch) {
                return (
                    isJahrgangArea(ch) ||
                    isElternJahrgangArea(ch) ||
                    isKlassenchatJahrgangArea(ch) ||
                    normStr(ch.refKind).toLowerCase() === 'jahrgang' ||
                    /^sec_jahrgang_/i.test(normStr(ch.id)) ||
                    /^sec_eltern_jahrgang_/i.test(normStr(ch.id)) ||
                    /^sec_klassenchat_jahrgang_/i.test(normStr(ch.id))
                );
            }));
    const nestArgesSideBySide =
        nested.length > 0 &&
        (normStr(area.id) === 'sec_arge' ||
            nested.every(function (ch) {
                return isArgeAreaNode(ch) || isArgeLeiterArea(ch);
            }));
    // Mit Maximalbreite: alle Unterbereiche fließen (Wrap), sonst Jahrgangs-/ARGE-Boards
    const flowNested =
        nested.length > 0 && (nestJahrgaengeSideBySide || nestArgesSideBySide || maxW != null);
    const wrapNested = flowNested && maxW != null;

    // Unterbereiche zuerst – bei fester/enger Breite Zeilen umbrechen (Fenster-Inhalt)
    if (flowNested) {
        let colX = absX + AREA_PAD;
        let rowY = cursorY;
        let rowH = 0;
        nested.forEach(function (child) {
            // Probe-Layout mit derselben Kind-Einschränkung
            const probe = layoutAreaTree(base, child, 0, 0, area.id, childOpts);
            const childW = probe.w;
            if (wrapNested && colX > absX + AREA_PAD && colX + childW > innerMaxRight + 0.5) {
                colX = absX + AREA_PAD;
                rowY += rowH + NEST_AREA_GAP;
                rowH = 0;
            }
            const nestedLayout = layoutAreaTree(base, child, colX, rowY, area.id, childOpts);
            childBoxes.push.apply(childBoxes, nestedLayout.boxes);
            rowH = Math.max(rowH, nestedLayout.h);
            maxRight = Math.max(maxRight, colX + nestedLayout.w + AREA_PAD);
            maxBottom = Math.max(maxBottom, rowY + nestedLayout.h + AREA_PAD);
            colX += nestedLayout.w + NEST_AREA_GAP;
        });
        cursorY = rowY + rowH + NEST_AREA_GAP;
    } else {
        nested.forEach(function (child) {
            const rx = parseSchulOrgCanvasCoord(child.x);
            const ry = parseSchulOrgCanvasCoord(child.y);
            const useFree = rx != null && ry != null && rx <= 400 && ry <= 300;
            const cx = useFree ? absX + AREA_PAD + rx : absX + AREA_PAD;
            const cy = useFree ? contentTop(absY) + ry : cursorY;
            const nestedLayout = layoutAreaTree(base, child, cx, cy, area.id, childOpts);
            childBoxes.push.apply(childBoxes, nestedLayout.boxes);
            maxRight = Math.max(maxRight, cx + nestedLayout.w + AREA_PAD);
            maxBottom = Math.max(maxBottom, cy + nestedLayout.h + AREA_PAD);
            cursorY = Math.max(cursorY, cy + nestedLayout.h + NEST_AREA_GAP);
        });
    }

    classNests.forEach(function (cls) {
        const nest = layoutClassWithKursteams(
            base,
            cls,
            absX + AREA_PAD,
            cursorY,
            area.id,
            childConstrain != null ? childConstrain : undefined
        );
        childBoxes.push.apply(childBoxes, nest.boxes);
        maxRight = Math.max(maxRight, absX + AREA_PAD + nest.w + AREA_PAD);
        maxBottom = Math.max(maxBottom, cursorY + nest.h + AREA_PAD);
        cursorY += nest.h + NEST_AREA_GAP;
    });

    const leafW = Math.max(120, baseW - AREA_PAD * 2);
    plainLeaves.forEach(function (leaf) {
        childBoxes.push({
            node: leaf,
            x: absX + AREA_PAD,
            y: cursorY,
            w: leafW,
            h: LEAF_H,
            kind: 'leaf',
            parentAreaId: area.id
        });
        cursorY += LEAF_H + NEST_GAP;
        maxBottom = Math.max(maxBottom, cursorY);
    });

    // Personen-Chips wie bei Lehrkräften – auch in IT-Team / Verwaltungsbereichen
    let previewCount = 0;
    let previewH = 0;
    let wantsPeoplePreview = false;
    if (areaAllowsPeoplePreview(area)) {
        const sammel = sammelgruppeKeyForArea(area);
        if (opts && typeof opts.peopleCountForArea === 'function') {
            previewCount = Math.max(0, Number(opts.peopleCountForArea(area)) || 0);
        } else if (sammel === 'lehrer') {
            previewCount = PREVIEW_CHIP_MAX;
        }
        previewH = peoplePreviewStripHeight(previewCount);
        wantsPeoplePreview = previewH > 0;
        if (wantsPeoplePreview) {
            maxBottom = Math.max(maxBottom, cursorY + previewH + AREA_PAD);
            maxRight = Math.max(maxRight, absX + baseW);
        }
    }

    // Sichtbare Innen-Inhalte (nicht: Kursteams, die in der Übersicht ausgeblendet sind).
    // Steuert CSS has-leaves: absolute Preview nur wenn wirklich Platz für Leaves nötig ist.
    const visibleInnerCount = nested.length + classNests.length + plainLeaves.length;

    childBoxes.forEach(function (b) {
        maxRight = Math.max(maxRight, (b.x || 0) + (b.w || 0) + AREA_PAD);
        maxBottom = Math.max(maxBottom, (b.y || 0) + (b.h || 0) + AREA_PAD);
    });

    const contentW = Math.max(baseW, Math.ceil(maxRight - absX));
    const contentH = Math.max(HEAD_H + CONTENT_GAP + 36, Math.ceil(maxBottom - absY));
    // Root-Resize: exakte Breite; verschachtelt: Inhalt, aber nie breiter als maxW
    const w =
        forcedW != null
            ? forcedW
            : maxW != null
              ? Math.max(160, Math.min(contentW, maxW))
              : contentW;
    const h =
        isRoot && userH != null
            ? Math.max(contentH, Math.max(AREA_RESIZE_MIN_H, userH))
            : contentH;
    childBoxes.forEach(function (b) {
        if (b.kind === 'leaf' && b.parentAreaId === area.id) {
            b.w = Math.max(120, w - AREA_PAD * 2);
        }
    });

    return {
        boxes: [
            {
                node: area,
                x: absX,
                y: absY,
                w: w,
                h: h,
                kind: 'area',
                parentAreaId: parentAreaId,
                peoplePreview: wantsPeoplePreview,
                previewH: previewH,
                previewCount: previewCount,
                collapsed: false,
                contentCount: contentCount,
                visibleInnerCount: visibleInnerCount,
                resizable: isRoot,
                contentH: contentH,
                widthLocked: widthLocked
            }
        ].concat(childBoxes),
        w: w,
        h: h
    };
}

/**
 * Nach dem Packen: Geschwister unter demselben Parent vertikal entflechten
 * (keine Überlappung von Bereichen/Personen-Kacheln).
 * @param {object[]} boxes
 */
export function resolveLandkarteOverlaps(boxes) {
    const list = Array.isArray(boxes) ? boxes : [];
    /** @type {Map<string, object[]>} */
    const byParent = new Map();
    list.forEach(function (b) {
        if (!b || b.kind === 'person-chip') return;
        const pid = b.parentAreaId == null ? '' : String(b.parentAreaId);
        if (!byParent.has(pid)) byParent.set(pid, []);
        byParent.get(pid).push(b);
    });

    function areaBox(id) {
        return list.find(function (b) {
            return b.kind === 'area' && b.node && b.node.id === id;
        });
    }

    function isJahrgangBox(b) {
        const n = b && b.node;
        if (!n) return false;
        const rk = normStr(n.refKind).toLowerCase();
        return (
            rk === 'jahrgang' ||
            rk === 'elternjahrgang' ||
            rk === 'klassenchatjahrgang' ||
            isJahrgangArea(n) ||
            isElternJahrgangArea(n) ||
            isKlassenchatJahrgangArea(n) ||
            /^sec_jahrgang_/i.test(normStr(n.id)) ||
            /^sec_eltern_jahrgang_/i.test(normStr(n.id)) ||
            /^sec_klassenchat_jahrgang_/i.test(normStr(n.id))
        );
    }

    function isWidthLocked(box) {
        if (!box) return false;
        if (box.widthLocked) return true;
        const isRoot = box.parentAreaId == null || box.parentAreaId === '';
        if (!isRoot || box.classNest) return false;
        return parseSchulOrgCanvasCoord(box.node && box.node.w) != null;
    }

    /**
     * Bereich und Nachfahren auf maxW begrenzen (nur schrumpfen, nie aufblasen).
     * @param {object} areaBox
     * @param {number} maxW
     */
    function fitSubtreeToWidth(areaBox, maxW) {
        if (!areaBox || !areaBox.node || maxW == null || !Number.isFinite(maxW)) return;
        const edgePad = areaBox.classNest ? 10 : AREA_PAD;
        const cap = Math.max(areaBox.classNest ? 120 : 160, maxW);
        const cur = Number(areaBox.w) || cap;
        areaBox.w = Math.min(cur, cap);
        areaBox.widthLocked = true;
        const inner = Math.max(120, areaBox.w - edgePad * 2);
        const kids = byParent.get(String(areaBox.node.id)) || [];
        kids.forEach(function (ch) {
            if (ch.kind === 'leaf') {
                ch.w = Math.max(120, Math.min(Number(ch.w) || inner, inner));
                return;
            }
            if (ch.kind === 'area') {
                fitSubtreeToWidth(ch, inner);
            }
        });
    }

    function shiftSubtree(all, root, dx, dy) {
        if (!root || (!dx && !dy)) return;
        /** @type {Set<string>} */
        const ids = new Set();
        if (root.node && root.node.id) ids.add(String(root.node.id));
        root.x += dx;
        root.y += dy;
        let grew = true;
        let guard = 0;
        while (grew && guard < 50) {
            grew = false;
            guard += 1;
            all.forEach(function (b) {
                if (!b.parentAreaId || !ids.has(String(b.parentAreaId))) return;
                if (b._shifted) return;
                b.x += dx;
                b.y += dy;
                b._shifted = true;
                if (b.kind === 'area' && b.node && b.node.id) {
                    ids.add(String(b.node.id));
                    grew = true;
                }
            });
        }
        all.forEach(function (b) {
            delete b._shifted;
        });
    }

    function reflowUnder(parentId) {
        const kids = (byParent.get(parentId == null ? '' : String(parentId)) || []).slice();
        if (!kids.length) return;
        const nested = kids
            .filter(function (b) {
                return b.kind === 'area';
            })
            .sort(function (a, b) {
                return (a.y || 0) - (b.y || 0) || (a.x || 0) - (b.x || 0);
            });
        const leaves = kids
            .filter(function (b) {
                return b.kind === 'leaf';
            })
            .sort(function (a, b) {
                return (a.y || 0) - (b.y || 0);
            });

        // Zuerst Kinder (deren Innenleben), dann Geschwister unter diesem Parent
        nested.forEach(function (ch) {
            if (ch.node && ch.node.id) reflowUnder(ch.node.id);
        });

        const parent = parentId ? areaBox(parentId) : null;
        const parentNode = parent && parent.node;
        const parentSammel = parentNode ? sammelgruppeKeyForArea(parentNode) : '';
        const edgePad = parent && parent.classNest ? 10 : AREA_PAD;
        const locked = parent ? isWidthLocked(parent) : false;
        const targetW = parent ? Number(parent.w) || 0 : 0;
        const maxChildW =
            parent && targetW > 0 ? Math.max(120, targetW - edgePad * 2) : null;

        // Unterbereiche an Elternbreite anpassen, bevor wir sie stapeln/umfließen
        if (maxChildW != null) {
            nested.forEach(function (ch) {
                if (ch.w > maxChildW + 0.5) fitSubtreeToWidth(ch, maxChildW);
            });
        }

        const flowSideBySide =
            nested.length > 0 &&
            ((parentSammel === 'schueler' ||
                parentSammel === 'eltern' ||
                parentSammel === 'klassenchats') ||
                nested.every(isJahrgangBox) ||
                (locked && nested.length > 1));

        let cursorY = parent
            ? parent.classNest
                ? parent.y + CLASS_NEST_HEAD_H + 8
                : contentTop(parent.y)
            : kids[0]
              ? kids[0].y
              : 0;
        const baseX = parent ? parent.x + edgePad : null;
        const wrapLimit =
            parent && locked && !parent.classNest && targetW > 0
                ? parent.x + targetW - edgePad
                : null;

        if (flowSideBySide) {
            let colX = baseX != null ? baseX : nested[0].x;
            let rowY = cursorY;
            let rowH = 0;
            nested.forEach(function (ch) {
                if (
                    wrapLimit != null &&
                    colX > (baseX != null ? baseX : colX) &&
                    colX + ch.w > wrapLimit + 0.5
                ) {
                    colX = baseX != null ? baseX : nested[0].x;
                    rowY += rowH + NEST_AREA_GAP;
                    rowH = 0;
                }
                const dy = rowY - ch.y;
                const dx = colX - ch.x;
                if (dy || dx) shiftSubtree(list, ch, dx, dy);
                rowH = Math.max(rowH, ch.h);
                colX = ch.x + ch.w + NEST_AREA_GAP;
            });
            cursorY = rowY + rowH + NEST_AREA_GAP;
        } else {
            nested.forEach(function (ch) {
                const dy = cursorY - ch.y;
                const dx = baseX != null ? baseX - ch.x : 0;
                if (dy || dx) shiftSubtree(list, ch, dx, dy);
                cursorY = ch.y + ch.h + NEST_AREA_GAP;
            });
        }
        leaves.forEach(function (leaf) {
            if (baseX != null) leaf.x = baseX;
            leaf.y = cursorY;
            leaf.h = LEAF_H;
            if (parent) {
                leaf.w = Math.max(120, (Number(parent.w) || AREA_W) - edgePad * 2);
            }
            cursorY = leaf.y + LEAF_H + NEST_GAP;
        });

        if (parent) {
            // Eingeklappt: nur Kopfzeile – nicht wieder aufblasen
            if (parent.collapsed || (parent.node && parent.node.collapsed)) {
                if (!locked) parent.w = Math.max(AREA_W, parent.w || AREA_W);
                parent.h = HEAD_H;
                return;
            }
            const minHead = parent.classNest
                ? CLASS_NEST_HEAD_H + 12
                : HEAD_H + CONTENT_GAP + 36;
            let maxRight = parent.x + (parent.classNest ? parent.w : edgePad * 2);
            let maxBottom =
                parent.y + (parent.classNest ? CLASS_NEST_HEAD_H + 12 : HEAD_H + CONTENT_GAP + 24);
            kids.forEach(function (b) {
                maxRight = Math.max(maxRight, b.x + b.w + edgePad);
                maxBottom = Math.max(maxBottom, b.y + b.h + edgePad);
            });
            const previewH = Number(parent.previewH) || 0;
            if (previewH > 0) {
                maxBottom = Math.max(maxBottom, cursorY + previewH + edgePad);
            }
            if (locked) {
                // Breite bleibt vom Layout/Resize – Inhalt darf sie nicht aufblasen
                const fixedParentW = parseSchulOrgCanvasCoord(parent.node && parent.node.w);
                if (fixedParentW != null && !parent.classNest) {
                    parent.w = Math.max(AREA_RESIZE_MIN_W, fixedParentW);
                }
                // Nochmal strikt: Nachfahren in die finale Breite pressen
                fitSubtreeToWidth(parent, parent.w);
            } else {
                parent.w = Math.max(
                    parent.classNest ? 160 : AREA_W,
                    Math.ceil(maxRight - parent.x)
                );
            }
            const contentNeedH = Math.max(minHead, Math.ceil(maxBottom - parent.y));
            const fixedParentH = parseSchulOrgCanvasCoord(parent.node && parent.node.h);
            parent.h =
                fixedParentH != null && !parent.classNest
                    ? Math.max(contentNeedH, fixedParentH)
                    : contentNeedH;
            leaves.forEach(function (leaf) {
                leaf.w = Math.max(120, parent.w - edgePad * 2);
            });
        }
    }

    // Nur Inhalte innerhalb von Root-Areas neu stapeln (Root-Positionen bleiben frei)
    list.forEach(function (b) {
        if (b.kind === 'area' && b.node && (b.parentAreaId == null || b.parentAreaId === '')) {
            reflowUnder(b.node.id);
        }
    });
    // Rollen fest in Spalten einrasten
    snapLandkarteLeavesInAreas(list);

    // Root-Areas bewusst nicht entflechten: Nutzer-Anordnung (x/y) ist kanonisch und persistent.
    return list;
}

function normStr(v) {
    return String(v ?? '').trim();
}

/**
 * Ob ein gespeicherter Viewport vorliegt (nicht nur Schema-Default 0/0/1).
 * @param {{ x?: number, y?: number, k?: number }|null|undefined} view
 */
export function hasSavedLandkarteView(view) {
    if (!view || typeof view !== 'object') return false;
    const x = Number(view.x);
    const y = Number(view.y);
    const k = Number(view.k);
    if (!Number.isFinite(k) || k <= 0) return false;
    // Explizit gespeichert und vom Neutralwert (0,0,1) verschieden
    if (Math.abs(k - 1) > 0.02) return true;
    if (Number.isFinite(x) && Math.abs(x) > 1) return true;
    if (Number.isFinite(y) && Math.abs(y) > 1) return true;
    return false;
}

/**
 * Absolute Positionen für Root-Areas (x/y aus Node) + verschachtelte Kinder relativ.
 * @param {object} layout
 * @param {{ peopleCountForArea?: (area: object) => number }} [layoutOpts]
 */
export function computeLandkarteLayout(layout, layoutOpts) {
    const base = normalizeSchulOrganigramm(layout);
    const treeOpts =
        layoutOpts && typeof layoutOpts === 'object'
            ? {
                  peopleCountForArea:
                      typeof layoutOpts.peopleCountForArea === 'function'
                          ? layoutOpts.peopleCountForArea
                          : undefined
              }
            : {};
    const roots = childrenOf(base, null).filter(isAreaNode);
    /** @type {object[]} */
    const boxes = [];
    let maxX = 800;
    let maxY = 600;

    roots.forEach(function (area, idx) {
        const ax = Number.isFinite(Number(area.x)) ? Number(area.x) : 40 + (idx % 3) * 480;
        const ay = Number.isFinite(Number(area.y)) ? Number(area.y) : 40 + Math.floor(idx / 3) * 280;
        const tree = layoutAreaTree(base, area, ax, ay, null, treeOpts);
        boxes.push.apply(boxes, tree.boxes);
        maxX = Math.max(maxX, ax + tree.w + 80);
        maxY = Math.max(maxY, ay + tree.h + 80);
    });

    childrenOf(base, null)
        .filter(function (n) {
            return !isAreaNode(n) && n.type !== NODE.PERSON;
        })
        .forEach(function (n, i) {
            boxes.push({
                node: n,
                x: 40,
                y: maxY + i * 44,
                w: AREA_W,
                h: LEAF_H,
                kind: 'leaf',
                parentAreaId: null
            });
            maxY += 50;
        });

    resolveLandkarteOverlaps(boxes);
    boxes.forEach(function (b) {
        maxX = Math.max(maxX, (b.x || 0) + (b.w || 0) + 80);
        maxY = Math.max(maxY, (b.y || 0) + (b.h || 0) + 80);
    });

    return { boxes: boxes, width: maxX, height: maxY, layout: base };
}

/**
 * Leaves je Bereich links bündig untereinander – „snap to grid“.
 * @param {object[]} boxes
 */
export function snapLandkarteLeavesInAreas(boxes) {
    const list = Array.isArray(boxes) ? boxes : [];
    /** @type {Map<string, object>} */
    const areaById = new Map();
    /** @type {Map<string, object[]>} */
    const leavesByParent = new Map();
    /** @type {Map<string, object[]>} */
    const nestedByParent = new Map();
    list.forEach(function (b) {
        if (!b || !b.node) return;
        if (b.kind === 'area' && b.node.id) {
            areaById.set(String(b.node.id), b);
            const pid =
                b.parentAreaId == null || b.parentAreaId === '' ? '' : String(b.parentAreaId);
            if (pid) {
                if (!nestedByParent.has(pid)) nestedByParent.set(pid, []);
                nestedByParent.get(pid).push(b);
            }
        }
        if (b.kind === 'leaf' && b.parentAreaId) {
            const pid = String(b.parentAreaId);
            if (!leavesByParent.has(pid)) leavesByParent.set(pid, []);
            leavesByParent.get(pid).push(b);
        }
    });
    leavesByParent.forEach(function (leaves, pid) {
        const parent = areaById.get(pid);
        if (!parent || parent.collapsed || (parent.node && parent.node.collapsed)) return;
        const edgePad = parent.classNest ? 10 : AREA_PAD;
        const baseX = parent.x + edgePad;
        let cursorY = parent.classNest
            ? parent.y + CLASS_NEST_HEAD_H + 8
            : contentTop(parent.y);
        const nested = (nestedByParent.get(pid) || []).slice().sort(function (a, b) {
            return (a.y || 0) - (b.y || 0);
        });
        nested.forEach(function (ch) {
            cursorY = Math.max(cursorY, (ch.y || 0) + (ch.h || 0) + NEST_AREA_GAP);
        });
        leaves
            .slice()
            .sort(function (a, b) {
                const ao = a.node && typeof a.node.order === 'number' ? a.node.order : 0;
                const bo = b.node && typeof b.node.order === 'number' ? b.node.order : 0;
                // Primär order (nach Drop), dann Y – verhindert Überlappung bei gleichen Y
                return ao - bo || (a.y || 0) - (b.y || 0);
            })
            .forEach(function (leaf) {
                leaf.x = baseX;
                leaf.y = cursorY;
                leaf.h = LEAF_H;
                leaf.w = Math.max(120, (parent.w || AREA_W) - edgePad * 2);
                cursorY = leaf.y + LEAF_H + NEST_GAP;
            });
        const previewH = Number(parent.previewH) || 0;
        let maxBottom = cursorY + (previewH > 0 ? previewH + edgePad : edgePad);
        nested.forEach(function (ch) {
            maxBottom = Math.max(maxBottom, (ch.y || 0) + (ch.h || 0) + edgePad);
        });
        const minHead = parent.classNest
            ? CLASS_NEST_HEAD_H + 12
            : HEAD_H + CONTENT_GAP + 36;
        const contentNeedH = Math.max(minHead, Math.ceil(maxBottom - parent.y));
        const fixedParentW = parseSchulOrgCanvasCoord(parent.node && parent.node.w);
        if (fixedParentW != null && !parent.classNest) {
            parent.w = Math.max(AREA_RESIZE_MIN_W, fixedParentW);
            leaves.forEach(function (leaf) {
                leaf.w = Math.max(120, parent.w - edgePad * 2);
            });
        }
        const fixedParentH = parseSchulOrgCanvasCoord(parent.node && parent.node.h);
        parent.h =
            fixedParentH != null && !parent.classNest
                ? Math.max(contentNeedH, fixedParentH)
                : contentNeedH;
    });
}

/**
 * In-Canvas-Zoom: Fokus-Knoten wächst an seiner Position, Inhalt (Kinder/Personen) liegt darin.
 * Andere Knoten bleiben als Kontext (abgeblendet) auf derselben Landkarte.
 *
 * @param {object} layout
 * @param {object} focusNode
 * @param {object[]} people
 * @param {object[]} [childNodes]
 */
/**
 * Alle Nachfahren-IDs (ohne Personen) für Fokus-Hide sammeln.
 * @param {object} base
 * @param {string} rootId
 * @param {Set<string>} into
 */
function collectFocusHideIds(base, rootId, into) {
    childrenOf(base, rootId).forEach(function (ch) {
        if (!ch || ch.type === NODE.PERSON) return;
        const id = String(ch.id || '');
        if (!id || into.has(id)) return;
        into.add(id);
        if (isAreaNode(ch) || isClassNode(ch)) collectFocusHideIds(base, id, into);
    });
}

export function computeFocusLayout(layout, focusNode, people, childNodes, layoutOpts) {
    const overview = computeLandkarteLayout(layout, layoutOpts);
    const peopleList = Array.isArray(people) ? people : [];
    const kids = Array.isArray(childNodes) ? childNodes : [];
    const focusId = focusNode && focusNode.id ? String(focusNode.id) : '';
    const origin =
        overview.boxes.find(function (b) {
            return b.node && b.node.id === focusId;
        }) || null;
    const ox = origin ? Number(origin.x) : 40;
    const oy = origin ? Number(origin.y) : 40;

    const isLehrerFocus =
        normStr(focusNode && focusNode.refKind) === 'sammelgruppe' &&
        normStr(focusNode && focusNode.refKey) === 'lehrer';
    const leadPeople = peopleList.filter(function (p) {
        return normStr(p && p.kind).toLowerCase() === 'head';
    });
    const memberPeople = peopleList.filter(function (p) {
        return normStr(p && p.kind).toLowerCase() !== 'head';
    });
    const peopleCols = isLehrerFocus ? 1 : FOCUS_COLS;
    const chipGapX = 14;
    const chipGapY = 12;
    const childGapX = 14;
    const childGapY = 12;
    const leadChipW = Math.max(FOCUS_CHIP_W + 40, 280);
    /** @type {object[]} */
    const innerBoxes = [];
    /** @type {object} */
    let expandBox;

    // Klassen-Zoom: links Klassenvorstand + Schüler, rechts Unterrichtsteams
    if (isClassNode(focusNode)) {
        const focusTopPad = 10;
        const colGap = 18;
        const hasPeople = !!(leadPeople.length || memberPeople.length);
        const hasKids = kids.length > 0;
        const leftPeopleCols =
            hasPeople && hasKids
                ? memberPeople.length > 6
                    ? 2
                    : 1
                : peopleCols;
        const leftInnerW = Math.max(
            leadPeople.length ? Math.min(leadChipW + 80, 360) : 0,
            leftPeopleCols > 0
                ? leftPeopleCols * FOCUS_CHIP_W + (leftPeopleCols - 1) * chipGapX
                : 0,
            hasPeople ? 240 : 0
        );
        const rightLeafMinW = 300;
        const contentW =
            hasPeople && hasKids
                ? leftInnerW + colGap + rightLeafMinW
                : Math.max(
                      leftInnerW,
                      hasKids ? rightLeafMinW : 0,
                      isLehrerFocus ? FOCUS_CHIP_W + 40 : 400
                  );
        const EXPAND_W = Math.max(EXPAND_MIN_W, contentW + EXPAND_INNER_PAD * 2);
        const usableInner = EXPAND_W - EXPAND_INNER_PAD * 2;
        let leftColW = hasPeople ? leftInnerW : 0;
        let rightColW = 0;
        if (hasPeople && hasKids) {
            rightColW = Math.max(rightLeafMinW, usableInner - leftColW - colGap);
        } else if (hasKids) {
            rightColW = usableInner;
        } else {
            leftColW = usableInner;
        }

        const leftX = ox + EXPAND_INNER_PAD;
        const rightX = hasPeople && hasKids ? leftX + leftColW + colGap : leftX;
        let leftCursorY = oy + HEAD_H + focusTopPad;
        let rightCursorY = oy + HEAD_H + focusTopPad;

        if (hasPeople) {
            const leadW = Math.min(leadChipW + 80, leftColW);
            leadPeople.forEach(function (p, i) {
                innerBoxes.push({
                    person: p,
                    x: leftX,
                    y: leftCursorY + i * (FOCUS_CHIP_H + chipGapY),
                    w: leadW,
                    h: FOCUS_CHIP_H,
                    kind: 'person-chip',
                    parentAreaId: focusId,
                    leadSeat: true
                });
            });
            if (leadPeople.length) {
                leftCursorY += leadPeople.length * (FOCUS_CHIP_H + chipGapY) + 8;
            }
            const memberChipW =
                leftPeopleCols <= 1
                    ? leftColW
                    : Math.max(
                          FOCUS_CHIP_W,
                          Math.floor(
                              (leftColW - (leftPeopleCols - 1) * chipGapX) / leftPeopleCols
                          )
                      );
            memberPeople.forEach(function (p, i) {
                const col = i % leftPeopleCols;
                const row = Math.floor(i / leftPeopleCols);
                innerBoxes.push({
                    person: p,
                    x: leftX + col * (memberChipW + chipGapX),
                    y: leftCursorY + row * (FOCUS_CHIP_H + chipGapY),
                    w: memberChipW,
                    h: FOCUS_CHIP_H,
                    kind: 'person-chip',
                    parentAreaId: focusId
                });
            });
            if (memberPeople.length) {
                leftCursorY +=
                    Math.ceil(memberPeople.length / leftPeopleCols) * (FOCUS_CHIP_H + chipGapY);
            }
        }

        if (hasKids) {
            kids.forEach(function (n, i) {
                innerBoxes.push({
                    node: n,
                    x: rightX,
                    y: rightCursorY + i * (LEAF_H + childGapY),
                    w: rightColW,
                    h: LEAF_H,
                    kind: isAreaNode(n) ? 'area' : 'leaf',
                    focusChild: true,
                    parentAreaId: focusId
                });
            });
            rightCursorY += kids.length * (LEAF_H + childGapY);
        } else if (!hasPeople) {
            leftCursorY += 40;
        }

        const contentBottom = Math.max(leftCursorY, rightCursorY);
        expandBox = {
            node: focusNode,
            x: ox,
            y: oy,
            w: EXPAND_W,
            h: Math.max(160, contentBottom - oy + EXPAND_INNER_PAD),
            kind: 'expanded',
            parentAreaId: origin ? origin.parentAreaId : null
        };
    } else if (isAreaNode(focusNode)) {
        // Bereich-Zoom: dieselbe Pack-Struktur wie auf der Landkarte (Nested-Areas + Rollen)
        // Eingeklappte Bereiche im Fokus immer ausgeklappt zeigen
        const tree = layoutAreaTree(
            overview.layout,
            focusNode,
            ox,
            oy,
            origin ? origin.parentAreaId : null,
            Object.assign({}, layoutOpts || {}, { forceExpand: true })
        );
        let maxRight = ox + EXPAND_MIN_W;
        let maxBottom = oy + 160;
        tree.boxes.forEach(function (b) {
            if (!b) return;
            if (b.node && String(b.node.id) === focusId) {
                expandBox = Object.assign({}, b, {
                    kind: 'expanded',
                    w: Math.max(EXPAND_MIN_W, Number(b.w) || EXPAND_MIN_W),
                    h: Math.max(160, Number(b.h) || 160),
                    parentAreaId: origin ? origin.parentAreaId : null
                });
                maxRight = Math.max(maxRight, expandBox.x + expandBox.w);
                maxBottom = Math.max(maxBottom, expandBox.y + expandBox.h);
                return;
            }
            innerBoxes.push(Object.assign({}, b, { focusChild: true }));
            maxRight = Math.max(maxRight, (b.x || 0) + (b.w || 0));
            maxBottom = Math.max(maxBottom, (b.y || 0) + (b.h || 0));
        });
        if (!expandBox) {
            expandBox = {
                node: focusNode,
                x: ox,
                y: oy,
                w: EXPAND_MIN_W,
                h: 160,
                kind: 'expanded',
                parentAreaId: origin ? origin.parentAreaId : null
            };
        }
        // Personen-Chips unter dem gepackten Bereichsinhalt
        let cursorY = maxBottom + (leadPeople.length || memberPeople.length ? 16 : 0);
        const chipWidth = Math.max(expandBox.w, EXPAND_MIN_W);
        leadPeople.forEach(function (p, i) {
            innerBoxes.push({
                person: p,
                x: ox + EXPAND_INNER_PAD,
                y: cursorY + i * (FOCUS_CHIP_H + chipGapY),
                w: Math.min(leadChipW + 80, chipWidth - EXPAND_INNER_PAD * 2),
                h: FOCUS_CHIP_H,
                kind: 'person-chip',
                parentAreaId: focusId,
                leadSeat: true
            });
        });
        if (leadPeople.length) {
            cursorY += leadPeople.length * (FOCUS_CHIP_H + chipGapY) + 8;
        }
        memberPeople.forEach(function (p, i) {
            const col = i % peopleCols;
            const row = Math.floor(i / peopleCols);
            innerBoxes.push({
                person: p,
                x: ox + EXPAND_INNER_PAD + col * (FOCUS_CHIP_W + chipGapX),
                y: cursorY + row * (FOCUS_CHIP_H + chipGapY),
                w: FOCUS_CHIP_W,
                h: FOCUS_CHIP_H,
                kind: 'person-chip',
                parentAreaId: focusId
            });
        });
        if (memberPeople.length) {
            cursorY += Math.ceil(memberPeople.length / peopleCols) * (FOCUS_CHIP_H + chipGapY);
        }
        if (leadPeople.length || memberPeople.length) {
            expandBox.h = Math.max(expandBox.h, cursorY - oy + EXPAND_INNER_PAD);
            expandBox.w = Math.max(expandBox.w, chipWidth);
        }
    } else {
        // Klasse / Rolle / ARGE: Personen-Chips + flache Kind-Kacheln (Kursteams …)
        const contentWidth = Math.max(
            peopleCols * FOCUS_CHIP_W + (peopleCols - 1) * chipGapX,
            leadPeople.length ? leadChipW : 0,
            isLehrerFocus ? FOCUS_CHIP_W + 40 : 400,
            kids.length ? 400 : 0
        );
        const EXPAND_W = Math.max(EXPAND_MIN_W, contentWidth + EXPAND_INNER_PAD * 2 + 8);
        const focusTopPad = 10;
        let cursorY = oy + HEAD_H + focusTopPad;

        leadPeople.forEach(function (p, i) {
            innerBoxes.push({
                person: p,
                x: ox + EXPAND_INNER_PAD,
                y: cursorY + i * (FOCUS_CHIP_H + chipGapY),
                w: Math.min(leadChipW + 80, EXPAND_W - EXPAND_INNER_PAD * 2),
                h: FOCUS_CHIP_H,
                kind: 'person-chip',
                parentAreaId: focusId,
                leadSeat: true
            });
        });
        if (leadPeople.length) {
            cursorY += leadPeople.length * (FOCUS_CHIP_H + chipGapY) + 8;
        }

        memberPeople.forEach(function (p, i) {
            const col = i % peopleCols;
            const row = Math.floor(i / peopleCols);
            innerBoxes.push({
                person: p,
                x: ox + EXPAND_INNER_PAD + col * (FOCUS_CHIP_W + chipGapX),
                y: cursorY + row * (FOCUS_CHIP_H + chipGapY),
                w: FOCUS_CHIP_W,
                h: FOCUS_CHIP_H,
                kind: 'person-chip',
                parentAreaId: focusId
            });
        });
        if (memberPeople.length) {
            cursorY += Math.ceil(memberPeople.length / peopleCols) * (FOCUS_CHIP_H + chipGapY);
        }

        if (kids.length) {
            if (leadPeople.length || memberPeople.length) cursorY += 14;
            const compactChildW = 168;
            const compactChildH = 40;
            const kidCols = Math.max(
                2,
                Math.min(
                    4,
                    Math.floor((EXPAND_W - EXPAND_INNER_PAD * 2 + childGapX) / (compactChildW + childGapX))
                )
            );
            kids.forEach(function (n, i) {
                const col = i % kidCols;
                const row = Math.floor(i / kidCols);
                innerBoxes.push({
                    node: n,
                    x: ox + EXPAND_INNER_PAD + col * (compactChildW + childGapX),
                    y: cursorY + row * (compactChildH + childGapY),
                    w: compactChildW,
                    h: compactChildH,
                    kind: isAreaNode(n) ? 'area' : 'leaf',
                    focusChild: true,
                    parentAreaId: focusId
                });
            });
            cursorY += Math.ceil(kids.length / kidCols) * (compactChildH + childGapY);
        } else if (!leadPeople.length && !memberPeople.length) {
            cursorY += 40;
        }

        expandBox = {
            node: focusNode,
            x: ox,
            y: oy,
            w: EXPAND_W,
            h: Math.max(160, cursorY - oy + EXPAND_INNER_PAD),
            kind: 'expanded',
            parentAreaId: origin ? origin.parentAreaId : null
        };
    }

    // Kontext: Übersicht dimmen; Fokus-Knoten + gesamter Unterbaum neu zeichnen
    const hideIds = new Set([focusId]);
    if (isAreaNode(focusNode) || isClassNode(focusNode)) {
        collectFocusHideIds(overview.layout, focusId, hideIds);
    }

    /** @type {object[]} */
    const boxes = [];
    overview.boxes.forEach(function (b) {
        if (!b || !b.node) return;
        if (hideIds.has(b.node.id)) return;
        boxes.push(Object.assign({}, b, { dimmed: true }));
    });
    boxes.push(expandBox);
    boxes.push.apply(boxes, innerBoxes);

    const maxX = boxes.reduce(function (m, b) {
        return Math.max(m, (b.x || 0) + (b.w || 0));
    }, overview.width);
    const maxY = boxes.reduce(function (m, b) {
        return Math.max(m, (b.y || 0) + (b.h || 0));
    }, overview.height);

    return {
        boxes: boxes,
        width: Math.max(overview.width, maxX + 80),
        height: Math.max(overview.height, maxY + 80),
        focusBox: expandBox,
        layout: overview.layout
    };
}

export function typeLabel(node) {
    const t = String((node && node.type) || '');
    if (t === NODE.ROLE || t === 'role') return 'Rolle';
    if (t === NODE.ARGE || t === 'arge' || isArgeAreaNode(node)) return 'ARGE';
    if (t === NODE.CLASS || t === 'class' || isClassNode(node)) return 'Klasse';
    if (t === NODE.KURSTEAM || t === 'kursteam') return 'Unterrichtsteam';
    if (t === NODE.COUNCIL || t === 'council') {
        const rk = String((node && node.refKey) || '').toLowerCase();
        if (rk.indexOf('schulsprecher') !== -1 || String((node && node.id) || '').indexOf('sv_') === 0) {
            return 'Funktion';
        }
        return 'SV';
    }
    if (t === NODE.SGA || t === 'sga') return 'SGA';
    if (t === NODE.PERSON || t === 'person') return 'Person';
    if (isArgeLeiterArea(node)) return 'Bereich';
    if (isAreaNode(node)) return 'Bereich';
    return t || 'Eintrag';
}

/**
 * Ob Klick den Drill-down öffnen soll.
 * @param {object} node
 */
export function isDrillableNode(node) {
    if (!node) return false;
    if (isAreaNode(node)) return true;
    const t = String(node.type || '');
    return (
        t === NODE.ROLE ||
        t === NODE.ARGE ||
        t === NODE.CLASS ||
        t === NODE.KURSTEAM ||
        t === NODE.COUNCIL ||
        t === NODE.SGA ||
        t === 'role' ||
        t === 'arge' ||
        t === 'class' ||
        t === 'kursteam' ||
        t === 'council' ||
        t === 'sga'
    );
}

/**
 * @param {object} opts
 * @param {HTMLElement} opts.host
 * @param {object} opts.layout
 * @param {(next: object) => void} opts.onLayoutChange
 * @param {(node: object) => void} [opts.onEditNode]
 * @param {(person: object, contextNode: object|null) => void} [opts.onEditPerson]
 * @param {(opts: { parentId: string|null, x: number, y: number }) => void|Promise<void>} [opts.onCreateArea]
 * @param {(areaNode: object) => void|Promise<void>} [opts.onAddPerson]
 * @param {(areaNode: object) => void|Promise<void>} [opts.onCreateRole]
 * @param {(nodeId: string, newParentId: string, liveLayout?: object) => object|null|undefined} [opts.onMoveNode]
 * @param {(person: object, targetNode: object) => { ok?: boolean, message?: string }|void} [opts.onMovePerson]
 * @param {(node: object) => number} [opts.peopleCountForArea]
 * @param {(node: object) => object[]} [opts.peopleForNode]
 * @param {(node: object) => Promise<object[]>|object[]} [opts.fetchFocusPeople]
 * @param {(email: string, name: string) => string} [opts.userThumbHtml]
 * @param {(groupId: string, name: string) => string} [opts.groupThumbHtml]
 * @param {(node: object) => string} [opts.resolveGroupId]
 * @param {(root: HTMLElement) => void} [opts.hydratePhotos]
 * @param {{
 *   active?: boolean,
 *   email?: string,
 *   name?: string,
 *   hitIds?: Set<string>|string[],
 *   pathIds?: Set<string>|string[]
 * }|null} [opts.personFocus]
 */
export function mountSchulOrganigrammLandkarte(opts) {
    const host = opts.host;
    let layout = normalizeSchulOrganigramm(opts.layout);
    /** @type {{ active: boolean, email: string, name: string, hitIds: Set<string>, pathIds: Set<string> }|null} */
    let personFocus = normalizePersonFocusOpt(opts.personFocus);
    let view = Object.assign({ x: 40, y: 24, k: 1 }, layout.view || {});
    /** @type {{ id: string, label: string }[]} */
    let focusStack = [];
    let panSession = null;
    let areaDrag = null;
    /** @type {object|null} */
    let areaResize = null;
    /** @type {number} */
    let areaResizeRaf = 0;
    /** @type {object|null} */
    let leafDrag = null;
    /** @type {object|null} */
    let personDrag = null;
    let viewSaveTimer = null;
    let overviewView = null;
    /** @type {object[]} */
    let lastOverviewBoxes = [];
    /** Nach Pan keinen Auszoom-Klick auslösen */
    let suppressBackgroundClick = false;
    /** Nach Leaf-/Personen-Drag keinen Drill-Klick auslösen */
    let suppressDrillClick = false;
    /**
     * Einmalig einpassen – nicht nach jedem Speichern/Remount,
     * wenn Viewport bereits gespeichert ist (sonst wirkt die Karte „zusammengerückt“).
     */
    let needsOverviewFit = !hasSavedLandkarteView(layout.view);
    /** @type {HTMLElement|null} */
    let ctxMenuEl = null;
    /** @type {Map<string, { loading?: boolean, people?: object[], error?: string }>} */
    const focusPeopleCache = new Map();
    /** @type {string} */
    let dropHighlightId = '';

    function clearDropHighlight() {
        host.querySelectorAll('.is-drop-target, .is-drop-promote, .is-person-drop').forEach(
            function (el) {
                el.classList.remove('is-drop-target', 'is-drop-promote', 'is-person-drop');
            }
        );
        const mapHostEl = host.querySelector('#schulOrgMapHost');
        if (mapHostEl) mapHostEl.classList.remove('is-detach-target');
        dropHighlightId = '';
    }

    /**
     * CSS-Klassen für Personen-Fokus (Hit / Pfad / gedimmt).
     * @param {string} nodeId
     */
    function personFocusNodeClass(nodeId) {
        if (!personFocus) return '';
        const id = String(nodeId || '');
        if (personFocus.hitIds.has(id)) return ' is-person-hit';
        if (personFocus.pathIds.has(id)) return ' is-person-path';
        return ' is-person-dimmed';
    }

    /**
     * @param {{ email?: string, name?: string }|null|undefined} p
     */
    function personFocusMatchesPerson(p) {
        if (!personFocus || !p) return false;
        const em = String((p && p.email) || '')
            .trim()
            .toLowerCase();
        if (personFocus.email && em && personFocus.email === em) return true;
        const nm = String((p && p.name) || '')
            .trim()
            .toLowerCase();
        const fn = String(personFocus.name || '')
            .trim()
            .toLowerCase();
        return !!(fn && nm && fn === nm);
    }

    /**
     * Drop-Ziel unter dem Zeiger hervorheben (Rahmen), solange Area gezogen wird.
     * @param {number} clientX
     * @param {number} clientY
     * @param {string} dragId
     * @param {Set<string>} [blockedIds] gezogene Area + Kinder (kein Drop auf sich selbst)
     */
    function updateDropHighlight(clientX, clientY, dragId, blockedIds) {
        const blocked = blockedIds || new Set([String(dragId || '')]);
        const nest = resolveNestDropAt(dragId, clientX, clientY, blocked);
        const mapHostEl = /** @type {HTMLElement|null} */ (host.querySelector('#schulOrgMapHost'));
        // Loslösen auf die Landkarte (Root)
        if (nest.mode === 'root') {
            const node = layout.nodes.find(function (n) {
                return n.id === dragId;
            });
            const wasNested = !!(node && node.parentId);
            if (wasNested) {
                if (dropHighlightId === '__root__') {
                    if (mapHostEl) mapHostEl.classList.add('is-detach-target');
                    return;
                }
                clearDropHighlight();
                dropHighlightId = '__root__';
                if (mapHostEl) mapHostEl.classList.add('is-detach-target');
                return;
            }
            clearDropHighlight();
            return;
        }
        // Nur neues Nest-/Promote-Ziel hervorheben (nicht den bisherigen Parent)
        const dropId =
            nest.mode === 'nest' || nest.mode === 'promote' ? nest.parentId || '' : '';
        if (!dropId) {
            clearDropHighlight();
            return;
        }
        if (dropHighlightId === dropId) return;
        clearDropHighlight();
        dropHighlightId = dropId;
        const el = host.querySelector(
            '.schul-org-map__area[data-area-id="' + CSS.escape(dropId) + '"]'
        );
        if (el) {
            el.classList.add('is-drop-target');
            if (nest.mode === 'promote') el.classList.add('is-drop-promote');
        }
    }

    /**
     * Person-Drop: Rolle/ARGE/Klasse-Leaf oder Area.
     * @returns {{ id: string, el: HTMLElement }|null}
     */
    function resolvePersonDropAt(clientX, clientY, blockedIds) {
        const blocked = blockedIds instanceof Set ? blockedIds : new Set();
        const stack =
            typeof document.elementsFromPoint === 'function'
                ? document.elementsFromPoint(clientX, clientY)
                : [document.elementFromPoint(clientX, clientY)].filter(Boolean);
        for (let i = 0; i < stack.length; i++) {
            const el = stack[i];
            if (!el || !host.contains(el) || typeof el.closest !== 'function') continue;
            const leaf = el.closest('.schul-org-map__leaf[data-node-id]');
            if (leaf && host.contains(leaf)) {
                const id = leaf.getAttribute('data-node-id') || '';
                if (!id || blocked.has(id)) continue;
                const node = layout.nodes.find(function (n) {
                    return n.id === id;
                });
                if (node && isPersonDropTargetNode(node)) {
                    return { id: id, el: /** @type {HTMLElement} */ (leaf) };
                }
            }
            const area = el.closest('.schul-org-map__area[data-area-id]');
            if (area && host.contains(area)) {
                const id = area.getAttribute('data-area-id') || '';
                if (!id || blocked.has(id)) continue;
                const node = layout.nodes.find(function (n) {
                    return n.id === id;
                });
                if (node && isPersonDropTargetNode(node)) {
                    return { id: id, el: /** @type {HTMLElement} */ (area) };
                }
            }
        }
        return null;
    }

    function updatePersonDropHighlight(clientX, clientY, blockedIds) {
        const hit = resolvePersonDropAt(clientX, clientY, blockedIds);
        const dropId = hit ? hit.id : '';
        if (!dropId) {
            clearDropHighlight();
            return;
        }
        if (dropHighlightId === dropId) return;
        clearDropHighlight();
        dropHighlightId = dropId;
        hit.el.classList.add('is-drop-target', 'is-person-drop');
    }

    function focusNode() {
        if (!focusStack.length) return null;
        const top = focusStack[focusStack.length - 1];
        return (
            layout.nodes.find(function (n) {
                return n.id === top.id;
            }) || null
        );
    }

    function sanitizeView() {
        const k = Number(view && view.k);
        const x = Number(view && view.x);
        const y = Number(view && view.y);
        view = {
            k: Number.isFinite(k) && k >= 0.25 && k <= 2.8 ? k : 1,
            x: Number.isFinite(x) ? x : 40,
            y: Number.isFinite(y) ? y : 24
        };
    }

    function applyTransform() {
        sanitizeView();
        const inner = host.querySelector('#schulOrgMapInner');
        if (!inner) return;
        inner.style.transform = 'translate(' + view.x + 'px, ' + view.y + 'px) scale(' + view.k + ')';
        inner.style.transformOrigin = '0 0';
    }

    /**
     * Client-Koordinaten → Landkarten-Welt über #schulOrgMapInner (Drop/Nest).
     * @param {number} clientX
     * @param {number} clientY
     */
    function clientToWorldInner(clientX, clientY) {
        const inner = host.querySelector('#schulOrgMapInner');
        if (!inner) return { x: 0, y: 0 };
        const rect = inner.getBoundingClientRect();
        const k = view.k || 1;
        // transform-origin 0 0: rect.left/top = Screen-Position von Welt (0,0)
        return {
            x: (clientX - rect.left) / k,
            y: (clientY - rect.top) / k
        };
    }

    /**
     * Live-DOM: Area unter dem Zeiger (getBoundingClientRect), Kopf bevorzugt.
     * Zuverlässiger während Drag als elementsFromPoint / alte Box-Koordinaten.
     * @param {number} clientX
     * @param {number} clientY
     * @param {Set<string>} blockedIds
     * @returns {{ id: string, inHeader: boolean }|null}
     */
    function pickLiveDropAreaAt(clientX, clientY, blockedIds) {
        /** @type {Set<string>} */
        const blocked = blockedIds instanceof Set ? blockedIds : new Set();
        /** @type {{ id: string, area: number, inHeader: boolean, depth: number }[]} */
        const hits = [];
        /** @type {Map<string, object>} */
        const byId = new Map();
        layout.nodes.forEach(function (n) {
            if (n && n.id) byId.set(String(n.id), n);
        });
        function depthOf(id) {
            let d = 0;
            let walk = String(id || '');
            let guard = 0;
            while (walk && guard < 48) {
                guard += 1;
                const n = byId.get(walk);
                const pid =
                    n && n.parentId != null && n.parentId !== '' ? String(n.parentId) : '';
                if (!pid) break;
                d += 1;
                walk = pid;
            }
            return d;
        }
        host.querySelectorAll('.schul-org-map__area[data-area-id]').forEach(function (el) {
            if (!(el instanceof HTMLElement)) return;
            if (el.classList.contains('is-dragging') || el.classList.contains('is-dimmed')) return;
            const id = el.getAttribute('data-area-id') || '';
            if (!id || blocked.has(id)) return;
            if (!el.hasAttribute('data-drop-area')) return;
            const r = el.getBoundingClientRect();
            if (
                clientX < r.left ||
                clientX > r.right ||
                clientY < r.top ||
                clientY > r.bottom
            ) {
                return;
            }
            const head = el.querySelector('.schul-org-map__area-head');
            const hr = head ? head.getBoundingClientRect() : null;
            const inHeader = !!(
                hr &&
                clientX >= hr.left &&
                clientX <= hr.right &&
                clientY >= hr.top &&
                clientY <= hr.bottom
            );
            hits.push({
                id: id,
                area: Math.max(1, r.width * r.height),
                inHeader: inHeader,
                depth: depthOf(id)
            });
        });
        if (!hits.length) return null;
        const headerHits = hits.filter(function (h) {
            return h.inHeader;
        });
        const pool = headerHits.length ? headerHits : hits;
        pool.sort(function (a, b) {
            if (b.depth !== a.depth) return b.depth - a.depth;
            return a.area - b.area;
        });
        return { id: pool[0].id, inHeader: pool[0].inHeader };
    }

    /**
     * Drop-Ziel für Nest/Promote/Root: Live-DOM zuerst, dann Geometrie.
     * Verschachtelte Bereiche: außerhalb des Parents → eine Ebene hoch oder loslösen (Root),
     * bzw. in einen anderen Bereich nesten.
     * @param {string} dragId
     * @param {number} clientX
     * @param {number} clientY
     * @param {Set<string>} blockedIds
     */
    function resolveNestDropAt(dragId, clientX, clientY, blockedIds) {
        const node = layout.nodes.find(function (n) {
            return n.id === dragId;
        });
        const curParent =
            node && node.parentId != null && node.parentId !== ''
                ? String(node.parentId)
                : '';
        /** @type {Map<string, object>} */
        const byId = new Map();
        layout.nodes.forEach(function (n) {
            if (n && n.id) byId.set(String(n.id), n);
        });
        function modeForTarget(targetId) {
            if (!targetId || targetId === dragId) return null;
            if (targetId === curParent) return { parentId: curParent || null, mode: 'keep' };
            let walk = curParent;
            let guard = 0;
            while (walk && guard < 48) {
                guard += 1;
                if (walk === targetId) {
                    return { parentId: targetId, mode: 'promote' };
                }
                const n = byId.get(walk);
                walk =
                    n && n.parentId != null && n.parentId !== ''
                        ? String(n.parentId)
                        : '';
            }
            return { parentId: targetId, mode: 'nest' };
        }
        function areaBoxById(id) {
            if (!id) return null;
            return (
                lastOverviewBoxes.find(function (b) {
                    return b && b.kind === 'area' && b.node && String(b.node.id) === id;
                }) || null
            );
        }
        function pointInAreaBox(box, wx, wy, pad) {
            if (!box) return false;
            const m = typeof pad === 'number' ? pad : 0;
            return (
                wx >= box.x - m &&
                wx <= box.x + (box.w || 0) + m &&
                wy >= box.y - m &&
                wy <= box.y + (box.h || 0) + m
            );
        }
        /** Außerhalb curParent: Ahnen hochklettern oder Root */
        function escapeFromParent(wx, wy) {
            let walk = curParent;
            let guard = 0;
            while (walk && guard < 48) {
                guard += 1;
                const n = byId.get(walk);
                const gp =
                    n && n.parentId != null && n.parentId !== '' ? String(n.parentId) : '';
                if (!gp) return { parentId: null, mode: 'root' };
                if (pointInAreaBox(areaBoxById(gp), wx, wy, 2)) {
                    return { parentId: gp, mode: 'promote' };
                }
                walk = gp;
            }
            return { parentId: null, mode: 'root' };
        }

        const world = clientToWorldInner(clientX, clientY);
        const live = pickLiveDropAreaAt(clientX, clientY, blockedIds);

        // 0) Verschachtelter Bereich hat den Elternrahmen verlassen →
        //    in anderen Bereich legen ODER eine Ebene hoch / als Root lösen
        if (curParent) {
            const parentBox = areaBoxById(curParent);
            const stillInParent = pointInAreaBox(parentBox, world.x, world.y, 6);
            if (!stillInParent) {
                if (live && live.id && live.id !== curParent) {
                    const escapedLive = modeForTarget(live.id);
                    if (
                        escapedLive &&
                        (escapedLive.mode === 'nest' || escapedLive.mode === 'promote')
                    ) {
                        return escapedLive;
                    }
                }
                const domEsc = resolveLandkarteDropAreaId(clientX, clientY, blockedIds, host);
                if (domEsc && domEsc !== curParent && domEsc !== dragId) {
                    const escapedDom = modeForTarget(domEsc);
                    if (
                        escapedDom &&
                        (escapedDom.mode === 'nest' || escapedDom.mode === 'promote')
                    ) {
                        return escapedDom;
                    }
                }
                return escapeFromParent(world.x, world.y);
            }
        }

        // 1) Live-DOM (sichtbare Kacheln unter dem Zeiger) – primär während Drag
        if (live && live.id) {
            const fromLive = modeForTarget(live.id);
            // Root-Bereiche: Verschachteln nur bei Drop auf den Kopf – sonst frei platzieren
            if (!curParent && fromLive && fromLive.mode === 'nest') {
                if (live.inHeader) return fromLive;
                return { parentId: null, mode: 'root' };
            }
            // Verschachtelt → anderen Bereich: Körper oder Kopf genügt zum Umhängen
            if (curParent && fromLive && fromLive.mode === 'nest') {
                return fromLive;
            }
            if (fromLive && fromLive.mode !== 'keep') return fromLive;
            if (fromLive && fromLive.mode === 'keep' && live.inHeader) {
                // Kopf des eigenen Parents: kein Nest-Wechsel
                return fromLive;
            }
        }

        // 2) Geometrie aus letzter Übersicht
        const geo = resolveLandkarteNestParent({
            boxes: lastOverviewBoxes,
            nodes: layout.nodes,
            worldX: world.x,
            worldY: world.y,
            dragId: dragId,
            blockedIds: blockedIds,
            currentParentId: curParent
        });
        // Root frei verschieben: Geometrie darf nicht versehentlich nesten (nur Live-Kopf)
        if (!curParent && geo.mode === 'nest') {
            return { parentId: null, mode: 'root' };
        }
        if (geo.mode === 'nest' || geo.mode === 'promote' || geo.mode === 'root') {
            return geo;
        }

        // 3) elementsFromPoint-Fallback
        const domId = resolveLandkarteDropAreaId(clientX, clientY, blockedIds, host);
        const fromDom = modeForTarget(domId);
        if (!curParent && fromDom && fromDom.mode === 'nest') {
            return { parentId: null, mode: 'root' };
        }
        if (curParent && fromDom && fromDom.mode === 'nest') {
            return fromDom;
        }
        if (fromDom && fromDom.mode !== 'keep') return fromDom;

        return geo;
    }

    /**
     * Übersicht so einpassen, dass Areas im sichtbaren Host liegen.
     * @param {{ width: number, height: number }} computed
     */
    function fitOverview(computed) {
        const mapHost = host.querySelector('#schulOrgMapHost');
        if (!mapHost || !computed) return false;
        const cw = mapHost.clientWidth;
        const ch = mapHost.clientHeight;
        if (cw < 80 || ch < 80) return false;
        const pad = 32;
        const k = Math.min(1.05, (cw - pad) / Math.max(320, computed.width), (ch - pad) / Math.max(240, computed.height));
        view = {
            k: Math.max(0.35, k),
            x: Math.max(12, (cw - computed.width * k) / 2),
            y: Math.max(12, (ch - computed.height * k) / 2)
        };
        applyTransform();
        return true;
    }

    function scheduleViewSave() {
        if (focusStack.length) return;
        if (viewSaveTimer) window.clearTimeout(viewSaveTimer);
        viewSaveTimer = window.setTimeout(function () {
            layout = setSchulOrganigrammView(layout, view);
            if (typeof opts.onLayoutChange === 'function') opts.onLayoutChange(layout);
        }, 400);
    }

    /**
     * Viewport zentriert auf die expandierte Box (Reinzoom-Fokus).
     * @param {{ x: number, y: number, w: number, h: number }} box
     * @returns {boolean}
     */
    function fitToBox(box) {
        const mapHost = host.querySelector('#schulOrgMapHost');
        if (!mapHost || !box) return false;
        const cwHost = mapHost.clientWidth;
        const chHost = mapHost.clientHeight;
        if (cwHost < 80 || chHost < 80) return false;
        const padX = 56;
        const padY = 72;
        const cw = Math.max(120, cwHost - padX * 2);
        const ch = Math.max(120, chHost - padY * 2);
        const bw = Math.max(80, Number(box.w) || 80);
        const bh = Math.max(80, Number(box.h) || 80);
        const k = Math.min(1.55, Math.max(0.5, Math.min(cw / bw, ch / bh)));
        const cx = Number(box.x) + bw / 2;
        const cy = Number(box.y) + bh / 2;
        view = {
            k: k,
            x: cwHost / 2 - cx * k,
            y: chHost / 2 - cy * k
        };
        applyTransform();
        return true;
    }

    function peopleFor(node) {
        if (!node) return [];
        const cached = focusPeopleCache.get(String(node.id));
        if (cached && Array.isArray(cached.people)) return cached.people;
        if (typeof opts.peopleForNode === 'function') {
            return opts.peopleForNode(node) || [];
        }
        return [];
    }

    /**
     * Personenanzahl für Layout/Vorschau: Graph-Cache (nachgeladen) hat Vorrang.
     * @param {object} area
     */
    function peopleCountForLayout(area) {
        if (!area) return 0;
        const cached = focusPeopleCache.get(String(area.id));
        if (cached && Array.isArray(cached.people) && (cached.fetched || cached.loading)) {
            return cached.people.filter(function (p) {
                return !!(p && (p.name || p.email) && !p.vacant);
            }).length;
        }
        if (typeof opts.peopleCountForArea === 'function') {
            return Math.max(0, Number(opts.peopleCountForArea(area)) || 0);
        }
        return 0;
    }

    /**
     * Ob für diesen Knoten Graph-Mitglieder nachgeladen werden sollen.
     * Fokus-Knoten immer; in der Übersicht ARGEs (+ freie Bereiche mit Gruppe).
     * Klassen nur im Fokus – Stammdaten-Schüler reichen für die Vorschau,
     * sonst zu viele Graph-Calls.
     * @param {object} node
     */
    function shouldHydrateGraphPeople(node) {
        if (!node || !nodeUsesGraphFocusPeople(node)) return false;
        if (typeof opts.fetchFocusPeople !== 'function') return false;
        const focused = focusNode();
        if (focused && String(focused.id) === String(node.id)) return true;
        if (isClassAreaNode(node) || isClassNode(node)) return false;
        if (isArgeAreaNode(node)) return true;
        if (node.type === NODE.ARGE || node.type === 'arge') return true;
        if (areaHasLinkedM365Group(node)) return true;
        return false;
    }

    function ensureFocusPeopleLoaded(node) {
        if (!shouldHydrateGraphPeople(node)) return;
        const id = String(node.id);
        const prev = focusPeopleCache.get(id);
        // Wichtig: auch ohne Graph-Treffer (nur lokale Leitung) als fertig markieren –
        // sonst paint → fetch → paint in Endlosschleife (UI friert ein).
        if (prev && (prev.loading || prev.fetched)) return;
        const local =
            typeof opts.peopleForNode === 'function' ? opts.peopleForNode(node) || [] : [];
        focusPeopleCache.set(id, { loading: true, fetched: false, people: local });
        Promise.resolve()
            .then(function () {
                return opts.fetchFocusPeople(node);
            })
            .then(function (list) {
                focusPeopleCache.set(id, {
                    loading: false,
                    fetched: true,
                    people: Array.isArray(list) ? list : local
                });
                // Übersicht + nested ARGE-Kacheln brauchen ebenfalls Neuzeichnung
                // (nicht nur der fokussierte Knoten selbst). Graph-Antworten bündeln.
                schedulePaint();
            })
            .catch(function (err) {
                focusPeopleCache.set(id, {
                    loading: false,
                    fetched: true,
                    people: local,
                    error: err && err.message ? String(err.message) : 'Graph-Laden fehlgeschlagen'
                });
                schedulePaint();
            });
    }

    /**
     * Sichtbare ARGE-/Klassen-/Gruppen-Bereiche: Mitglieder aus Graph nachladen,
     * damit die Landkarten-Vorschau nicht leer bleibt.
     * @param {object[]} boxes
     */
    function ensureVisibleGraphPeople(boxes) {
        (Array.isArray(boxes) ? boxes : []).forEach(function (b) {
            if (!b || !b.node || b.dimmed) return;
            if (b.kind !== 'area' && b.kind !== 'leaf' && b.kind !== 'expanded') return;
            ensureFocusPeopleLoaded(b.node);
        });
    }

    function personThumb(email, name) {
        if (typeof opts.userThumbHtml === 'function') {
            return opts.userThumbHtml(email || '', name || '') || '';
        }
        const ini = String(name || email || '?')
            .trim()
            .slice(0, 2)
            .toUpperCase();
        return (
            '<span class="schul-org-map__person-avatar" aria-hidden="true">' +
            escapeHtml(ini) +
            '</span>'
        );
    }

    function groupThumb(node) {
        if (typeof opts.resolveGroupId !== 'function' || typeof opts.groupThumbHtml !== 'function') {
            return '';
        }
        const gid = opts.resolveGroupId(node) || '';
        if (!gid) return '';
        return opts.groupThumbHtml(gid, (node && node.label) || '') || '';
    }

    function afterPaint() {
        if (typeof opts.hydratePhotos === 'function') {
            try {
                opts.hydratePhotos(host);
            } catch {
                /* ignore */
            }
        }
    }

    function enterFocus(node) {
        if (!node || !isDrillableNode(node)) return;
        if (!focusStack.length) {
            overviewView = { x: view.x, y: view.y, k: view.k };
        }
        // Gleicher Knoten erneut → Editor statt nochmal zoomen
        if (focusStack.length && focusStack[focusStack.length - 1].id === node.id) {
            if (typeof opts.onEditNode === 'function') opts.onEditNode(node);
            return;
        }
        focusStack.push({ id: node.id, label: node.label || typeLabel(node) });
        ensureFocusPeopleLoaded(node);
        paint();
    }

    function exitFocus() {
        if (!focusStack.length) return;
        focusStack.pop();
        if (!focusStack.length && overviewView) {
            view = Object.assign({}, overviewView);
            overviewView = null;
        }
        paint();
    }

    function exitFocusAll() {
        focusStack = [];
        if (overviewView) {
            view = Object.assign({}, overviewView);
            overviewView = null;
        } else {
            needsOverviewFit = true;
        }
        paint();
    }

    /**
     * Während Fenster-Resize: Box-Geometrie aktualisieren ohne mapHost zu zerstören
     * (Pointer-Capture bleibt erhalten, Inhalte fließen neu).
     */
    function applyLiveLandkarteGeometry() {
        const layoutOpts = {
            peopleCountForArea: peopleCountForLayout
        };
        const computed = computeLandkarteLayout(layout, layoutOpts);
        lastOverviewBoxes = computed.boxes.slice();
        const inner = /** @type {HTMLElement|null} */ (host.querySelector('#schulOrgMapInner'));
        const mapHostEl = /** @type {HTMLElement|null} */ (host.querySelector('#schulOrgMapHost'));
        if (!inner) return;
        inner.style.width = computed.width + 'px';
        inner.style.height = computed.height + 'px';
        computed.boxes.forEach(function (box) {
            if (!box || !box.node || !box.node.id) return;
            if (box.kind !== 'area' && box.kind !== 'leaf') return;
            const sel =
                box.kind === 'area'
                    ? '.schul-org-map__area[data-area-id="' + CSS.escape(box.node.id) + '"]'
                    : '.schul-org-map__leaf[data-node-id="' + CSS.escape(box.node.id) + '"]';
            const el = /** @type {HTMLElement|null} */ (inner.querySelector(sel));
            if (!el) return;
            el.style.left = box.x + 'px';
            el.style.top = box.y + 'px';
            el.style.width = box.w + 'px';
            el.style.height = box.h + 'px';
        });
        if (mapHostEl && areaResize) {
            mapHostEl.classList.add('is-area-resizing');
            const active = /** @type {HTMLElement|null} */ (
                inner.querySelector(
                    '.schul-org-map__area[data-area-id="' + CSS.escape(areaResize.id) + '"]'
                )
            );
            if (active) active.classList.add('is-resizing');
        }
    }

    function scheduleAreaResizePaint() {
        if (areaResizeRaf) return;
        areaResizeRaf = window.requestAnimationFrame(function () {
            areaResizeRaf = 0;
            if (!areaResize) return;
            applyLiveLandkarteGeometry();
        });
    }

    let paintRaf = 0;
    let paintImmediate = false;

    function paintNow() {
        paintRaf = 0;
        paintImmediate = false;
        host.replaceChildren();
        host.classList.add('schul-org-map');
        paintCanvas();
        bindHostOnce();
    }

    /**
     * Sofort zeichnen (User-Interaktion: Zoomen, Drag, Menü).
     */
    function paint() {
        if (paintRaf) {
            window.cancelAnimationFrame(paintRaf);
            paintRaf = 0;
        }
        paintImmediate = true;
        paintNow();
    }

    /**
     * Graph-/Async-Updates: mehrere Completions im selben Frame → eine Zeichnung.
     */
    function schedulePaint() {
        if (paintImmediate || paintRaf) return;
        paintRaf = window.requestAnimationFrame(function () {
            paintRaf = 0;
            paintNow();
        });
    }

    function paintChrome(hintHtml) {
        const crumbs = focusStack
            .map(function (c, i) {
                const last = i === focusStack.length - 1;
                return last
                    ? '<span class="schul-org-map__crumb is-current">' + escapeHtml(c.label) + '</span>'
                    : '<button type="button" class="schul-org-map__crumb" data-focus-index="' +
                          i +
                          '">' +
                          escapeHtml(c.label) +
                          '</button>';
            })
            .join('<span class="schul-org-map__crumb-sep" aria-hidden="true">/</span>');

        return (
            '<div class="schul-org-map__toolbar">' +
            '<div class="schul-org-map__crumbs" aria-label="Navigation">' +
            (focusStack.length
                ? '<button type="button" class="schul-org-map__crumb schul-org-map__crumb--root" data-focus-root="1">Landkarte</button>' +
                  '<span class="schul-org-map__crumb-sep" aria-hidden="true">/</span>' +
                  crumbs
                : '<span class="schul-org-map__crumb is-current">Landkarte</span>') +
            '</div>' +
            (hintHtml
                ? '<p class="schul-org-map__hint muted">' + hintHtml + '</p>'
                : '') +
            '</div>'
        );
    }

    function wireChrome() {
        host.querySelectorAll('[data-focus-root]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                exitFocusAll();
            });
        });
        host.querySelectorAll('[data-focus-index]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const idx = Number(btn.getAttribute('data-focus-index'));
                if (!Number.isFinite(idx)) return;
                focusStack = focusStack.slice(0, idx + 1);
                paint();
            });
        });
        const back = host.querySelector('[data-focus-back]');
        if (back) {
            back.addEventListener('click', function () {
                exitFocus();
            });
        }
    }

    /**
     * Alle Boxen (Area + Nested + Leaves), die mit einer Area mitwandern.
     * @param {string} areaId
     */
    function collectDragGroupBoxes(areaId) {
        const rootId = String(areaId || '');
        if (!rootId) return [];
        const idSet = new Set([rootId]);
        let grew = true;
        while (grew) {
            grew = false;
            lastOverviewBoxes.forEach(function (b) {
                if (!b || !b.node) return;
                const pid = b.parentAreaId != null ? String(b.parentAreaId) : '';
                if (pid && idSet.has(pid) && !idSet.has(b.node.id)) {
                    idSet.add(b.node.id);
                    grew = true;
                }
            });
        }
        return lastOverviewBoxes.filter(function (b) {
            return b && b.node && idSet.has(b.node.id);
        });
    }

    /**
     * @param {string} areaId
     */
    function collectDragSatellites(areaId) {
        const boxes = collectDragGroupBoxes(areaId);
        /** @type {{ el: HTMLElement, startX: number, startY: number, isRoot: boolean }[]} */
        const sats = [];
        boxes.forEach(function (b) {
            let el = null;
            if (b.kind === 'area') {
                el = /** @type {HTMLElement|null} */ (
                    host.querySelector('.schul-org-map__area[data-area-id="' + CSS.escape(b.node.id) + '"]')
                );
            } else if (b.kind === 'leaf') {
                el = /** @type {HTMLElement|null} */ (
                    host.querySelector('.schul-org-map__leaf[data-node-id="' + CSS.escape(b.node.id) + '"]')
                );
            }
            if (!el) return;
            sats.push({
                el: el,
                startX: parseFloat(el.style.left) || Number(b.x) || 0,
                startY: parseFloat(el.style.top) || Number(b.y) || 0,
                isRoot: b.node.id === areaId
            });
        });
        return sats;
    }

    function paintCanvas() {
        const focused = focusNode();
        if (focused) ensureFocusPeopleLoaded(focused);
        const people = focused ? peopleFor(focused) : [];
        const focusCache = focused ? focusPeopleCache.get(String(focused.id)) : null;
        const childNodes = focused
            ? childrenOf(layout, focused.id).filter(function (n) {
                  return n.type !== NODE.PERSON;
              })
            : [];
        const layoutOpts = {
            peopleCountForArea: peopleCountForLayout
        };
        const computed = focused
            ? computeFocusLayout(layout, focused, people, childNodes, layoutOpts)
            : computeLandkarteLayout(layout, layoutOpts);
        // Im Fokus: teure zweite Landkarten-Berechnung vermeiden (Overview-Cache behalten)
        if (!focused) {
            lastOverviewBoxes = computed.boxes.slice();
        } else if (!lastOverviewBoxes.length) {
            lastOverviewBoxes = computeLandkarteLayout(layout, layoutOpts).boxes.slice();
        }
        // Nested ARGEs/Klassen mit M365-Gruppe: Mitglieder auch ohne Hineinzoomen laden
        ensureVisibleGraphPeople(computed.boxes);

        const hint =
            focused && focusCache && focusCache.loading
                ? 'Mitglieder werden aus Microsoft 365 geladen …'
                : '';

        host.innerHTML =
            paintChrome(hint) +
            (focused
                ? '<div class="schul-org-map__focus-bar">' +
                  '<button type="button" class="btn btn-sm" data-focus-back><i class="bi bi-zoom-out" aria-hidden="true"></i> Auszoomen</button>' +
                  '<button type="button" class="btn btn-sm btn-success" data-edit-id="' +
                  escapeHtml(focused.id) +
                  '"><i class="bi bi-pencil" aria-hidden="true"></i> Bearbeiten</button>' +
                  (focusCache && focusCache.loading
                      ? '<span class="muted schul-org-map__focus-loading">Mitglieder laden …</span>'
                      : '') +
                  '</div>'
                : '') +
            '<div class="schul-org-map__host' +
            (focused ? ' schul-org-map__host--zoomed' : '') +
            (personFocus ? ' schul-org-map__host--person-focus' : '') +
            '" id="schulOrgMapHost">' +
            '<div class="schul-org-map__inner" id="schulOrgMapInner" style="width:' +
            computed.width +
            'px;height:' +
            computed.height +
            'px"></div></div>';

        wireChrome();
        const inner = /** @type {HTMLElement} */ (host.querySelector('#schulOrgMapInner'));

        computed.boxes.forEach(function (box) {
            try {
            if (box.kind === 'person-chip' && box.person) {
                const p = box.person;
                const el = document.createElement('div');
                const isLead = !!box.leadSeat || normStr(p.kind).toLowerCase() === 'head';
                const vacant = !!p.vacant || !(p.email || p.name);
                const kindLc = normStr(p.kind).toLowerCase();
                const canDragPerson =
                    !vacant &&
                    !p.fromGraph &&
                    (p.movable === true ||
                        kindLc === 'admin' ||
                        kindLc === 'areaperson' ||
                        kindLc === 'head');
                const userLinked = !vacant && personHasM365UserLink(p);
                const pfHit = !vacant && personFocusMatchesPerson(p);
                el.className =
                    'schul-org-map__person' +
                    (isLead ? ' schul-org-map__person--lead' : '') +
                    (vacant ? ' is-vacant' : '') +
                    (canDragPerson ? ' is-draggable' : '') +
                    (userLinked ? ' is-m365-user' : '') +
                    (personFocus
                        ? pfHit
                            ? ' is-person-hit'
                            : ' is-person-dimmed'
                        : '');
                el.style.left = box.x + 'px';
                el.style.top = box.y + 'px';
                el.style.width = box.w + 'px';
                el.style.height = box.h + 'px';
                el.dataset.personEmail = p.email || '';
                el.dataset.personName = p.name || p.label || '';
                el.dataset.personKind = p.kind || '';
                if (userLinked) el.dataset.m365User = '1';
                if (p.seatLabel) el.dataset.seatLabel = p.seatLabel;
                if (box.parentAreaId) el.dataset.parentAreaId = box.parentAreaId;
                if (focused && focused.id) el.dataset.focusNodeId = focused.id;
                if (canDragPerson) {
                    el.dataset.personDrag = '1';
                    const srcNode =
                        layout.nodes.find(function (n) {
                            return n.id === (box.parentAreaId || (focused && focused.id) || '');
                        }) || focused;
                    if (srcNode) {
                        el.dataset.sourceNodeId = srcNode.id || '';
                        el.dataset.sourceRefKind = srcNode.refKind || '';
                        el.dataset.sourceRefKey = srcNode.refKey || '';
                    }
                }
                const editTitle = isLead
                    ? vacant
                        ? (p.seatLabel || 'Leitung') + ' zuordnen'
                        : (p.seatLabel || 'Leitung') + ' bearbeiten'
                    : 'Person bearbeiten';
                el.title =
                    (p.seatLabel ? p.seatLabel + ': ' : '') +
                    (p.email || p.name || 'offen') +
                    (userLinked ? ' · M365-Konto' : vacant ? '' : ' · nur lokal') +
                    (canDragPerson ? ' · ziehen = verschieben' : '') +
                    ' · Stift = bearbeiten';
                el.innerHTML =
                    (p.seatLabel || isLead
                        ? '<span class="schul-org-map__person-seat' +
                          (isLead ? ' is-lead-seat' : '') +
                          '">' +
                          escapeHtml(p.seatLabel || (isLead ? 'Leitung' : '')) +
                          '</span>'
                        : '') +
                    personThumb(p.email, p.name || p.label) +
                    '<span class="schul-org-map__person-text">' +
                    '<strong>' +
                    escapeHtml(
                        p.name || p.email || (vacant ? 'Noch nicht besetzt – zuordnen' : 'Person')
                    ) +
                    '</strong>' +
                    (p.email
                        ? '<span class="muted">' + escapeHtml(p.email) + '</span>'
                        : vacant
                          ? '<span class="muted">Stift = zuordnen</span>'
                          : '') +
                    '</span>' +
                    personM365MarkHtml(p) +
                    '<button type="button" class="schul-org-map__edit schul-org-map__edit--person" data-edit-person="1" title="' +
                    escapeHtml(editTitle) +
                    '"><i class="bi bi-pencil" aria-hidden="true"></i></button>';
                inner.appendChild(el);
                return;
            }

            if (box.kind === 'expanded' && box.node) {
                const m365 = box.node.m365 || {};
                const m365Parts = [];
                if (normStr(box.node.meta && box.node.meta.chatId)) {
                    m365Parts.push('Teams-Gruppenchat');
                }
                if (
                    m365.groupMode &&
                    m365.groupMode !== 'none' &&
                    m365.groupId &&
                    normalizeSchulOrgGroupKind(m365.groupKind) !== 'groupChat'
                ) {
                    m365Parts.push(
                        schulOrgGroupKindLabel(m365.groupKind) +
                            (m365.groupLabel ? ': ' + m365.groupLabel : '')
                    );
                }
                if (m365.mailboxMode && m365.mailboxMode !== 'none' && m365.mailboxUpn) {
                    m365Parts.push(m365.mailboxLabel || 'Postfach');
                }
                const gThumb = groupThumb(box.node);
                const el = document.createElement('div');
                el.className =
                    'schul-org-map__area schul-org-map__area--expanded schul-org-map__area--' +
                    (box.node.tone || box.node.type || 'custom');
                el.dataset.areaId = box.node.id;
                el.dataset.expandedId = box.node.id;
                el.style.left = box.x + 'px';
                el.style.top = box.y + 'px';
                el.style.width = box.w + 'px';
                el.style.height = box.h + 'px';
                el.innerHTML =
                    '<div class="schul-org-map__area-head">' +
                    (gThumb ? '<span class="schul-org-map__group-thumb">' + gThumb + '</span>' : '') +
                    '<div class="schul-org-map__area-head-text">' +
                    '<strong>' +
                    escapeHtml(box.node.label) +
                    '</strong>' +
                    '<span class="muted">' +
                    escapeHtml(
                        [
                            typeLabel(box.node),
                            people.length
                                ? people.length +
                                  (people.length === 1 ? ' Person' : ' Personen')
                                : '',
                            childNodes.length
                                ? childNodes.length +
                                  (childNodes.length === 1 ? ' Eintrag' : ' Einträge')
                                : '',
                            m365Parts.join(' · ')
                        ]
                            .filter(Boolean)
                            .join(' · ')
                    ) +
                    '</span></div>' +
                    '<button type="button" class="schul-org-map__edit" data-edit-id="' +
                    escapeHtml(box.node.id) +
                    '" title="Bearbeiten"><i class="bi bi-pencil" aria-hidden="true"></i></button>' +
                    '</div>' +
                    '<div class="schul-org-map__area-body schul-org-map__area-body--expanded"></div>';
                inner.appendChild(el);
                return;
            }

            if (box.kind === 'area' && box.node) {
                const el = document.createElement('div');
                // Legacy classNest nur noch für alte Boxen; Klassen sind echte Bereiche
                const classNest = !!box.classNest;
                const isCollapsed = !!(box.collapsed || box.node.collapsed);
                const contentCount =
                    typeof box.contentCount === 'number'
                        ? box.contentCount
                        : countAreaMapContents(computed.layout || layout, box.node.id);
                const visibleInnerCount =
                    typeof box.visibleInnerCount === 'number'
                        ? box.visibleInnerCount
                        : contentCount;
                const kursCount = classNest
                    ? typeof box.kursCount === 'number'
                        ? box.kursCount
                        : childrenOf(computed.layout || layout, box.node.id).filter(isKursteamNode)
                              .length
                    : 0;
                const resolvedGid =
                    typeof opts.resolveGroupId === 'function' ? opts.resolveGroupId(box.node) || '' : '';
                const resolvedChatId =
                    typeof opts.resolveChatId === 'function' ? opts.resolveChatId(box.node) || '' : '';
                const m365Flags = nodeM365LinkFlags(box.node, resolvedGid, resolvedChatId);
                const canResize =
                    !!box.resizable &&
                    !focused &&
                    !box.dimmed &&
                    !classNest &&
                    isAreaNode(box.node);
                el.className =
                    'schul-org-map__area schul-org-map__area--' +
                    (box.node.tone || box.node.type || 'custom') +
                    (classNest ? ' schul-org-map__area--class-nest' : '') +
                    (isCollapsed ? ' schul-org-map__area--collapsed' : '') +
                    (visibleInnerCount > 0 && !classNest ? ' schul-org-map__area--has-leaves' : '') +
                    (box.peoplePreview ? ' schul-org-map__area--has-preview' : '') +
                    (m365Flags.linked ? ' is-m365-linked' : '') +
                    (box.dimmed ? ' is-dimmed' : '') +
                    (canResize ? ' is-resizable' : '') +
                    personFocusNodeClass(box.node.id) +
                    (box.focusChild ? ' schul-org-map__area--inside is-drillable' : '') +
                    (!box.focusChild && classNest ? ' is-drillable' : '');
                el.dataset.areaId = box.node.id;
                if (canResize) el.dataset.resizable = '1';
                if (!classNest) el.dataset.dropArea = '1';
                el.dataset.drillId = box.node.id;
                if (isCollapsed) el.dataset.collapsed = '1';
                if (box.parentAreaId) el.dataset.parentAreaId = box.parentAreaId;
                if (m365Flags.linked) el.dataset.m365Linked = '1';
                el.style.left = box.x + 'px';
                el.style.top = box.y + 'px';
                el.style.width = box.w + 'px';
                el.style.height = box.h + 'px';
                const m365 = box.node.m365 || {};
                const m365Parts = [];
                if (Array.isArray(m365Flags.marks) && m365Flags.marks.length) {
                    m365Flags.marks.forEach(function (m) {
                        m365Parts.push(m.title.replace(/\s+verknüpft$/i, '') || m.kind);
                    });
                } else {
                    if (m365Flags.hasChat || m365Flags.showChatMark) {
                        m365Parts.push('Teams-Gruppenchat');
                    }
                    if (m365Flags.hasGroup) {
                        m365Parts.push(
                            m365.groupId
                                ? schulOrgGroupKindLabel(m365.groupKind)
                                : 'Gruppe'
                        );
                    }
                    if (m365Flags.hasMailbox) {
                        m365Parts.push('Postfach');
                    }
                }
                let peopleHint = '';
                {
                    const pc = peopleCountForLayout(box.node);
                    if (pc > 0) peopleHint = pc + (pc === 1 ? ' Person' : ' Personen');
                }
                if (classNest) {
                    if (kursCount > 0) {
                        peopleHint =
                            kursCount +
                            (kursCount === 1 ? ' Unterrichtsteam' : ' Unterrichtsteams');
                    } else {
                        const hm = normStr(box.node.meta && box.node.meta.headName);
                        const he = normStr(box.node.meta && box.node.meta.headEmail);
                        peopleHint = hm || he || 'keine Unterrichtsteams';
                    }
                }
                const entryHint =
                    !classNest && contentCount > 0
                        ? contentCount + (contentCount === 1 ? ' Eintrag' : ' Einträge')
                        : '';
                const metaLine = [
                    typeLabel(box.node),
                    isCollapsed ? entryHint || peopleHint : peopleHint || entryHint,
                    m365Parts.join(' · ')
                ]
                    .filter(Boolean)
                    .join(' · ');
                const aThumb = groupThumb(box.node);
                let previewHtml = '';
                if (box.peoplePreview && !box.dimmed && !classNest && !isCollapsed) {
                    // Personen, die schon an Kind-Rollen hängen, nicht nochmal im Bereich listen
                    /** @type {Set<string>} */
                    const onRoles = new Set();
                    childrenOf(computed.layout || layout, box.node.id).forEach(function (ch) {
                        const ct = String((ch && ch.type) || '').toLowerCase();
                        if (ct !== NODE.ROLE && ct !== 'role') return;
                        peopleFor(ch).forEach(function (p) {
                            const em = String((p && p.email) || '')
                                .trim()
                                .toLowerCase();
                            if (em) onRoles.add(em);
                        });
                    });
                    const allPeople = peopleFor(box.node).filter(function (p) {
                        const em = String((p && p.email) || '')
                            .trim()
                            .toLowerCase();
                        if (em && onRoles.has(em)) return false;
                        return !!(p && (p.name || p.email) && !p.vacant);
                    });
                    const leadPeople = allPeople.filter(function (p) {
                        return String(p.kind || '').toLowerCase() === 'head';
                    });
                    const memberPeople = allPeople.filter(function (p) {
                        return String(p.kind || '').toLowerCase() !== 'head';
                    });
                    const leadShown = leadPeople.slice(0, 1);
                    const memberSlots = Math.max(0, PREVIEW_CHIP_MAX - leadShown.length);
                    const memberShown = memberPeople.slice(0, memberSlots);
                    const more = Math.max(
                        0,
                        allPeople.length - leadShown.length - memberShown.length
                    );

                    const membersLabel =
                        leadShown.length && memberShown.length
                            ? String(
                                  (memberShown[0] && memberShown[0].seatLabel) || 'Mitglieder'
                              ).trim() || 'Mitglieder'
                            : '';

                    /**
                     * @param {object} p
                     * @param {{ isLead?: boolean, showSeat?: boolean }} opts
                     */
                    function previewChipHtml(p, opts) {
                        const isLead = !!(opts && opts.isLead);
                        const showSeat = opts && opts.showSeat !== false;
                        const pKind = String(p.kind || 'areaPerson').toLowerCase();
                        const seatLabel = String(p.seatLabel || '').trim();
                        const seatVisible = showSeat && !!seatLabel;
                        const canDragPreview =
                            p.movable !== false &&
                            (pKind === 'areaperson' ||
                                pKind === 'admin' ||
                                pKind === 'person' ||
                                pKind === 'head');
                        const userLinked = personHasM365UserLink(p);
                        const pfHit = personFocusMatchesPerson(p);
                        return (
                            '<span class="schul-org-map__preview-chip' +
                            (isLead ? ' is-lead' : ' is-member') +
                            (canDragPreview ? ' is-draggable' : '') +
                            (userLinked ? ' is-m365-user' : '') +
                            (personFocus
                                ? pfHit
                                    ? ' is-person-hit'
                                    : ' is-person-dimmed'
                                : '') +
                            '" ' +
                            'data-person-email="' +
                            escapeHtml(p.email || '') +
                            '" data-person-name="' +
                            escapeHtml(p.name || p.label || '') +
                            '" data-person-kind="' +
                            escapeHtml(p.kind || 'areaPerson') +
                            '"' +
                            (seatLabel
                                ? ' data-seat-label="' + escapeHtml(seatLabel) + '"'
                                : '') +
                            ' data-parent-area-id="' +
                            escapeHtml(box.node.id) +
                            '"' +
                            (userLinked ? ' data-m365-user="1"' : '') +
                            (canDragPreview
                                ? ' data-person-drag="1" data-source-node-id="' +
                                  escapeHtml(box.node.id) +
                                  '" data-source-ref-kind="area" data-source-ref-key="' +
                                  escapeHtml(box.node.id) +
                                  '"'
                                : '') +
                            ' title="' +
                            escapeHtml(
                                (seatLabel ? seatLabel + ': ' : '') +
                                    (p.name || p.email || '') +
                                    (userLinked ? ' · M365-Konto' : ' · nur lokal') +
                                    (canDragPreview
                                        ? ' · ziehen auf Rolle = zuordnen'
                                        : '') +
                                    ' · Stift = bearbeiten'
                            ) +
                            '">' +
                            personThumb(p.email, p.name || p.label) +
                            '<span class="schul-org-map__preview-chip-text">' +
                            (seatVisible
                                ? '<span class="schul-org-map__preview-chip-seat' +
                                  (isLead ? ' is-lead-seat' : '') +
                                  '">' +
                                  escapeHtml(seatLabel) +
                                  '</span>'
                                : '') +
                            '<span class="schul-org-map__preview-chip-name">' +
                            escapeHtml(p.name || p.email || '?') +
                            '</span></span>' +
                            personM365MarkHtml(p) +
                            '<button type="button" class="schul-org-map__edit schul-org-map__edit--preview" data-edit-person="1" title="Person bearbeiten">' +
                            '<i class="bi bi-pencil" aria-hidden="true"></i></button>' +
                            '</span>'
                        );
                    }

                    if (leadShown.length || memberShown.length || more) {
                        previewHtml =
                            '<div class="schul-org-map__people-preview' +
                            (leadShown.length ? ' has-lead' : '') +
                            (memberShown.length ? ' has-members' : '') +
                            '" data-drill-id="' +
                            escapeHtml(box.node.id) +
                            '">' +
                            (leadShown.length
                                ? '<div class="schul-org-map__preview-lead">' +
                                  leadShown
                                      .map(function (p) {
                                          return previewChipHtml(p, {
                                              isLead: true,
                                              showSeat: true
                                          });
                                      })
                                      .join('') +
                                  '</div>'
                                : '') +
                            (memberShown.length
                                ? '<div class="schul-org-map__preview-members">' +
                                  (membersLabel
                                      ? '<span class="schul-org-map__preview-members-label">' +
                                        escapeHtml(membersLabel) +
                                        '</span>'
                                      : '') +
                                  memberShown
                                      .map(function (p) {
                                          return previewChipHtml(p, {
                                              isLead: false,
                                              // Gruppenlabel oben – Seat nicht auf jeder Pille wiederholen
                                              showSeat: !membersLabel
                                          });
                                      })
                                      .join('') +
                                  '</div>'
                                : '') +
                            (more
                                ? '<span class="schul-org-map__preview-more">+' + more + '</span>'
                                : '') +
                            '</div>';
                    }
                }
                const allowAreaDrag = isAreaNode(box.node) && !box.dimmed && !focused && !classNest;
                const peopleCountForCollapse =
                    typeof box.previewCount === 'number'
                        ? box.previewCount
                        : peopleCountForLayout(box.node);
                // Auch reine Personen-Bereiche (Lehrkräfte, Kooperationsteam, …) einklappbar
                const showCollapse =
                    !classNest &&
                    !box.dimmed &&
                    (contentCount > 0 ||
                        peopleCountForCollapse > 0 ||
                        !!box.peoplePreview ||
                        isCollapsed);
                const collapseBtn = showCollapse
                    ? '<button type="button" class="schul-org-map__collapse" data-collapse-id="' +
                      escapeHtml(box.node.id) +
                      '" title="' +
                      (isCollapsed ? 'Ausklappen' : 'Einklappen') +
                      '" aria-expanded="' +
                      (isCollapsed ? 'false' : 'true') +
                      '" aria-label="' +
                      (isCollapsed ? 'Bereich ausklappen' : 'Bereich einklappen') +
                      '"><i class="bi ' +
                      (isCollapsed ? 'bi-chevron-right' : 'bi-chevron-down') +
                      '" aria-hidden="true"></i></button>'
                    : '';
                el.innerHTML =
                    '<div class="schul-org-map__area-head"' +
                    (allowAreaDrag ? ' data-area-drag="' + escapeHtml(box.node.id) + '"' : '') +
                    (classNest ? ' data-drill-id="' + escapeHtml(box.node.id) + '"' : '') +
                    '>' +
                    (aThumb ? '<span class="schul-org-map__group-thumb">' + aThumb + '</span>' : '') +
                    '<div class="schul-org-map__area-head-text">' +
                    '<strong>' +
                    escapeHtml(box.node.label) +
                    m365MarkHtml(m365Flags, 'schul-org-map__m365-mark--inline') +
                    '</strong>' +
                    '<span class="muted">' +
                    escapeHtml(metaLine) +
                    '</span></div>' +
                    '<div class="schul-org-map__area-actions">' +
                    collapseBtn +
                    '<button type="button" class="schul-org-map__edit" data-edit-id="' +
                    escapeHtml(box.node.id) +
                    '" title="' +
                    (classNest ? 'Klasse bearbeiten' : 'Bereich bearbeiten') +
                    '"><i class="bi bi-pencil" aria-hidden="true"></i></button>' +
                    '</div></div>' +
                    (isCollapsed
                        ? ''
                        : '<div class="schul-org-map__area-body" data-drill-id="' +
                          escapeHtml(box.node.id) +
                          '" title="Hineinzoomen"></div>') +
                    // Vorschau nach dem Body: bei Rollen/Leaves unten im Container
                    // (Layout reserviert den Streifen unter den Leaves – nicht unter dem Kopf)
                    previewHtml +
                    (canResize
                        ? '<span class="schul-org-map__resize schul-org-map__resize--e" data-area-resize="' +
                          escapeHtml(box.node.id) +
                          '" data-resize-edge="e" title="Breite ändern"></span>' +
                          '<span class="schul-org-map__resize schul-org-map__resize--s" data-area-resize="' +
                          escapeHtml(box.node.id) +
                          '" data-resize-edge="s" title="Höhe ändern"></span>' +
                          '<span class="schul-org-map__resize schul-org-map__resize--se" data-area-resize="' +
                          escapeHtml(box.node.id) +
                          '" data-resize-edge="se" title="Größe ändern"></span>'
                        : '');
                inner.appendChild(el);
                return;
            }

            if (box.kind === 'leaf' && box.node) {
                const el = document.createElement('div');
                // Rollen auch im Fokus/Zoom verschiebbar (nicht nur in der Übersicht)
                const leafDraggable = isMapLeafDraggable(box.node) && !box.dimmed;
                const leafType = String(box.node.type || '').toLowerCase();
                const seatLeaf =
                    leafType === NODE.ROLE ||
                    leafType === 'role' ||
                    leafType === NODE.ARGE ||
                    leafType === 'arge' ||
                    leafType === NODE.CLASS ||
                    leafType === 'class' ||
                    leafType === NODE.KURSTEAM ||
                    leafType === 'kursteam';
                const leafGid =
                    typeof opts.resolveGroupId === 'function' ? opts.resolveGroupId(box.node) || '' : '';
                const leafChatId =
                    typeof opts.resolveChatId === 'function' ? opts.resolveChatId(box.node) || '' : '';
                const leafM365 = nodeM365LinkFlags(box.node, leafGid, leafChatId);
                el.className =
                    'schul-org-map__leaf schul-org-map__leaf--' +
                    (box.node.tone || box.node.type || 'custom') +
                    ' is-drillable' +
                    (box.dimmed ? ' is-dimmed' : '') +
                    (box.focusChild ? ' schul-org-map__leaf--inside' : '') +
                    (leafDraggable ? ' is-draggable' : '') +
                    (seatLeaf ? ' schul-org-map__leaf--with-people' : '') +
                    (leafM365.linked ? ' is-m365-linked' : '') +
                    personFocusNodeClass(box.node.id);
                el.dataset.nodeId = box.node.id;
                el.dataset.drillId = box.node.id;
                if (box.parentAreaId) el.dataset.parentAreaId = box.parentAreaId;
                if (leafDraggable) el.dataset.leafDrag = box.node.id;
                if (leafM365.linked) el.dataset.m365Linked = '1';
                el.style.left = box.x + 'px';
                el.style.top = box.y + 'px';
                el.style.width = box.w + 'px';
                el.style.height = box.h + 'px';
                const people = peopleFor(box.node);
                const filled = people.filter(function (p) {
                    return p && !p.vacant && (p.name || p.email);
                });
                const vacantSeat = people.find(function (p) {
                    return p && p.vacant;
                });
                const personHint =
                    (box.node.meta && box.node.meta.personName) ||
                    (filled[0] && (filled[0].name || filled[0].email)) ||
                    '';
                const lThumb = groupThumb(box.node);
                const tl = typeLabel(box.node);
                let peopleHtml = '';
                if (seatLeaf) {
                    if (filled.length) {
                        const shownP = filled.slice(0, 2);
                        const moreP = filled.length - shownP.length;
                        peopleHtml =
                            '<span class="schul-org-map__leaf-people">' +
                            shownP
                                .map(function (p) {
                                    const ul = personHasM365UserLink(p);
                                    const pfHit = personFocusMatchesPerson(p);
                                    return (
                                        '<span class="schul-org-map__leaf-person' +
                                        (ul ? ' is-m365-user' : '') +
                                        (personFocus
                                            ? pfHit
                                                ? ' is-person-hit'
                                                : ' is-person-dimmed'
                                            : '') +
                                        '" title="' +
                                        escapeHtml(
                                            (p.name || p.email || '') +
                                                (ul ? ' · M365-Konto' : '')
                                        ) +
                                        '">' +
                                        personThumb(p.email, p.name || p.label) +
                                        '<span class="schul-org-map__leaf-person-name">' +
                                        escapeHtml(p.name || p.email || '?') +
                                        '</span>' +
                                        personM365MarkHtml(p) +
                                        '</span>'
                                    );
                                })
                                .join('') +
                            (moreP > 0
                                ? '<span class="schul-org-map__leaf-count">+' + moreP + '</span>'
                                : '') +
                            '</span>';
                    } else if (vacantSeat) {
                        peopleHtml =
                            '<span class="schul-org-map__leaf-people schul-org-map__leaf-people--vacant">' +
                            '<span class="muted">' +
                            escapeHtml(vacantSeat.seatLabel || 'offen') +
                            '</span></span>';
                    }
                } else if (filled.length && tl !== 'Funktion') {
                    peopleHtml =
                        '<span class="schul-org-map__leaf-count">' + filled.length + '</span>';
                }
                el.innerHTML =
                    (lThumb
                        ? '<span class="schul-org-map__group-thumb schul-org-map__group-thumb--leaf">' +
                          lThumb +
                          '</span>'
                        : '') +
                    '<span class="schul-org-map__leaf-type">' +
                    escapeHtml(tl) +
                    '</span>' +
                    '<span class="schul-org-map__leaf-label">' +
                    escapeHtml(box.node.label) +
                    (personHint && tl === 'Funktion'
                        ? '<span class="muted"> · ' + escapeHtml(personHint) + '</span>'
                        : '') +
                    '</span>' +
                    m365MarkHtml(leafM365) +
                    peopleHtml +
                    '<button type="button" class="schul-org-map__edit schul-org-map__edit--leaf" data-edit-id="' +
                    escapeHtml(box.node.id) +
                    '" title="Bearbeiten"><i class="bi bi-pencil" aria-hidden="true"></i></button>';
                const names = filled
                    .map(function (p) {
                        return p.name || p.email;
                    })
                    .filter(Boolean)
                    .join(', ');
                el.title =
                    (leafM365.linked ? 'M365 verknüpft · ' : '') +
                    (names
                        ? box.node.label +
                          ': ' +
                          names +
                          (leafDraggable
                              ? ' · Klick = hineinzoomen · Ziehen = verschieben'
                              : ' · Klick = hineinzoomen')
                        : leafDraggable
                          ? 'Klick = hineinzoomen · Ziehen = verschieben · Stift = bearbeiten'
                          : 'Klick = hineinzoomen · Stift = bearbeiten');
                inner.appendChild(el);
            }
            } catch (err) {
                try {
                    console.warn('[schul-org-map] Box übersprungen', box && box.kind, err);
                } catch {
                    /* ignore */
                }
            }
        });

        wireEditAndDrill(computed.boxes);
        afterPaint();
        requestAnimationFrame(function () {
            if (focused && computed.focusBox) {
                if (fitToBox(computed.focusBox)) return;
                applyTransform();
                window.setTimeout(function () {
                    if (focusStack.length && computed.focusBox) fitToBox(computed.focusBox);
                }, 60);
                return;
            }
            const mapHost = host.querySelector('#schulOrgMapHost');
            const hostOk = mapHost && mapHost.clientWidth >= 80 && mapHost.clientHeight >= 80;
            if (needsOverviewFit || !hostOk) {
                if (fitOverview(computed)) {
                    needsOverviewFit = false;
                } else {
                    applyTransform();
                    window.setTimeout(function () {
                        if (!focusStack.length && fitOverview(computed)) needsOverviewFit = false;
                    }, 60);
                }
                return;
            }
            applyTransform();
        });
    }

    function wireEditAndDrill(boxes) {
        host.querySelectorAll('[data-collapse-id]').forEach(function (btn) {
            btn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                const id = btn.getAttribute('data-collapse-id') || '';
                if (!id) return;
                layout = setSchulOrganigrammAreaCollapsed(layout, id);
                layout = setSchulOrganigrammView(layout, view);
                if (typeof opts.onLayoutChange === 'function') opts.onLayoutChange(layout);
                paint();
            });
        });
        host.querySelectorAll('[data-edit-id]').forEach(function (btn) {
            btn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                const id = btn.getAttribute('data-edit-id') || '';
                const node = layout.nodes.find(function (n) {
                    return n.id === id;
                });
                if (node && typeof opts.onEditNode === 'function') opts.onEditNode(node);
            });
        });
        host.querySelectorAll('[data-edit-person]').forEach(function (btn) {
            btn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                const personEl = /** @type {HTMLElement|null} */ (
                    btn.closest('.schul-org-map__person') ||
                    btn.closest('.schul-org-map__preview-chip')
                );
                if (!personEl) return;
                const person = {
                    email: personEl.getAttribute('data-person-email') || '',
                    name: personEl.getAttribute('data-person-name') || '',
                    kind: personEl.getAttribute('data-person-kind') || '',
                    seatLabel: personEl.getAttribute('data-seat-label') || ''
                };
                const parentId =
                    personEl.getAttribute('data-parent-area-id') ||
                    personEl.getAttribute('data-focus-node-id') ||
                    (focusNode() && focusNode().id) ||
                    '';
                const contextNode =
                    layout.nodes.find(function (n) {
                        return n.id === parentId;
                    }) || focusNode();
                if (typeof opts.onEditPerson === 'function') {
                    opts.onEditPerson(person, contextNode);
                    return;
                }
                if (contextNode && typeof opts.onEditNode === 'function') {
                    opts.onEditNode(contextNode);
                }
            });
        });
        host.querySelectorAll('[data-drill-id]').forEach(function (el) {
            el.addEventListener('click', function (ev) {
                if (suppressDrillClick) {
                    suppressDrillClick = false;
                    return;
                }
                if (
                    ev.target.closest('[data-edit-id]') ||
                    ev.target.closest('[data-edit-person]') ||
                    ev.target.closest('[data-collapse-id]')
                ) {
                    return;
                }
                // Kopf mit Drag: Zoom kommt aus pointerup (Klick ohne Bewegung)
                if (
                    el.classList.contains('schul-org-map__area-head') &&
                    el.hasAttribute('data-area-drag')
                ) {
                    return;
                }
                const id = el.getAttribute('data-drill-id') || '';
                drillIntoId(id);
            });
        });
    }

    /**
     * Kurz-Klick → Reinzoomen (gemeinsam für Click-Handler und Drag-ohne-Bewegung).
     * @param {string} id
     */
    function drillIntoId(id) {
        const node = layout.nodes.find(function (n) {
            return n.id === id;
        });
        if (!node || !isDrillableNode(node)) return false;
        enterFocus(node);
        return true;
    }

    function closeContextMenu() {
        if (ctxMenuEl && ctxMenuEl.parentNode) ctxMenuEl.parentNode.removeChild(ctxMenuEl);
        ctxMenuEl = null;
        document.removeEventListener('pointerdown', onCtxOutside, true);
        document.removeEventListener('keydown', onCtxEsc, true);
    }

    function onCtxOutside(ev) {
        if (ctxMenuEl && !ctxMenuEl.contains(/** @type {Node} */ (ev.target))) {
            closeContextMenu();
        }
    }

    function onCtxEsc(ev) {
        if (ev.key === 'Escape') closeContextMenu();
    }

    /**
     * @param {number} clientX
     * @param {number} clientY
     * @param {{ label: string, icon?: string, action: () => void }[]} items
     */
    function openContextMenu(clientX, clientY, items) {
        closeContextMenu();
        if (!items || !items.length) return;
        const menu = document.createElement('div');
        menu.className = 'schul-org-map__ctx';
        menu.setAttribute('role', 'menu');
        menu.innerHTML = items
            .map(function (it, i) {
                return (
                    '<button type="button" class="schul-org-map__ctx-item" role="menuitem" data-ctx-i="' +
                    i +
                    '">' +
                    (it.icon
                        ? '<i class="bi ' + it.icon + '" aria-hidden="true"></i>'
                        : '') +
                    '<span>' +
                    escapeHtml(it.label) +
                    '</span></button>'
                );
            })
            .join('');
        document.body.appendChild(menu);
        ctxMenuEl = menu;
        const pad = 8;
        const mw = menu.offsetWidth;
        const mh = menu.offsetHeight;
        let left = clientX;
        let top = clientY;
        if (left + mw > window.innerWidth - pad) left = Math.max(pad, window.innerWidth - mw - pad);
        if (top + mh > window.innerHeight - pad) top = Math.max(pad, window.innerHeight - mh - pad);
        menu.style.left = left + 'px';
        menu.style.top = top + 'px';
        menu.querySelectorAll('[data-ctx-i]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const i = Number(btn.getAttribute('data-ctx-i'));
                const item = items[i];
                closeContextMenu();
                if (item && typeof item.action === 'function') item.action();
            });
        });
        document.addEventListener('pointerdown', onCtxOutside, true);
        document.addEventListener('keydown', onCtxEsc, true);
    }

    function clientToWorld(mapHost, clientX, clientY) {
        const rect = mapHost.getBoundingClientRect();
        const k = view.k || 1;
        return {
            x: (clientX - rect.left - view.x) / k,
            y: (clientY - rect.top - view.y) / k
        };
    }

    function bindHostOnce() {
        const mapHost = host.querySelector('#schulOrgMapHost');
        if (!mapHost) return;
        // Nach jedem paint neu binden (Host wird ersetzt)
        if (mapHost.dataset.bound === '1') return;
        mapHost.dataset.bound = '1';

        mapHost.addEventListener(
            'wheel',
            function (ev) {
                ev.preventDefault();
                closeContextMenu();
                const delta = ev.deltaY < 0 ? 1.08 : 1 / 1.08;
                const rect = mapHost.getBoundingClientRect();
                const mx = ev.clientX - rect.left;
                const my = ev.clientY - rect.top;
                const nx = mx - (mx - view.x) * delta;
                const ny = my - (my - view.y) * delta;
                view.k = Math.min(2.5, Math.max(0.35, view.k * delta));
                view.x = nx;
                view.y = ny;
                applyTransform();
                scheduleViewSave();
            },
            { passive: false }
        );

        mapHost.addEventListener('contextmenu', function (ev) {
            ev.preventDefault();
            const world = clientToWorld(mapHost, ev.clientX, ev.clientY);
            const leafEl = /** @type {HTMLElement|null} */ (ev.target.closest('.schul-org-map__leaf'));
            const areaEl = /** @type {HTMLElement|null} */ (ev.target.closest('.schul-org-map__area'));
            const personEl = /** @type {HTMLElement|null} */ (
                ev.target.closest('.schul-org-map__person')
            );
            /** @type {{ label: string, icon?: string, action: () => void }[]} */
            const items = [];

            function pushEdit(node) {
                if (!node || typeof opts.onEditNode !== 'function') return;
                items.push({
                    label: isAreaNode(node) ? 'Bereich bearbeiten' : 'Bearbeiten',
                    icon: 'bi-pencil-square',
                    action: function () {
                        opts.onEditNode(node);
                    }
                });
            }

            function pushAreaActions(areaNode) {
                if (!areaNode) return;
                pushEdit(areaNode);
                const contentN = countAreaMapContents(layout, areaNode.id);
                const peopleN = peopleCountForLayout(areaNode);
                if (contentN > 0 || peopleN > 0 || areaNode.collapsed === true) {
                    const collapsed = areaNode.collapsed === true;
                    items.push({
                        label: collapsed ? 'Ausklappen' : 'Einklappen',
                        icon: collapsed ? 'bi-chevron-right' : 'bi-chevron-down',
                        action: function () {
                            layout = setSchulOrganigrammAreaCollapsed(layout, areaNode.id);
                            layout = setSchulOrganigrammView(layout, view);
                            if (typeof opts.onLayoutChange === 'function') {
                                opts.onLayoutChange(layout);
                            }
                            paint();
                        }
                    });
                }
                if (typeof opts.onAddPerson === 'function') {
                    items.push({
                        label: 'Person hinzufügen',
                        icon: 'bi-person-plus',
                        action: function () {
                            void opts.onAddPerson(areaNode);
                        }
                    });
                }
                if (typeof opts.onCreateRole === 'function') {
                    items.push({
                        label: 'Funktion hinzufügen',
                        icon: 'bi-person-vcard',
                        action: function () {
                            void opts.onCreateRole(areaNode);
                        }
                    });
                }
                if (typeof opts.onCreateArea === 'function') {
                    items.push({
                        label: 'Unterbereich anlegen',
                        icon: 'bi-plus-square',
                        action: function () {
                            void opts.onCreateArea({
                                parentId: areaNode.id
                            });
                        }
                    });
                }
            }

            if (personEl) {
                const person = {
                    email: personEl.getAttribute('data-person-email') || '',
                    name: personEl.getAttribute('data-person-name') || '',
                    kind: personEl.getAttribute('data-person-kind') || '',
                    seatLabel: personEl.getAttribute('data-seat-label') || ''
                };
                const parentId =
                    personEl.getAttribute('data-parent-area-id') ||
                    personEl.getAttribute('data-focus-node-id') ||
                    (focusNode() && focusNode().id) ||
                    '';
                const contextNode =
                    layout.nodes.find(function (n) {
                        return n.id === parentId;
                    }) || focusNode();
                const isLeadSeat = String(person.kind).toLowerCase() === 'head';
                const leadVacant = isLeadSeat && !(person.email || person.name);
                if (typeof opts.onEditPerson === 'function') {
                    items.push({
                        label: leadVacant
                            ? (person.seatLabel || 'Leitung') + ' zuordnen'
                            : isLeadSeat
                              ? (person.seatLabel || 'Leitung') + ' bearbeiten'
                              : 'Person bearbeiten',
                        icon: leadVacant ? 'bi-person-plus' : 'bi-person-gear',
                        action: function () {
                            opts.onEditPerson(person, contextNode);
                        }
                    });
                }
                if (contextNode && typeof opts.onEditNode === 'function') {
                    items.push({
                        label: isAreaNode(contextNode)
                            ? 'Bereich bearbeiten'
                            : contextNode.type === NODE.CLASS
                              ? 'Klasse bearbeiten'
                              : contextNode.type === NODE.ARGE
                                ? 'ARGE bearbeiten'
                                : contextNode.type === NODE.ROLE
                                  ? 'Rolle bearbeiten'
                                  : 'Eintrag bearbeiten',
                        icon: 'bi-pencil-square',
                        action: function () {
                            opts.onEditNode(contextNode);
                        }
                    });
                }
            } else if (leafEl && !areaEl) {
                const id = leafEl.getAttribute('data-node-id') || leafEl.dataset.drillId || '';
                const node = layout.nodes.find(function (n) {
                    return n.id === id;
                });
                pushEdit(node);
                // Leaf in einem Bereich: Personen/Unterbereich am Parent anbieten
                const parentId = leafEl.getAttribute('data-parent-area-id') || '';
                const parentArea = parentId
                    ? layout.nodes.find(function (n) {
                          return n.id === parentId && isAreaNode(n);
                      })
                    : null;
                if (parentArea) {
                    if (typeof opts.onAddPerson === 'function') {
                        items.push({
                            label: 'Person in Bereich',
                            icon: 'bi-person-plus',
                            action: function () {
                                void opts.onAddPerson(parentArea);
                            }
                        });
                    }
                    if (typeof opts.onCreateRole === 'function') {
                        items.push({
                            label: 'Funktion hinzufügen',
                            icon: 'bi-person-vcard',
                            action: function () {
                                void opts.onCreateRole(parentArea);
                            }
                        });
                    }
                    if (typeof opts.onCreateArea === 'function') {
                        items.push({
                            label: 'Unterbereich anlegen',
                            icon: 'bi-plus-square',
                            action: function () {
                                void opts.onCreateArea({
                                    parentId: parentArea.id
                                });
                            }
                        });
                    }
                }
            } else if (areaEl) {
                const id = areaEl.getAttribute('data-area-id') || '';
                const areaNode = layout.nodes.find(function (n) {
                    return n.id === id && isAreaNode(n);
                });
                pushAreaActions(areaNode);
            } else if (!focusStack.length) {
                if (typeof opts.onCreateArea === 'function') {
                    items.push({
                        label: 'Neuer Bereich',
                        icon: 'bi-plus-lg',
                        action: function () {
                            void opts.onCreateArea({
                                parentId: null,
                                x: Math.round(world.x),
                                y: Math.round(world.y)
                            });
                        }
                    });
                }
                if (typeof opts.onCreateRole === 'function') {
                    items.push({
                        label: 'Funktion hinzufügen…',
                        icon: 'bi-person-vcard',
                        action: function () {
                            void opts.onCreateRole(null);
                        }
                    });
                }
                if (typeof opts.onAddPerson === 'function') {
                    items.push({
                        label: 'Person (Bereich wählen)…',
                        icon: 'bi-person-plus',
                        action: function () {
                            void opts.onAddPerson(null);
                        }
                    });
                }
            }

            openContextMenu(ev.clientX, ev.clientY, items);
        });

        mapHost.addEventListener('click', function (ev) {
            if (!focusStack.length) return;
            if (suppressBackgroundClick) {
                suppressBackgroundClick = false;
                return;
            }
            if (
                ev.target.closest('[data-edit-id]') ||
                ev.target.closest('[data-drill-id]') ||
                ev.target.closest('[data-focus-back]') ||
                ev.target.closest('button') ||
                ev.target.closest('.schul-org-map__area--expanded') ||
                ev.target.closest('.schul-org-map__person') ||
                ev.target.closest('.schul-org-map__leaf--inside') ||
                ev.target.closest('.schul-org-map__focus-child')
            ) {
                return;
            }
            // Klick auf Leinwand / außerhalb des geöffneten Bereichs → auszoomen
            exitFocus();
        });

        mapHost.addEventListener('pointerdown', function (ev) {
            if (ev.target.closest('[data-edit-id]') || ev.target.closest('button') || ev.target.closest('[data-edit-person]'))
                return;
            const resizeHandle = /** @type {HTMLElement|null} */ (
                ev.target.closest('[data-area-resize]')
            );
            if (resizeHandle && !focusStack.length) {
                const id = resizeHandle.getAttribute('data-area-resize') || '';
                const edge = resizeHandle.getAttribute('data-resize-edge') || 'se';
                const areaEl = /** @type {HTMLElement|null} */ (
                    resizeHandle.closest('.schul-org-map__area')
                );
                const box = lastOverviewBoxes.find(function (b) {
                    return b.kind === 'area' && b.node && b.node.id === id;
                });
                const startW = box
                    ? box.w
                    : areaEl
                      ? parseFloat(areaEl.style.width) || AREA_W
                      : AREA_W;
                const startH = box
                    ? box.h
                    : areaEl
                      ? parseFloat(areaEl.style.height) || AREA_RESIZE_MIN_H
                      : AREA_RESIZE_MIN_H;
                areaResize = {
                    id: id,
                    edge: edge,
                    pointerId: ev.pointerId,
                    startClientX: ev.clientX,
                    startClientY: ev.clientY,
                    startW: startW,
                    startH: startH,
                    lastW: startW,
                    lastH: startH,
                    moved: false
                };
                if (areaEl) areaEl.classList.add('is-resizing');
                mapHost.classList.add('is-area-resizing');
                try {
                    mapHost.setPointerCapture(ev.pointerId);
                } catch {
                    /* ignore */
                }
                ev.preventDefault();
                ev.stopPropagation();
                return;
            }
            const head = /** @type {HTMLElement|null} */ (ev.target.closest('[data-area-drag]'));
            if (head && !focusStack.length) {
                const id = head.getAttribute('data-area-drag') || '';
                const areaEl = head.closest('.schul-org-map__area');
                const satellites = collectDragSatellites(id);
                const blocked = new Set([id]);
                satellites.forEach(function (s) {
                    const sid = s.el.getAttribute('data-area-id') || s.el.getAttribute('data-node-id') || '';
                    if (sid) blocked.add(sid);
                });
                // Pending: erst ab Schwelle als Drag; darunter = Klick/Zoom
                areaDrag = {
                    id: id,
                    pointerId: ev.pointerId,
                    startClientX: ev.clientX,
                    startClientY: ev.clientY,
                    startX: areaEl ? parseFloat(areaEl.style.left) || 0 : 0,
                    startY: areaEl ? parseFloat(areaEl.style.top) || 0 : 0,
                    moved: false,
                    satellites: satellites,
                    blockedIds: blocked,
                    areaEl: areaEl
                };
                try {
                    mapHost.setPointerCapture(ev.pointerId);
                } catch {
                    /* ignore */
                }
                ev.preventDefault();
                return;
            }
            const personEl = /** @type {HTMLElement|null} */ (
                ev.target.closest('[data-person-drag]')
            );
            if (personEl) {
                // Vorschau-Chips sitzen im Area (overflow:hidden) → zum Ziehen auf die Map heben
                const mapInner = /** @type {HTMLElement|null} */ (
                    host.querySelector('#schulOrgMapInner')
                );
                let startX = parseFloat(personEl.style.left);
                let startY = parseFloat(personEl.style.top);
                if (
                    mapInner &&
                    (!Number.isFinite(startX) ||
                        !Number.isFinite(startY) ||
                        personEl.classList.contains('schul-org-map__preview-chip'))
                ) {
                    const chipRect = personEl.getBoundingClientRect();
                    const world = clientToWorld(mapHost, chipRect.left, chipRect.top);
                    const k0 = view.k || 1;
                    startX = world.x;
                    startY = world.y;
                    personEl.style.position = 'absolute';
                    personEl.style.left = startX + 'px';
                    personEl.style.top = startY + 'px';
                    personEl.style.width = Math.max(120, Math.round(chipRect.width / k0)) + 'px';
                    personEl.style.zIndex = '30';
                    personEl.classList.add('schul-org-map__preview-chip--floating');
                    if (personEl.parentElement !== mapInner) {
                        mapInner.appendChild(personEl);
                    }
                }
                if (!Number.isFinite(startX)) startX = 0;
                if (!Number.isFinite(startY)) startY = 0;
                personDrag = {
                    pointerId: ev.pointerId,
                    startClientX: ev.clientX,
                    startClientY: ev.clientY,
                    startX: startX,
                    startY: startY,
                    moved: false,
                    el: personEl,
                    person: {
                        kind: personEl.getAttribute('data-person-kind') || 'admin',
                        name: personEl.getAttribute('data-person-name') || '',
                        email: personEl.getAttribute('data-person-email') || '',
                        label: personEl.getAttribute('data-person-name') || '',
                        seatLabel: personEl.getAttribute('data-seat-label') || '',
                        sourceNodeId: personEl.getAttribute('data-source-node-id') || '',
                        sourceRefKind: personEl.getAttribute('data-source-ref-kind') || '',
                        sourceRefKey: personEl.getAttribute('data-source-ref-key') || ''
                    },
                    blockedIds: new Set(
                        [
                            personEl.getAttribute('data-source-node-id') || '',
                            personEl.getAttribute('data-focus-node-id') || '',
                            personEl.getAttribute('data-parent-area-id') || ''
                        ].filter(Boolean)
                    )
                };
                try {
                    mapHost.setPointerCapture(ev.pointerId);
                } catch {
                    /* ignore */
                }
                ev.preventDefault();
                return;
            }
            const leafEl = /** @type {HTMLElement|null} */ (ev.target.closest('[data-leaf-drag]'));
            if (leafEl) {
                const id = leafEl.getAttribute('data-leaf-drag') || leafEl.getAttribute('data-node-id') || '';
                leafDrag = {
                    id: id,
                    pointerId: ev.pointerId,
                    startClientX: ev.clientX,
                    startClientY: ev.clientY,
                    startX: parseFloat(leafEl.style.left) || 0,
                    startY: parseFloat(leafEl.style.top) || 0,
                    moved: false,
                    el: leafEl,
                    blockedIds: new Set([id])
                };
                try {
                    mapHost.setPointerCapture(ev.pointerId);
                } catch {
                    /* ignore */
                }
                ev.preventDefault();
                return;
            }
            // Drillable Leaf ohne Drag-Handle: Pan nicht starten (Klick = Zoom)
            if (ev.target.closest('.schul-org-map__leaf') || ev.target.closest('.schul-org-map__person')) {
                return;
            }
            if (ev.target.closest('.schul-org-map__area-body')) return;
            // Klassen-Nest / Area ohne Kopf-Drag: Klick = Zoom, kein Pan vom Body/Kopf
            if (ev.target.closest('.schul-org-map__area.is-drillable')) return;
            panSession = {
                x: ev.clientX,
                y: ev.clientY,
                tx: view.x,
                ty: view.y,
                pointerId: ev.pointerId
            };
            mapHost.classList.add('is-panning');
            try {
                mapHost.setPointerCapture(ev.pointerId);
            } catch {
                /* ignore */
            }
            // Sonst markiert der Browser beim Verschieben Text (blaue Selection)
            ev.preventDefault();
            try {
                const sel = window.getSelection && window.getSelection();
                if (sel && sel.removeAllRanges) sel.removeAllRanges();
            } catch {
                /* ignore */
            }
        });

        mapHost.addEventListener('selectstart', function (ev) {
            // Textauswahl auf der Landkarte unterbinden (Pan/Drag)
            ev.preventDefault();
        });

        mapHost.addEventListener('pointermove', function (ev) {
            if (areaResize && ev.pointerId === areaResize.pointerId) {
                const k = view.k || 1;
                const dx = (ev.clientX - areaResize.startClientX) / k;
                const dy = (ev.clientY - areaResize.startClientY) / k;
                const dist = Math.hypot(
                    ev.clientX - areaResize.startClientX,
                    ev.clientY - areaResize.startClientY
                );
                if (!areaResize.moved && dist >= DRAG_THRESHOLD_PX) {
                    areaResize.moved = true;
                }
                if (!areaResize.moved) return;
                let nextW = areaResize.startW;
                let nextH = areaResize.startH;
                const edge = areaResize.edge || 'se';
                if (edge === 'e' || edge === 'se') nextW = areaResize.startW + dx;
                if (edge === 's' || edge === 'se') nextH = areaResize.startH + dy;
                const clamped = clampSchulOrganigrammAreaSize(nextW, nextH);
                const w = clamped.w != null ? clamped.w : areaResize.startW;
                const h = clamped.h != null ? clamped.h : areaResize.startH;
                if (w === areaResize.lastW && h === areaResize.lastH) return;
                areaResize.lastW = w;
                areaResize.lastH = h;
                layout = setSchulOrganigrammNodeSize(layout, areaResize.id, w, h);
                scheduleAreaResizePaint();
                return;
            }
            if (personDrag && ev.pointerId === personDrag.pointerId) {
                const k = view.k || 1;
                const dx = (ev.clientX - personDrag.startClientX) / k;
                const dy = (ev.clientY - personDrag.startClientY) / k;
                const dist = Math.hypot(
                    ev.clientX - personDrag.startClientX,
                    ev.clientY - personDrag.startClientY
                );
                if (!personDrag.moved && dist >= DRAG_THRESHOLD_PX) {
                    personDrag.moved = true;
                    personDrag.el.classList.add('is-dragging');
                    mapHost.classList.add('is-person-dragging');
                }
                if (!personDrag.moved) return;
                personDrag.el.style.left = personDrag.startX + dx + 'px';
                personDrag.el.style.top = personDrag.startY + dy + 'px';
                updatePersonDropHighlight(ev.clientX, ev.clientY, personDrag.blockedIds);
                return;
            }
            if (leafDrag && ev.pointerId === leafDrag.pointerId) {
                const k = view.k || 1;
                const dx = (ev.clientX - leafDrag.startClientX) / k;
                const dy = (ev.clientY - leafDrag.startClientY) / k;
                const dist = Math.hypot(
                    ev.clientX - leafDrag.startClientX,
                    ev.clientY - leafDrag.startClientY
                );
                if (!leafDrag.moved && dist >= DRAG_THRESHOLD_PX) {
                    leafDrag.moved = true;
                    leafDrag.el.classList.add('is-dragging');
                    mapHost.classList.add('is-leaf-dragging');
                    // Nach oben heben, damit die gezogene Kachel überall sichtbar bleibt
                    const mapInner = host.querySelector('#schulOrgMapInner');
                    if (mapInner && leafDrag.el.parentElement === mapInner) {
                        mapInner.appendChild(leafDrag.el);
                    }
                    leafDrag.el.style.zIndex = '80';
                    leafDrag.el.style.pointerEvents = 'none';
                }
                if (!leafDrag.moved) return;
                // Frei verschieben (auch in andere Bereiche); Einrasten erst beim Drop
                leafDrag.el.style.left = leafDrag.startX + dx + 'px';
                leafDrag.el.style.top = leafDrag.startY + dy + 'px';
                updateDropHighlight(ev.clientX, ev.clientY, leafDrag.id, leafDrag.blockedIds);
                return;
            }
            if (areaDrag && ev.pointerId === areaDrag.pointerId) {
                const k = view.k || 1;
                const dx = (ev.clientX - areaDrag.startClientX) / k;
                const dy = (ev.clientY - areaDrag.startClientY) / k;
                const dist = Math.hypot(
                    ev.clientX - areaDrag.startClientX,
                    ev.clientY - areaDrag.startClientY
                );
                if (!areaDrag.moved && dist >= DRAG_THRESHOLD_PX) {
                    areaDrag.moved = true;
                    (areaDrag.satellites || []).forEach(function (s) {
                        s.el.classList.add('is-dragging');
                    });
                    if (areaDrag.areaEl) areaDrag.areaEl.classList.add('is-dragging');
                    mapHost.classList.add('is-area-dragging');
                }
                if (!areaDrag.moved) return;
                const sats = areaDrag.satellites || [];
                if (sats.length) {
                    sats.forEach(function (s) {
                        s.el.style.left = s.startX + dx + 'px';
                        s.el.style.top = s.startY + dy + 'px';
                    });
                } else if (areaDrag.areaEl) {
                    areaDrag.areaEl.style.left = areaDrag.startX + dx + 'px';
                    areaDrag.areaEl.style.top = areaDrag.startY + dy + 'px';
                }
                updateDropHighlight(
                    ev.clientX,
                    ev.clientY,
                    areaDrag.id,
                    areaDrag.blockedIds
                );
                return;
            }
            if (!panSession || ev.pointerId !== panSession.pointerId) return;
            const pdx = ev.clientX - panSession.x;
            const pdy = ev.clientY - panSession.y;
            if (Math.abs(pdx) > 4 || Math.abs(pdy) > 4) suppressBackgroundClick = true;
            view.x = panSession.tx + pdx;
            view.y = panSession.ty + pdy;
            applyTransform();
        });

        function endPointer(ev) {
            if (areaResize && ev.pointerId === areaResize.pointerId) {
                const session = areaResize;
                areaResize = null;
                if (areaResizeRaf) {
                    window.cancelAnimationFrame(areaResizeRaf);
                    areaResizeRaf = 0;
                }
                mapHost.classList.remove('is-area-resizing');
                host.querySelectorAll('.schul-org-map__area.is-resizing').forEach(function (el) {
                    el.classList.remove('is-resizing');
                });
                try {
                    mapHost.releasePointerCapture(ev.pointerId);
                } catch {
                    /* ignore */
                }
                if (!session.moved) {
                    return;
                }
                suppressBackgroundClick = true;
                suppressDrillClick = true;
                const clamped = clampSchulOrganigrammAreaSize(session.lastW, session.lastH);
                layout = setSchulOrganigrammNodeSize(
                    layout,
                    session.id,
                    clamped.w != null ? clamped.w : session.startW,
                    clamped.h != null ? clamped.h : session.startH
                );
                const snapshot = computeLandkarteLayout(layout, {
                    peopleCountForArea: peopleCountForLayout
                });
                layout = applyLandkartePositionsFromBoxes(layout, snapshot.boxes, {
                    rootsOnly: true
                });
                layout = setSchulOrganigrammView(layout, view);
                if (typeof opts.onLayoutChange === 'function') opts.onLayoutChange(layout);
                paint();
                return;
            }
            if (personDrag && ev.pointerId === personDrag.pointerId) {
                const drag = personDrag;
                personDrag = null;
                clearDropHighlight();
                mapHost.classList.remove('is-person-dragging');
                if (drag.el) drag.el.classList.remove('is-dragging');
                try {
                    mapHost.releasePointerCapture(ev.pointerId);
                } catch {
                    /* ignore */
                }
                if (!drag.moved) {
                    // Kurz-Klick auf Person → Bearbeiten (Zoom bleibt über Bereich/Rolle)
                    suppressDrillClick = true;
                    const parentId =
                        (drag.el &&
                            (drag.el.getAttribute('data-parent-area-id') ||
                                drag.el.getAttribute('data-focus-node-id'))) ||
                        '';
                    const contextNode =
                        layout.nodes.find(function (n) {
                            return n.id === parentId;
                        }) || focusNode();
                    if (typeof opts.onEditPerson === 'function') {
                        opts.onEditPerson(drag.person, contextNode);
                    }
                    return;
                }
                suppressBackgroundClick = true;
                suppressDrillClick = true;
                const hit = resolvePersonDropAt(ev.clientX, ev.clientY, drag.blockedIds);
                if (hit) {
                    const target = layout.nodes.find(function (n) {
                        return n.id === hit.id;
                    });
                    if (target && typeof opts.onMovePerson === 'function') {
                        opts.onMovePerson(drag.person, target);
                        // UI rendert ggf. neu; Layout aus Settings nachladen
                        const cur =
                            typeof window !== 'undefined' &&
                            typeof window.ms365TenantSettingsLoad === 'function'
                                ? window.ms365TenantSettingsLoad()
                                : null;
                        if (cur && cur.schulOrganigramm) {
                            layout = normalizeSchulOrganigramm(cur.schulOrganigramm);
                        }
                    }
                }
                paint();
                return;
            }
            if (leafDrag && ev.pointerId === leafDrag.pointerId) {
                const drag = leafDrag;
                leafDrag = null;
                clearDropHighlight();
                mapHost.classList.remove('is-leaf-dragging');
                if (drag.el) drag.el.classList.remove('is-dragging');
                try {
                    mapHost.releasePointerCapture(ev.pointerId);
                } catch {
                    /* ignore */
                }
                if (!drag.moved) {
                    suppressDrillClick = true;
                    drillIntoId(drag.id);
                    return;
                }
                suppressBackgroundClick = true;
                suppressDrillClick = true;
                const blocked = drag.blockedIds || new Set([String(drag.id || '')]);
                const nest = resolveNestDropAt(drag.id, ev.clientX, ev.clientY, blocked);
                const nodeNow = layout.nodes.find(function (n) {
                    return n.id === drag.id;
                });
                const curParent =
                    nodeNow && nodeNow.parentId != null && nodeNow.parentId !== ''
                        ? String(nodeNow.parentId)
                        : '';
                let dropId = nest.mode !== 'keep' ? nest.parentId || '' : '';
                // Keep / kein Treffer: Live-DOM (auch gleicher Parent = Umsortieren)
                if (!dropId) {
                    const live = pickLiveDropAreaAt(ev.clientX, ev.clientY, blocked);
                    if (live && live.id && live.id !== drag.id) {
                        dropId = live.id;
                    }
                }
                if (!dropId) dropId = curParent;
                const world = clientToWorld(mapHost, ev.clientX, ev.clientY);
                const beforeId = dropId
                    ? findInsertBeforeSiblingId(
                          lastOverviewBoxes,
                          dropId,
                          drag.id,
                          world.y,
                          'leaf'
                      )
                    : '';
                if (dropId) {
                    let next = null;
                    if (dropId !== curParent) {
                        if (typeof opts.onMoveNode === 'function') {
                            next = opts.onMoveNode(drag.id, dropId, layout);
                        }
                        if (!next || !next.nodes) {
                            next = moveSchulOrganigrammNode(layout, drag.id, dropId, beforeId);
                        } else {
                            // Stammdaten-Move ok → Reihenfolge im Zielbereich nach Drop-Y
                            next = moveSchulOrganigrammNode(next, drag.id, dropId, beforeId);
                        }
                    } else {
                        // Gleicher Bereich: nur neu einreihen + packen
                        next = moveSchulOrganigrammNode(layout, drag.id, dropId, beforeId);
                    }
                    if (next && next.nodes) {
                        const after = next.nodes.find(function (n) {
                            return n.id === drag.id;
                        });
                        if (after && String(after.parentId || '') === String(dropId)) {
                            layout = normalizeSchulOrganigramm(next);
                            layout = setSchulOrganigrammView(layout, view);
                            if (typeof opts.onLayoutChange === 'function') {
                                opts.onLayoutChange(layout);
                            }
                        }
                    }
                }
                paint();
                return;
            }
            if (areaDrag && ev.pointerId === areaDrag.pointerId) {
                const drag = areaDrag;
                areaDrag = null;
                clearDropHighlight();
                mapHost.classList.remove('is-area-dragging');
                const areaEl =
                    drag.areaEl ||
                    host.querySelector(
                        '.schul-org-map__area[data-area-id="' + CSS.escape(drag.id) + '"]'
                    );
                (drag.satellites || []).forEach(function (s) {
                    s.el.classList.remove('is-dragging');
                });
                if (areaEl) areaEl.classList.remove('is-dragging');
                try {
                    mapHost.releasePointerCapture(ev.pointerId);
                } catch {
                    /* ignore */
                }
                if (!drag.moved) {
                    suppressDrillClick = true;
                    drillIntoId(drag.id);
                    return;
                }
                suppressBackgroundClick = true;
                suppressDrillClick = true;
                const k = view.k || 1;
                const nx = drag.startX + (ev.clientX - drag.startClientX) / k;
                const ny = drag.startY + (ev.clientY - drag.startClientY) / k;

                const blocked = drag.blockedIds || new Set([String(drag.id || '')]);
                const nest = resolveNestDropAt(drag.id, ev.clientX, ev.clientY, blocked);
                const dropId = nest.parentId;
                const nodeBefore = layout.nodes.find(function (n) {
                    return n.id === drag.id;
                });
                const prevParent =
                    nodeBefore && nodeBefore.parentId != null && nodeBefore.parentId !== ''
                        ? String(nodeBefore.parentId)
                        : '';

                const worldArea = clientToWorld(mapHost, ev.clientX, ev.clientY);
                const dropX = Math.round(nx);
                const dropY = Math.round(ny);
                let freeRootPlace = false;
                if (nest.mode === 'nest' || nest.mode === 'promote') {
                    const beforeNest = findInsertBeforeSiblingId(
                        lastOverviewBoxes,
                        dropId || '',
                        drag.id,
                        worldArea.y,
                        'area'
                    );
                    layout = moveSchulOrganigrammArea(layout, drag.id, dropId, beforeNest);
                    // Nested: freie Koordinaten weg – Packer richtet unter dem Kopf aus
                } else if (nest.mode === 'root' || !prevParent) {
                    // Loslösen auf die Landkarte (höhere Ebene / Root) oder Root verschieben
                    if (prevParent || nest.mode === 'root') {
                        layout = moveSchulOrganigrammArea(layout, drag.id, null);
                    }
                    // Exakte Drop-Position → Stammdaten (Anordnung bleibt beim nächsten Öffnen)
                    layout = setSchulOrganigrammNodePosition(layout, drag.id, dropX, dropY);
                    freeRootPlace = true;
                } else {
                    // keep: im gleichen Parent – alternativ anderer Bereich unter dem Zeiger?
                    const domId = resolveLandkarteDropAreaId(
                        ev.clientX,
                        ev.clientY,
                        blocked,
                        host
                    );
                    if (domId && domId !== drag.id && domId !== prevParent) {
                        const beforeDom = findInsertBeforeSiblingId(
                            lastOverviewBoxes,
                            domId,
                            drag.id,
                            worldArea.y,
                            'area'
                        );
                        layout = moveSchulOrganigrammArea(layout, drag.id, domId, beforeDom);
                    } else {
                        const beforeKeep = findInsertBeforeSiblingId(
                            lastOverviewBoxes,
                            prevParent,
                            drag.id,
                            worldArea.y,
                            'area'
                        );
                        layout = moveSchulOrganigrammArea(
                            layout,
                            drag.id,
                            prevParent,
                            beforeKeep
                        );
                    }
                }
                // Alle sichtbaren Root-Orte aus dem DOM sichern (nicht nur neu berechnen)
                const fromDom = collectRootBoxesFromDom();
                if (fromDom.length) {
                    layout = applyLandkartePositionsFromBoxes(layout, fromDom, {
                        rootsOnly: true
                    });
                } else {
                    const snapshot = computeLandkarteLayout(layout, {
                        peopleCountForArea: peopleCountForLayout
                    });
                    layout = applyLandkartePositionsFromBoxes(layout, snapshot.boxes, {
                        rootsOnly: true
                    });
                }
                if (freeRootPlace) {
                    // Drop-Ziel hat Vorrang, falls Packer/Messung abweicht
                    layout = setSchulOrganigrammNodePosition(layout, drag.id, dropX, dropY);
                }
                layout = setSchulOrganigrammView(layout, view);
                if (typeof opts.onLayoutChange === 'function') opts.onLayoutChange(layout);
                paint();
                return;
            }
            if (!panSession || ev.pointerId !== panSession.pointerId) return;
            panSession = null;
            mapHost.classList.remove('is-panning');
            scheduleViewSave();
            try {
                mapHost.releasePointerCapture(ev.pointerId);
            } catch {
                /* ignore */
            }
        }
        mapHost.addEventListener('pointerup', endPointer);
        mapHost.addEventListener('pointercancel', endPointer);
    }

    paint();

    /**
     * Sichtbare Root-Bereichs-Geometrie aus dem DOM lesen (was der Nutzer gerade sieht).
     * Wichtiger als computeLandkarteLayout: nach Drag kann die Anzeige der Wahrheit entsprechen.
     * @returns {object[]}
     */
    function collectRootBoxesFromDom() {
        /** @type {object[]} */
        const boxes = [];
        const inner = host.querySelector('#schulOrgMapInner');
        if (!inner) return boxes;
        inner.querySelectorAll('.schul-org-map__area[data-area-id]').forEach(function (el) {
            const htmlEl = /** @type {HTMLElement} */ (el);
            // Nested Areas haben data-parent-area-id – nur Top-Level persistieren
            if (htmlEl.getAttribute('data-parent-area-id')) return;
            if (htmlEl.classList.contains('schul-org-map__area--class-nest')) return;
            const id = htmlEl.getAttribute('data-area-id') || '';
            if (!id) return;
            const node = layout.nodes.find(function (n) {
                return n.id === id && isAreaNode(n);
            });
            if (!node) return;
            if (node.parentId != null && node.parentId !== '') return;
            const x = parseFloat(htmlEl.style.left);
            const y = parseFloat(htmlEl.style.top);
            const w = parseFloat(htmlEl.style.width);
            const h = parseFloat(htmlEl.style.height);
            if (!Number.isFinite(x) || !Number.isFinite(y)) return;
            boxes.push({
                kind: 'area',
                node: node,
                x: Math.round(x),
                y: Math.round(y),
                w: Number.isFinite(w) ? Math.round(w) : undefined,
                h: Number.isFinite(h) ? Math.round(h) : undefined,
                parentAreaId: null
            });
        });
        return boxes;
    }

    /**
     * Aktuelle Box-Positionen (Root absolut) + Viewport ins Layout schreiben.
     * Vor Speichern / Remount: zuerst DOM (sichtbare Orte), sonst berechnete Boxes.
     */
    function flushLandkartePositions() {
        const fromDom = collectRootBoxesFromDom();
        if (fromDom.length) {
            layout = applyLandkartePositionsFromBoxes(layout, fromDom, { rootsOnly: true });
            // Fenstergröße nur fortschreiben, wenn der Nutzer sie schon gesetzt hat
            fromDom.forEach(function (b) {
                if (!b || !b.node || !b.node.id) return;
                if (b.w == null || b.h == null) return;
                const node = layout.nodes.find(function (n) {
                    return n.id === b.node.id;
                });
                if (!node || (node.w == null && node.h == null)) return;
                layout = setSchulOrganigrammNodeSize(layout, b.node.id, b.w, b.h);
            });
        } else {
            const snapshot = computeLandkarteLayout(layout, {
                peopleCountForArea: peopleCountForLayout
            });
            layout = applyLandkartePositionsFromBoxes(layout, snapshot.boxes, { rootsOnly: true });
        }
        layout = setSchulOrganigrammView(layout, view);
        return layout;
    }

    return {
        setLayout: function (next) {
            layout = normalizeSchulOrganigramm(next);
            focusPeopleCache.clear();
            if (!focusStack.length) {
                view = Object.assign(view, layout.view || {});
            }
            paint();
        },
        getLayout: function () {
            return setSchulOrganigrammView(layout, view);
        },
        /**
         * Nach display:none → visible: Viewport/Fit nachziehen ohne Layout-Reset.
         */
        relayout: function () {
            const mapHost = host.querySelector('#schulOrgMapHost');
            if (!mapHost) return;
            const cw = mapHost.clientWidth;
            const ch = mapHost.clientHeight;
            if (cw < 80 || ch < 80) return;
            if (!focusStack.length) {
                needsOverviewFit = true;
                paint();
            }
        },
        flushLandkartePositions: flushLandkartePositions,
        /**
         * Viewport auf Weltkoordinate zentrieren (z. B. neuer Bereich).
         * @param {number} wx
         * @param {number} wy
         * @param {{ w?: number, h?: number }} [size]
         */
        focusWorldPoint: function (wx, wy, size) {
            const mapHost = host.querySelector('#schulOrgMapHost');
            if (!mapHost) return false;
            const cwHost = mapHost.clientWidth;
            const chHost = mapHost.clientHeight;
            if (cwHost < 80 || chHost < 80) return false;
            const bw = Math.max(120, Number(size && size.w) || 280);
            const bh = Math.max(80, Number(size && size.h) || 160);
            return fitToBox({
                x: Number(wx) || 0,
                y: Number(wy) || 0,
                w: bw,
                h: bh
            });
        },
        getFocusStack: function () {
            return focusStack.slice();
        },
        enterFocus: function (nodeId) {
            const node = layout.nodes.find(function (n) {
                return n.id === nodeId;
            });
            if (node) enterFocus(node, null);
        },
        exitFocus: exitFocusAll,
        setPersonFocus: function (next) {
            personFocus = normalizePersonFocusOpt(next);
            paint();
        },
        destroy: function () {
            if (viewSaveTimer) window.clearTimeout(viewSaveTimer);
            if (paintRaf) {
                window.cancelAnimationFrame(paintRaf);
                paintRaf = 0;
            }
            if (areaResizeRaf) {
                window.cancelAnimationFrame(areaResizeRaf);
                areaResizeRaf = 0;
            }
            areaResize = null;
            closeContextMenu();
            host.replaceChildren();
        }
    };
}
