/**
 * SVG/Canvas-Größe nach Blockverschiebung aktualisieren (ohne vollständiges Re-Render).
 */

/**
 * @param {object} state
 */
export function syncDatenlandkarteCanvasDom(state) {
    const layout = state.layout;
    if (!layout) return;
    const inner = document.getElementById('dlCanvasInner');
    const svg = inner && inner.querySelector('.dl-svg');
    if (!inner || !svg) return;

    inner.style.width = `${layout.width}px`;
    inner.style.height = `${layout.height}px`;
    svg.setAttribute('width', String(layout.width));
    svg.setAttribute('height', String(layout.height));

    const sel = state.selectedBlockId || '';
    const selLink = state.selectedLinkId || '';
    const paths = svg.querySelectorAll('path[data-dl-link]');
    const linkMap = new Map((layout.links || []).map((l) => [l.id, l]));
    paths.forEach((el) => {
        const id = el.getAttribute('data-dl-link');
        const l = id ? linkMap.get(id) : null;
        if (!l) return;
        el.setAttribute('d', l.path);
        const hi = sel && (l.from === sel || l.to === sel);
        el.classList.toggle('dl-link--hi', !!hi);
    });
}
