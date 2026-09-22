/**
 * Gemeinsames Rendern von Release-Notes (Admin + Welcome).
 */
(function (global) {
    'use strict';

    function kindLabel(kind) {
        if (kind === 'feature') return 'Neu';
        if (kind === 'fix') return 'Fix';
        return 'Info';
    }

    function kindClass(kind) {
        if (kind === 'feature') return 'rn-kind--feature';
        if (kind === 'fix') return 'rn-kind--fix';
        return 'rn-kind--other';
    }

    /**
     * @param {object} note
     * @param {{ editable?: boolean, onDelete?: function, onEdit?: function }} [opts]
     */
    function renderNoteCard(note, opts) {
        opts = opts || {};
        const article = document.createElement('article');
        article.className = 'rn-card';
        article.dataset.id = note.id || '';

        const head = document.createElement('div');
        head.className = 'rn-card__head';

        const badge = document.createElement('span');
        badge.className = 'rn-kind ' + kindClass(note.kind);
        badge.textContent = kindLabel(note.kind);

        const title = document.createElement('h4');
        title.className = 'rn-card__title';
        title.textContent = note.title || '(ohne Titel)';

        head.appendChild(badge);
        head.appendChild(title);

        if (opts.editable) {
            const actions = document.createElement('div');
            actions.className = 'rn-card__actions';
            if (typeof opts.onEdit === 'function') {
                const editBtn = document.createElement('button');
                editBtn.type = 'button';
                editBtn.className = 'btn btn-sm alt';
                editBtn.innerHTML = '<i class="bi bi-pencil" aria-hidden="true"></i>';
                editBtn.title = 'Bearbeiten';
                editBtn.addEventListener('click', function () {
                    opts.onEdit(note);
                });
                actions.appendChild(editBtn);
            }
            if (typeof opts.onDelete === 'function') {
                const delBtn = document.createElement('button');
                delBtn.type = 'button';
                delBtn.className = 'btn btn-sm btn-danger';
                delBtn.innerHTML = '<i class="bi bi-trash" aria-hidden="true"></i>';
                delBtn.title = 'Löschen';
                delBtn.addEventListener('click', function () {
                    opts.onDelete(note);
                });
                actions.appendChild(delBtn);
            }
            head.appendChild(actions);
        }

        const meta = document.createElement('div');
        meta.className = 'rn-card__meta';
        const parts = [];
        if (note.at && !Number.isNaN(new Date(note.at).getTime())) {
            parts.push(
                new Intl.DateTimeFormat('de-AT', {
                    day: '2-digit',
                    month: '2-digit',
                    year: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit'
                }).format(new Date(note.at))
            );
        }
        if (note.source === 'github') parts.push('GitHub');
        else if (note.source === 'local') parts.push('Lokal');
        meta.textContent = parts.join(' · ');

        const body = document.createElement('div');
        body.className = 'rn-card__body';
        const html = note.bodyHtml || '';
        if (html) body.innerHTML = html;
        else if (note.body) {
            const pre = document.createElement('pre');
            pre.className = 'rn-card__pre';
            pre.textContent = note.body;
            body.appendChild(pre);
        }

        article.appendChild(head);
        article.appendChild(meta);
        article.appendChild(body);

        if (Array.isArray(note.images) && note.images.length) {
            const gallery = document.createElement('div');
            gallery.className = 'rn-card__gallery';
            note.images.forEach(function (img) {
                const figure = document.createElement('figure');
                figure.className = 'rn-card__figure';
                const image = document.createElement('img');
                image.src = img.src;
                image.alt = img.alt || 'Screenshot';
                image.loading = 'lazy';
                image.addEventListener('click', function () {
                    if (typeof global.ms365OpenImageLightbox === 'function') {
                        global.ms365OpenImageLightbox(img.src, img.alt);
                    } else {
                        global.open(img.src, '_blank', 'noopener');
                    }
                });
                figure.appendChild(image);
                gallery.appendChild(figure);
            });
            article.appendChild(gallery);
        }

        return article;
    }

    /**
     * Screenshot komprimieren (JPEG), max. Kantenlänge.
     * @returns {Promise<{ src: string, alt: string, bytes: number }>}
     */
    function compressImageFile(file, alt, maxEdge) {
        maxEdge = maxEdge || 1280;
        return new Promise(function (resolve, reject) {
            const url = URL.createObjectURL(file);
            const img = new Image();
            img.onload = function () {
                try {
                    let w = img.naturalWidth || img.width;
                    let h = img.naturalHeight || img.height;
                    const scale = Math.min(1, maxEdge / Math.max(w, h));
                    w = Math.max(1, Math.round(w * scale));
                    h = Math.max(1, Math.round(h * scale));
                    const canvas = document.createElement('canvas');
                    canvas.width = w;
                    canvas.height = h;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, w, h);
                    const dataUrl = canvas.toDataURL('image/jpeg', 0.72);
                    URL.revokeObjectURL(url);
                    resolve({
                        src: dataUrl,
                        alt: alt || file.name || 'Screenshot',
                        bytes: Math.round((dataUrl.length * 3) / 4)
                    });
                } catch (e) {
                    URL.revokeObjectURL(url);
                    reject(e);
                }
            };
            img.onerror = function () {
                URL.revokeObjectURL(url);
                reject(new Error('Bild konnte nicht geladen werden.'));
            };
            img.src = url;
        });
    }

    function openLightbox(src, alt) {
        let overlay = document.getElementById('rnLightbox');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'rnLightbox';
            overlay.className = 'rn-lightbox';
            overlay.hidden = true;
            overlay.innerHTML =
                '<button type="button" class="rn-lightbox__close" aria-label="Schließen">&times;</button>' +
                '<img class="rn-lightbox__img" alt="">';
            document.body.appendChild(overlay);
            overlay.addEventListener('click', function (ev) {
                if (ev.target === overlay || (ev.target && ev.target.classList.contains('rn-lightbox__close'))) {
                    overlay.hidden = true;
                }
            });
        }
        const image = overlay.querySelector('.rn-lightbox__img');
        if (image) {
            image.src = src;
            image.alt = alt || '';
        }
        overlay.hidden = false;
    }

    global.ms365ReleaseNotesUi = {
        renderNoteCard: renderNoteCard,
        compressImageFile: compressImageFile,
        kindLabel: kindLabel
    };
    global.ms365OpenImageLightbox = openLightbox;
})(typeof window !== 'undefined' ? window : globalThis);
