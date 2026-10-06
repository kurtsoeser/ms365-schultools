/**
 * Tenant-Panel: WebUntis-Handoff, Review, Apply + Adapter-Registry (Phase 7a).
 */
import { peekWebuntisImportPayload, clearWebuntisImportPayload } from './webuntis-stammdaten-import-handoff.js';
import { summarizeWebuntisMergeResult } from './webuntis-stammdaten-wizard-logic.js';
import { buildWebuntisTenantReview } from './stammdaten-import-review.js';
import {
    mergeStammdatenImport,
    ADAPTER_WEBUNTIS_HANDOFF,
    ADAPTER_SIS_FILE
} from './stammdaten-import-pipeline.js';
import './stammdaten-import-adapters.js';

function requestRegisterTab(tabBtnId) {
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-register-tab-request', { detail: { tabBtnId: tabBtnId || 'tabMainSchueler' } })
        );
        window.dispatchEvent(
            new CustomEvent('ms365-register-mode-request', { detail: { mode: 'import' } })
        );
    } catch {
        /* ignore */
    }
}

/**
 * @param {object} ctx
 * @param {() => object} ctx.getCurrentLines
 * @param {() => object} ctx.getMergeDeps
 * @param {(merged: object, payload?: object) => Promise<void>|void} ctx.onApplyMerged
 * @param {{ textContent?: string, className?: string }|null} [ctx.summary]
 * @param {() => object|null} [ctx.sisApi]
 */
export function mountTenantImportPanel(ctx) {
    if (!ctx || typeof ctx.getCurrentLines !== 'function' || typeof ctx.onApplyMerged !== 'function') {
        return;
    }

    let pendingWebuntisPayload = null;

    function hideWebuntisImportReview() {
        pendingWebuntisPayload = null;
        const box = document.getElementById('tenantWebuntisImportReview');
        if (box) box.hidden = true;
    }

    function mergePayload(payload) {
        const deps = typeof ctx.getMergeDeps === 'function' ? ctx.getMergeDeps() : {};
        const lines = typeof ctx.getCurrentLines === 'function' ? ctx.getCurrentLines() : {};
        return mergeStammdatenImport(ADAPTER_WEBUNTIS_HANDOFF, lines, payload, deps);
    }

    function showWebuntisImportReview(payload) {
        if (!payload) return;
        pendingWebuntisPayload = payload;
        const merged = mergePayload(payload);
        const sis = ctx.sisApi ? ctx.sisApi() : null;
        const lines = typeof ctx.getCurrentLines === 'function' ? ctx.getCurrentLines() : {};
        const deps = typeof ctx.getMergeDeps === 'function' ? ctx.getMergeDeps() : {};
        const review = buildWebuntisTenantReview(merged, payload, sis || {}, {
            existingTeachersLines: lines.teachersLines,
            parseTeachersLines: deps.parseTeachersLines
        });
        const box = document.getElementById('tenantWebuntisImportReview');
        const titleEl = document.getElementById('tenantWebuntisImportReviewTitle');
        const ul = document.getElementById('tenantWebuntisImportReviewBullets');
        const details = document.getElementById('tenantWebuntisImportReviewDetails');
        const stuBox = document.getElementById('tenantWebuntisImportReviewStudents');
        const summary = ctx.summary;
        if (!box || !ul) return;
        if (titleEl) titleEl.textContent = review.title;
        ul.replaceChildren();
        (review.bullets || []).forEach(function (line) {
            const li = document.createElement('li');
            li.textContent = line;
            ul.appendChild(li);
        });
        if (stuBox && details) {
            stuBox.replaceChildren();
            const teacherRows = review.teacherRows || [];
            const studentRows = review.studentRows || [];
            const rows = teacherRows.concat(studentRows);
            if (rows.length) {
                details.hidden = false;
                if (teacherRows.length) {
                    const h = document.createElement('div');
                    h.className = 'ts-import-review-section-label';
                    h.textContent = 'Lehrer (Kürzel)';
                    h.style.cssText = 'font-weight:600;margin:8px 0 4px;';
                    stuBox.appendChild(h);
                    teacherRows.forEach(function (row) {
                        const div = document.createElement('div');
                        div.className = 'kind-' + (row.kind || 'update');
                        div.textContent = row.text;
                        stuBox.appendChild(div);
                    });
                }
                if (studentRows.length) {
                    const h2 = document.createElement('div');
                    h2.className = 'ts-import-review-section-label';
                    h2.textContent = 'Schüler';
                    h2.style.cssText = 'font-weight:600;margin:12px 0 4px;';
                    stuBox.appendChild(h2);
                    studentRows.forEach(function (row) {
                        const div = document.createElement('div');
                        div.className = 'kind-' + (row.kind || 'update');
                        div.textContent = row.text;
                        stuBox.appendChild(div);
                    });
                }
            } else {
                details.hidden = true;
            }
        }
        box.hidden = false;
        requestRegisterTab('tabMainSchueler');
        try {
            document.getElementById('tenantWebuntisImportReview')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } catch {
            /* ignore */
        }
        if (summary) {
            summary.textContent = review.hasConflicts
                ? 'WebUntis-Import wartet auf Ihre Bestätigung – bitte Konflikte prüfen und „Übernehmen“ klicken.'
                : 'WebUntis-Import wartet auf Ihre Bestätigung – bitte „Übernehmen“ oder „Verwerfen“.';
            summary.className = review.hasConflicts ? 'summary warn' : 'summary info';
        }
    }

    const btnWuApply = document.getElementById('tenantWebuntisImportApply');
    if (btnWuApply && !btnWuApply.dataset.bound) {
        btnWuApply.dataset.bound = '1';
        btnWuApply.addEventListener('click', function () {
            if (!pendingWebuntisPayload) return;
            const payload = pendingWebuntisPayload;
            hideWebuntisImportReview();
            clearWebuntisImportPayload();
            void Promise.resolve(ctx.onApplyMerged(mergePayload(payload), payload));
        });
    }
    const btnWuDiscard = document.getElementById('tenantWebuntisImportDiscard');
    if (btnWuDiscard && !btnWuDiscard.dataset.bound) {
        btnWuDiscard.dataset.bound = '1';
        btnWuDiscard.addEventListener('click', function () {
            clearWebuntisImportPayload();
            hideWebuntisImportReview();
            const summary = ctx.summary;
            if (summary) {
                summary.textContent = 'WebUntis-Import verworfen – keine Änderung an den Listen.';
                summary.className = 'summary info';
            }
        });
    }

    const pendingWebuntis = peekWebuntisImportPayload();
    if (pendingWebuntis) {
        if (pendingWebuntis.skipTenantReview) {
            clearWebuntisImportPayload();
            void Promise.resolve(ctx.onApplyMerged(mergePayload(pendingWebuntis), pendingWebuntis));
        } else {
            showWebuntisImportReview(pendingWebuntis);
        }
    }
}

/**
 * SIS-Datei-Merge über den gemeinsamen Adapter (für Tenant/UI).
 * @param {object} existingLines
 * @param {{ records: object[], lines?: string, mode?: string, source?: string }} sisResult
 * @param {object} deps
 */
export function mergeSisImportViaAdapter(existingLines, sisResult, deps) {
    return mergeStammdatenImport(ADAPTER_SIS_FILE, existingLines, sisResult, deps);
}

export { summarizeWebuntisMergeResult, ADAPTER_SIS_FILE, ADAPTER_WEBUNTIS_HANDOFF };
