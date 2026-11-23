/**
 * tenant-settings-ui-bind-events-writeback.js
 * @param {Record<string, any>} ctx
 */
export function installTenantBindEventsWriteback(ctx) {
    ctx.mergeIntoSubjects = function mergeIntoSubjects(existing, codesToEnsure) {
        const out = (existing || []).slice();
        const seen = new Set(out.map((s) => String(s.code || '').toUpperCase()).filter(Boolean));
        (codesToEnsure || []).forEach((c) => {
            const code = ctx.normCode(c);
            if (!code || seen.has(code)) return;
            seen.add(code);
            out.push({ code, name: '' });
        });
        out.sort((a, b) => ctx.normCode(a.code).localeCompare(ctx.normCode(b.code)));
        return out;
    }

    ctx.mergeIntoClasses = function mergeIntoClasses(existing, codesToEnsure) {
        const out = (existing || []).slice();
        const seen = new Set(out.map((c) => ctx.normCode(c.code || c.name || '')).filter(Boolean));
        (codesToEnsure || []).forEach((c) => {
            const code = ctx.normCode(c);
            if (!code || seen.has(code)) return;
            seen.add(code);
            out.push({ code, name: `Klasse ${code}`, year: '', headName: '', headEmail: '' });
        });
        out.sort((a, b) => ctx.normCode(a.code).localeCompare(ctx.normCode(b.code)));
        return out;
    }

    ctx.mergeIntoArges = function mergeIntoArges(existing, codesToEnsure) {
        const out = (existing || []).slice();
        const seen = new Set(out.map((a) => ctx.normCode(a.code || '')).filter(Boolean));
        (codesToEnsure || []).forEach((c) => {
            const code = ctx.normCode(c);
            if (!code || seen.has(code)) return;
            seen.add(code);
            out.push({ code, name: code, subjects: [] });
        });
        out.sort((a, b) => ctx.normCode(a.code).localeCompare(ctx.normCode(b.code)));
        return out;
    }

    ctx.applyWriteback = function applyWriteback(detail) {
        if (!detail || !detail.writeback) return;
        if (typeof ctx.load !== 'function' || typeof ctx.save !== 'function') return;
        const current = ctx.load();
        const next = Object.assign({}, current);
        if (detail.writeback.subjectCodes) {
            next.subjects = ctx.mergeIntoSubjects(current.subjects || [], detail.writeback.subjectCodes);
        }
        if (detail.writeback.argeCodes) {
            next.arges = ctx.mergeIntoArges(current.arges || [], detail.writeback.argeCodes);
        }
        if (detail.writeback.classCodes) {
            next.classes = ctx.mergeIntoClasses(current.classes || [], detail.writeback.classCodes);
        }
        ctx.__syncGuard++;
        const saved = ctx.save(next);
        // UI aktualisieren
        try {
            if (ctx.taSubjects) ctx.taSubjects.value = ctx.subjectsToLines(saved.subjects || []);
            if (ctx.taArges) ctx.taArges.value = ctx.argesToLines(saved.arges || []);
            if (ctx.taClasses) ctx.taClasses.value = ctx.classesToLines(saved.classes || []);
            ctx.renderSubjectsTableFromTextarea();
            ctx.renderArgesTableFromTextarea();
            ctx.renderClassesTableFromTextarea();
            ctx.setSummary('Listen wurden aus der Struktur ergänzt.', 'ok');
        } catch {
            // ignore
        }
        ctx.__syncGuard--;
        ctx.dispatchTenantSettingsChanged(saved, 'writeback');
    }

    try {
        window.addEventListener('ms365-structure-changed', (ev) => {
            if (ctx.__syncGuard) return;
            ctx.applyWriteback(ev && ev.detail ? ev.detail : null);
            try {
                if (typeof window.__ms365TenantStepsRefreshMatch === 'function') window.__ms365TenantStepsRefreshMatch();
            } catch {
                // ignore
            }
        });
        window.addEventListener('ms365-match-links-changed', () => {
            try {
                if (typeof window.__ms365TenantStepsRefreshMatch === 'function') window.__ms365TenantStepsRefreshMatch();
            } catch {
                // ignore
            }
        });
    } catch {
        // ignore
    }
}
