(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-tenant-settings-v1';
    const CURRENT_VERSION = 2;

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function loadRaw() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return null;
            return safeJsonParse(raw);
        } catch {
            return null;
        }
    }

    function normalizeSettings(obj) {
        const o = obj && typeof obj === 'object' ? obj : {};
        const domain =
            typeof window.ms365GetSchoolDomainNoAt === 'function'
                ? window.ms365GetSchoolDomainNoAt()
                : normStr(o.domain);

        const subjectsIn = Array.isArray(o.subjects) ? o.subjects : [];
        const teachersIn = Array.isArray(o.teachers) ? o.teachers : [];
        const studentsIn = Array.isArray(o.students) ? o.students : [];
        const classesIn = Array.isArray(o.classes) ? o.classes : [];

        const subjectsSeen = new Set();
        const subjects = [];
        subjectsIn.forEach((s) => {
            const code = normCode(s?.code);
            const name = normStr(s?.name);
            if (!code) return;
            const key = code.toLowerCase();
            if (subjectsSeen.has(key)) return;
            subjectsSeen.add(key);
            subjects.push({ code, name });
        });

        const teachersSeen = new Set();
        const teachers = [];
        teachersIn.forEach((t) => {
            const code = normCode(t?.code);
            const name = normStr(t?.name);
            const email = normStr(t?.email).toLowerCase();
            if (!code) return;
            const key = code.toLowerCase();
            if (teachersSeen.has(key)) return;
            teachersSeen.add(key);
            teachers.push({ code, name, email });
        });

        const students = [];
        studentsIn.forEach((s) => {
            const klasse = normStr(s?.klasse || s?.class || s?.group || s?.Klassse || s?.Klasse);
            const name = normStr(s?.name);
            const email = normStr(s?.email).toLowerCase();
            if (!klasse && !name && !email) return;
            students.push({ klasse, name, email });
        });

        const classesSeen = new Set();
        const classes = [];
        classesIn.forEach((c) => {
            const code = normCode(c?.code);
            const name = normStr(c?.name || c?.klasse || c?.Klasse);
            const yearRaw = normStr(c?.year || c?.abschlussjahr || c?.Abschlussjahr || c?.graduationYear || '');
            const year = /^\d{4}$/.test(yearRaw) ? yearRaw : '';
            const headName = normStr(c?.headName || c?.klassenvorstandName || c?.kvName);
            const headEmail = normStr(c?.headEmail || c?.klassenvorstandEmail || c?.kvEmail).toLowerCase();
            if (!code && !name && !year && !headName && !headEmail) return;
            const key = (code || name).toLowerCase();
            if (classesSeen.has(key)) return;
            classesSeen.add(key);
            classes.push({ code, name, year, headName, headEmail });
        });

        return {
            version: CURRENT_VERSION,
            domain: normStr(domain),
            subjects,
            teachers,
            students,
            classes
        };
    }

    function save(settings) {
        const normalized = normalizeSettings(settings);
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized));
        } catch {
            // ignore
        }
        if (typeof window.ms365SetSchoolDomainNoAt === 'function' && normalized.domain) {
            window.ms365SetSchoolDomainNoAt(normalized.domain);
        }
        return normalized;
    }

    function load() {
        const raw = loadRaw();
        const normalized = normalizeSettings(raw || {});
        return normalized;
    }

    function getTeacherEmailMap() {
        const s = load();
        const map = {};
        s.teachers.forEach((t) => {
            if (t.code && t.email) map[t.code] = t.email;
        });
        return map;
    }

    function parseDelimitedLines(text) {
        const lines = String(text || '').split(/\r\n|\n|\r/);
        const out = [];
        lines.forEach((line) => {
            const t = normStr(line);
            if (!t || t.startsWith('#')) return;
            const parts = t
                .split(/[;\t,|]/)
                .map((x) => normStr(x))
                .filter(Boolean);
            if (!parts.length) return;
            out.push(parts);
        });
        return out;
    }

    function parseLinesToSubjects(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts.slice(1).join(' '));
            if (!code) return;
            out.push({ code, name });
        });
        return out;
    }

    function parseLinesToTeachers(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!code) return;
            out.push({ code, name, email });
        });
        return out;
    }

    function parseLinesToStudents(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const klasse = normStr(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!klasse && !name && !email) return;
            out.push({ klasse, name, email });
        });
        return out;
    }

    function parseLinesToClasses(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            // Unterstützte Formate:
            // - code;name;headName;headEmail (alt)
            // - code;year;name;headName;headEmail (neu)
            // - code;name;year;headName;headEmail (tolerant)
            let year = '';
            let name = '';
            let headName = '';
            let headEmail = '';

            if (parts.length >= 2 && /^\d{4}$/.test(normStr(parts[1] || ''))) {
                year = normStr(parts[1] || '');
                name = normStr(parts[2] || '');
                headName = normStr(parts[3] || '');
                headEmail = normStr(parts[4] || '').toLowerCase();
            } else if (parts.length >= 3 && /^\d{4}$/.test(normStr(parts[2] || ''))) {
                name = normStr(parts[1] || '');
                year = normStr(parts[2] || '');
                headName = normStr(parts[3] || '');
                headEmail = normStr(parts[4] || '').toLowerCase();
            } else {
                name = normStr(parts[1] || '');
                headName = normStr(parts[2] || '');
                headEmail = normStr(parts[3] || '').toLowerCase();
            }

            const y = /^\d{4}$/.test(year) ? year : '';
            if (!code && !name && !y && !headName && !headEmail) return;
            out.push({ code, name, year: y, headName, headEmail });
        });
        return out;
    }

    // Public API (kompatibel zu bisher)
    window.ms365TenantSettingsLoad = load;
    window.ms365TenantSettingsSave = save;
    window.ms365TenantSettingsGetTeacherEmailMap = getTeacherEmailMap;
    window.ms365TenantSettingsParseSubjectsLines = parseLinesToSubjects;
    window.ms365TenantSettingsParseTeachersLines = parseLinesToTeachers;
    window.ms365TenantSettingsParseStudentsLines = parseLinesToStudents;
    window.ms365TenantSettingsParseClassesLines = parseLinesToClasses;
})();

