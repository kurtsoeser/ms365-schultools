import { describe, expect, it } from 'vitest';
import {
    pdfImportTargetFromFilename,
    resolveRegisterImportTarget,
    spreadsheetImportTargetFromAoa
} from '../src/shared/register-webuntis-file-import.js';

describe('register-webuntis-file-import', () => {
    it('erkennt WebUntis-PDFs am Dateinamen', () => {
        expect(pdfImportTargetFromFilename('Subject_2025.pdf')).toBe('subjects');
        expect(pdfImportTargetFromFilename('Class_2025.pdf')).toBe('classes');
    });

    it('erkennt WebUntis-Tabellen am Kopf', () => {
        const teacherAoa = [['Lehrkraft', 'Familienname', 'Vorname']];
        const subjectAoa = [['name', 'longName']];
        const studentAoa = [['forename', 'longname', 'klasse.name', 'id']];
        const detect = (aoa) => {
            const h = (aoa[0] || []).join(',').toLowerCase();
            if (h.includes('lehrkraft')) return 'teacher';
            if (h === 'name,longname') return 'subject';
            if (h.includes('forename')) return 'student';
            return '';
        };
        expect(spreadsheetImportTargetFromAoa(teacherAoa, detect)).toBe('teachers');
        expect(spreadsheetImportTargetFromAoa(subjectAoa, detect)).toBe('subjects');
        expect(spreadsheetImportTargetFromAoa(studentAoa, detect)).toBe('students');
    });

    it('nutzt Tab-Fallback bei unbekanntem Format', () => {
        expect(resolveRegisterImportTarget('subjects', 'generic')).toBe('subjects');
        expect(resolveRegisterImportTarget('subjects', 'teachers')).toBe('teachers');
    });
});
