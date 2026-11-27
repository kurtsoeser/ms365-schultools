import { describe, expect, it } from 'vitest';
import { isInteractionRequired } from '../src/shared/it-app-auth.js';

describe('IT-App isInteractionRequired', () => {
    it('behandelt den stillen Iframe-Timeout als Anmeldung, die ein Fenster braucht', () => {
        expect(
            isInteractionRequired({
                errorCode: 'monitor_window_timeout',
                message:
                    'monitor_window_timeout: Token acquisition in iframe failed due to timeout.'
            })
        ).toBe(true);
        expect(
            isInteractionRequired({
                message:
                    'monitor_window_timeout: Token acquisition in iframe failed due to timeout. For more information visit: aka.ms/msaljs/browser-errors'
            })
        ).toBe(true);
    });

    it('lässt fremde Fehler stehen', () => {
        expect(isInteractionRequired({ message: 'Network error' })).toBe(false);
        expect(isInteractionRequired(null)).toBe(false);
    });
});
