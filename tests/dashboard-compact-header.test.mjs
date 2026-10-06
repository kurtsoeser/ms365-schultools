import { describe, expect, it } from 'vitest';
import {
    dashboardHeroParallaxStyle,
    dashboardHeroScrollPhase
} from '../src/shared/dashboard-compact-header.js';

describe('dashboard-compact-header', () => {
    it('klassifiziert Scroll-Phasen', () => {
        expect(dashboardHeroScrollPhase(0)).toBe('top');
        expect(dashboardHeroScrollPhase(40)).toBe('transition');
        expect(dashboardHeroScrollPhase(120)).toBe('scrolled');
    });

    it('berechnet Parallax-Styles', () => {
        const top = dashboardHeroParallaxStyle(0);
        expect(top.opacity).toBe('1');
        expect(top.transform).toContain('translate3d');

        const low = dashboardHeroParallaxStyle(200, { fadeDistance: 200 });
        expect(parseFloat(low.opacity)).toBe(0);
    });
});
