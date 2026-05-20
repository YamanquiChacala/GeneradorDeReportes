import { getStep } from "./math";

export interface RGBColor {
    r: number;
    g: number;
    b: number;
}

export interface HSLColor {
    h: number;
    s: number;
    l: number;
}

/**
 * Transform a Hex color into it's rgb [0,1] componenet.
 */
export function hexToRgb(hex: string): RGBColor | null {
    const match = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    if (!match) return null;

    const [, r, g, b] = match as RegExpExecArray & [string, string, string, string];

    return {
        r: parseInt(r, 16) / 255,
        g: parseInt(g, 16) / 255,
        b: parseInt(b, 16) / 255,
    };
}

/**
 * Transforms an rgb color into a hsl color.
 */
export function rgbToHsl({ r, g, b }: RGBColor): HSLColor {
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h = 0,
        s = 0,
        l = (max + min) / 2;

    // Not grey
    if (max !== min) {
        const dist = max - min;
        s = l > 0.5 ? dist / (2 - max - min) : dist / (max + min);
        switch (max) {
            case r:
                h = (g - b) / dist + (g < b ? 6 : 0);
                break;
            case g:
                h = (b - r) / dist + 2;
                break;
            case b:
                h = (r - g) / dist + 4;
                break;
        }
        h /= 6;
    }
    return { h, s, l };
}

/**
 * Transforms an hsl color into an rgb color.
 */
export function hslToRgb({ h, s, l }: HSLColor): RGBColor {
    let r: number, g: number, b: number;
    if (s === 0) {
        r = g = b = l; // gray
    } else {
        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        r = hue2rgb(p, q, h + 1 / 3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1 / 3);
    }
    return { r: r, g: g, b: b };
}

function hue2rgb(p: number, q: number, t: number): number {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
}

/**
 * Transfroms an rgb color to it's hex representation.
 */
export function rgbToHex({ r, g, b }: RGBColor): string {
    return (
        "#" +
        [r, g, b]
            .map((x) => {
                var h = Math.round(x * 255).toString(16);
                return h.length === 1 ? `0${h}` : h;
            })
            .join("")
    ).toUpperCase();
}

/**
 * Distributes how `many` hues between [0,1) as distinct as possible.
 */
export function getDistinctHues(many: number, first = 0): number[] {
    if (many <= 0) return [];

    const step = getStep(many);
    const hues: number[] = [];

    let currentPosition = 0;

    for (let i = 0; i < many; i++) {
        // Calculate the base fraction between 0 and 1
        const fraction = currentPosition / many;

        // Add the 'first' offset and wrap around using modulo 1
        let hue = (fraction + first) % 1;

        // Handle potential negative 'first' values to ensure [0, 1) bounds
        if (hue < 0) {
            hue += 1;
        }

        hues.push(hue);

        // Advance by the optimal step, wrapping around the total number of elements
        currentPosition = (currentPosition + step) % many;
    }

    return hues;
}
