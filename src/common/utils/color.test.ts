import { getDistinctHues, hexToRgb, hslToRgb, rgbToHex, rgbToHsl } from "./color";
import { getStep } from "./math";

describe("Color Conversion Utilities", () => {
    describe("hexToRgb", () => {
        it("should correctly convert standard 6-character hex codes", () => {
            expect(hexToRgb("#ff0000")).toEqual({ r: 1, g: 0, b: 0 });
            expect(hexToRgb("#00ff00")).toEqual({ r: 0, g: 1, b: 0 });
            expect(hexToRgb("#0000ff")).toEqual({ r: 0, g: 0, b: 1 });
            expect(hexToRgb("#ffffff")).toEqual({ r: 1, g: 1, b: 1 });
            expect(hexToRgb("#000000")).toEqual({ r: 0, g: 0, b: 0 });
        });

        it("should work without the # prefix", () => {
            expect(hexToRgb("ff0000")).toEqual({ r: 1, g: 0, b: 0 });
        });

        it("should handle case insensitivity", () => {
            expect(hexToRgb("#FF0000")).toEqual({ r: 1, g: 0, b: 0 });
            expect(hexToRgb("#fF0000")).toEqual({ r: 1, g: 0, b: 0 });
        });

        it("should return null for invalid hex strings", () => {
            expect(hexToRgb("#ff00")).toBeNull(); // Too short
            expect(hexToRgb("#ff000000")).toBeNull(); // Too long
            expect(hexToRgb("#gg0000")).toBeNull(); // Invalid characters
            expect(hexToRgb("randomString")).toBeNull();
        });
    });

    describe("rgbToHex", () => {
        it("should correctly convert RGB [0,1] objects to hex", () => {
            expect(rgbToHex({ r: 1, g: 0, b: 0 })).toBe("#FF0000");
            expect(rgbToHex({ r: 0, g: 1, b: 0 })).toBe("#00FF00");
            expect(rgbToHex({ r: 0, g: 0, b: 1 })).toBe("#0000FF");
        });

        it("should correctly pad single-character hex values with a leading zero", () => {
            // 15/255 ≈ 0.0588, which should be '0f' in hex
            expect(rgbToHex({ r: 15 / 255, g: 0, b: 0 })).toBe("#0F0000");
        });
    });

    describe("rgbToHsl", () => {
        it("should correctly convert primary colors", () => {
            expect(rgbToHsl({ r: 1, g: 0, b: 0 })).toEqual({ h: 0, s: 1, l: 0.5 }); // Red
            expect(rgbToHsl({ r: 0, g: 1, b: 0 })).toEqual({ h: 1 / 3, s: 1, l: 0.5 }); // Green
            expect(rgbToHsl({ r: 0, g: 0, b: 1 })).toEqual({ h: 2 / 3, s: 1, l: 0.5 }); // Blue
        });

        it("should correctly handle grayscale (saturation = 0)", () => {
            expect(rgbToHsl({ r: 0.5, g: 0.5, b: 0.5 })).toEqual({ h: 0, s: 0, l: 0.5 });
            expect(rgbToHsl({ r: 0, g: 0, b: 0 })).toEqual({ h: 0, s: 0, l: 0 }); // Black
            expect(rgbToHsl({ r: 1, g: 1, b: 1 })).toEqual({ h: 0, s: 0, l: 1 }); // White
        });

        it("should correctly calculate hue when red is max and green is less than blue", () => {
            // r=1, g=0, b=0.5 (Magenta/Pink-ish)
            const hsl = rgbToHsl({ r: 1, g: 0, b: 0.5 });

            expect(hsl.h).toBeCloseTo(5.5 / 6); // (0 - 0.5) / 1 + 6 = 5.5, then divided by 6
            expect(hsl.s).toBe(1);
            expect(hsl.l).toBe(0.5);
        });
    });

    describe("hslToRgb", () => {
        it("should correctly convert primary colors", () => {
            expect(hslToRgb({ h: 0, s: 1, l: 0.5 })).toEqual({ r: 1, g: 0, b: 0 });
            expect(hslToRgb({ h: 1 / 3, s: 1, l: 0.5 })).toEqual({ r: 0, g: 1, b: 0 });
            expect(hslToRgb({ h: 2 / 3, s: 1, l: 0.5 })).toEqual({ r: 0, g: 0, b: 1 });
        });

        it("should correctly handle grayscale inputs", () => {
            expect(hslToRgb({ h: 0.5, s: 0, l: 0.5 })).toEqual({ r: 0.5, g: 0.5, b: 0.5 });
        });
    });

    describe("Round Trip Conversions (Hex -> RGB -> HSL -> RGB -> Hex)", () => {
        // A robust list covering primaries, darks, lights, pastels, and grayscales
        const testColors = [
            "#FF0000", // Red
            "#00FF00", // Green
            "#0000FF", // Blue
            "#FFFFFF", // White
            "#000000", // Black
            "#808080", // Gray
            "#F2A65A", // Random pastel orange
            "#4CA5D9", // Random light blue
            "#0A1B2C", // Very dark color
            "#FDFDFD", // Very light, almost white
        ];

        // test.each creates an individual test case for every item in the array
        test.each(testColors)("should return %s after a full conversion cycle", (hex) => {
            // 1. Hex to RGB
            const rgb1 = hexToRgb(hex);
            expect(rgb1).not.toBeNull(); // Guard against bad parsing

            // 2. RGB to HSL
            // biome-ignore lint/style/noNonNullAssertion: Testing
            const hsl = rgbToHsl(rgb1!);

            // 3. HSL back to RGB
            const rgb2 = hslToRgb(hsl);

            // 4. RGB back to Hex
            const finalHex = rgbToHex(rgb2);

            expect(finalHex).toBe(hex);
        });
    });

    describe("getStep", () => {
        it("returns 1 for small values of n", () => {
            expect(getStep(1)).toBe(0);
            expect(getStep(2)).toBe(1);
            expect(getStep(3)).toBe(1);
        });

        it("returns the largest coprime near n/2 for even numbers", () => {
            expect(getStep(4)).toBe(1); // gcd(4, 2) is 2, so it falls back to 1
            expect(getStep(6)).toBe(1); // gcd(6, 3) is 3, gcd(6, 2) is 2, gcd(6, 4) is 2, falls back to 1
            expect(getStep(8)).toBe(3); // target 4. gcd(8, 4) is 4. delta 1: gcd(8, 3) is 1. returns 3
            expect(getStep(10)).toBe(3); // target 5. gcd(10, 5) is 5, gcd(10, 4) is 2, gcd(10, 3) is 1
        });

        it("returns the largest coprime near n/2 for odd numbers", () => {
            expect(getStep(5)).toBe(2); // target 2. gcd(5, 2) is 1. returns 2
            expect(getStep(7)).toBe(3); // target 3. gcd(7, 3) is 1. returns 3
            expect(getStep(9)).toBe(4); // target 4. gcd(9, 4) is 1. returns 4
        });
    });

    describe("getDistinctHues", () => {
        it("returns an empty array if many <= 0", () => {
            expect(getDistinctHues(0)).toEqual([]);
            expect(getDistinctHues(-5)).toEqual([]);
        });

        it("generates values starting from 0 by default", () => {
            const result = getDistinctHues(5);
            expect(result).toHaveLength(5);
            expect(result[0]).toBeCloseTo(0);
            // Step for 5 is 2. Positions: 0, 2, 4, 1, 3
            expect(result[1]).toBeCloseTo(0.4);
            expect(result[2]).toBeCloseTo(0.8);
            expect(result[3]).toBeCloseTo(0.2);
            expect(result[4]).toBeCloseTo(0.6);
        });

        it("applies the `first` offset correctly", () => {
            const offset = 0.1;
            const result = getDistinctHues(5, offset);

            expect(result[0]).toBeCloseTo(0 + offset);
            expect(result[1]).toBeCloseTo(0.4 + offset);
            expect(result[2]).toBeCloseTo(0.8 + offset);
            expect(result[3]).toBeCloseTo(0.2 + offset);
            expect(result[4]).toBeCloseTo(0.6 + offset);
        });

        it("wraps around correctly if `first` pushes value >= 1", () => {
            const result = getDistinctHues(2, 0.8);
            // Step for 2 is 1. Positions: 0, 1 -> Fractions: 0, 0.5
            // + 0.8 -> 0.8, 1.3 -> wrapped: 0.8, 0.3
            expect(result[0]).toBeCloseTo(0.8);
            expect(result[1]).toBeCloseTo(0.3);
        });

        it("handles negative `first` values correctly, keeping output in [0, 1)", () => {
            const result = getDistinctHues(2, -0.2);
            // Fractions: 0, 0.5
            // - 0.2 -> -0.2, 0.3 -> wrapped: 0.8, 0.3
            expect(result[0]).toBeCloseTo(0.8);
            expect(result[1]).toBeCloseTo(0.3);
        });

        it("always generates strictly unique elements within [0, 1)", () => {
            const many = 12; // Test an arbitrary larger number
            const result = getDistinctHues(many);

            // 1. Check bounds
            result.forEach((hue) => {
                expect(hue).toBeGreaterThanOrEqual(0);
                expect(hue).toBeLessThan(1);
            });

            // 2. Check uniqueness (using toFixed to avoid JS precision edge cases in Set)
            const uniqueHues = new Set(result.map((h) => h.toFixed(5)));
            expect(uniqueHues.size).toBe(many);
        });
    });
});
