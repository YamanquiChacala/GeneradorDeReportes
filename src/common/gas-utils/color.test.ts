import { hslToRgb } from "../utils";
import { colorToHex, createBanding, hexToColor } from "./color";
import { BAND_LIGHT, FOOTER_LIGH, FOOTER_SAT, HEADER_LIGH, HEADER_SAT, LIGHT_GREY_COLOR } from "./constants";

describe("GAS Utils, Colors", () => {
    describe("createBanding", () => {
        it("should create default banding properties (no header/footer)", () => {
            const hue = 0.4;
            const expectedSecondBandRgb = hslToRgb({ h: hue, s: 1, l: BAND_LIGHT });
            const result = createBanding(hue);

            expect(result).toEqual({
                firstBandColorStyle: { rgbColor: LIGHT_GREY_COLOR },
                secondBandColorStyle: {
                    rgbColor: {
                        red: expectedSecondBandRgb.r,
                        green: expectedSecondBandRgb.g,
                        blue: expectedSecondBandRgb.b,
                        alpha: 1,
                    },
                },
            });

            // Verify optional properties are strictly excluded, not just undefined
            expect("headerColor" in result).toBe(false);
            expect("footerColor" in result).toBe(false);
        });

        it("should include header color when header is true", () => {
            const hue = 0.4;
            const expectedHeaderRgb = hslToRgb({ h: hue, s: HEADER_SAT, l: HEADER_LIGH });
            const result = createBanding(hue, true, false);

            expect(result.headerColorStyle?.rgbColor).toEqual({
                red: expectedHeaderRgb.r,
                green: expectedHeaderRgb.g,
                blue: expectedHeaderRgb.b,
                alpha: 1,
            });
            expect("footerColor" in result).toBe(false);
        });

        it("should include footer color when footer is true", () => {
            const hue = 0.4;
            const expectedFooterRgb = hslToRgb({ h: hue, s: FOOTER_SAT, l: FOOTER_LIGH });
            const result = createBanding(hue, false, true);

            expect(result.footerColorStyle?.rgbColor).toEqual({
                red: expectedFooterRgb.r,
                green: expectedFooterRgb.g,
                blue: expectedFooterRgb.b,
                alpha: 1,
            });
            expect("headerColor" in result).toBe(false);
        });

        it("should include both header and footer colors when both are true", () => {
            const hue = 0.7;
            const expectedHeaderRgb = hslToRgb({ h: hue, s: HEADER_SAT, l: HEADER_LIGH });
            const expectedFooterRgb = hslToRgb({ h: hue, s: FOOTER_SAT, l: FOOTER_LIGH });

            const result = createBanding(hue, true, true);

            expect(result.headerColorStyle?.rgbColor).toEqual({
                red: expectedHeaderRgb.r,
                green: expectedHeaderRgb.g,
                blue: expectedHeaderRgb.b,
                alpha: 1,
            });
            expect(result.footerColorStyle?.rgbColor).toEqual({
                red: expectedFooterRgb.r,
                green: expectedFooterRgb.g,
                blue: expectedFooterRgb.b,
                alpha: 1,
            });
        });

        it("should handle boundary hue values [0, 1]", () => {
            // Testing hue = 0 (Red boundary)
            const resultLow = createBanding(0);
            const expectedLowSecondBand = hslToRgb({ h: 0, s: 1, l: BAND_LIGHT });

            expect(resultLow.secondBandColorStyle).toEqual({
                rgbColor: {
                    red: expectedLowSecondBand.r,
                    green: expectedLowSecondBand.g,
                    blue: expectedLowSecondBand.b,
                    alpha: 1,
                },
            });

            // Testing hue = 1 (Red boundary, full rotation)
            const resultHigh = createBanding(1);
            const expectedHighSecondBand = hslToRgb({ h: 1, s: 1, l: BAND_LIGHT });

            expect(resultHigh.secondBandColorStyle).toEqual({
                rgbColor: {
                    red: expectedHighSecondBand.r,
                    green: expectedHighSecondBand.g,
                    blue: expectedHighSecondBand.b,
                    alpha: 1,
                },
            });
        });
    });

    describe("colorToHex (Integration with rgbToHex)", () => {
        it("should convert an RGB color to hex correctly", () => {
            const color: GoogleAppsScript.Sheets.Schema.Color = {
                red: 1, // 255
                green: 0.50196, // ~128
                blue: 0, // 0
            };
            expect(colorToHex(color)).toBe("#FF8000");
        });

        it("should convert explicit bounds (pure white and pure black)", () => {
            const white: GoogleAppsScript.Sheets.Schema.Color = { red: 1, green: 1, blue: 1 };
            expect(colorToHex(white)).toBe("#FFFFFF");

            const black: GoogleAppsScript.Sheets.Schema.Color = { red: 0, green: 0, blue: 0 };
            expect(colorToHex(black)).toBe("#000000");
        });

        it("should treat missing channels as 0", () => {
            const redOnly: GoogleAppsScript.Sheets.Schema.Color = { red: 1 };
            expect(colorToHex(redOnly)).toBe("#FF0000");

            const greenOnly: GoogleAppsScript.Sheets.Schema.Color = { green: 1 };
            expect(colorToHex(greenOnly)).toBe("#00FF00");

            const blueOnly: GoogleAppsScript.Sheets.Schema.Color = { blue: 1 };
            expect(colorToHex(blueOnly)).toBe("#0000FF");
        });

        it("should treat missing all channels as black", () => {
            const emptyColor: GoogleAppsScript.Sheets.Schema.Color = {};
            expect(colorToHex(emptyColor)).toBe("#000000");
        });

        it("should fallback to the provided fallback if color is undefined", () => {
            expect(colorToHex(undefined, "#AAAAAA")).toBe("#AAAAAA");
        });

        it("should fallback to pure white if color is undefined and no fallback is provided", () => {
            expect(colorToHex(undefined)).toBe("#FFFFFF");
        });
    });

    describe("hexToColor", () => {
        it("should convert a valid hex to a color", () => {
            expect(hexToColor("#000000")).toEqual({ red: 0, green: 0, blue: 0, alpha: 1 });
            expect(hexToColor("#ff0000")).toEqual({ red: 1, green: 0, blue: 0, alpha: 1 });
            expect(hexToColor("#00FF00")).toEqual({ red: 0, green: 1, blue: 0, alpha: 1 });
            expect(hexToColor("#ffffff")).toEqual({ red: 1, green: 1, blue: 1, alpha: 1 });
        });

        it("should return null on invalid input", () => {
            expect(hexToColor("bad")).toBeNull();
        });
    });
});
