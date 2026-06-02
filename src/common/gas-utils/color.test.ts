import { hslToRgb } from "../utils";
import { BAND_LIGHT, colorToHex, createBanding, FOOTER_LIGH, FOOTER_SAT, HEADER_LIGH, HEADER_SAT, LIGHT_GREY_COLOR } from ".";

describe("Colors", () => {
    describe("createBanding", () => {
        it("should create default banding properties (no header/footer)", () => {
            const hue = 0.4;
            const expectedSecondBandRgb = hslToRgb({ h: hue, s: 1, l: BAND_LIGHT });
            const result = createBanding(hue);

            expect(result).toEqual({
                firstBandColor: LIGHT_GREY_COLOR,
                secondBandColor: {
                    red: expectedSecondBandRgb.r,
                    green: expectedSecondBandRgb.g,
                    blue: expectedSecondBandRgb.b,
                },
            });
            expect(result.headerColor).toBeUndefined();
            expect(result.footerColor).toBeUndefined();
        });

        it("should include header color when header is true", () => {
            const hue = 0.4;
            const expectedHeaderRgb = hslToRgb({ h: hue, s: HEADER_SAT, l: HEADER_LIGH });
            const result = createBanding(hue, true, false);

            expect(result.headerColor).toEqual({
                red: expectedHeaderRgb.r,
                green: expectedHeaderRgb.g,
                blue: expectedHeaderRgb.b,
            });
        });

        it("should include footer color when footer is true", () => {
            const hue = 0.4;
            const expectedFooterRgb = hslToRgb({ h: hue, s: FOOTER_SAT, l: FOOTER_LIGH });
            const result = createBanding(hue, false, true);

            expect(result.footerColor).toEqual({
                red: expectedFooterRgb.r,
                green: expectedFooterRgb.g,
                blue: expectedFooterRgb.b,
            });
        });

        it("should handle boundary hue values", () => {
            // Testing hue = 0
            const resultLow = createBanding(0, true, true);
            const expectedLowSecondBand = hslToRgb({ h: 0, s: 1, l: BAND_LIGHT });

            expect(resultLow.secondBandColor).toEqual({
                red: expectedLowSecondBand.r,
                green: expectedLowSecondBand.g,
                blue: expectedLowSecondBand.b,
            });

            // Testing hue = 1
            const resultHigh = createBanding(1, true, true);
            const expectedHighSecondBand = hslToRgb({ h: 1, s: 1, l: BAND_LIGHT });

            expect(resultHigh.secondBandColor).toEqual({
                red: expectedHighSecondBand.r,
                green: expectedHighSecondBand.g,
                blue: expectedHighSecondBand.b,
            });
        });
    });

    describe("colorToHex", () => {
        it("should convert an RGB color to hex correctly", () => {
            const color: GoogleAppsScript.Sheets.Schema.Color = {
                red: 1, // 255
                green: 0.50196, // ~128
                blue: 0, // 0
            };
            expect(colorToHex(color)).toBe("#FF8000");
        });

        it("should treat missing channels as 0", () => {
            const red: GoogleAppsScript.Sheets.Schema.Color = { red: 1 };
            expect(colorToHex(red)).toBe("#FF0000");

            const green: GoogleAppsScript.Sheets.Schema.Color = { green: 1 };
            expect(colorToHex(green)).toBe("#00FF00");
        });

        it("should treat missing all channels as black", () => {
            const color: GoogleAppsScript.Sheets.Schema.Color = {};
            expect(colorToHex(color)).toBe("#000000");
        });

        it("should fallback to the provided fallback if color is undefined", () => {
            expect(colorToHex(undefined, "#aaaaaa")).toBe("#aaaaaa");
            expect(colorToHex(undefined)).toBe("#FFFFFF"); // Default fallback
        });
    });
});
