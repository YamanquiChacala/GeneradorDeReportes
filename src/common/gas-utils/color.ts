import { type HSLColor, hexToRgb, hslToRgb, type RGBColor, rgbToHex } from "../utils";
import { BAND_LIGHT, FOOTER_LIGH, FOOTER_SAT, HEADER_LIGH, HEADER_SAT, LIGHT_GREY_COLOR } from "./constants";

/**
 * Creates a banding for alternating cell backgrounds in a given hue [0,1]
 */
export function createBanding(hue: number, header = false, footer = false): GoogleAppsScript.Sheets.Schema.BandingProperties {
    const banding: GoogleAppsScript.Sheets.Schema.BandingProperties = {
        firstBandColorStyle: { rgbColor: LIGHT_GREY_COLOR },
        secondBandColorStyle: { rgbColor: hslToColor({ h: hue, s: 1, l: BAND_LIGHT }) },
    };
    if (header) banding.headerColorStyle = { rgbColor: hslToColor({ h: hue, s: HEADER_SAT, l: HEADER_LIGH }) };
    if (footer) banding.footerColorStyle = { rgbColor: hslToColor({ h: hue, s: FOOTER_SAT, l: FOOTER_LIGH }) };

    return banding;
}

/** Helper function to convert Sheets API Color to a Hex string */
export function colorToHex(color?: GoogleAppsScript.Sheets.Schema.Color, fallback = "#FFFFFF"): string {
    // If there's no color provided, default to white as requested
    if (!color) return fallback;

    const rbgColor: RGBColor = {
        r: color.red ?? 0,
        g: color.green ?? 0,
        b: color.blue ?? 0,
    };

    return rgbToHex(rbgColor);
}

/** Helper function to convert a HSL color into Google Sheets color */
function hslToColor(color: HSLColor): GoogleAppsScript.Sheets.Schema.Color {
    const rgbColor = hslToRgb(color);
    return {
        red: rgbColor.r,
        green: rgbColor.g,
        blue: rgbColor.b,
        alpha: 1,
    };
}

export function hexToColor(hex: string): GoogleAppsScript.Sheets.Schema.Color | null {
    const rgb = hexToRgb(hex);
    if (rgb == null) {
        return null;
    }
    return {
        red: rgb.r,
        green: rgb.g,
        blue: rgb.b,
        alpha: 1,
    };
}
