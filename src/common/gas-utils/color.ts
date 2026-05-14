import { type HSLColor, hslToRgb, type RGBColor, rgbToHex } from "../utils/color-utils";

/**
 * Creates a banding for alternating cell backgrounds in a given hue [0,1]
 */
export function createBanding(hue: number, header = false, footer = false): GoogleAppsScript.Sheets.Schema.BandingProperties {
    const banding: GoogleAppsScript.Sheets.Schema.BandingProperties = {
        firstBandColor: { red: 0.98, green: 0.98, blue: 0.98 },
        secondBandColor: hslToColor({ h: hue, s: 1, l: 0.95 }),
    };
    if (header) banding.headerColor = hslToColor({ h: hue, s: 0.75, l: 0.2 });
    if (footer) banding.footerColor = hslToColor({ h: hue, s: 0.5, l: 0.7 });

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
    };
}
