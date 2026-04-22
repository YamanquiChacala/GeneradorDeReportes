import { CssColorMap, isColorKey } from "../enums";

/**
 * Normalizes unicode, collapses multiple spaces, and trims edges.
 */
function baseSanitize(input: string): string {
    if (!input) return "";

    return input
        .normalize("NFKC") // Standardize composed characters (keeps Spanish accents intact)
        .replace(/\s+/g, " ") // Replace multiple spaces with a single space
        .trim(); // Remove leading/trailing spaces
}

/**
 * Sanitizes a string for safe use as a file name across Drive, Windows, Mac and Linux.
 */
export function sanitizeFileName(input?: string, fallback = "Grupo"): string {
    if (!input) return fallback;
    // Remove OS-prohibited characters: < > : " / \ | ? *
    const sanitized = input.replace(/[<>:"/\\|?*]/g, "");

    return baseSanitize(sanitized) || fallback;
}

/**
 * Sanitizes a string for use as a Google Sheets tab name.
 */
export function sanitizeSheetName(input: string): string {
    // Handle Apostrophes and Quotes, leaving a space behind: '"`
    let sanitized = input.replace(/['"`]/g, " ");

    // Remove Spreadsheet-prohibited characters: []:*?/\|
    sanitized = sanitized.replace(/[[\]:*?/\\|]/g, "");

    sanitized = baseSanitize(sanitized);

    // 4. Enforce Character Limits for downloading to Excell
    if (sanitized.length > 31) {
        sanitized = sanitized.substring(0, 31).trim();
    }

    return sanitized;
}

/**
 * Validates a web color string and normalizes it to a 6-digit hex format.
 * Translates standard CSS color names into their hex equivalents.
 */
export function webColor(input?: string): string | null {
    const trimmedInput = input?.trim();
    if (!trimmedInput) return null;

    // 1. Check for Hex Format (#RGB or #RRGGBB)
    const hexRegex = /^#([A-Fa-f0-9]{3}|[A-Fa-f0-9]{6})$/;
    if (hexRegex.test(trimmedInput)) {
        let hex = trimmedInput.toUpperCase();

        // Convert 3-digit hex (#FA0) to 6-digit hex (#FFAA00)
        if (hex.length === 4) {
            hex = `#${hex[1]}${hex[1]}${hex[2]}${hex[2]}${hex[3]}${hex[3]}`;
        }

        return hex;
    }

    // 2. Check for Named Colors and map to Hex
    const lowerInput = trimmedInput.toLowerCase();
    if (isColorKey(lowerInput)) {
        return CssColorMap[lowerInput];
    }
    return null;
}
