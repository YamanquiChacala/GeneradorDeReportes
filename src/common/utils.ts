import type { Icon } from "./enums";

export interface IconifyParams {
    iconName: Icon | string;
    color?: string;
    width?: number;
    height?: number;
    box?: boolean;
}

/**
 * Helper method to generate the URL for an Iconify icon.
 */
export function iconifyUrl({ iconName, color, width, height, box }: IconifyParams): { name: string; url: string } {
    // Sanitize the color (replace # with %23 for the URL)
    const safeColor = color?.startsWith("#") ? `%23${color.substring(1)}` : color;

    // Collect provided options into a parameters array
    const params = [];
    if (safeColor) params.push(`color=${safeColor}`);
    if (width) params.push(`width=${width}`);
    if (height) params.push(`height=${height}`);
    if (box) params.push(`box=1`);

    // Construct the base URL
    let url = `https://api.iconify.design/${iconName}.svg`;

    // 4. Append query string only if there are parameters
    if (params.length > 0) {
        url += `?${params.join("&")}`;
    }

    // Extract the alt text from the last element of the name
    const nameParts = iconName.split("/");
    const name = nameParts[nameParts.length - 1] || "";

    return { url, name };
}

/**
 * Normalizes unicode, collapses multiple spaces, and trims edges.
 *
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
