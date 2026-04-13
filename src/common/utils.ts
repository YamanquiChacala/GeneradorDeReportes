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

/**
 * Flattens the nested GAS formInputs object into a clean dictionary.
 * Uses a generic <T> so you can strongly type the returned object!
 */
export function flattenFormInputs<T = Record<string, any>>(inputs: Record<string, any> | undefined): Partial<T> {
    const flat: Record<string, any> = {};

    // If the entire form was empty/omitted, return an empty object safely
    if (!inputs) return flat as Partial<T>;

    // Object.entries avoids the "noPropertyAccessFromIndexSignature" error!
    for (const [key, obj] of Object.entries(inputs)) {
        if (!obj) continue;

        if (obj.stringInputs?.value !== undefined) {
            const values = obj.stringInputs.value;
            // If it's a single input (text/radio), return the string.
            // If it's a multiselect checkbox, return the array of strings.
            flat[key] = values.length === 1 ? values[0] : values;
        } else if (obj.dateInput?.msSinceEpoch !== undefined) {
            // Extract the timestamp for dates
            flat[key] = parseInt(obj.dateInput.msSinceEpoch, 10);
        } else if (obj.dateTimeInput?.msSinceEpoch !== undefined) {
            // Extract the timestamp for DateTimes
            flat[key] = parseInt(obj.dateTimeInput.msSinceEpoch, 10);
        } else {
            // Fallback for anything unexpected
            flat[key] = obj;
        }
    }

    return flat as Partial<T>;
}
