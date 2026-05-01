import type { Icon } from "../enums";
import { webColor } from "./text";

export interface IconifyParams {
    iconName: Icon;
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
    const safeColor = webColor(color)?.replace(/^#/, "%23");

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
