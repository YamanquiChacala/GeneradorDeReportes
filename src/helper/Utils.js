const Utils = {
    /**
     * Helper method to generate the URL for an Iconify icon.
     * @param {Object} options
     * @param {Icon} options.iconName
     * @param {string} [options.color]
     * @param {number} [options.width]
     * @param {number} [options.height]
     * @param {boolean} [options.box]
     * @returns {{ name: string, url: string}} The iconify API url
     */
    iconifyUrl({ iconName, color, width, height, box }) {
        // Sanitize the color (replace # with %23 for the URL)
        let safeColor = color?.startsWith("#") ? "%23" + color.substring(1) : color;

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
        const name = nameParts[nameParts.length - 1];

        return ({ url, name });
    }
}