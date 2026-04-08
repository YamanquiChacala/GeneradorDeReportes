const CardParts = {

    /**
     * Normal header for a Card.
     * 
     * @param {Object} options
     * @param {string} options.title The title on the header
     * @param {string} [options.subtitle] The subtitle on the header
     * @param {string} [options.icon] The name of the image to use as icon. A file called <icon>_48.png should exist in images folder.
     * @param {boolean} [options.circle] Crop the icon into a circle?
     * @returns {GoogleAppsScript.Card_Service.CardHeader} The header, ready to insert on a Card
     */
    header({ title, subtitle, icon = 'icon', circle = false }) {
        const imageURL = `https://media.githubusercontent.com/media/YamanquiChacala/GeneradorDeReportes/refs/heads/main/images/${icon}_48.png`;
        let imageStyle = CardService.ImageStyle.SQUARE;
        if (circle) {
            imageStyle = CardService.ImageStyle.CIRCLE;
        }
        const header = CardService.newCardHeader()
            .setTitle(title)
            .setImageUrl(imageURL)
            .setImageStyle(imageStyle);
        if (subtitle) {
            header.setSubtitle(subtitle);
        }
        return header;
    },

    /**
     * Builds am IconImage using the Iconify API.
     * 
     * @param {Object} options
     * @param {Icon} options.name 
     * @param {string} [options.color] 
     * @param {number} [options.width] 
     * @param {number} [options.height] 
     * @param {boolean} [options.box] 
     * @returns {GoogleAppsScript.Card_Service.IconImage}
     */
    icon({ name, color, width, height, box }) {
        // Sanitize the color (replace # with %23 for the URL)
        let safeColor = color?.startsWith("#") ? "%23" + color.substring(1) : color;

        // Collect provided options into a parameters array
        const params = [];
        if (safeColor) params.push(`color=${safeColor}`);
        if (width) params.push(`width=${width}`);
        if (height) params.push(`height=${height}`);
        if (box) params.push(`box=1`);

        // Construct the base URL
        let url = `https://api.iconify.design/${name}.svg`;

        // 4. Append query string only if there are parameters
        if (params.length > 0) {
            url += `?${params.join("&")}`;
        }

        // Extract the alt text from the last element of the name
        const nameParts = name.split("/");
        const altText = nameParts[nameParts.length - 1];

        // Return the constructed IconImage
        return CardService.newIconImage()
            .setIconUrl(url)
            .setAltText(altText);
    }
}

/** @enum {string} */
const Icon = {
    warning: "material-symbols/warning-rounded",
    folder_question: "mdi/folder-question",
}