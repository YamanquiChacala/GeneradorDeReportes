const CardParts = {

    /**
     * Base builder for headers.
     * @param {Object} options
     * @param {string} options.title
     * @param {string} [options.subtitle]
     * @param {string} options.imageUrl 
     * @param {string} options.imageAltText 
     * @param {boolean} [options.cropCircle]
     * @returns {GoogleAppsScript.Card_Service.CardHeader} The header, ready to insert on a Card
     */
    _header({ title, subtitle, imageUrl, imageAltText, cropCircle }) {
        const imageStyle = cropCircle ? CardService.ImageStyle.CIRCLE : CardService.ImageStyle.SQUARE;
        const newHeader = CardService.newCardHeader()
            .setTitle(title)
            .setImageUrl(imageUrl)
            .setImageAltText(imageAltText)
            .setImageStyle(imageStyle);
        if (subtitle) {
            newHeader.setSubtitle(subtitle);
        }
        return newHeader;
    },

    /**
     * Header with a local hosted image.
     * @param {Object} options
     * @param {string} options.title The title on the header
     * @param {string} [options.subtitle] The subtitle on the header
     * @param {string} [options.image] The name of the image to use as icon. A file called <image>_48.png should exist in images folder.
     * @param {boolean} [options.cropCircle] Crop the icon into a circle?
     * @returns {GoogleAppsScript.Card_Service.CardHeader} The header, ready to insert on a Card
     */
    headerImage({ title, subtitle, image = 'icon', cropCircle = false }) {
        const imageUrl = `https://media.githubusercontent.com/media/YamanquiChacala/GeneradorDeReportes/refs/heads/main/images/${image}_64.png`;

        return CardParts._header({ title, subtitle, imageUrl, imageAltText: image, cropCircle });
    },

    /**
     * Header using an Iconify icon.
     * @param {Object} options
     * @param {string} options.title 
     * @param {string} [options.subtitle] 
     * @param {boolean} [options.cropCircle] 
     * @param {Icon} options.iconName 
     * @param {string} [options.color] 
     * @param {number} [options.width] 
     * @param {number} [options.height] 
     * @param {boolean} [options.box]
     * @returns {GoogleAppsScript.Card_Service.CardHeader}
     */
    headerIcon({ title, subtitle, cropCircle, iconName, color, width, height, box }) {
        const { url: imageUrl, name: imageAltText } = Utils.iconifyUrl({ iconName, color, width, height, box });

        return CardParts._header({ title, subtitle, imageUrl, imageAltText, cropCircle });
    },

    /**
     * Builds am IconImage using the Iconify API.
     * @param {Object} options
     * @param {Icon} options.iconName 
     * @param {string} [options.color] 
     * @param {number} [options.width] 
     * @param {number} [options.height] 
     * @param {boolean} [options.box] 
     * @returns {GoogleAppsScript.Card_Service.IconImage}
     */
    icon({ iconName, color, width, height, box }) {
        const { url, name: displayName } = Utils.iconifyUrl({ iconName, color, width, height, box });

        return CardService.newIconImage()
            .setIconUrl(url)
            .setAltText(displayName);
    },

    /**
     * Build a Card with a header and text.
     * @param {GoogleAppsScript.Card_Service.CardHeader} header 
     * @param {string} htmlText 
     * @returns {GoogleAppsScript.Card_Service.Card}
     */
    buildParagraphCard(header, htmlText) {
        return CardService.newCardBuilder()
            .setHeader(header)
            .addSection(CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(htmlText)))
            .build();
    },
}