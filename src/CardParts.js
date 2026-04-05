const CardParts = {

    /**
     * Normal header for a Card.
     * 
     * @param {string} title The title on the header
     * @param {string} [subtitle] The subtitle on the header
     * @param {string} [icon] The name of the image to use as icon. A file called <icon>_48.png should exist in images folder.
     * @param {boolean} [circle] Crop the icon into a circle? Defaults to true.
     * @returns {GoogleAppsScript.Card_Service.CardHeader} The header, ready to insert on a Card
     */
    header(title, subtitle, icon = 'icon', circle = true) {
        const imageURL = `https://media.githubusercontent.com/media/YamanquiChacala/Cats/refs/heads/main/images/${icon}_48.png`;
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
    }
}