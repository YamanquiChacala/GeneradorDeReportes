import { type Icon, Urls } from "./enums";
import { type IconifyParams, iconifyUrl } from "./utils/image";

interface BaseHeaderParams {
    title: string;
    subtitle?: string;
    imageUrl: string;
    imageAltText: string;
    cropCircle?: boolean;
}

interface HeaderImageParams {
    title: string;
    subtitle?: string;
    image?: string;
    cropCircle?: boolean;
}

interface HeaderIconParams {
    title: string;
    subtitle?: string;
    cropCircle?: boolean;
    iconName: Icon;
    color?: string;
    width?: number;
    height?: number;
    box?: boolean;
}

/**
 * Base builder for headers.
 */
function baseHeader({ title, subtitle, imageUrl, imageAltText, cropCircle }: BaseHeaderParams): GoogleAppsScript.Card_Service.CardHeader {
    const imageStyle = cropCircle ? CardService.ImageStyle.CIRCLE : CardService.ImageStyle.SQUARE;
    const newHeader = CardService.newCardHeader().setTitle(title).setImageUrl(imageUrl).setImageAltText(imageAltText).setImageStyle(imageStyle);
    if (subtitle) {
        newHeader.setSubtitle(subtitle);
    }
    return newHeader;
}

/**
 * Header with a local hosted image.
 */
export function headerImage({ title, subtitle, image = "school", cropCircle = false }: HeaderImageParams): GoogleAppsScript.Card_Service.CardHeader {
    const imageUrl = `${Urls.MEDIA_SERVER}images/${image}_64.png`;

    return baseHeader({ title, subtitle, imageUrl, imageAltText: image, cropCircle });
}

/**
 * Header using an Iconify icon.
 */
export function headerIcon({ title, subtitle, cropCircle, iconName, color, width, height, box }: HeaderIconParams): GoogleAppsScript.Card_Service.CardHeader {
    const { url: imageUrl, name: imageAltText } = iconifyUrl({ iconName, color, width, height, box });

    return baseHeader({ title, subtitle, imageUrl, imageAltText, cropCircle });
}

/**
 * Builds am IconImage using the Iconify API.
 */
export function icon({ iconName, color, width, height, box }: IconifyParams): GoogleAppsScript.Card_Service.IconImage {
    const { url, name: displayName } = iconifyUrl({ iconName, color, width, height, box });

    return CardService.newIconImage().setIconUrl(url).setAltText(displayName);
}
