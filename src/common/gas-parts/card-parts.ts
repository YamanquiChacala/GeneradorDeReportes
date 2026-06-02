import { Urls } from "../gas-utils/types";
import type { Icon, IconifyParams } from "../utils";
import { iconifyUrl, webColor } from "../utils";

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

interface TextButtonParams {
    text: string;
    action: GoogleAppsScript.Card_Service.Action;
    style?: GoogleAppsScript.Card_Service.TextButtonStyle;
    backgroundColor?: string;
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
export function icon(args: IconifyParams): GoogleAppsScript.Card_Service.IconImage {
    const { url, name: displayName } = iconifyUrl(args);

    return CardService.newIconImage().setIconUrl(url).setAltText(displayName);
}

/**
 * Build a `TextButton` with the given action.
 */
export function textButton({ text, action, style, backgroundColor }: TextButtonParams): GoogleAppsScript.Card_Service.TextButton {
    const button = CardService.newTextButton().setText(text).setOnClickAction(action);
    if (style) button.setTextButtonStyle(style);
    const goodColor = webColor(backgroundColor);
    if (goodColor) button.setBackgroundColor(goodColor);
    return button;
}
