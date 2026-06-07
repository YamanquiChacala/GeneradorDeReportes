import { Urls } from "../gas-utils/types";
import type { IconifyParams } from "../utils";
import { iconifyUrl, webColor } from "../utils";

interface CommonHeaderParams {
    readonly title: string;
    readonly subtitle?: string;
    readonly cropCircle?: boolean;
}

interface BaseHeaderParams extends CommonHeaderParams {
    readonly imageUrl: string;
    readonly imageAltText: string;
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

interface HeaderImageParams extends CommonHeaderParams {
    readonly image?: string;
}

/**
 * Header with a local hosted image.
 */
export function headerImage({ title, subtitle, cropCircle = false, image = "school" }: HeaderImageParams): GoogleAppsScript.Card_Service.CardHeader {
    const imageUrl = `${Urls.MEDIA_SERVER}images/${image}_64.png`;

    return baseHeader({ title, subtitle, imageUrl, imageAltText: image, cropCircle });
}

type HeaderIconParams = CommonHeaderParams & IconifyParams;

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
export function iconImage(args: IconifyParams): GoogleAppsScript.Card_Service.IconImage {
    const { url, name: displayName } = iconifyUrl(args);

    return CardService.newIconImage().setIconUrl(url).setAltText(displayName);
}

interface TextButtonParams {
    readonly text: string;
    readonly action: GoogleAppsScript.Card_Service.Action;
    readonly style?: GoogleAppsScript.Card_Service.TextButtonStyle;
    readonly backgroundColor?: string;
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
