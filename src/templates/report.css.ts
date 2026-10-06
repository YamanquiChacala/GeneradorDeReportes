import { Base64Fonts } from "../common/utils/base64-constants";

export const REPORT_CSS_STYLE = `@font-face { font-family: "Montserrat"; font-style: normal; font-weight: 400; src: url("${ Base64Fonts.MONTSERRAT_REGULAR }") format("truetype");}@page { size: letter landscape; margin: 25px;}body { font-family: "Montserrat", Arial, sans-serif; font-size: 0.6rem; color: #333;}`;