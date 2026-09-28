import { headerIcon, iconImage } from "../common/gas-parts";
import { CssColorMap, Icon } from "../common/utils";

export function buildReportOptionsMainCard(_reportFileId: string): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Asistencia y Reportes", subtitle: "Montessori Chacala", iconName: Icon.CHART }));

    const cardSection = CardService.newCardSection().setHeader("Tablero de Administración");

    const action = CardService.newAction().setFunctionName(openLinkCallback.name);

    const reports = CardService.newDecoratedText()
        .setText("Reportes")
        .setBottomLabel("Generar reportes individuales o por grupo")
        .setStartIcon(iconImage({ iconName: Icon.PAGE }))
        .setEndIcon(iconImage({ iconName: Icon.FORWARD, color: CssColorMap.black }))
        .setOnClickAction(action);

    cardSection.addWidget(reports);

    return card.addSection(cardSection).build();
}

function openLinkCallback() {
    return CardService.newActionResponseBuilder().setOpenLink(CardService.newOpenLink().setUrl("https://www.google.com")).build();
}
