import { decoratedTextWithCallback, headerIcon, iconImage } from "../common/gas-parts";
import { defineActionParameters, ParamType } from "../common/gas-utils";
import { Icon } from "../common/utils";

export const ReportFileIdParam = defineActionParameters({
    reportFileId: ParamType.STRING,
} as const);

/**
 * Main Card for a report admin interface
 */
export function buildReportOptionsMainCard(reportFileId: string): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Asistencia y Reportes", subtitle: "Montessori Chacala", iconName: Icon.CHART }));

    const spacer = CardService.newTextParagraph().setText("&nbsp;");
    const cardSection = CardService.newCardSection()
        .setHeader("Tablero de Administración")
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: openLinkCallback.name, // TODO: Callback to card with details
                parameters: ReportFileIdParam.build({ reportFileId }),
                text: "Reportes",
                bottomText: "Generar reportes individuales o por grupo",
                startIcon: iconImage({ iconName: Icon.PAGE }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                parameters: ReportFileIdParam.build({ reportFileId }),
                text: "Periodos y Calendario",
                bottomText: "Configurar fechas, permisos y ciclo actual.",
                startIcon: iconImage({ iconName: Icon.CALENDAR }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                parameters: ReportFileIdParam.build({ reportFileId }),
                text: "Alumnos",
                bottomText: "Agregar, dar de baja o reordenar alumnos",
                startIcon: iconImage({ iconName: Icon.STUDENT }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                parameters: ReportFileIdParam.build({ reportFileId }),
                text: "Materias y Áreas",
                bottomText: "Administrar el plan de estudios y posiciones",
                startIcon: iconImage({ iconName: Icon.BOOKS }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        )
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                parameters: ReportFileIdParam.build({ reportFileId }),
                text: "Ponderaciones",
                bottomText: "Ajustar cálculo de califiaciones general y por materia",
                startIcon: iconImage({ iconName: Icon.WEIGHTS }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        );

    const peekHeader = headerIcon({ title: "Tablero de Administración", subtitle: "Asistencia y Reportes", iconName: Icon.CHART });

    return card.addSection(cardSection).setPeekCardHeader(peekHeader).build();
}

export function reportCreateReportsCard(_reportFileId: string): GoogleAppsScript.Card_Service.Card {
    const card = CardService.newCardBuilder().setHeader(headerIcon({ title: "Generación de Reportes", iconName: Icon.PAGE }));

    const spacer = CardService.newTextParagraph().setText("&nbsp;");
    const cardSection = CardService.newCardSection()
        .setHeader("Tablero de Administración")
        .addWidget(spacer)
        .addWidget(
            decoratedTextWithCallback({
                callback: "", // TODO: Callback to card with details
                text: "Reportes",
                bottomText: "Generar reportes individuales o por grupo",
                startIcon: iconImage({ iconName: Icon.PAGE }),
                endIcon: iconImage({ iconName: Icon.FORWARD }),
            }),
        );

    return card.addSection(cardSection).build();
}

function openLinkCallback() {
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().pushCard(reportCreateReportsCard("")))
        .build();
}
