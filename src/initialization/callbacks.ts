import { onPopCardStack } from "../common/callbacks";
import { headerImage } from "../common/cardParts";
import { flattenFormInputs, sanitizeFileName } from "../common/utils";
import type { InitFileData } from "./code";
import { createInitializationFile, generateCalendar } from "./code";

/**
 * Callback to the button to create a new Initialization Group File.
 */
export function onCreateInitializationFile(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const { destinationFolder: folderId } = e.commonEventObject.parameters;

    const {
        groupName,
        attendancePerClass = false,
        averagePerField = false,
        dateStart,
        dateEndTrimester1,
        dateEndTrimester2,
        dateEnd,
    } = flattenFormInputs<InitFileData>(e.commonEventObject.formInputs);

    if (!folderId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("❌ Error creando Registro de Grupo: Falta carpeta de destino."))
            .build();
    }
    if (!groupName || !dateStart || !dateEndTrimester1 || !dateEndTrimester2 || !dateEnd) {
        return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("❌ Error creando Registro de Grupo: Faltan datos.")).build();
    }

    const groupNameSanitized = sanitizeFileName(groupName);

    if (!(dateStart < dateEndTrimester1 && dateEndTrimester1 < dateEndTrimester2 && dateEndTrimester2 < dateEnd)) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("❌ Error creando Registro de Grupo: Las fechas deben estar en orden ascendente."))
            .build();
    }

    const moreThanAYear = 400 * 24 * 60 * 60 * 1000;
    if (dateEnd - dateStart > moreThanAYear) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("❌ Error creando Registro de Grupo: Periodo demasiado largo."))
            .build();
    }

    const initData: InitFileData = {
        folderId,
        groupName: groupNameSanitized,
        attendancePerClass,
        averagePerField,
        dateStart,
        dateEndTrimester1,
        dateEndTrimester2,
        dateEnd,
    };

    try {
        createInitializationFile(initData);
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error(errorMessage);
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`❌ Error creando Registro de Grupo: ${errorMessage}`))
            .build();
    }

    const successCard = CardService.newCardBuilder()
        .setHeader(headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", image: "school" }))
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    CardService.newTextParagraph().setText(
                        `✅ Creación del Registro para "<b>${groupName}</b>" en proceso.<br><br>El archivo aparecerá la carpeta de Drive en un momento.<br><br>(Por favor espera al menos 1 minúto antes de intentarlo de nuevo.)`,
                    ),
                )
                .addWidget(CardService.newTextButton().setText("Regresar al inicio").setOnClickAction(CardService.newAction().setFunctionName(onPopCardStack.name))),
        )
        .build();

    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(successCard)).build();
}

/**
 *
 */
export function onGenerateCalendar(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const { fileId } = e.commonEventObject.parameters;

    if (!fileId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`❌ Error generando calendario: No se encuentra el archivo.`))
            .build();
    }

    try {
        generateCalendar(fileId);
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error(errorMessage);
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`❌ Error generando calendario: ${errorMessage}`))
            .build();
    }

    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("✅ Calendario creado.")).build();
}
