import { onPopCardStack } from "../common/callbacks";
import { headerImage, textButton } from "../common/cardParts";
import { Numbers } from "../common/enums";
import { buildUtilityCard } from "../common/premadeCards";
import { flattenFormInputs } from "../common/utils/googleAPI";
import { sanitizeFileName } from "../common/utils/text";
import type { InitFileData } from "./code";
import { createInitializationFile, generateCalendar } from "./code";

/**
 * Callback to the button to create a new Initialization Group File.
 */
export function onCreateInitializationFile(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error creando Registro de Grupo: ";

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
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Falta carpeta de destino.`))
            .build();
    }
    if (!groupName || !dateStart || !dateEndTrimester1 || !dateEndTrimester2 || !dateEnd) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Faltan datos.`))
            .build();
    }

    const groupNameSanitized = sanitizeFileName(groupName);

    if (!(dateStart < dateEndTrimester1 && dateEndTrimester1 < dateEndTrimester2 && dateEndTrimester2 < dateEnd)) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Las fechas deben estar en orden ascendente.`))
            .build();
    }

    if (dateEnd - dateStart > Numbers.MORE_THAN_A_YEAR) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Periodo demasiado largo.`))
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
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}${errorMessage}`))
            .build();
    }

    const successCard = buildUtilityCard({
        header: headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", image: "school" }),
        title: `✅ Creación del Registro Inicial para "<b>${groupName}</b>" listo.`,
        points: ["El archivo aparecerá la carpeta de Drive en un momento.", "Puedes seguir generando otros Registros."],
        button: textButton({ text: "Regresar al inicio", action: CardService.newAction().setFunctionName(onPopCardStack.name) }),
    });

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
