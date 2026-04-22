import { onPopCardStack } from "../common/callbacks";
import { headerImage, textButton } from "../common/cardParts";
import { Colors, Numbers } from "../common/enums";
import { buildUtilityCard } from "../common/premadeCards";
import { getInputs } from "../common/utils/googleAPI";
import { sanitizeFileName } from "../common/utils/text";
import { CreateSetupFileInputs, CreateSetupFileParams, EditSetupFileInputs } from "./cards";
import type { SetupFileData } from "./code";
import { createSetupFile, generateCalendar } from "./code";

/**
 * Callback to the button to create a new Initialization Group File.
 */
export function onCreateSetupFile(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error creando Registro de Grupo: ";

    const { folderId } = CreateSetupFileParams.parse(e.commonEventObject.parameters);

    const { groupName, attendancePerClass, averagePerField, dateStart, dateEndTrimester1, dateEndTrimester2, dateEnd } = getInputs(
        e.commonEventObject.formInputs,
        CreateSetupFileInputs.schema,
    );

    if (!folderId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Falta carpeta de destino.`))
            .build();
    }
    if (!groupName || !dateStart || !dateEndTrimester1 || !dateEndTrimester2 || !dateEnd || attendancePerClass == null || averagePerField == null) {
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

    const initData: SetupFileData = {
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
        createSetupFile(initData);
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error(errorMessage);
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}${errorMessage}`))
            .build();
    }

    const successCard = buildUtilityCard({
        header: headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessori Chacala" }),
        title: `<font color="${Colors.LOGO_OSCURO}"><b>✅ ¡Registro creado con éxito!</b></font>`,
        message: `El archivo de configuración para el grupo "<b>${groupName}</b>" ya está listo.`,
        points: ["El archivo aparecerá en la carpeta de Drive en un momento.", "Puedes seguir generando otros Registros."],
        button: textButton({ text: "Regresar al inicio", action: CardService.newAction().setFunctionName(onPopCardStack.name) }),
    });

    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(successCard)).build();
}

/**
 * Callback to regenerate the calendar in a Setup file.
 */
export function onGenerateCalendar(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error generando calendario: ";
    const { fileId } = e.commonEventObject.parameters;

    if (!fileId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}No se encuentra el archivo.`))
            .build();
    }

    try {
        generateCalendar(fileId);
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error(errorMessage);
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}${errorMessage}`))
            .build();
    }

    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("✅ Calendario creado.")).build();
}

/**
 * Callback to make a copy of the current Setup file.
 */
export function onCopySetup(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error copiando archivo: ";
    const { fileId } = e.commonEventObject.parameters;

    const { groupName, folderId } = getInputs(e.commonEventObject.formInputs, EditSetupFileInputs.schema);

    if (!fileId || !groupName || !folderId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Fantan parámetros`))
            .build();
    }

    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("✅ Copia creada.")).build();
}
