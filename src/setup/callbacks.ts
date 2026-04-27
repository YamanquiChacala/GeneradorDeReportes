import { onPopCardStack } from "../common/callbacks";
import { headerIcon, headerImage, textButton } from "../common/card-parts";
import { Colors, Icon, Numbers } from "../common/enums";
import { buildUtilityCard } from "../common/premade-cards";
import { getInputs } from "../common/utils/api-types";
import { sanitizeFileName } from "../common/utils/text";
import { CopySetupFileInputs, CopySetupFileParams, CreateSetupFileInputs, CreateSetupFileParams, GenerateCalendarParams, InitializeReportParams } from "./cards";
import type { SetupFileData } from "./code";
import { copySetupFile, createSetupFile, generateCalendar, initializeReport } from "./code";

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
    const { setupFileId } = GenerateCalendarParams.parse(e.commonEventObject.parameters);

    if (!setupFileId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}No se encuentra el archivo.`))
            .build();
    }

    try {
        generateCalendar(setupFileId);
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
export function onCopySetupFile(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error copiando archivo: ";
    const { setupFileId: fileId } = CopySetupFileParams.parse(e.commonEventObject.parameters);

    const { groupName, folderId } = getInputs(e.commonEventObject.formInputs, CopySetupFileInputs.schema);

    if (!fileId || !groupName || !folderId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Fantan parámetros`))
            .build();
    }

    const validGroupName = sanitizeFileName(groupName);
    const realFolderId = folderId === "null" ? undefined : folderId;

    try {
        copySetupFile(fileId, realFolderId, validGroupName);
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}${errorMessage}`))
            .build();
    }

    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("✅ Copia creada.")).build();
}

/**
 * Callback to initializa a brand new Report based on a Setup file.
 */
export function onInitializeReport(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error creando archivo de reportes: ";
    const { setupFileId, parentId } = InitializeReportParams.parse(e.commonEventObject.parameters);

    if (!setupFileId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}No se encuentra el Registro Inicial de grupo.`))
            .build();
    }
    if (!parentId) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Carpeta no valida para crear Reporte.`))
            .build();
    }

    try {
        initializeReport(setupFileId, parentId);
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}${errorMessage}`))
            .build();
    }

    const successCard = buildUtilityCard({
        header: headerIcon({ title: "Registro Inicial de Grupos", subtitle: "Montessori Chacala", iconName: Icon.CLIPBOARD }),
        title: `<font color="${Colors.LOGO_OSCURO}"><b>✅ ¡Archivo de reportes del grupo creado con éxito!</b></font>`,
        message: `📊 El archivo de asistencias, calificaciones y reportes para el grupo ya está listo.`,
        points: [
            "El archivo aparecerá en la carpeta de Drive en un momento.",
            "Ya puedes cerrar este archivo",
            "Una vez que revises que el archivo de calificaciones, asistencias y reportes es correcto, puedes borrar este archivo.",
        ],
    });

    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(successCard)).build();
}
