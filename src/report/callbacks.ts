import { getInputs } from "../common/gas-utils";
import { reportCreateReportsCard, StudentReportInputs } from "./cards";
import { loadReportPersistentData } from "./load";

/**
 * Builds and pushes on the stack the Report Menu Card
 */
export function onPushReportMenuCard(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const reportFileId = e?.sheets?.id;

    if (reportFileId == null) throw new Error("Unexpected error, file should be defined.");

    const persistentData = loadReportPersistentData(reportFileId);

    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().pushCard(reportCreateReportsCard(persistentData)))
        .build();
}

export function onGenerateIndividualReport(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error creando reporte individual: ";

    const inputs = getInputs(e.commonEventObject.formInputs, StudentReportInputs.schema);

    if (inputs.studentIndex == null) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Falta seleccionar alumno`))
            .build();
    }

    const fileName = inputs.fileName ? inputs.fileName : "Nombre del estudiante";

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`Index: ${inputs.studentIndex}: ${fileName}`))
        .build();
}
