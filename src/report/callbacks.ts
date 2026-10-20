import { getInputs } from "../common/gas-utils";
import { buildReportCreateReportsCard, StudentReportInputs } from "./cards";
import { generateSingleReport } from "./generateReports";
import { loadReportPersistentData } from "./load";

/**
 * Builds and pushes on the stack the Report Menu Card
 */
export function onPushReportMenuCard(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const reportFileId = e?.sheets?.id;

    if (reportFileId == null) throw new Error("Unexpected error, file should be defined.");

    const persistentData = loadReportPersistentData(reportFileId);

    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().pushCard(buildReportCreateReportsCard(persistentData)))
        .build();
}

/**
 *
 */
export function onGenerateIndividualReport(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error creando reporte individual: ";

    const { studentIndex, fileName } = getInputs(e.commonEventObject.formInputs, StudentReportInputs.schema);

    if (studentIndex == null) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Ningún alumno seleccionado`))
            .build();
    }

    const reportFileId = e?.sheets?.id;
    if (reportFileId == null) throw new Error("Unexpected error, file should be defined.");

    const parent = DriveApp.getFileById(reportFileId).getParents().next();

    const { html, name } = generateSingleReport(reportFileId, studentIndex);

    const pdfName = fileName ?? name;

    const pdfBlob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);

    // parent.createFile(`Reporte_${data.studentName}.html`, htmlString, MimeType.HTML);
    parent.createFile(pdfBlob.setName(`${pdfName}.pdf`));

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`Reporte "${pdfName}" creado`))
        .build();
}
