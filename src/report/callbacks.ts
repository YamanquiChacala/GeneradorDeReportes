import { getInputs } from "../common/gas-utils";
import { Base64Images } from "../common/utils/base64-constants";
import { type ReportData, renderReport } from "../templates/reportTemplate";
import { buildReportCreateReportsCard, StudentReportInputs } from "./cards";
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

export function onGenerateIndividualReport(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error creando reporte individual: ";

    const inputs = getInputs(e.commonEventObject.formInputs, StudentReportInputs.schema);

    if (inputs.studentIndex == null) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(`${mainErrorMessage}Falta seleccionar alumno`))
            .build();
    }

    const reportFileId = e?.sheets?.id;
    if (reportFileId == null) throw new Error("Unexpected error, file should be defined.");

    const parent = DriveApp.getFileById(reportFileId).getParents().next();

    const fileName = inputs.fileName ? inputs.fileName : "Nombre del estudiante";

    const data: ReportData = {
        studentName: "Yamanqui",
        photoBase64: Base64Images.SCHOOL,
        grades: [
            { subject: "Español", score: 9 },
            { subject: "Matemáticas", score: 5 },
        ],
    };

    const htmlString = renderReport(data);

    const pdfBlob = HtmlService.createHtmlOutput(htmlString).getAs(MimeType.PDF);

    parent.createFile(`Reporte_${data.studentName}.html`, htmlString, MimeType.HTML);
    parent.createFile(pdfBlob.setName(`Reporte_${data.studentName}.pdf`));

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`Index: ${inputs.studentIndex}: ${fileName}`))
        .build();
}
