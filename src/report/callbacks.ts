import { buildUtilityCard, headerImage, notify, onPopCardStack, textButton } from "../common/gas-parts";
import { Colors, getInputs } from "../common/gas-utils";
import { Icon } from "../common/utils";
import { buildReportCreateReportsCard, StudentReportInputs } from "./cards";
import { generateAllReports, generateSingleReport } from "./generateReports";
import { loadReportPersistentData } from "./load";

const JOB_PREFIX = "reportJob:";
const STALE_AFTER_MS = 15 * 60 * 1000;
const BUDGET_MS = 4.5 * 60 * 1000;

interface JobState {
    fileId: string;
    next: number;
    updatedAt: number;
}

function pendingJobs(props: GoogleAppsScript.Properties.Properties): [string, JobState][] {
    return Object.entries(props.getProperties())
        .filter(([key]) => key.startsWith(JOB_PREFIX))
        .map(([key, raw]) => [key, JSON.parse(raw) as JobState]);
}

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

export function onGenerateAllReports(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    try {
        const fileId = e?.sheets?.id;
        if (fileId == null) throw new Error("Unexpected error, file should be defined.");

        const props = PropertiesService.getUserProperties();
        const key = JOB_PREFIX + fileId;
        const existing = props.getProperty(key);
        if (existing && Date.now() - (JSON.parse(existing) as JobState).updatedAt < STALE_AFTER_MS) {
            return notify("Ya hay una generación en curso para este archivo.");
        }

        props.setProperty(key, JSON.stringify({ fileId, next: 0, updatedAt: Date.now() } satisfies JobState));
        ScriptApp.newTrigger(triggerProcessGenerateAllReports.name).timeBased().after(1000).create();

        const successCard = buildUtilityCard({
            header: headerImage({ title: "Generación de Reportes", image: Icon.PAGE }),
            title: `<font color="${Colors.LOGO_OSCURO}"><b>⏳ Generando reportes en segundo plano</b></font>`,
            message: "Tus reportes se están creando. Puedes cerrar esta ventana, el proceso continuará por su cuenta.",
            points: [
                "Puede tardar hasta <b>10 minutos</b> en terminar.",
                "Los reportes aparecerán como archivos PDF en la <b>misma carpeta de Drive</b> donde está este archivo.",
                "<b>No inicies otra generación</b> hasta que todos los reportes hayan aparecido en la carpeta.",
            ],
            button: textButton({ text: "Regresar al inicio", action: CardService.newAction().setFunctionName(onPopCardStack.name) }),
        });

        return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(successCard)).build();

        // return notify("Generando reportes en segundo plano. Aparecerán en la carpeta en unos minutos.");
    } catch (error) {
        console.error(error);
        return notify(`❌ Error creando reportes: ${error instanceof Error ? error.message : String(error)}`);
    }
}

/**
 * Generates a PDF report for the selected student and saves it in the spreadsheet's parent folder.
 * Uses the optional custom file name when provided, otherwise the generated report name.
 * Returns a notification if no student is selected or after the report is created.
 */
export function onGenerateIndividualReport(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    const mainErrorMessage = "❌ Error creando reporte individual: ";

    try {
        const { studentIndex, fileName } = getInputs(e.commonEventObject.formInputs, StudentReportInputs.schema);
        if (studentIndex == null) {
            throw new Error("Ningún alumno seleccionado");
        }

        const reportFileId = e?.sheets?.id;
        if (reportFileId == null) throw new Error("Unexpected error, file should be defined.");

        const parent = DriveApp.getFileById(reportFileId).getParents().next();
        const { html, name } = generateSingleReport(reportFileId, studentIndex);

        const pdfName = fileName ?? name;
        const pdfBlob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);
        parent.createFile(pdfBlob.setName(`${pdfName}.pdf`));

        return notify(`Reporte "${pdfName}" creado`);
    } catch (error) {
        console.error(error);
        const errorMessage = error instanceof Error ? error.message : String(error);
        return notify(`${mainErrorMessage}${errorMessage}`);
    }
}

export function triggerProcessGenerateAllReports() {
    // One worker per user at a time; if another is running, it will pick up our job when it rescans
    const lock = LockService.getUserLock();
    if (!lock.tryLock(0)) return;

    try {
        for (const t of ScriptApp.getProjectTriggers()) {
            if (t.getHandlerFunction() === triggerProcessGenerateAllReports.name) ScriptApp.deleteTrigger(t);
        }

        const props = PropertiesService.getUserProperties();
        const start = Date.now();

        for (const [key, state] of pendingJobs(props)) {
            if (Date.now() - start >= BUDGET_MS) break;
            try {
                const parent = DriveApp.getFileById(state.fileId).getParents().next();
                const reports = generateAllReports(state.fileId);

                while (state.next < reports.length && Date.now() - start < BUDGET_MS) {
                    // biome-ignore lint/style/noNonNullAssertion: Guaranteed to be in range
                    const { html, name } = reports[state.next]!;
                    parent.createFile(HtmlService.createHtmlOutput(html).getAs(MimeType.PDF).setName(`${name}.pdf`));
                    state.next++;
                    state.updatedAt = Date.now();
                    props.setProperty(key, JSON.stringify(state)); // checkpoint
                }

                if (state.next >= reports.length) props.deleteProperty(key);
            } catch (error) {
                console.error(`Job ${key} failed`, error);
                props.deleteProperty(key); // don't let a broken job block the queue
            }
        }

        // Rescan before releasing the lock: covers unfinished jobs and ones enqueued meanwhile
        if (pendingJobs(props).length > 0) {
            ScriptApp.newTrigger(triggerProcessGenerateAllReports.name).timeBased().after(1000).create();
        }
    } finally {
        lock.releaseLock();
    }
}
