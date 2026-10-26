import type { Period } from "../common/report-utils";
import { loadReport } from "./load";

export function switchPeriod(reportFileId: string, newPeriod: Period, resetComments: boolean) {
    const { parsedReport, persistentData } = loadReport(reportFileId);

    if (persistentData.currentPeriod === newPeriod && !resetComments) return; // Nothing to change

    // Update each student's data

    // Reset abilities and comments

    // Update unprotected ranges

    // Hide calendar columns

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, reportFileId);
}
