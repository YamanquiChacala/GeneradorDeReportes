import { ReportSheetSchema } from "../common/gas-parts";
import { buildFieldsMask, type ParsedSpreadsheet, PERSISTENT_DATA_KEY, parseSpreadsheet } from "../common/gas-utils";
import type { ReportPersistentData } from "../common/report-utils";

export function loadReport(reportFileId: string): { parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>; persistentData: ReportPersistentData } {
    const reportFieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>(
        "sheets.properties.sheetId",
        "sheets.properties.title",
        "sheets.properties.gridProperties.rowCount",
        "sheets.properties.gridProperties.columnCount",
        "sheets.protectedRanges.protectedRangeId",
        "namedRanges",
        "developerMetadata",
    );

    const reportSpreadsheet = Sheets?.Spreadsheets.get(reportFileId, { fields: reportFieldsMask });
    const parsedReport = parseSpreadsheet(reportSpreadsheet, ReportSheetSchema);

    const metadataObj = reportSpreadsheet?.developerMetadata?.find((m) => m.metadataKey === PERSISTENT_DATA_KEY);
    if (metadataObj == null) throw new Error("Missing persistent data in metadata");

    const persistentData: ReportPersistentData = JSON.parse(metadataObj.metadataValue ?? "");

    return { parsedReport, persistentData };
}

export function loadReportPersistentData(reportFileId: string): ReportPersistentData {
    const reportFieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>("developerMetadata");

    const reportSpreadsheet = Sheets?.Spreadsheets.get(reportFileId, { fields: reportFieldsMask });
    const metadataObj = reportSpreadsheet?.developerMetadata?.find((m) => m.metadataKey === PERSISTENT_DATA_KEY);
    if (metadataObj == null) throw new Error("Missing persistent data in metadata");

    const persistentData: ReportPersistentData = JSON.parse(metadataObj.metadataValue ?? "");

    return persistentData;
}
