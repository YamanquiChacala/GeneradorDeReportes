import { ReportSheetSchema } from "../common/gas-parts";
import {
    buildFieldsMask,
    createRequiredGetter,
    type ExtractRangeNames,
    getA1Notation,
    offsetGridRange,
    type ParsedSpreadsheet,
    PERSISTENT_DATA_KEY,
    parseSpreadsheet,
} from "../common/gas-utils";
import type { ReportPersistentData, Student } from "../common/report-utils";

type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

/**
 * Load the full report spreadsheet into memory. Includes all the values of the cells, so it can be big and slow.
 */
export function loadReport(reportFileId: string): { parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>; persistentData: ReportPersistentData } {
    const reportFieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>(
        "sheets.properties.sheetId",
        "sheets.properties.title",
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

/**
 * Load only the metadata PersistentData of the report. Fast, but no actual report data.
 */
export function loadReportPersistentData(reportFileId: string): ReportPersistentData {
    const reportFieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>("developerMetadata");

    const reportSpreadsheet = Sheets?.Spreadsheets.get(reportFileId, { fields: reportFieldsMask });
    const metadataObj = reportSpreadsheet?.developerMetadata?.find((m) => m.metadataKey === PERSISTENT_DATA_KEY);
    if (metadataObj == null) throw new Error("Missing persistent data in metadata");

    const persistentData: ReportPersistentData = JSON.parse(metadataObj.metadataValue ?? "");

    return persistentData;
}

export type StudentRangeKey = "generalInfo" | "trim1" | "trim1Totals" | "trim2" | "trim2Totals" | "trim3" | "trim3Totals" | "abilities" | "comments";

export type LoadedStudentData = {
    student: Student;
    ranges: Record<StudentRangeKey, string[][]>;
};

const BATCH_SIZE = 50;

export function loadStudentData(reportFileId: string, parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>, students: Student[]): LoadedStudentData[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    const getLastColumsRange = (rangeName: ExtractRangeNames<typeof ReportSheetSchema>, width: number) => {
        const origin = getMappedRange(rangeName).namedRange.range;
        return offsetGridRange({ origin, colOffset: (origin.endColumnIndex ?? 0) - (origin.startColumnIndex ?? 0) - width, width });
    };

    const template = ReportSheetSchema.sheets.studentTemplate.ranges;

    const originalRanges: Record<StudentRangeKey, GridRange> = {
        generalInfo: getMappedRange(template.generalInfo).namedRange.range,
        trim1: getLastColumsRange(template.trim1Subjects, 2),
        trim1Totals: getLastColumsRange(template.trim1Totals, 1),
        trim2: getLastColumsRange(template.trim2Subjects, 2),
        trim2Totals: getLastColumsRange(template.trim2Totals, 1),
        trim3: getLastColumsRange(template.trim3Subjects, 3),
        trim3Totals: getLastColumsRange(template.trim3Totals, 2),
        abilities: offsetGridRange({ origin: getMappedRange(template.abilities).namedRange.range, colOffset: 1, width: 4 }),
        comments: offsetGridRange({ origin: getMappedRange(template.comments).namedRange.range, colOffset: 1, width: 1 }),
    };

    const rangeKeys = Object.keys(originalRanges) as StudentRangeKey[];

    const studentsBySheet = new Map(students.map((s) => [s.sheetName, s]));

    const targets = parsedReport.extraSheets.flatMap((sheet) => {
        const sheetName = sheet.properties?.title;
        const student = sheetName ? studentsBySheet.get(sheetName) : undefined;
        if (!sheetName || !student) return [];
        return [{ student, sheetName, ranges: {} as Record<StudentRangeKey, string[][]> }];
    });

    const requests = targets.flatMap((target) =>
        rangeKeys.map((key) => ({
            target,
            key,
            a1: getA1Notation({ range: originalRanges[key], customSheetName: target.sheetName, includeSheetName: true }),
        })),
    );

    const fetched: (string[][] | undefined)[] = [];
    for (let i = 0; i < requests.length; i += BATCH_SIZE) {
        const chunk = requests.slice(i, i + BATCH_SIZE);
        const response = Sheets?.Spreadsheets.Values.batchGet(reportFileId, {
            ranges: chunk.map((r) => r.a1),
            valueRenderOption: "FORMATTED_VALUE",
        });
        const valueRanges = response?.valueRanges ?? [];
        chunk.forEach((_, j) => {
            fetched.push(valueRanges[j]?.values as string[][] | undefined);
        });
    }

    requests.forEach((req, i) => {
        req.target.ranges[req.key] = fetched[i] ?? [];
    });

    return targets.map(({ student, ranges }) => ({ student, ranges }));
}
