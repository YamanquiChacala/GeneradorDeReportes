import { ReportSheetSchema } from "../common/gas-parts";
import { createRequiredGetter, getStringRangeData, offsetGridRange, type ParsedSpreadsheet } from "../common/gas-utils";
import { type ReportPersistentData, StudentRowType } from "../common/report-utils";

export function generateSingleReport(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>, persistenData: ReportPersistentData, studentRowIndex: number) {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");
    const student = persistenData.students[studentRowIndex];

    if (student == null || student.type === StudentRowType.SEPARATOR) throw new Error("Alumno no válido.");

    const sheetName = student.sheetName;

    const sheet = parsedReport.extraSheets.find((sheet) => sheet.properties?.title === sheetName);

    if (sheet == null) throw new Error("No existe hoja del alumno.");

    const startDate = persistenData.configData.dates[0];
    const endDate = persistenData.configData.dates[3];

    const startYear = new Date(startDate).getUTCFullYear().toString();
    const endYear = new Date(endDate).getUTCFullYear().toString();

    const generalInfoRange = offsetGridRange({
        origin: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.generalInfo).namedRange.range,
        sheetId: sheet.properties?.sheetId,
    });
    const generalInfo = getStringRangeData(sheet, generalInfoRange);
}
