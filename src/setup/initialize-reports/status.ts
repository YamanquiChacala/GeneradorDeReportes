import { ReportSheetSchema } from "../../common/gas-parts";
import {
    buildFieldsMask,
    buildTransferRequests,
    //buildUnmergeCellsRequest,
    buildUpdateSheetPropertiesRequest,
    //createRange,
    createRequiredGetter,
    getA1Notation,
    type ParsedSpreadsheet,
    RangeBehavior,
    Style,
} from "../../common/gas-utils";
import { createSetupRowValidFormula, createSetupTextValidationFormula, type ReportPersistentData, StudentRowType } from "../../common/report-utils";

/**
 * Fills in the Status sheet with the subjects, students and the formulas to put ✔️ or ❌ everywhere.
 */
export function prepareStatusSheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    let rowOffset = 0;

    // Prepare the sheet
    const prepareRequests = prepareSheet(parsedReport, persistentData);

    // Fill in the data
    const { requests: generalInfoRequests, newRowOffset: generalInfoRowOffset } = fillGeneralInfo(parsedReport, persistentData, rowOffset);
    rowOffset = generalInfoRowOffset;

    // Fill in Habilities
    const { requests: habilitiesRequests, newRowOffset: habilitiesRowOffset } = fillHabilities(parsedReport, persistentData, rowOffset);
    rowOffset = habilitiesRowOffset;

    // Fill in the data
    return [...prepareRequests, ...generalInfoRequests, ...habilitiesRequests];
}

/**
 * Resizes, sorts and gets the sheet ready to input information.
 */
function prepareSheet(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const getSheet = createRequiredGetter(parsedReport.mappedSheets, "hoja de reporte");
    const statusSheetId = getSheet(ReportSheetSchema.sheets.status.sheetName).properties?.sheetId ?? 0;

    const frozenCols = 4;

    // Calculate width
    const width = frozenCols + Math.max(9, 4 * persistentData.subjects.length);

    // Set the properties
    const propertiesRequest = buildUpdateSheetPropertiesRequest({ sheetId: statusSheetId, columnCount: width, hidden: false, frozenColumnCount: frozenCols, index: 1 });

    // Remove merged cells to avoid problems with data input
    //const unfrozenRange = createRange(statusSheetId, 0, frozenCols);
    //const unmergeRequest = buildUnmergeCellsRequest(unfrozenRange);

    return [propertiesRequest]; // , unmergeRequest];
}

/**
 * Fill the General Data section of the sheet
 */
function fillGeneralInfo(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    const statusInfoRange = getMappedRange(ReportSheetSchema.sheets.status.ranges.info);
    const studentInfoRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.generalInfo);

    const userEnteredFormat: GoogleAppsScript.Sheets.Schema.CellFormat = {
        borders: {
            left: {
                style: Style.SOLID,
                colorStyle: { rgbColor: { red: 204, green: 204, blue: 204 } },
            },
        },
    };

    // Header
    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
        [
            { userEnteredValue: { stringValue: "Datos Generales" } },
            {},
            {},
            {},
            { userEnteredValue: { stringValue: "CURP" }, userEnteredFormat },
            {},
            {},
            { userEnteredValue: { stringValue: "Grado" }, userEnteredFormat },
            {},
            {},
            { userEnteredValue: { stringValue: "Nivel" }, userEnteredFormat },
            {},
            {},
        ],
    ];

    for (const [index, studentRow] of persistentData.students.entries()) {
        const rowData: GoogleAppsScript.Sheets.Schema.CellData[] = [];

        if (studentRow.type === StudentRowType.STUDENT) {
            const curpA1Cell = getA1Notation({
                mappedRange: studentInfoRange,
                includeSheetName: true,
                customSheetName: studentRow.sheetName,
                rowOffset: 0,
                width: 1,
                height: 1,
            });
            const gradoA1Cell = getA1Notation({
                mappedRange: studentInfoRange,
                includeSheetName: true,
                customSheetName: studentRow.sheetName,
                rowOffset: 1,
                width: 1,
                height: 1,
            });
            const nivelA1Cell = getA1Notation({
                mappedRange: studentInfoRange,
                includeSheetName: true,
                customSheetName: studentRow.sheetName,
                rowOffset: 2,
                width: 1,
                height: 1,
            });

            rowData.push(
                { userEnteredValue: { formulaValue: createSetupRowValidFormula(statusInfoRange, 1 + rowOffset + index, 4) } },
                { userEnteredValue: { numberValue: studentRow.id } },
                { userEnteredValue: { stringValue: studentRow.firstName } },
                { userEnteredValue: { stringValue: studentRow.lastName } },
                { userEnteredFormat },
                { userEnteredValue: { formulaValue: createSetupTextValidationFormula(curpA1Cell) } },
                {},
                { userEnteredFormat },
                { userEnteredValue: { formulaValue: createSetupTextValidationFormula(gradoA1Cell) } },
                {},
                { userEnteredFormat },
                { userEnteredValue: { formulaValue: createSetupTextValidationFormula(nivelA1Cell) } },
                {},
            );
        }

        data.push(rowData);
    }

    const result = buildTransferRequests({
        destination: statusInfoRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    return { requests: result.requests, newRowOffset: result.rowOffset };
}

/**
 * Fill the Status habilities formulas
 */
function fillHabilities(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: "Habilidades de aprendizaje" } }, {}, {}, {}];

    for (const subject of persistentData.subjects) {
    }

    return {};
}
