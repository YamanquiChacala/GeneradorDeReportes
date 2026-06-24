import { ReportSheetSchema } from "../../common/gas-parts";
import {
    buildFieldsMask,
    buildMergeCellsRequest,
    buildRightBorderRequest,
    buildTransferRequests,
    buildUnmergeCellsRequest,
    buildUpdateSheetPropertiesRequest,
    createRange,
    createRequiredGetter,
    getA1Notation,
    offsetGridRange,
    type ParsedSpreadsheet,
    RangeBehavior,
} from "../../common/gas-utils";
import {
    createSetupAbilityValidationFormula,
    createSetupRowValidFormula,
    createSetupTextValidationFormula,
    type ReportPersistentData,
    StudentRowType,
} from "../../common/report-utils";
import { CssColorMap } from "../../common/utils";

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

    // Fill in abilities
    const { requests: habilitiesRequests, newRowOffset: habilitiesRowOffset } = fillAbilities(parsedReport, persistentData, rowOffset);
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
    const propertiesRequest = buildUpdateSheetPropertiesRequest({ sheetId: statusSheetId, hidden: false, columnCount: width, frozenColumnCount: frozenCols, index: 1 });

    // Remove merged cells to avoid problems with data input
    const unfrozenRange = createRange(statusSheetId, 0, frozenCols);
    const unmergeRequest = buildUnmergeCellsRequest(unfrozenRange);

    return [propertiesRequest, unmergeRequest];
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

    // Header
    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
        [
            { userEnteredValue: { stringValue: "Datos Generales" } },
            {},
            {},
            {},
            { userEnteredValue: { stringValue: "CURP" } },
            {},
            {},
            { userEnteredValue: { stringValue: "Grado" } },
            {},
            {},
            { userEnteredValue: { stringValue: "Nivel" } },
            {},
            {},
        ],
    ];

    for (const [index, studentRow] of persistentData.students.entries()) {
        if (studentRow.type === StudentRowType.SEPARATOR) {
            data.push([]);
            continue;
        }

        const rowData: GoogleAppsScript.Sheets.Schema.CellData[] = [];

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
            {},
            { userEnteredValue: { formulaValue: createSetupTextValidationFormula(curpA1Cell) } },
            {},
            {},
            { userEnteredValue: { formulaValue: createSetupTextValidationFormula(gradoA1Cell) } },
            {},
            {},
            { userEnteredValue: { formulaValue: createSetupTextValidationFormula(nivelA1Cell) } },
            {},
        );

        data.push(rowData);
    }

    const dataTransferResult = buildTransferRequests({
        destination: statusInfoRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    for (let i = 0; i < 3; i++) {
        const borderRange = offsetGridRange({
            origin: statusInfoRange.namedRange.range,
            colOffset: 4 + 3 * i,
            width: 3,
        });
        const mergeRange = offsetGridRange({
            origin: borderRange,
            height: 1,
        });
        formatRequests.push(buildMergeCellsRequest(mergeRange));
        formatRequests.push(buildRightBorderRequest(borderRange, CssColorMap.lightgray));
    }

    return { requests: [...dataTransferResult.requests, ...formatRequests], newRowOffset: dataTransferResult.rowOffset };
}

/**
 * Fill the Status abilities formulas
 */
function fillAbilities(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");
    const statusAbilitiesRange = getMappedRange(ReportSheetSchema.sheets.status.ranges.abilities);
    const studentAbilitiesRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedAbilities);

    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    // Header
    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: "Habilidades de aprendizaje" } }, {}, {}, {}];
    for (const subject of persistentData.subjects) {
        header.push({ userEnteredValue: { stringValue: subject.subject } }, {}, {}, {});
    }
    data.push(header);

    for (const [studentIndex, studentRow] of persistentData.students.entries()) {
        if (studentRow.type === StudentRowType.SEPARATOR) {
            data.push([]);
            continue;
        }

        const studentData: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { formulaValue: createSetupRowValidFormula(statusAbilitiesRange, 1 + rowOffset + studentIndex, 4) } },
            { userEnteredValue: { numberValue: studentRow.id } },
            { userEnteredValue: { stringValue: studentRow.firstName } },
            { userEnteredValue: { stringValue: studentRow.lastName } },
        ];

        persistentData.subjects.forEach((_, subjectIndex) => {
            for (let i = 0; i < 4; i++) {
                const a1Cell = getA1Notation({
                    mappedRange: studentAbilitiesRange,
                    includeSheetName: true,
                    customSheetName: studentRow.sheetName,
                    rowOffset: subjectIndex,
                    colOffset: i,
                    height: 1,
                    width: 1,
                    lockRows: true,
                });
                studentData.push({ userEnteredValue: { formulaValue: createSetupAbilityValidationFormula(a1Cell) } });
            }
        });

        data.push(studentData);
    }

    const dataTransferResult = buildTransferRequests({
        destination: statusAbilitiesRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    for (let i = 0; i < persistentData.subjects.length; i++) {
        const borderRange = offsetGridRange({
            origin: statusAbilitiesRange.namedRange.range,
            colOffset: 4 + 4 * i,
            width: 4,
        });
        const mergeRange = offsetGridRange({
            origin: borderRange,
            height: 1,
        });
        formatRequests.push(buildMergeCellsRequest(mergeRange));
        formatRequests.push(buildRightBorderRequest(borderRange, CssColorMap.lightgray));
    }

    return { requests: [...dataTransferResult.requests, ...formatRequests], newRowOffset: dataTransferResult.rowOffset };
}
