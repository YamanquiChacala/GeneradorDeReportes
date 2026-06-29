import { ReportSheetSchema } from "../../common/gas-parts";
import {
    type A1NotationParams,
    buildFieldsMask,
    buildMergeCellsRequest,
    buildRightBorderRequest,
    buildTransferRequests,
    buildUnmergeCellsRequest,
    buildUpdateSheetPropertiesRequest,
    createRange,
    createRequiredGetter,
    getA1Notation,
    type MappedNamedRange,
    MergeType,
    offsetGridRange,
    type ParsedSpreadsheet,
    RangeBehavior,
} from "../../common/gas-utils";
import {
    createSetupAbilityValidationFormula,
    createSetupCommentValidationFormula,
    createSetupGradeValidationFormula,
    createSetupRowValidFormula,
    createSetupTextValidationFormula,
    type ReportPersistentData,
    type StudentRow,
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
    const { propertiesRequest, unmergeRequest } = prepareSheet(parsedReport, persistentData);

    // Fill in the data
    const generalInfoResult = fillGeneralInfo(parsedReport, persistentData, rowOffset);
    rowOffset = generalInfoResult.newRowOffset;

    // Fill in abilities
    const abilitiesResult = fillAbilities(parsedReport, persistentData, rowOffset);
    rowOffset = abilitiesResult.newRowOffset;

    // Fill in comments
    const commentsResult = fillComments(parsedReport, persistentData, rowOffset);
    rowOffset = commentsResult.newRowOffset;

    // Fill in Periods
    const periodsResult = fillPeriods(parsedReport, persistentData, rowOffset);
    rowOffset = periodsResult.newRowOffset;

    // Fill in the data
    return [
        propertiesRequest,
        unmergeRequest,
        ...generalInfoResult.requests,
        ...abilitiesResult.requests,
        ...commentsResult.requests,
        ...periodsResult.requests,
        propertiesRequest,
    ];
}

/**
 * Resizes, sorts and gets the sheet ready to input information.
 */
function prepareSheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): { propertiesRequest: GoogleAppsScript.Sheets.Schema.Request; unmergeRequest: GoogleAppsScript.Sheets.Schema.Request } {
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

    return { propertiesRequest, unmergeRequest };
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

    const headers = ["CURP", "Grado", "Nivel"];

    return fillStatusSection(
        statusInfoRange,
        studentInfoRange,
        "Datos Generales",
        headers,
        persistentData.students,
        3,
        rowOffset,
        false,
        createSetupTextValidationFormula,
    );
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

    const subjects = persistentData.subjects.map((weightedSubject) => weightedSubject.subject);

    return fillStatusSection(
        statusAbilitiesRange,
        studentAbilitiesRange,
        "Habilidades de aprendizaje",
        subjects,
        persistentData.students,
        4,
        rowOffset,
        true,
        createSetupAbilityValidationFormula,
    );
}

/**
 * Fill the Status comments formulas
 */
function fillComments(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");
    const statusCommentsRange = getMappedRange(ReportSheetSchema.sheets.status.ranges.comments);
    const studentCommentsRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedComments);

    const subjects = persistentData.subjects.map((weightedSubject) => weightedSubject.subject);

    return fillStatusSection(
        statusCommentsRange,
        studentCommentsRange,
        "Observaciones",
        subjects,
        persistentData.students,
        3,
        rowOffset,
        false,
        createSetupCommentValidationFormula,
    );
}

/**
 * Fill the Status Periods formulas
 */
function fillPeriods(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    const periodOp: Array<{ title: string; statusRange: MappedNamedRange; studentRange: MappedNamedRange }> = [
        {
            title: "1er Periodo",
            statusRange: getMappedRange(ReportSheetSchema.sheets.status.ranges.trim1),
            studentRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim1),
        },
        {
            title: "2do Periodo",
            statusRange: getMappedRange(ReportSheetSchema.sheets.status.ranges.trim2),
            studentRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim2),
        },
        {
            title: "3er Periodo",
            statusRange: getMappedRange(ReportSheetSchema.sheets.status.ranges.trim3),
            studentRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim3),
        },
    ];

    const subjects = persistentData.subjects.map((weightedSubject) => weightedSubject.subject);

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    let newRowOffset = rowOffset;

    for (const { title, statusRange, studentRange } of periodOp) {
        const periodResult = fillStatusSection(
            statusRange,
            studentRange,
            title,
            subjects,
            persistentData.students,
            3,
            newRowOffset,
            true,
            createSetupGradeValidationFormula,
        );
        requests.push(...periodResult.requests);
        newRowOffset = periodResult.newRowOffset;
    }

    return { requests, newRowOffset };
}

/**
 * Builds the data for the sections of the Status sheet
 */
function buildStatusSectionData(
    statusRange: MappedNamedRange,
    studentRange: MappedNamedRange,
    title: string,
    headers: string[],
    students: StudentRow[],
    colsPerItem: number,
    rowOffset: number,
    mergeOnlyheader: boolean,
    formulaFunction: (a1Cell: string) => string,
): GoogleAppsScript.Sheets.Schema.CellData[][] {
    
}

/**
 * Helper to generate requests for each status sheet section
 */
function fillStatusSection(
    statusRange: MappedNamedRange,
    studentRange: MappedNamedRange,
    title: string,
    headers: string[],
    students: StudentRow[],
    colsPerItem: number,
    rowOffset: number,
    mergeOnlyheader: boolean,
    formulaFunction: (a1Cell: string) => string,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    // Header
    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: title } }, {}, {}, {}];
    for (const label of headers) {
        header.push({ userEnteredValue: { stringValue: label } });
        for (let i = 1; i < colsPerItem; i++) {
            header.push({});
        }
    }
    data.push(header);

    // Student rows
    for (const [studentIndex, studentRow] of students.entries()) {
        if (studentRow.type === StudentRowType.SEPARATOR) {
            data.push([]);
            continue;
        }

        const studentDataRow: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { formulaValue: createSetupRowValidFormula(statusRange, 1 + rowOffset + studentIndex, 4) } },
            { userEnteredValue: { numberValue: studentRow.id } },
            { userEnteredValue: { stringValue: studentRow.firstName } },
            { userEnteredValue: { stringValue: studentRow.lastName } },
        ];

        headers.forEach((_, headerIndex) => {
            const a1Params: A1NotationParams = {
                mappedRange: studentRange,
                includeSheetName: true,
                customSheetName: studentRow.sheetName,
                rowOffset: headerIndex,
                height: 1,
                width: 1,
                lockColumns: true,
                lockRows: true,
            };
            if (mergeOnlyheader) {
                for (let i = 0; i < colsPerItem; i++) {
                    a1Params.colOffset = i;
                    const a1Cell = getA1Notation(a1Params);
                    studentDataRow.push({ userEnteredValue: { formulaValue: formulaFunction(a1Cell) } });
                }
            } else {
                const a1Cell = getA1Notation(a1Params);
                studentDataRow.push({ userEnteredValue: { formulaValue: formulaFunction(a1Cell) } });
                for (let i = 1; i < colsPerItem; i++) {
                    studentDataRow.push({});
                }
            }
        });

        data.push(studentDataRow);
    }

    // Data transfer
    const dataTransferResult = buildTransferRequests({
        destination: statusRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    headers.forEach((_, headerIndex) => {
        const range = offsetGridRange({
            origin: statusRange.namedRange.range,
            colOffset: 4 + colsPerItem * headerIndex,
            width: colsPerItem,
        });
        if (mergeOnlyheader) {
            const mergeRange = offsetGridRange({ origin: range, height: 1 });
            formatRequests.push(buildMergeCellsRequest(mergeRange));
        } else {
            formatRequests.push(buildMergeCellsRequest(range, MergeType.MERGE_ROWS));
        }
        formatRequests.push(buildRightBorderRequest(range, CssColorMap.lightgray));
    });

    return { requests: [...dataTransferResult.requests, ...formatRequests], newRowOffset: dataTransferResult.rowOffset };
}
