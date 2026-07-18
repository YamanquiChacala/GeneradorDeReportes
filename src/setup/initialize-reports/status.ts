import { ReportSheetSchema } from "../../common/gas-parts";
import {
    BorderSide,
    buildBorderRequest,
    buildFieldsMask,
    buildMergeCellsRequest,
    buildTransferRequests,
    buildUnmergeCellsRequest,
    buildUpdateSheetPropertiesRequest,
    createRange,
    createRequiredGetter,
    type MappedNamedRange,
    MergeType,
    type ParsedSpreadsheet,
    RangeBehavior,
} from "../../common/gas-utils";
import {
    createSetupAbilityValidationFormula,
    createSetupCommentValidationFormula,
    createSetupGradeValidationFormula,
    createSetupSepCommentLenghtValidationFormula,
    createSetupTextValidationFormula,
    type ReportPersistentData,
} from "../../common/report-utils";
import { buildStatusSectionData, type CellFormulaClosure } from "../../common/setup-utils";
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

    return fillStatusSection({
        statusRange: statusInfoRange,
        studentRange: studentInfoRange,
        markColOffsets: [0],
        title: "Datos Generales",
        headers,
        persistentData,
        currentRowOffset: rowOffset,
        formulaFunction: createSetupTextValidationFormula,
    });
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

    return fillStatusSection({
        statusRange: statusAbilitiesRange,
        studentRange: studentAbilitiesRange,
        markColOffsets: [0, 1, 2, 3],
        title: "Habilidades de aprendizaje",
        headers: subjects,
        persistentData,
        currentRowOffset: rowOffset,
        formulaFunction: createSetupAbilityValidationFormula,
    });
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

    return fillStatusSection({
        statusRange: statusCommentsRange,
        studentRange: studentCommentsRange,
        markColOffsets: [0, 6],
        markUsesFieldValue: [false, true],
        title: "Observaciones",
        headers: subjects,
        persistentData,
        currentRowOffset: rowOffset,
        formulaFunction: [createSetupCommentValidationFormula, createSetupSepCommentLenghtValidationFormula],
    });
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
        const periodResult = fillStatusSection({
            statusRange,
            studentRange,
            markColOffsets: [0, 1, 2],
            title,
            headers: subjects,
            persistentData,
            currentRowOffset: newRowOffset,
            formulaFunction: createSetupGradeValidationFormula,
        });
        requests.push(...periodResult.requests);
        newRowOffset = periodResult.newRowOffset;
    }

    return { requests, newRowOffset };
}

interface FillStatusSectionParams {
    statusRange: MappedNamedRange;
    studentRange: MappedNamedRange;
    markColOffsets: number[];
    markUsesFieldValue?: boolean[];
    title: string;
    headers: string[];
    persistentData: ReportPersistentData;
    currentRowOffset: number;
    formulaFunction: CellFormulaClosure | CellFormulaClosure[];
}

/**
 * Helper to generate requests for each status sheet section
 */
function fillStatusSection({
    statusRange,
    studentRange,
    markColOffsets,
    markUsesFieldValue,
    title,
    headers,
    persistentData,
    currentRowOffset,
    formulaFunction,
}: FillStatusSectionParams): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    newRowOffset: number;
} {
    const fieldSubjects = persistentData.academicFields.reduce((acc: number[], field) => {
        const startIndex = acc.length;
        const newElements = Array(field.subjects).fill(startIndex);
        return acc.concat(newElements);
    }, []);

    // const fieldSubjects = persistentData.academicFields.flatMap((field, index) => Array.from({ length: field.subjects }, () => index));

    const { data, mergeRanges, borderRanges } = buildStatusSectionData({
        statusRange,
        studentRange,
        markColOffsets,
        markUsesFieldValue,
        title,
        headers,
        studentRows: persistentData.students,
        averagePerField: persistentData.configData.averagePerField,
        fieldSubjects,
        statusRowOffset: currentRowOffset,
        formulaFunction,
    });

    // Data transfer
    const dataTransferResult = buildTransferRequests({
        destination: statusRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset: currentRowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // Merge cells
    for (const mergeRange of mergeRanges) {
        formatRequests.push(buildMergeCellsRequest(mergeRange, MergeType.MERGE_ROWS));
    }

    // Borders
    for (const borderRange of borderRanges) {
        formatRequests.push(buildBorderRequest(borderRange, CssColorMap.lightgrey, BorderSide.RIGHT));
    }

    return { requests: [...dataTransferResult.requests, ...formatRequests], newRowOffset: dataTransferResult.rowOffset };
}
