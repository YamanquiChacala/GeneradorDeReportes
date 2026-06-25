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
    type Student,
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

    // Fill in Period1
    const { requests: period1Requests, newRowOffset: period1RowOffset } = fillPeriod1(parsedReport, persistentData, rowOffset);
    rowOffset = period1RowOffset;

    // Fill in Period2
    const { requests: period2Requests, newRowOffset: period2RowOffset } = fillPeriod2(parsedReport, persistentData, rowOffset);
    rowOffset = period2RowOffset;

    // Fill in Period3
    const { requests: period3Requests, newRowOffset: period3RowOffset } = fillPeriod3(parsedReport, persistentData, rowOffset);
    rowOffset = period3RowOffset;

    // Fill in the data
    return [
        propertiesRequest,
        unmergeRequest,
        ...generalInfoResult.requests,
        ...abilitiesResult.requests,
        ...commentsResult.requests,
        ...period1Requests,
        ...period2Requests,
        ...period3Requests,
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
            { userEnteredValue: { formulaValue: createSetupTextValidationFormula(curpA1Cell) } },
            {},
            {},
            { userEnteredValue: { formulaValue: createSetupTextValidationFormula(gradoA1Cell) } },
            {},
            {},
            { userEnteredValue: { formulaValue: createSetupTextValidationFormula(nivelA1Cell) } },
            {},
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
        const range = offsetGridRange({
            origin: statusInfoRange.namedRange.range,
            colOffset: 4 + 3 * i,
            width: 3,
        });
        formatRequests.push(buildMergeCellsRequest(range, MergeType.MERGE_ROWS));
        formatRequests.push(buildRightBorderRequest(range, CssColorMap.lightgray));
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

    return fillSection(statusAbilitiesRange, persistentData, "Habilidades de aprendizaje", 4, rowOffset, true, (studentRow, subjectIndex) => {
        const cells: GoogleAppsScript.Sheets.Schema.CellData[] = [];
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
                lockColumns: true,
            });
            cells.push({ userEnteredValue: { formulaValue: createSetupAbilityValidationFormula(a1Cell) } });
        }
        return cells;
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

    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    // Header
    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: "Observaciones" } }, {}, {}, {}];
    for (const subject of persistentData.subjects) {
        header.push({ userEnteredValue: { stringValue: subject.subject } }, {}, {});
    }
    data.push(header);

    for (const [studentIndex, studentRow] of persistentData.students.entries()) {
        if (studentRow.type === StudentRowType.SEPARATOR) {
            data.push([]);
            continue;
        }

        const studentData: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { formulaValue: createSetupRowValidFormula(statusCommentsRange, 1 + rowOffset + studentIndex, 4) } },
            { userEnteredValue: { numberValue: studentRow.id } },
            { userEnteredValue: { stringValue: studentRow.firstName } },
            { userEnteredValue: { stringValue: studentRow.lastName } },
        ];

        persistentData.subjects.forEach((_, subjectIndex) => {
            const a1Cell = getA1Notation({
                mappedRange: studentCommentsRange,
                includeSheetName: true,
                customSheetName: studentRow.sheetName,
                rowOffset: subjectIndex,
                height: 1,
                width: 1,
                lockRows: true,
                lockColumns: true,
            });
            studentData.push({ userEnteredValue: { formulaValue: createSetupCommentValidationFormula(a1Cell) } }, {}, {});
        });

        data.push(studentData);
    }

    const dataTransferResult = buildTransferRequests({
        destination: statusCommentsRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    for (let i = 0; i < persistentData.subjects.length; i++) {
        const range = offsetGridRange({
            origin: statusCommentsRange.namedRange.range,
            colOffset: 4 + 3 * i,
            width: 3,
        });
        formatRequests.push(buildMergeCellsRequest(range, MergeType.MERGE_ROWS));
        formatRequests.push(buildRightBorderRequest(range, CssColorMap.lightgray));
    }

    return { requests: [...dataTransferResult.requests, ...formatRequests], newRowOffset: dataTransferResult.rowOffset };
}

/**
 * Fill the Status period 1 formulas
 */
function fillPeriod1(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");
    const statusPeriodRange = getMappedRange(ReportSheetSchema.sheets.status.ranges.trim1);
    const studentPeriodRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim1);

    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    // Header
    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: "1er Periodo" } }, {}, {}, {}];
    for (const subject of persistentData.subjects) {
        header.push({ userEnteredValue: { stringValue: subject.subject } }, {}, {});
    }
    data.push(header);

    for (const [studentIndex, studentRow] of persistentData.students.entries()) {
        if (studentRow.type === StudentRowType.SEPARATOR) {
            data.push([]);
            continue;
        }

        const studentData: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { formulaValue: createSetupRowValidFormula(statusPeriodRange, 1 + rowOffset + studentIndex, 4) } },
            { userEnteredValue: { numberValue: studentRow.id } },
            { userEnteredValue: { stringValue: studentRow.firstName } },
            { userEnteredValue: { stringValue: studentRow.lastName } },
        ];

        persistentData.subjects.forEach((_, subjectIndex) => {
            for (let i = 0; i < 3; i++) {
                const a1Cell = getA1Notation({
                    mappedRange: studentPeriodRange,
                    includeSheetName: true,
                    customSheetName: studentRow.sheetName,
                    rowOffset: subjectIndex,
                    colOffset: i,
                    height: 1,
                    width: 1,
                    lockRows: true,
                });
                studentData.push({ userEnteredValue: { formulaValue: createSetupGradeValidationFormula(a1Cell) } });
            }
        });

        data.push(studentData);
    }

    const dataTransferResult = buildTransferRequests({
        destination: statusPeriodRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    for (let i = 0; i < persistentData.subjects.length; i++) {
        const borderRange = offsetGridRange({
            origin: statusPeriodRange.namedRange.range,
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
 * Fill the Status period 2 formulas
 */
function fillPeriod2(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");
    const statusPeriodRange = getMappedRange(ReportSheetSchema.sheets.status.ranges.trim2);
    const studentPeriodRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim2);

    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    // Header
    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: "2do Periodo" } }, {}, {}, {}];
    for (const subject of persistentData.subjects) {
        header.push({ userEnteredValue: { stringValue: subject.subject } }, {}, {});
    }
    data.push(header);

    for (const [studentIndex, studentRow] of persistentData.students.entries()) {
        if (studentRow.type === StudentRowType.SEPARATOR) {
            data.push([]);
            continue;
        }

        const studentData: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { formulaValue: createSetupRowValidFormula(statusPeriodRange, 1 + rowOffset + studentIndex, 4) } },
            { userEnteredValue: { numberValue: studentRow.id } },
            { userEnteredValue: { stringValue: studentRow.firstName } },
            { userEnteredValue: { stringValue: studentRow.lastName } },
        ];

        persistentData.subjects.forEach((_, subjectIndex) => {
            for (let i = 0; i < 3; i++) {
                const a1Cell = getA1Notation({
                    mappedRange: studentPeriodRange,
                    includeSheetName: true,
                    customSheetName: studentRow.sheetName,
                    rowOffset: subjectIndex,
                    colOffset: i,
                    height: 1,
                    width: 1,
                    lockRows: true,
                });
                studentData.push({ userEnteredValue: { formulaValue: createSetupGradeValidationFormula(a1Cell) } });
            }
        });

        data.push(studentData);
    }

    const dataTransferResult = buildTransferRequests({
        destination: statusPeriodRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    for (let i = 0; i < persistentData.subjects.length; i++) {
        const borderRange = offsetGridRange({
            origin: statusPeriodRange.namedRange.range,
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
 * Fill the Status period 3 formulas
 */
function fillPeriod3(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");
    const statusPeriodRange = getMappedRange(ReportSheetSchema.sheets.status.ranges.trim3);
    const studentPeriodRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim3);

    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    // Header
    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: "3er Periodo" } }, {}, {}, {}];
    for (const subject of persistentData.subjects) {
        header.push({ userEnteredValue: { stringValue: subject.subject } }, {}, {});
    }
    data.push(header);

    for (const [studentIndex, studentRow] of persistentData.students.entries()) {
        if (studentRow.type === StudentRowType.SEPARATOR) {
            data.push([]);
            continue;
        }

        const studentData: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { formulaValue: createSetupRowValidFormula(statusPeriodRange, 1 + rowOffset + studentIndex, 4) } },
            { userEnteredValue: { numberValue: studentRow.id } },
            { userEnteredValue: { stringValue: studentRow.firstName } },
            { userEnteredValue: { stringValue: studentRow.lastName } },
        ];

        persistentData.subjects.forEach((_, subjectIndex) => {
            for (let i = 0; i < 3; i++) {
                const a1Cell = getA1Notation({
                    mappedRange: studentPeriodRange,
                    includeSheetName: true,
                    customSheetName: studentRow.sheetName,
                    rowOffset: subjectIndex,
                    colOffset: i,
                    height: 1,
                    width: 1,
                    lockRows: true,
                });
                studentData.push({ userEnteredValue: { formulaValue: createSetupGradeValidationFormula(a1Cell) } });
            }
        });

        data.push(studentData);
    }

    const dataTransferResult = buildTransferRequests({
        destination: statusPeriodRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowOffset,
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.INSERT_DELETE_CELLS,
    });

    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    for (let i = 0; i < persistentData.subjects.length; i++) {
        const borderRange = offsetGridRange({
            origin: statusPeriodRange.namedRange.range,
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
 * Helper to generate requests for each section
 */
function fillSection(
    statusRange: MappedNamedRange,
    persistentData: ReportPersistentData,
    title: string,
    colsPerSubject: number,
    rowOffset: number,
    mergeOnlyheader: boolean,
    getStudentSubjectDataRow: (studentRow: Student, subjectIndex: number) => GoogleAppsScript.Sheets.Schema.CellData[],
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    // Header
    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: title } }, {}, {}, {}];
    for (const subject of persistentData.subjects) {
        header.push({ userEnteredValue: { stringValue: subject.subject } });
        for (let i = 1; i < colsPerSubject; i++) {
            header.push({});
        }
    }
    data.push(header);

    // Student rows
    for (const [studentIndex, studentRow] of persistentData.students.entries()) {
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

        persistentData.subjects.forEach((_, subjectIndex) => {
            studentDataRow.push(...getStudentSubjectDataRow(studentRow, subjectIndex));
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
    persistentData.subjects.forEach((_, subjectIndex) => {
        const range = offsetGridRange({
            origin: statusRange.namedRange.range,
            colOffset: 4 + colsPerSubject * subjectIndex,
            width: colsPerSubject,
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
