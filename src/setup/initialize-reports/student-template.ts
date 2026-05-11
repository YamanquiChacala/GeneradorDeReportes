import { DEFAULT_COMMENT, TRIMESTER_NAMES } from "../../common/constants";
import { Dimension, MergeType } from "../../common/gas-enums";
import { ReportSheetSchema } from "../../common/sheet-schema";
import { buildFieldsMask } from "../../common/utils/gas-types";
import { buildAddNamedRangeRequest, buildMergeCellsRequest, buildTransferRequests, getColumnLetter, offsetGridRange } from "../../common/utils/gas-utils";
import type { ExtractRangeNames, MappedNamedRange, ParsedSpreadsheet } from "../../common/utils/mapped-name-range";
import { createStudentAsistanceFormula, createStudentAsistancePerSubjectFormula, generatePeriodString, type ReportPersistentData } from "../../common/utils/report-utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Prepares the student template sheet, filling in subjects, formulas, etc, so it can be duplicated for each student.
 */
export function prepareStudentTemplate(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    // Adapt the sheet's ranges
    const adaptSheetRangesRequests = adaptSizeAndRanges(parsedReport.namedRanges, persistentData);

    // Fill in data
    const infoRequests = prepareInfo(parsedReport.namedRanges, persistentData);

    // Prepare abilities
    const abilitiesRequests = prepareAbilities(parsedReport.namedRanges, persistentData);

    // Prepare comments
    const commentsRequests = prepareComments(parsedReport.namedRanges, persistentData);

    // Prepare Trimester 1
    const subjectRequests = prepareSubjects(parsedReport.namedRanges, persistentData);

    // Build requests
    return [...adaptSheetRangesRequests, ...infoRequests, ...abilitiesRequests, ...commentsRequests, ...subjectRequests];
}

/**
 * Adapts the size of the template sheet, and updates the named ranges
 */
function adaptSizeAndRanges(namedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    let rowOffset = 0;

    const subjectCount = persistentData.subjects.length;
    const fieldCount = persistentData.configData.averagePerField ? persistentData.academicFields.length : 0;

    const ranges = ReportSheetSchema.sheets.studentTemplate.ranges;

    const getRange = (name: RangeName): MappedNamedRange => {
        const range = namedRanges[name];
        if (!range) throw new Error(`Falta en template de estudiante: ${name}`);
        return range;
    };

    // Handle absences
    if (persistentData.configData.attendancePerClass) {
        const info = getRange(ranges.generalInfo);
        const height = (info.range.endRowIndex ?? 0) - (info.range.startRowIndex ?? 0) || 1;
        const { resizeRequests, newRange, newRowOffset } = updateSheetAndRange(info.range, rowOffset, height - 1);
        info.range = newRange;
        rowOffset = newRowOffset;
        requests.push(...resizeRequests);
    } else {
        requests.push(removeRangeColumns(getRange(ranges.trim1Absences).range));
        requests.push(removeRangeColumns(getRange(ranges.trim2Absences).range));
        requests.push(removeRangeColumns(getRange(ranges.trim3Absences).range));
    }

    const resizeOperations: Array<{ name: RangeName; count?: number }> = [
        { name: ranges.abilities, count: subjectCount },
        { name: ranges.comments, count: subjectCount },

        { name: ranges.trim1Subjects, count: subjectCount },
        { name: ranges.trim1Fields, count: fieldCount },
        { name: ranges.trim1Absences },
        { name: ranges.trim1Totals },

        { name: ranges.trim2Subjects, count: subjectCount },
        { name: ranges.trim2Fields, count: fieldCount },
        { name: ranges.trim2Absences },
        { name: ranges.trim2Totals },

        { name: ranges.trim3Subjects, count: subjectCount },
        { name: ranges.trim3Fields, count: fieldCount },
        { name: ranges.trim3Absences },
        { name: ranges.trim3Totals },
    ];

    for (const op of resizeOperations) {
        const mapped = getRange(op.name);

        if (op.count !== undefined) {
            const { resizeRequests, newRange, newRowOffset } = updateSheetAndRange(mapped.range, rowOffset, op.count);
            mapped.range = newRange;
            rowOffset = newRowOffset;
            requests.push(...resizeRequests);
        } else {
            mapped.range = offsetGridRange({ origin: mapped.range, rowOffset });
        }
    }

    // Fix merged cells
    const userCommentRange = offsetGridRange({ origin: getRange(ranges.comments).range, colOffset: 1, width: 3 });
    const simpleCommentRange = offsetGridRange({ origin: getRange(ranges.comments).range, colOffset: 4, width: 3 });

    const userCommentRangeMerge = buildMergeCellsRequest(userCommentRange, MergeType.MERGE_ROWS);
    if (userCommentRangeMerge) requests.push(userCommentRangeMerge);

    const simpleCommentRangeMerge = buildMergeCellsRequest(simpleCommentRange, MergeType.MERGE_ROWS);
    if (simpleCommentRangeMerge) requests.push(simpleCommentRangeMerge);

    // Add ranges for unprotected parts of the sheet
    const unprotectedRangeOperations: Array<{ origin: RangeName; width: number; name: RangeName }> = [
        { origin: ranges.abilities, width: 4, name: ranges.unprotectedAbilities },
        { origin: ranges.comments, width: 3, name: ranges.unprotectedComments },
        { origin: ranges.trim1Subjects, width: 2, name: ranges.unprotectedTrim1 },
        { origin: ranges.trim2Subjects, width: 2, name: ranges.unprotectedTrim2 },
        { origin: ranges.trim3Subjects, width: 2, name: ranges.unprotectedTrim3 },
    ];

    for (const op of unprotectedRangeOperations) {
        const origin = getRange(op.origin);
        const newRange = offsetGridRange({ origin: origin.range, colOffset: 1, width: op.width });
        const request = buildAddNamedRangeRequest<typeof ReportSheetSchema>(op.name, newRange);
        if (request) requests.push(request);
        namedRanges[op.name] = {
            range: newRange,
            sheet: origin.sheet,
        };
    }

    return requests;
}

/**
 * Helper function to delete a local range's columns, not affecting the rest of the sheet.
 */
function removeRangeColumns(range: GoogleAppsScript.Sheets.Schema.GridRange): GoogleAppsScript.Sheets.Schema.Request {
    return {
        deleteRange: {
            range,
            shiftDimension: Dimension.COLUMNS,
        },
    };
}

/**
 * Helper function to update the rows of a range both in the sheet and in the local range.
 */
function updateSheetAndRange(
    range: GoogleAppsScript.Sheets.Schema.GridRange,
    rowOffset: number,
    newRowCount: number,
): { resizeRequests: GoogleAppsScript.Sheets.Schema.Request[]; newRange: GoogleAppsScript.Sheets.Schema.GridRange; newRowOffset: number } {
    const resizeRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const shiftedRange = offsetGridRange({ origin: range, rowOffset });
    const startRow = shiftedRange.startRowIndex ?? 0;
    const endRow = shiftedRange.endRowIndex ?? 0;

    const currentRowCount = endRow - startRow;

    const diff = newRowCount - currentRowCount;

    if (diff > 0) {
        // Insert before last row so it automatically enlarges everything in the sheet.
        const insertIndex = endRow - 1;
        resizeRequests.push({
            insertDimension: {
                range: {
                    sheetId: shiftedRange.sheetId,
                    dimension: Dimension.ROWS,
                    startIndex: insertIndex,
                    endIndex: insertIndex + diff,
                },
                inheritFromBefore: true,
            },
        });
    } else if (diff < 0) {
        // Delete from the button of the range.
        const deleteCount = Math.abs(diff);
        resizeRequests.push({
            deleteDimension: {
                range: {
                    sheetId: shiftedRange.sheetId,
                    dimension: Dimension.ROWS,
                    startIndex: endRow - deleteCount,
                    endIndex: endRow,
                },
            },
        });
    }

    const newRange = offsetGridRange({ origin: shiftedRange, height: newRowCount });

    const newRowOffset = rowOffset + diff;

    return { resizeRequests, newRange, newRowOffset };
}

/**
 * Fills the general data, setting things for 1st trimester.
 */
function prepareInfo(namedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const dataRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.generalInfo]?.range;

    if (!dataRange) throw new Error("Falta rango para datos.");

    const period = 0;

    const dataData: GoogleAppsScript.Sheets.Schema.CellData[][] = [[], [], []]; // First three rows empty.

    dataData.push([{ userEnteredValue: { stringValue: TRIMESTER_NAMES[period] } }]);
    dataData.push([{ userEnteredValue: { stringValue: generatePeriodString(persistentData, period) } }]);

    if (!persistentData.configData.attendancePerClass) {
        const firstNameRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.firstName]?.range;
        const lastNameRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.lastName]?.range;

        if (!firstNameRange || !lastNameRange) throw new Error("Falta rango del nombre.");

        dataData.push([
            {
                userEnteredValue: {
                    formulaValue: createStudentAsistanceFormula(period, firstNameRange, lastNameRange, ReportSheetSchema.sheets.attendance.sheetName, true),
                },
            },
        ]);
    }

    return buildTransferRequests({
        destination: namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.generalInfo]?.range,
        data: dataData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
    });
}

/**
 * Fills the subjects for abilities.
 */
function prepareAbilities(namedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const abilitiesData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    const abilitiesRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.abilities]?.range;

    if (!abilitiesRange) throw new Error("Falta rango para habilidades.");

    for (const weightedSubject of persistenData.subjects) {
        abilitiesData.push([{ userEnteredValue: { stringValue: weightedSubject.subject } }]);
    }

    return buildTransferRequests({
        destination: abilitiesRange,
        data: abilitiesData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue"),
    });
}

/**
 * Prepares the comments, adding subjects and the formulas.
 */
function prepareComments(namedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const commentData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    const commentsRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.comments]?.range;
    if (!commentsRange) throw new Error("Falta rango para observaciones.");

    const firstRow = commentsRange.startRowIndex ?? 0;
    const firstCol = commentsRange.startColumnIndex ?? 0;

    for (const [index, weightedSubject] of persistenData.subjects.entries()) {
        commentData.push([
            { userEnteredValue: { stringValue: weightedSubject.subject } },
            { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
            {},
            {},
            { userEnteredValue: { formulaValue: getShortCommentFormula(firstRow + index, firstCol + 1) } },
        ]);
    }

    return buildTransferRequests({
        destination: commentsRange,
        data: commentData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
    });
}

/**
 * Helper function to build the formula for shorter comments.
 */
function getShortCommentFormula(sourceRow: number, sourceColumn: number): string {
    const sourceColumnLetters = getColumnLetter(sourceColumn);
    return `=LET(
  raw_text, ${sourceColumnLetters}${sourceRow + 1},
  lower_text, LOWER(raw_text),
  accent_map, {"á","a"; "é","e"; "í","i"; "ó","o"; "ú","u"; "ñ","n"; "ü","u"},
  cleaned_accents, REDUCE(lower_text, SEQUENCE(ROWS(accent_map)), LAMBDA(acc, i, SUBSTITUTE(acc, INDEX(accent_map, i, 1), INDEX(accent_map, i, 2)))),
  truncated, LEFT(cleaned_accents, MIN(250, IFERROR(FIND(CHAR(10), cleaned_accents) - 1, 250))),
  upper_text, UPPER(truncated),
  final_ascii, REGEXREPLACE(upper_text, "[^ -~]", ""),
  final_ascii
)`;
}

/**
 * Prepares the first trimester
 */
function prepareSubjects(namedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const subjectRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.trim1Subjects];
    const firstNameRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.firstName]?.range;
    const lastNameRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.lastName]?.range;
    if (!subjectRange || !firstNameRange || !lastNameRange) throw new Error("Faltan rangos de materias.");

    const period = 0;

    const subjectData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    for (const [index, weightedSubject] of persistenData.subjects.entries()) {
        const subjectRow: GoogleAppsScript.Sheets.Schema.CellData[] = [];
        subjectRow.push({ userEnteredValue: { stringValue: weightedSubject.subject } }, {}, {});
        if (persistenData.configData.attendancePerClass) {
            const subjectCell = offsetGridRange({ origin: subjectRange.range, rowOffset: index, height: 1, width: 1 });
            subjectRow.push(
                {
                    userEnteredValue: {
                        formulaValue: createStudentAsistancePerSubjectFormula(
                            period,
                            subjectCell,
                            firstNameRange,
                            lastNameRange,
                            ReportSheetSchema.sheets.attendance.sheetName,
                        ),
                    },
                },
                {},
            );
        } else {
            subjectRow.push({
                userEnteredValue: { formulaValue: createStudentAsistanceFormula(period, firstNameRange, lastNameRange, ReportSheetSchema.sheets.attendance.sheetName) },
            });
        }
        const valuesRange = offsetGridRange({ origin: subjectRange.range, rowOffset: index, colOffset: 1, height: 1, width: 3 });
        subjectRow.push({ userEnteredValue: { formulaValue: createSubjectAverageFormula(valuesRange) } });
        subjectData.push(subjectRow);
    }

    return buildTransferRequests({
        destination: subjectRange.range,
        data: subjectData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
    });
}

function createSubjectAverageFormula(valuesRange: GoogleAppsScript.Sheets.Schema.GridRange) {
    const row = valuesRange.startRowIndex ?? 0;
    const startCol = valuesRange.startColumnIndex ?? 0;

    const val1 = `${getColumnLetter(startCol)}${row + 1}`;
    const val2 = `${getColumnLetter(startCol + 1)}${row + 1}`;
    const val3 = `${getColumnLetter(startCol + 2)}${row + 1}`;

    return `=IF(COUNT(${val1}, ${val2}, ${val3})=3, ROUND(0.7 * ${val1} + 0.2 * ${val2} + ${val3}), "")`;
}
