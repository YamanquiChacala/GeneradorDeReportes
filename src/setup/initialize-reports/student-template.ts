import { TRIMESTER_NAMES } from "../../common/constants";
import { Dimension } from "../../common/gas-enums";
import { ReportSheetSchema } from "../../common/sheet-schema";
import { buildFieldsMask } from "../../common/utils/gas-types";
import { buildTransferRequest, offsetGridRange } from "../../common/utils/gas-utils";
import type { ExtractRangeNames, MappedNamedRange, ParsedSpreadsheet } from "../../common/utils/mapped-name-range";
import { generatePeriodString, type ReportPersistentData } from "../../common/utils/report-utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Prepares the student template sheet, filling in subjects, formulas, etc, so it can be duplicated for each student.
 */
export function prepareStudentTemplate(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // Adapt the sheet's ranges

    const adaptSheetRangesRequests = adaptSizeAndRanges(parsedReport.namedRanges, persistentData);

    // Fill in data

    const dataRequests = fillData(parsedReport.namedRanges, persistentData);

    // Build requests

    requests.push(...adaptSheetRangesRequests, ...dataRequests);

    return requests;
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
        const generalAbsences = getRange(ranges.generalAbsences);
        const { resizeRequests, newRange, newRowOffset } = updateSheetAndRange(generalAbsences.range, rowOffset, 0);
        generalAbsences.range = newRange;
        rowOffset = newRowOffset;
        requests.push(...resizeRequests);
    } else {
        requests.push(removeRangeColumns(getRange(ranges.trim1Absences).range));
        requests.push(removeRangeColumns(getRange(ranges.trim2Absences).range));
        requests.push(removeRangeColumns(getRange(ranges.trim3Absences).range));
    }

    const operations: Array<{ name: RangeName; count?: number }> = [
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

    for (const op of operations) {
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
 * Fills the subjects, fields and formulas.
 */
function fillData(namedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const period = 0;

    const dataData: GoogleAppsScript.Sheets.Schema.CellData[][] = [[], [], []]; // First three rows empty.

    dataData.push([{ userEnteredValue: { stringValue: TRIMESTER_NAMES[period] } }]);
    dataData.push([{ userEnteredValue: { stringValue: generatePeriodString(persistentData, period) } }]);

    return buildTransferRequest({
        destination: namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.generalInfo]?.range,
        data: dataData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue"),
    });
}
