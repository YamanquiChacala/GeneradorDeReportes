import { DEFAULT_COMMENT, Dimension, MergeType, TRIMESTER_NAMES } from "../../common/constants";
import { ReportSheetSchema } from "../../common/gas-parts";
import type { ExtractRangeNames, ParsedSpreadsheet } from "../../common/gas-utils";
import {
    buildAddNamedRangeRequest,
    buildFieldsMask,
    buildMergeCellsRequest,
    buildTransferRequests,
    buildUpdateCellsRequest,
    createRequiredGetter,
    getA1Notation,
    getColumnLetter,
    type MappedNamedRange,
    offsetGridRange,
    resizeMappedRange,
} from "../../common/gas-utils";
import {
    type AcademicField,
    createFinalSubjectAverageFormula,
    createStudentAsistanceFormula,
    createStudentAsistancePerSubjectFormula,
    createSubjectAverageFormula,
    generatePeriodString,
    getShortCommentFormula,
    type ReportPersistentData,
} from "../../common/report-utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Prepares the student template sheet, filling in subjects, formulas, etc, so it can be duplicated for each student.
 */
export function prepareStudentTemplate(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    // Adapt the sheet's ranges
    const adaptSheetRangesRequests = adaptSizeAndRanges(parsedReport.mappedRanges, persistentData);

    // Fill in data
    const infoRequests = prepareInfo(parsedReport.mappedRanges, persistentData);

    // Prepare abilities
    const abilitiesRequests = prepareAbilities(parsedReport.mappedRanges, persistentData);

    // Prepare comments
    const commentsRequests = prepareComments(parsedReport.mappedRanges, persistentData);

    // Prepare Trimester 1
    const subjectRequests = prepareSubjects(parsedReport.mappedRanges, persistentData);

    // Prepare Fields
    const fieldRequests = prepareFields(parsedReport.mappedRanges, persistentData);

    // Build requests
    return [...adaptSheetRangesRequests, ...infoRequests, ...abilitiesRequests, ...commentsRequests, ...subjectRequests, ...fieldRequests];
}

/**
 * Adapts the size of the template sheet, and updates the named ranges
 */
function adaptSizeAndRanges(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    let rowOffset = 0;

    const subjectCount = persistentData.subjects.length;
    const fieldCount = persistentData.configData.averagePerField ? persistentData.academicFields.length : 0;

    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "template de estudiante");

    // Handle absences
    if (persistentData.configData.attendancePerClass) {
        // Remove general absences.
        const info = getMappedRange(rangeNames.generalInfo);
        const height = (info.range.endRowIndex ?? 0) - (info.range.startRowIndex ?? 0) || 1;
        const { requests: infoRequests, rowOffset: infoRowOffset } = resizeMappedRange({ target: info, targetRows: height - 1, rowOffset });
        rowOffset = infoRowOffset;
        requests.push(...infoRequests);
    } else {
        // Remove individual class absences.
        requests.push(removeRangeColumns(getMappedRange(rangeNames.trim1Absences).range));
        requests.push(removeRangeColumns(getMappedRange(rangeNames.trim2Absences).range));
        requests.push(removeRangeColumns(getMappedRange(rangeNames.trim3Absences).range));
    }

    const resizeOperations: Array<{ name: RangeName; count?: number }> = [
        { name: rangeNames.abilities, count: subjectCount },
        { name: rangeNames.comments, count: subjectCount },

        { name: rangeNames.trim1Subjects, count: subjectCount },
        { name: rangeNames.trim1Fields, count: fieldCount },
        { name: rangeNames.trim1Absences },
        { name: rangeNames.trim1Totals },

        { name: rangeNames.trim2Subjects, count: subjectCount },
        { name: rangeNames.trim2Fields, count: fieldCount },
        { name: rangeNames.trim2Absences },
        { name: rangeNames.trim2Totals },

        { name: rangeNames.trim3Subjects, count: subjectCount },
        { name: rangeNames.trim3Fields, count: fieldCount },
        { name: rangeNames.trim3Absences },
        { name: rangeNames.trim3Totals },
    ];

    for (const op of resizeOperations) {
        const mapped = getMappedRange(op.name);

        const { requests: resizeRequests, rowOffset: newRowOffset } = resizeMappedRange({ target: mapped, rowOffset, targetRows: op.count });
        requests.push(...resizeRequests);
        rowOffset = newRowOffset;
    }

    // Fix merged cells
    const userCommentRange = offsetGridRange({ origin: getMappedRange(rangeNames.comments).range, colOffset: 1, width: 3 });
    const simpleCommentRange = offsetGridRange({ origin: getMappedRange(rangeNames.comments).range, colOffset: 4, width: 3 });

    const userCommentRangeMerge = buildMergeCellsRequest(userCommentRange, MergeType.MERGE_ROWS);
    if (userCommentRangeMerge) requests.push(userCommentRangeMerge);

    const simpleCommentRangeMerge = buildMergeCellsRequest(simpleCommentRange, MergeType.MERGE_ROWS);
    if (simpleCommentRangeMerge) requests.push(simpleCommentRangeMerge);

    // Add ranges for unprotected parts of the sheet
    const unprotectedRangeOperations: Array<{ origin: RangeName; width: number; name: RangeName }> = [
        { origin: rangeNames.abilities, width: 4, name: rangeNames.unprotectedAbilities },
        { origin: rangeNames.comments, width: 3, name: rangeNames.unprotectedComments },
        { origin: rangeNames.trim1Subjects, width: 2, name: rangeNames.unprotectedTrim1 },
        { origin: rangeNames.trim2Subjects, width: 2, name: rangeNames.unprotectedTrim2 },
        { origin: rangeNames.trim3Subjects, width: 2, name: rangeNames.unprotectedTrim3 },
    ];

    for (const op of unprotectedRangeOperations) {
        const origin = getMappedRange(op.origin);
        const newRange = offsetGridRange({ origin: origin.range, colOffset: 1, width: op.width });
        const request = buildAddNamedRangeRequest<typeof ReportSheetSchema>(op.name, newRange);
        if (request) requests.push(request);
        mappedRanges[op.name] = {
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
 * Fills the general data, setting things for 1st trimester.
 */
function prepareInfo(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango de formato de estudiante");

    const period = 0;

    const dataData: GoogleAppsScript.Sheets.Schema.CellData[][] = [[], [], []]; // First three rows empty.

    dataData.push([{ userEnteredValue: { stringValue: TRIMESTER_NAMES[period] } }]);
    dataData.push([{ userEnteredValue: { stringValue: generatePeriodString(persistentData, period) } }]);

    if (!persistentData.configData.attendancePerClass) {
        const firstNameRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.firstName);
        const lastNameRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.lastName);

        dataData.push([
            {
                userEnteredValue: {
                    formulaValue: createStudentAsistanceFormula(period, firstNameRange, lastNameRange, ReportSheetSchema.sheets.attendance.sheetName, true),
                },
            },
        ]);
    }

    const mappedInfo = getMappedRange(rangeNames.generalInfo);

    const { requests } = buildTransferRequests({
        destination: mappedInfo,
        data: dataData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
    });

    return requests;
}

/**
 * Fills the subjects for abilities.
 */
function prepareAbilities(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango de formato de estudiante");

    const abilitiesData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    const abilitiesRange = getMappedRange(rangeNames.abilities);

    for (const weightedSubject of persistenData.subjects) {
        abilitiesData.push([{ userEnteredValue: { stringValue: weightedSubject.subject } }]);
    }

    const { requests } = buildTransferRequests({
        destination: abilitiesRange,
        data: abilitiesData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue"),
    });

    return requests;
}

/**
 * Prepares the comments, adding subjects and the formulas.
 */
function prepareComments(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango de formato de estudiante");

    const commentData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    const commentsRange = getMappedRange(rangeNames.comments);

    for (const [index, weightedSubject] of persistenData.subjects.entries()) {
        commentData.push([
            { userEnteredValue: { stringValue: weightedSubject.subject } },
            { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
            {},
            {},
            { userEnteredValue: { formulaValue: getShortCommentFormula(commentsRange, index, 1) } },
        ]);
    }

    const { requests } = buildTransferRequests({
        destination: commentsRange,
        data: commentData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
    });

    return requests;
}

/**
 * Prepares the subject lists and formulas
 */
function prepareSubjects(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "template de estudiante");

    const firstNameMappedRange = getMappedRange(rangeNames.firstName);
    const lastNameMappedRange = getMappedRange(rangeNames.lastName);
    const gradingWeightsMappedRange = getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjectGradingWeights);

    const persistentDataSheetName = ReportSheetSchema.sheets.persistentData.sheetName;

    const attendancePerClass = persistenData.configData.attendancePerClass;

    const subjects: MappedNamedRange[] = [getMappedRange(rangeNames.trim1Subjects), getMappedRange(rangeNames.trim2Subjects), getMappedRange(rangeNames.trim3Subjects)];

    for (const [periodIndex, subjectRange] of subjects.entries()) {
        const period = periodIndex as 0 | 1 | 2;
        const subjectData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

        for (const [index, weightedSubject] of persistenData.subjects.entries()) {
            const subjectRow: GoogleAppsScript.Sheets.Schema.CellData[] = [];
            subjectRow.push({ userEnteredValue: { stringValue: weightedSubject.subject } }, {}, {});
            if (attendancePerClass) {
                subjectRow.push(
                    {
                        userEnteredValue: {
                            formulaValue: createStudentAsistancePerSubjectFormula(
                                period,
                                subjectRange,
                                index,
                                firstNameMappedRange,
                                lastNameMappedRange,
                                ReportSheetSchema.sheets.attendance.sheetName,
                            ),
                        },
                    },
                    {},
                );
            } else {
                subjectRow.push({
                    userEnteredValue: {
                        formulaValue: createStudentAsistanceFormula(period, firstNameMappedRange, lastNameMappedRange, ReportSheetSchema.sheets.attendance.sheetName),
                    },
                });
            }
            // Average for every subject
            subjectRow.push({
                userEnteredValue: { formulaValue: createSubjectAverageFormula(subjectRange, index, gradingWeightsMappedRange) },
            });

            // Average for the last trimester
            if (period === 2) {
                subjectRow.push({ userEnteredValue: { formulaValue: createFinalSubjectAverageFormula(...subjects, index) } });
            }

            subjectData.push(subjectRow);
        }

        requests.push(
            ...buildUpdateCellsRequest({
                destination: subjectRange,
                data: subjectData,
                fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
            }),
        );
    }

    return requests;
}

/**
 * Prepares the fields list and formulas
 */
function prepareFields(namedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    if (!persistenData.configData.averagePerField) return [];

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const attendancePerClass = persistenData.configData.attendancePerClass;

    const ranges = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(namedRanges, "template de estudiante");

    const weightsSheetName = ReportSheetSchema.sheets.persistentData.sheetName;
    const weightsRange = offsetGridRange({ origin: getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjects).range, colOffset: 1, width: 1 });

    const gradingWeight1 = getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjectGrading1).range;
    const gradingWeight2 = getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjectGrading2).range;
    const gradingWeight3 = getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjectGrading3).range;

    const trim1SubbjectsRange = getMappedRange(ranges.trim1Subjects).range;
    const trim2SubbjectsRange = getMappedRange(ranges.trim2Subjects).range;
    const trim3SubbjectsRange = getMappedRange(ranges.trim3Subjects).range;

    const trim1FieldsRange = getMappedRange(ranges.trim1Fields).range;
    const trim2FieldsRange = getMappedRange(ranges.trim2Fields).range;
    const trim3FieldsRange = getMappedRange(ranges.trim3Fields).range;

    const fieldGroups: Array<{ fieldsRange: GoogleAppsScript.Sheets.Schema.GridRange; subjectsRange: GoogleAppsScript.Sheets.Schema.GridRange }> = [
        { fieldsRange: trim1FieldsRange, subjectsRange: trim1SubbjectsRange },
        { fieldsRange: trim2FieldsRange, subjectsRange: trim2SubbjectsRange },
        { fieldsRange: trim3FieldsRange, subjectsRange: trim3SubbjectsRange },
    ];

    for (const [periodIndex, { fieldsRange, subjectsRange }] of fieldGroups.entries()) {
        const period = periodIndex as 0 | 1 | 2;

        const fieldData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

        let subjectOffset = 0;
        for (const [index, field] of persistenData.academicFields.entries()) {
            const fieldRow: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: field.name } }];
            for (let i = 1; i < 4; i++) {
                fieldRow.push({
                    userEnteredValue: {
                        formulaValue: createFieldFunction(
                            field,
                            offsetGridRange({ origin: subjectsRange, colOffset: i, width: 1 }),
                            weightsSheetName,
                            weightsRange,
                            subjectOffset,
                        ),
                    },
                });
            }

            if (attendancePerClass) {
                fieldRow.push({});
            }
            // Average
            const valuesRange = offsetGridRange({ origin: fieldsRange, rowOffset: index, colOffset: 1, height: 1, width: 3 });
            fieldRow.push({
                userEnteredValue: { formulaValue: createSubjectAverageFormula(valuesRange, weightsSheetName, gradingWeight1, gradingWeight2, gradingWeight3) },
            });

            // Final average
            if (period === 2) {
                const trim1Prom = offsetGridRange({ origin: trim1FieldsRange, rowOffset: index, colOffset: attendancePerClass ? 5 : 4, height: 1, width: 1 });
                const trim2Prom = offsetGridRange({ origin: trim2FieldsRange, rowOffset: index, colOffset: attendancePerClass ? 5 : 4, height: 1, width: 1 });
                const trim3Prom = offsetGridRange({ origin: trim3FieldsRange, rowOffset: index, colOffset: attendancePerClass ? 5 : 4, height: 1, width: 1 });

                fieldRow.push({ userEnteredValue: { formulaValue: createFinalSubjectAverageFormula(trim1Prom, trim2Prom, trim3Prom) } });
            }

            subjectOffset += field.subjects;
            fieldData.push(fieldRow);
        }

        requests.push(
            ...buildUpdateCellsRequest({
                destination: fieldsRange,
                data: fieldData,
                fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
            }),
        );
    }

    return requests;
}

/**
 * Helper function to calculate the field value from the subjects.
 */
function createFieldFunction(
    field: AcademicField,
    subjectsRange: GoogleAppsScript.Sheets.Schema.GridRange,
    weightsSheetName: string,
    weightsRange: GoogleAppsScript.Sheets.Schema.GridRange,
    subjectOffset: number,
): string {
    const valuesRange = offsetGridRange({ origin: subjectsRange, rowOffset: subjectOffset, height: field.subjects });
    const valuesA1 = getA1Notation(valuesRange, false, true, false);

    const weights = offsetGridRange({ origin: weightsRange, rowOffset: subjectOffset, height: field.subjects });
    const weightsA1 = getA1Notation(weights, false, true, true);

    return `=IFERROR(
    AVERAGE.WEIGHTED(${valuesA1}, '${weightsSheetName}'!${weightsA1})
)`;
}
