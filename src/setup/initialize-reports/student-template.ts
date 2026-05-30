import { DEFAULT_COMMENT, Dimension, MergeType, TRIMESTER_NAMES } from "../../common/constants";
import { ReportSheetSchema } from "../../common/gas-parts";
import type { ExtractRangeNames, ParsedSpreadsheet } from "../../common/gas-utils";
import {
    buildAddNamedRangeRequest,
    buildFieldsMask,
    buildMergeCellsRequest,
    buildTransferRequests,
    createRequiredGetter,
    getRangeHeight,
    getRangeWidth,
    insertNewNamedRangeToMemory,
    type MappedNamedRange,
    offsetGridRange,
    resizeMappedRange,
    shrinkRangeWidth,
} from "../../common/gas-utils";
import {
    createAllSubjectsAverageFormula,
    createFieldAverageFormula,
    createFieldFormula,
    createFinalSubjectAverageFormula,
    createIndividualSubjectAverageFormula,
    createStudentGeneralAttendanceFormula,
    createStudentPerSubjectAttendanceFormula,
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
    const adaptSheetRangesRequests = adaptSizeAndRanges(parsedReport, persistentData);

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

    // Prepare averages
    const averageRequests = prepareAverages(parsedReport.mappedRanges, persistentData);

    // Build requests
    return [...adaptSheetRangesRequests, ...infoRequests, ...abilitiesRequests, ...commentsRequests, ...subjectRequests, ...fieldRequests, ...averageRequests];
}

/**
 * Adapts the size of the template sheet, and updates the named ranges
 */
function adaptSizeAndRanges(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    let rowOffset = 0;

    const subjectCount = persistentData.subjects.length;
    const fieldCount = persistentData.configData.averagePerField ? persistentData.academicFields.length : 0;

    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "template de estudiante");

    // Handle absences
    if (persistentData.configData.attendancePerClass) {
        // Remove general absences.
        const info = getMappedRange(rangeNames.generalInfo);
        const height = getRangeHeight(info.namedRange.range) || 1;
        const { requests: infoRequests, rowOffset: infoRowOffset } = resizeMappedRange({ target: info, targetRows: height - 1, rowOffset });
        rowOffset = infoRowOffset;
        requests.push(...infoRequests);
    } else {
        // Remove individual class absences.
        requests.push(
            removeRangeColumns(
                getMappedRange(rangeNames.trim1Absences),
                getMappedRange(rangeNames.trim1Subjects),
                getMappedRange(rangeNames.trim1Fields),
                getMappedRange(rangeNames.trim1Totals),
            ),
        );
        requests.push(
            removeRangeColumns(
                getMappedRange(rangeNames.trim2Absences),
                getMappedRange(rangeNames.trim2Subjects),
                getMappedRange(rangeNames.trim2Fields),
                getMappedRange(rangeNames.trim2Totals),
            ),
        );
        requests.push(
            removeRangeColumns(
                getMappedRange(rangeNames.trim3Absences),
                getMappedRange(rangeNames.trim3Subjects),
                getMappedRange(rangeNames.trim3Fields),
                getMappedRange(rangeNames.trim3Totals),
            ),
        );
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
    const userCommentRange = offsetGridRange({ origin: getMappedRange(rangeNames.comments).namedRange.range, colOffset: 1, width: 3 });
    const simpleCommentRange = offsetGridRange({ origin: getMappedRange(rangeNames.comments).namedRange.range, colOffset: 4, width: 3 });

    requests.push(buildMergeCellsRequest(userCommentRange, MergeType.MERGE_ROWS));
    requests.push(buildMergeCellsRequest(simpleCommentRange, MergeType.MERGE_ROWS));

    // Add ranges for unprotected parts of the sheet
    const unprotectedRangeOperations: Array<{ origin: RangeName; width: number; name: RangeName }> = [
        { origin: rangeNames.abilities, width: 4, name: rangeNames.unprotectedAbilities },
        { origin: rangeNames.comments, width: 3, name: rangeNames.unprotectedComments },
        { origin: rangeNames.trim1Subjects, width: 2, name: rangeNames.unprotectedTrim1 },
        { origin: rangeNames.trim2Subjects, width: 2, name: rangeNames.unprotectedTrim2 },
        { origin: rangeNames.trim3Subjects, width: 2, name: rangeNames.unprotectedTrim3 },
    ];

    for (const op of unprotectedRangeOperations) {
        const rangeNameId = Utilities.getUuid();
        const origin = getMappedRange(op.origin);
        const newRange = offsetGridRange({ origin: origin.namedRange.range, colOffset: 1, width: op.width });
        requests.push(buildAddNamedRangeRequest<typeof ReportSheetSchema>(op.name, newRange, rangeNameId));

        insertNewNamedRangeToMemory({
            parsedData: parsedReport,
            sheetTitle: ReportSheetSchema.sheets.studentTemplate.sheetName,
            rangeNameId,
            rangeName: op.name,
            gridRange: newRange,
            staticRangeKey: op.name,
        });
    }

    return requests;
}

/**
 * Helper function to delete a local range's columns, not affecting the rest of the sheet.
 */
function removeRangeColumns(
    absences: MappedNamedRange,
    subjects: MappedNamedRange,
    fields: MappedNamedRange,
    averages: MappedNamedRange,
): GoogleAppsScript.Sheets.Schema.Request {
    const removedCols = getRangeWidth(absences.namedRange.range);

    const request: GoogleAppsScript.Sheets.Schema.Request = {
        deleteRange: {
            range: absences.namedRange.range,
            shiftDimension: Dimension.COLUMNS,
        },
    };

    absences.namedRange.range = shrinkRangeWidth(absences.namedRange.range, removedCols);
    subjects.namedRange.range = shrinkRangeWidth(subjects.namedRange.range, removedCols);
    fields.namedRange.range = shrinkRangeWidth(fields.namedRange.range, removedCols);
    averages.namedRange.range = shrinkRangeWidth(averages.namedRange.range, removedCols);

    return request;
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
                    formulaValue: createStudentGeneralAttendanceFormula(period, firstNameRange, lastNameRange, ReportSheetSchema.sheets.attendance.sheetName, true),
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
    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "template de estudiante");

    const firstNameMappedRange = getMappedRange(rangeNames.firstName);
    const lastNameMappedRange = getMappedRange(rangeNames.lastName);
    const gradingWeightsMappedRange = getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjectGradingWeights);

    const attendancePerClass = persistenData.configData.attendancePerClass;

    const subjects: [MappedNamedRange, MappedNamedRange, MappedNamedRange] = [
        getMappedRange(rangeNames.trim1Subjects),
        getMappedRange(rangeNames.trim2Subjects),
        getMappedRange(rangeNames.trim3Subjects),
    ];

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

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
                            formulaValue: createStudentPerSubjectAttendanceFormula(
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
                        formulaValue: createStudentGeneralAttendanceFormula(
                            period,
                            firstNameMappedRange,
                            lastNameMappedRange,
                            ReportSheetSchema.sheets.attendance.sheetName,
                        ),
                    },
                });
            }
            // Average for every subject
            subjectRow.push({
                userEnteredValue: { formulaValue: createIndividualSubjectAverageFormula(subjectRange, index, gradingWeightsMappedRange) },
            });

            // Average for the last trimester
            if (period === 2) {
                subjectRow.push({ userEnteredValue: { formulaValue: createFinalSubjectAverageFormula(...subjects, index) } });
            }

            subjectData.push(subjectRow);
        }

        const { requests } = buildTransferRequests({
            destination: subjectRange,
            data: subjectData,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
        });

        apiRequests.push(...requests);
    }

    return apiRequests;
}

/**
 * Prepares the fields list and formulas
 */
function prepareFields(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    if (!persistenData.configData.averagePerField) return [];

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const attendancePerClass = persistenData.configData.attendancePerClass;

    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "template de estudiante");

    const weightedSubjectsRange = getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjects);
    const gradingWeithgtsRange = getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjectGradingWeights);

    interface SubjectField {
        academidFieldRange: MappedNamedRange;
        subjectsRange: MappedNamedRange;
    }

    const periodOperations: [SubjectField, SubjectField, SubjectField] = [
        { academidFieldRange: getMappedRange(rangeNames.trim1Fields), subjectsRange: getMappedRange(rangeNames.trim1Subjects) },
        { academidFieldRange: getMappedRange(rangeNames.trim2Fields), subjectsRange: getMappedRange(rangeNames.trim2Subjects) },
        { academidFieldRange: getMappedRange(rangeNames.trim3Fields), subjectsRange: getMappedRange(rangeNames.trim3Subjects) },
    ];

    for (const [periodIndex, { academidFieldRange, subjectsRange }] of periodOperations.entries()) {
        const period = periodIndex as 0 | 1 | 2;
        let subjectOffset = 0;

        const fieldData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

        for (const [index, field] of persistenData.academicFields.entries()) {
            const fieldRow: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: field.name } }];
            const colOffsets = [1, 2, 3];
            for (const colOffset of colOffsets) {
                fieldRow.push({ userEnteredValue: { formulaValue: createFieldFormula(field, colOffset, subjectOffset, subjectsRange, weightedSubjectsRange) } });
            }
            if (attendancePerClass) {
                fieldRow.push({});
            }
            fieldRow.push({ userEnteredValue: { formulaValue: createIndividualSubjectAverageFormula(academidFieldRange, index, gradingWeithgtsRange) } });

            if (period === 2) {
                fieldRow.push({
                    userEnteredValue: {
                        formulaValue: createFinalSubjectAverageFormula(
                            periodOperations[0].academidFieldRange,
                            periodOperations[1].academidFieldRange,
                            periodOperations[2].academidFieldRange,
                            index,
                        ),
                    },
                });
            }

            subjectOffset += field.subjects;
            fieldData.push(fieldRow);
        }

        const { requests } = buildTransferRequests({
            destination: academidFieldRange,
            data: fieldData,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue", "userEnteredValue.formulaValue"),
        });
        apiRequests.push(...requests);
    }

    return apiRequests;
}

/**
 * Prepares the period and final averages
 */
function prepareAverages(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>, persistenData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "template de estudiante");

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    if (persistenData.configData.averagePerField) {
        const trimOperations: Array<{ grades: RangeName; dest: RangeName; lastBig: boolean }> = [
            { grades: rangeNames.trim1Fields, dest: rangeNames.trim1Totals, lastBig: false },
            { grades: rangeNames.trim2Fields, dest: rangeNames.trim2Totals, lastBig: false },
            { grades: rangeNames.trim3Fields, dest: rangeNames.trim3Totals, lastBig: true },
        ];

        for (const { grades, dest, lastBig } of trimOperations) {
            const rowData: GoogleAppsScript.Sheets.Schema.CellData[] = [];
            const gradesRange = getMappedRange(grades);
            const destRange = getMappedRange(dest);

            const totalColumns = getRangeWidth(destRange.namedRange.range);
            for (let colOffset = 0; colOffset < totalColumns; colOffset++) {
                rowData.push({
                    userEnteredValue: { formulaValue: createFieldAverageFormula(gradesRange, colOffset + 1, colOffset === totalColumns - 1 && lastBig ? 2 : 1) },
                });
            }

            const { requests } = buildTransferRequests({
                destination: destRange,
                data: [rowData],
                fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.formulaValue"),
            });

            apiRequests.push(...requests);
        }
    } else {
        const trimOperations: Array<{ grades: RangeName; dest: RangeName; lastBig: boolean }> = [
            { grades: rangeNames.trim1Subjects, dest: rangeNames.trim1Totals, lastBig: false },
            { grades: rangeNames.trim2Subjects, dest: rangeNames.trim2Totals, lastBig: false },
            { grades: rangeNames.trim3Subjects, dest: rangeNames.trim3Totals, lastBig: true },
        ];

        for (const { grades, dest, lastBig } of trimOperations) {
            const rowData: GoogleAppsScript.Sheets.Schema.CellData[] = [];
            const gradesRange = getMappedRange(grades);
            const destRange = getMappedRange(dest);
            const weightsRange = getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.subjects);

            const totalColumns = getRangeWidth(destRange.namedRange.range);
            for (let colOffset = 0; colOffset < totalColumns; colOffset++) {
                rowData.push({
                    userEnteredValue: {
                        formulaValue: createAllSubjectsAverageFormula(gradesRange, weightsRange, colOffset + 1, colOffset === totalColumns - 1 && lastBig ? 2 : 1),
                    },
                });
            }

            const { requests } = buildTransferRequests({
                destination: destRange,
                data: [rowData],
                fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.formulaValue"),
            });

            apiRequests.push(...requests);
        }
    }

    return apiRequests;
}
