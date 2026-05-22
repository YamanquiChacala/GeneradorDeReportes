import { MS_PER_DAY } from "../../common/constants";
import { ReportSheetSchema, SetupSheetSchema } from "../../common/gas-parts";
import type { ExtractRangeNames } from "../../common/gas-utils";
import {
    buildFieldsMask,
    buildTransferRequests,
    colorToHex,
    createRequiredGetter,
    getCellBoolean,
    getCellDataArray,
    getCellNumber,
    getCellUnixEpoch,
    getSheetsDate,
    type MappedNamedRange,
    makeUserEntered,
} from "../../common/gas-utils";
import type { AcademicField, ConfigData, ReportPersistentData, Student, StudentRow, WeightedSubject } from "../../common/report-utils";
import { normalizeSubjectWeights, normalizeTrimesterWeights, parseAcademicFieldsAndSubjects } from "../../common/setup-utils";
import { sanitizeSheetName, typedEntries } from "../../common/utils";

type SetupRangeName = ExtractRangeNames<typeof SetupSheetSchema>;
type ReportRangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Dumps the setup data into Persistent data in the report.
 */
export function fillPersistentData(
    setupMappedRanges: Partial<Record<SetupRangeName, MappedNamedRange>>,
    reportMappedRanges: Partial<Record<ReportRangeName, MappedNamedRange>>,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    persistentData: ReportPersistentData;
} {
    // How much we've pushed or pulled rows as we move from top to buttom.
    let rowOffset = 0;

    // Update general configuration data
    const { requests: configRequests, configData } = getConfigData(setupMappedRanges, reportMappedRanges);

    // Update Academic Fields and subjects
    const {
        requests: subjectRequests,
        newRowOffset: subjectsOffset,
        academicFields,
        subjects,
    } = getSubjects(setupMappedRanges, reportMappedRanges, configData.averagePerField, rowOffset);
    rowOffset = subjectsOffset;

    // Update Student list
    const { requests: studentRequests, newRowOffset: studentOffset, students } = getStudents(setupMappedRanges, reportMappedRanges, rowOffset);
    rowOffset = studentOffset;

    // Update calendar days
    const { requests: calendarDaysRequests, calendar } = getCalendarDays(setupMappedRanges, reportMappedRanges, rowOffset);

    // Build response
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [...configRequests, ...subjectRequests, ...studentRequests, ...calendarDaysRequests];

    // Build memory version of the sheet, for use without calling get again.
    const persistentData: ReportPersistentData = {
        protectedSections: {
            habilities: false,
            comments: false,
            trimesters: [false, true, true],
        },
        configData,
        calendar,
        students,
        academicFields,
        subjects,
    };

    return { persistentData, requests };
}

/**
 * Gets the basic configuration data from the setup file.
 */
function getConfigData(
    setupRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>,
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    configData: ConfigData;
} {
    type CellMapDefinition = {
        [K in keyof ConfigData]: {
            kind: "boolean" | "dateArray" | "numberArray";
            source: SetupRangeName;
            dest: ReportRangeName;
        };
    };

    const getSetupRange = createRequiredGetter(setupRanges);
    const getReportRange = createRequiredGetter(reportRanges);

    const cellMapping: CellMapDefinition = {
        attendancePerClass: {
            kind: "boolean",
            source: SetupSheetSchema.sheets.groupData.ranges.attendancePerClass,
            dest: ReportSheetSchema.sheets.persistentData.ranges.attendancePerClass,
        },
        averagePerField: {
            kind: "boolean",
            source: SetupSheetSchema.sheets.groupData.ranges.averagePerField,
            dest: ReportSheetSchema.sheets.persistentData.ranges.averagePerField,
        },
        dates: {
            kind: "dateArray",
            source: SetupSheetSchema.sheets.groupData.ranges.dates,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dates,
        },
        subjectGradingWeights: {
            kind: "numberArray",
            source: SetupSheetSchema.sheets.groupData.ranges.weights,
            dest: ReportSheetSchema.sheets.persistentData.ranges.subjectGradingWeights,
        },
    };

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const configData: Partial<ConfigData> = {};

    for (const [name, { kind, source, dest }] of typedEntries(cellMapping)) {
        const sourceMappedRange = getSetupRange(source);
        let data: GoogleAppsScript.Sheets.Schema.CellData[][];
        switch (kind) {
            case "boolean":
                Object.assign(configData, { [name]: getCellBoolean({ mappedRange: sourceMappedRange }) });
                data = makeUserEntered(getCellDataArray(sourceMappedRange), true);
                break;
            case "numberArray": {
                const num0 = getCellNumber({ mappedRange: sourceMappedRange, rowOffset: 0 });
                const num1 = getCellNumber({ mappedRange: sourceMappedRange, rowOffset: 1 });
                const num2 = getCellNumber({ mappedRange: sourceMappedRange, rowOffset: 2 });

                const [norm0, norm1, norm2] = normalizeTrimesterWeights(num0, num1, num2);

                Object.assign(configData, {
                    [name]: [norm0, norm1, norm2],
                });
                data = [[{ userEnteredValue: { numberValue: norm0 } }], [{ userEnteredValue: { numberValue: norm1 } }], [{ userEnteredValue: { numberValue: norm2 } }]];
                break;
            }
            case "dateArray":
                Object.assign(configData, {
                    [name]: [
                        getCellUnixEpoch({ mappedRange: sourceMappedRange, rowOffset: 0 }),
                        getCellUnixEpoch({ mappedRange: sourceMappedRange, rowOffset: 1 }),
                        getCellUnixEpoch({ mappedRange: sourceMappedRange, rowOffset: 2 }),
                        getCellUnixEpoch({ mappedRange: sourceMappedRange, rowOffset: 3 }),
                    ],
                });
                data = makeUserEntered(getCellDataArray(sourceMappedRange), true);
                break;
        }
        const { requests } = buildTransferRequests({
            destination: getReportRange(dest),
            data,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        });
        apiRequests.push(...requests);
    }
    return { configData: configData as ConfigData, requests: apiRequests };
}

/**
 * Gets the academic fields and subjects.
 */
function getSubjects(
    setupMappedRanges: Partial<Record<SetupRangeName, MappedNamedRange>>,
    reportMappedRanges: Partial<Record<ReportRangeName, MappedNamedRange>>,
    averagePerField: boolean,
    rowOffset: number,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    academicFields: AcademicField[];
    subjects: WeightedSubject[];
    newRowOffset: number;
} {
    const getSetupMappedRange = createRequiredGetter(setupMappedRanges, "rango en registro inicial");

    const setupSubjectData = getCellDataArray(getSetupMappedRange(SetupSheetSchema.sheets.groupData.ranges.subjects));

    const { academicFields, rawSubjects } = parseAcademicFieldsAndSubjects(setupSubjectData);

    const subjects = normalizeSubjectWeights(rawSubjects, academicFields, averagePerField);

    const { requests, newRowOffset } = buildReportFieldsAndSubjects(reportMappedRanges, academicFields, rawSubjects, rowOffset);

    return { requests, academicFields, subjects, newRowOffset };
}

/**
 * Creates the requests to fields and subjects into the persistent data.
 */
function buildReportFieldsAndSubjects(
    mappedRanges: Partial<Record<ReportRangeName, MappedNamedRange>>,
    academicFields: AcademicField[],
    subjects: WeightedSubject[],
    rowOffset: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newRowOffset: number } {
    const rangeNames = ReportSheetSchema.sheets.persistentData.ranges;
    const getNamedRange = createRequiredGetter(mappedRanges);

    const fieldsData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];
    const subjectsData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    for (const field of academicFields) {
        const fieldsDataRow: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { stringValue: field.color } },
            { userEnteredValue: { stringValue: field.name } },
            { userEnteredValue: { numberValue: field.subjects } },
        ];
        fieldsData.push(fieldsDataRow);
    }

    for (const weightedSubject of subjects) {
        const weightedSubjectDataRow: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { stringValue: weightedSubject.subject } },
            { userEnteredValue: { numberValue: weightedSubject.weight } },
        ];
        subjectsData.push(weightedSubjectDataRow);
    }

    const { requests: fieldRequests, rowOffset: fieldsRowOffset } = buildTransferRequests({
        destination: getNamedRange(rangeNames.fields),
        data: fieldsData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        adaptRange: true,
        rowOffset,
    });

    const { requests: subjectRequests, rowOffset: subjectRowOffset } = buildTransferRequests({
        destination: getNamedRange(rangeNames.subjects),
        data: subjectsData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        adaptRange: true,
        rowOffset: fieldsRowOffset,
    });

    return { requests: [...fieldRequests, ...subjectRequests], newRowOffset: subjectRowOffset };
}

/**
 * Gets the students with all their data.
 */
function getStudents(
    setupNamedRanges: Partial<Record<SetupRangeName, MappedNamedRange>>,
    reportNamedRanges: Partial<Record<ReportRangeName, MappedNamedRange>>,
    rowOffset: number,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    students: StudentRow[];
    newRowOffset: number;
} {
    const students: StudentRow[] = [];

    const getSetupNamedRange = createRequiredGetter(setupNamedRanges, "rango de registro inicial");
    const getReportNamedRange = createRequiredGetter(reportNamedRanges, "rango de reporte");

    // Unbound rows to get all students, even if the sheet grew past the range.
    const studentSetupData = getCellDataArray(getSetupNamedRange(SetupSheetSchema.sheets.groupData.ranges.students), true);

    const studentReportData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    let studentNumber = 0;
    let firstTime = true;
    let emptyBefore = false;
    for (const studentRow of studentSetupData) {
        const firstName = studentRow[0]?.effectiveValue?.stringValue;
        const lastName = studentRow[1]?.effectiveValue?.stringValue;
        if (!firstName || !lastName) {
            emptyBefore = true;
            studentNumber = 0;
            continue;
        }

        studentNumber++;
        const sheetName = sanitizeSheetName(`${firstName} ${lastName}`);

        const student: Student = {
            type: "student",
            id: studentNumber,
            firstName,
            lastName,
            sheetName,
            // TODO: Maybe mark these as optional.
            sex: studentRow[2]?.effectiveValue?.stringValue ?? "",
            level:
                studentRow[3]?.effectiveValue?.stringValue ??
                (studentRow[3]?.effectiveValue?.numberValue != null ? `${studentRow[3]?.effectiveValue?.numberValue}º` : ""),
            grade: studentRow[4]?.effectiveValue?.stringValue ?? "",
        };

        const studentReportDataRow: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { numberValue: studentNumber } },
            { userEnteredValue: { stringValue: firstName } },
            { userEnteredValue: { stringValue: lastName } },
            { userEnteredValue: { stringValue: sheetName } },
        ];
        if (emptyBefore && !firstTime) {
            students.push({ type: "separator" });
            studentReportData.push([]);
            emptyBefore = false;
        }

        students.push(student);
        studentReportData.push(studentReportDataRow);
        if (firstTime) {
            firstTime = false;
            emptyBefore = false;
        }
    }

    const { requests, rowOffset: studentRowOffset } = buildTransferRequests({
        destination: getReportNamedRange(ReportSheetSchema.sheets.persistentData.ranges.students),
        data: studentReportData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        adaptRange: true,
        rowOffset,
    });

    return { requests, students, newRowOffset: studentRowOffset };
}

/**
 * Gets the days of the calendar.
 */
function getCalendarDays(
    setupNamedRanges: Partial<Record<SetupRangeName, MappedNamedRange>>,
    reportNamedRanges: Partial<Record<ReportRangeName, MappedNamedRange>>,
    rowOffset: number,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    calendar: number[];
    newRowOffset: number;
} {
    const getSetupNamedRange = createRequiredGetter(setupNamedRanges, "rango de registro inicial");
    const getReportNamedRange = createRequiredGetter(reportNamedRanges, "rango de reporte");

    const calendar: number[] = [];

    const initialMillisecond = getCellNumber({ mappedRange: getSetupNamedRange(SetupSheetSchema.sheets.calendar.ranges.start) });
    const calendarRawData = getCellDataArray(getSetupNamedRange(SetupSheetSchema.sheets.calendar.ranges.calendar));

    if (new Date(initialMillisecond).getUTCDay() !== 0) throw new Error("Calendario no inicia en Domingo.");

    const sheetDays: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    for (const [rowIndex, row] of calendarRawData.entries()) {
        if (!row) continue;

        for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
            const columnIndex = dayIndex * 2 + 1;
            const cell = row[columnIndex];

            const isChecked = cell?.effectiveValue?.boolValue === true;

            if (isChecked) {
                const day = initialMillisecond + (rowIndex * 7 + dayIndex) * MS_PER_DAY;
                calendar.push(day);
                sheetDays.push([
                    {
                        userEnteredValue: { numberValue: getSheetsDate(day) },
                    },
                ]);
            }
        }
    }

    const { requests, rowOffset: calendaRowOffset } = buildTransferRequests({
        destination: getReportNamedRange(ReportSheetSchema.sheets.persistentData.ranges.calendarDates),
        data: sheetDays,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.numberValue"),
        adaptRange: true,
        rowOffset,
    });

    return { requests, calendar, newRowOffset: calendaRowOffset };
}
