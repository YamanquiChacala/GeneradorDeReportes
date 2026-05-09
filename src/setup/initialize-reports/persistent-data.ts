import { MS_PER_DAY } from "../../common/constants";
import { ReportSheetSchema, SetupSheetSchema } from "../../common/sheet-schema";
import { buildFieldsMask } from "../../common/utils/gas-types";
import { buildTransferRequest, colorToHex, getEpochDate, getSheetsDate } from "../../common/utils/gas-utils";
import { type ExtractRangeNames, MappedNamedRange } from "../../common/utils/mapped-name-range";
import type { AcademicField, ConfigData, ReportPersistentData, Student, StudentRow, WeightedSubject } from "../../common/utils/report-utils";
import { sanitizeSheetName } from "../../common/utils/text";

/**
 * Dumps the setup data into Persistent data in the report.
 */
export function fillPersistentData(
    setupRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>,
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    data: ReportPersistentData;
} {
    // Get general configuration data
    const { requests: configRequests, configData } = getConfigData(setupRanges, reportRanges);

    // Get calendar days
    const { requests: calendarDaysRequests, calendar } = getCalendarDays(setupRanges, reportRanges);

    // Get Students
    const { requests: studentRequests, students } = getStudents(setupRanges, reportRanges);

    // Get Academic Fields and subjects
    const { requests: subjectRequests, academicFields, subjects } = getSubjects(setupRanges, reportRanges, configData.averagePerField);

    // Build response
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [...configRequests, ...calendarDaysRequests, ...studentRequests, ...subjectRequests];

    const data: ReportPersistentData = {
        protectedSections: {
            data: true,
            habilities: false,
            comments: false,
            trim1: false,
            trim2: true,
            trim3: true,
        },
        configData,
        calendar,
        students,
        academicFields,
        subjects,
    };

    return { data: data, requests };
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
            kind: "boolean" | "number" | "date";
            source: ExtractRangeNames<typeof SetupSheetSchema>;
            dest: ExtractRangeNames<typeof ReportSheetSchema>;
        };
    };

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
        dateStart: {
            kind: "date",
            source: SetupSheetSchema.sheets.groupData.ranges.dateStart,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateStart,
        },
        dateTrim1: {
            kind: "date",
            source: SetupSheetSchema.sheets.groupData.ranges.dateTrim1,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateTrim1,
        },
        dateTrim2: {
            kind: "date",
            source: SetupSheetSchema.sheets.groupData.ranges.dateTrim2,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateTrim2,
        },
        dateEnd: {
            kind: "date",
            source: SetupSheetSchema.sheets.groupData.ranges.dateEnd,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateEnd,
        },
    };

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const configData: Partial<ConfigData> = {};

    const entries = Object.entries(cellMapping) as [keyof ConfigData, CellMapDefinition[keyof ConfigData]][];

    for (const [name, { kind, source, dest }] of entries) {
        const extendedValue = MappedNamedRange.getCellEffectiveValue({ mappedRange: setupRanges[source] });
        switch (kind) {
            case "boolean":
                Object.assign(configData, { [name]: extendedValue?.boolValue ?? false });
                break;
            case "number":
                Object.assign(configData, { [name]: extendedValue?.numberValue ?? 0 });
                break;
            case "date":
                Object.assign(configData, { [name]: getEpochDate(extendedValue?.numberValue ?? 0) });
        }
        requests.push({
            repeatCell: {
                cell: { userEnteredValue: extendedValue },
                range: reportRanges[dest]?.range,
                fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
            },
        });
    }
    return { configData: configData as ConfigData, requests };
}

/**
 * Gets the days of the calendar.
 */
function getCalendarDays(
    setupRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>,
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    calendar: number[];
} {
    const calendar: number[] = [];

    const initialMillisecond = MappedNamedRange.getCellNumber({ mappedRange: setupRanges[SetupSheetSchema.sheets.calendar.ranges.start] }) ?? 0;
    const calendarRawData = MappedNamedRange.getCellDataArray(setupRanges[SetupSheetSchema.sheets.calendar.ranges.calendar]);

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

    const requests = buildTransferRequest({
        destination: reportRanges[ReportSheetSchema.sheets.persistentData.ranges.calendarDates]?.range,
        data: sheetDays,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.numberValue"),
        adaptRange: true,
    });

    return { requests, calendar };
}

/**
 * Gets the students with all their data.
 */
function getStudents(
    setupRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>,
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    students: StudentRow[];
} {
    const students: StudentRow[] = [];

    const studentSetupData = MappedNamedRange.getCellDataArray(setupRanges[SetupSheetSchema.sheets.groupData.ranges.students], true);

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

    const requests = buildTransferRequest({
        destination: reportRanges[ReportSheetSchema.sheets.persistentData.ranges.students]?.range,
        data: studentReportData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        adaptRange: true,
    });

    return { requests, students };
}

/**
 * Gets the academic fields and subjects.
 */
function getSubjects(
    setupRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>,
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
    averagePerField: boolean,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    academicFields: AcademicField[];
    subjects: WeightedSubject[];
} {
    const academicFields: AcademicField[] = [];
    const subjects: WeightedSubject[] = [];

    const setupSubjectData = MappedNamedRange.getCellDataArray(setupRanges[SetupSheetSchema.sheets.groupData.ranges.subjects]);

    let currentField: AcademicField = {
        name: "",
        color: "",
        subjects: 0,
    };

    for (const dataRow of setupSubjectData) {
        const iterator = dataRow[Symbol.iterator]();
        const first = iterator.next();

        if (first.done) continue;

        const fieldCell = first.value;
        const fieldName = (fieldCell?.effectiveValue?.stringValue ?? "").trim();

        if (fieldName !== "") {
            // Save previous Field
            if (currentField.name !== "" && currentField.color !== "" && currentField.subjects > 0) academicFields.push(currentField);

            // Start a new one
            const bgColor = fieldCell?.effectiveFormat?.backgroundColor;
            currentField = {
                name: fieldName,
                color: colorToHex(bgColor),
                subjects: 0,
            };
        }
        if (currentField.name !== "") {
            let weight = 1;
            let subject = "";
            for (const cellData of iterator) {
                // Find weight
                const cellNumVal = cellData.effectiveValue?.numberValue;
                if (cellNumVal) weight = cellNumVal;
                // Find subject name
                const cellStringVal = (cellData.effectiveValue?.stringValue ?? "").trim();
                if (cellStringVal !== "") {
                    subject = cellStringVal;
                    break;
                }
            }
            if (subject !== "") {
                currentField.subjects++;
                subjects.push({ weight, subject });
            }
        }
    }
    // Save last Field
    if (currentField.name !== "" && currentField.color !== "" && currentField.subjects > 0) academicFields.push(currentField);

    // Normalize weights
    if (averagePerField) {
        // Normalize weights per field
        let subjectIndex = 0; // Keep track of where we are in the flat subjects array

        for (const field of academicFields) {
            const fieldSubjectCount = field.subjects;
            let totalFieldWeight = 0;

            // 1. Calculate the total weight for this specific block of subjects
            for (let i = 0; i < fieldSubjectCount; i++) {
                const currentSubject = subjects[subjectIndex + i];
                if (currentSubject) {
                    totalFieldWeight += currentSubject.weight;
                }
            }

            // 2. Normalize the weights for this block
            if (totalFieldWeight > 0) {
                for (let i = 0; i < fieldSubjectCount; i++) {
                    const currentSubject = subjects[subjectIndex + i];
                    if (currentSubject) {
                        currentSubject.weight /= totalFieldWeight;
                    }
                }
            }

            // 3. Move the index forward to the start of the next field's subjects
            subjectIndex += fieldSubjectCount;
        }
    } else {
        // Normalize weights globally across the entire flat array
        const grandTotalWeight = subjects.reduce((sum, subj) => sum + subj.weight, 0);

        if (grandTotalWeight > 0) {
            for (const subj of subjects) {
                subj.weight /= grandTotalWeight;
            }
        }
    }

    const requests = buildReportFieldsAndSubjects(reportRanges, academicFields, subjects);

    return { requests, academicFields, subjects };
}

/**
 * Creates the requests to fields and subjects into the persistent data.
 */
function buildReportFieldsAndSubjects(
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
    academicFields: AcademicField[],
    subjects: WeightedSubject[],
): GoogleAppsScript.Sheets.Schema.Request[] {
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

    const subjectsRequests = buildTransferRequest({
        destination: reportRanges[ReportSheetSchema.sheets.persistentData.ranges.subjects]?.range,
        data: subjectsData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        adaptRange: true,
    });

    const fieldsRequests = buildTransferRequest({
        destination: reportRanges[ReportSheetSchema.sheets.persistentData.ranges.fields]?.range,
        data: fieldsData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        adaptRange: true,
    });

    return [...subjectsRequests, ...fieldsRequests];
}
