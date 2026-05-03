import { MS_PER_DAY } from "../../common/constants";
import { ReportSheetSchema, SetupSheetSchema } from "../../common/sheet-schema";
import { buildFieldsMask } from "../../common/utils/gas-types";
import { buildTransferRequest, colorToHex, getSheetsDate } from "../../common/utils/gas-utils";
import { type ExtractRangeNames, MappedNamedRange } from "../../common/utils/mapped-name-range";
import { sanitizeSheetName } from "../../common/utils/text";

export interface ReportPersistentData {
    configData: ConfigData;
    protectedSections: ProtectedSections;
    academicFields: AcademicField[];
    students: Student[];
    calendar: number[];
}

interface ConfigData {
    attendancePerClass: boolean;
    averagePerField: boolean;
    dateStart: number;
    dateTrim1: number;
    dateTrim2: number;
    dateEnd: number;
}

interface ProtectedSections {
    data: boolean;
    habilities: boolean;
    comments: boolean;
    trim1: boolean;
    trim2: boolean;
    trim3: boolean;
}

interface AcademicField {
    name: string;
    color: string;
    subjects: string[];
}

interface Student {
    id: number;
    firstName: string;
    lastName: string;
    sheetName: string;
    sex: string;
    level: string;
    grade: string;
}

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
    const data: Partial<ReportPersistentData> = {
        protectedSections: {
            data: true,
            habilities: false,
            comments: false,
            trim1: false,
            trim2: true,
            trim3: true,
        },
    };

    // Get general configuration data
    const { requests: configRequests, configData } = getConfigData(setupRanges, reportRanges);

    // Get calendar days
    const { requests: calendarDaysRequests, calendarDays } = getCalendarDays(setupRanges, reportRanges);

    // Get Students
    const { requests: studentRequests, students } = getStudents(setupRanges, reportRanges);

    // Get Academic Fields and subjects
    const { requests: subjectRequests, academicFields } = getSubjects(setupRanges, reportRanges);

    // Build response
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [...configRequests, ...calendarDaysRequests, ...studentRequests, ...subjectRequests];

    data.configData = configData;
    data.calendar = calendarDays;
    data.students = students;
    data.academicFields = academicFields;

    return { data: data as ReportPersistentData, requests };
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
            kind: "boolean" | "number";
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
            kind: "number",
            source: SetupSheetSchema.sheets.groupData.ranges.dateStart,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateStart,
        },
        dateTrim1: {
            kind: "number",
            source: SetupSheetSchema.sheets.groupData.ranges.dateTrim1,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateTrim1,
        },
        dateTrim2: {
            kind: "number",
            source: SetupSheetSchema.sheets.groupData.ranges.dateTrim2,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateTrim2,
        },
        dateEnd: {
            kind: "number",
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
                // @ts-expect-error
                configData[name] = extendedValue?.boolValue ?? false;
                break;
            case "number":
                // @ts-expect-error
                configData[name] = extendedValue?.numberValue ?? 0;
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
    calendarDays: number[];
} {
    const calendarDays: number[] = [];

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
                calendarDays.push(day);
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

    return { requests, calendarDays };
}

/**
 * Gets the students with all their data.
 */
function getStudents(
    setupRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>,
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    students: Student[];
} {
    const students: Student[] = [];

    const studentSetupData = MappedNamedRange.getCellDataArray(setupRanges[SetupSheetSchema.sheets.groupData.ranges.students], true);

    console.log("rows: ", studentSetupData.length);

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
            id: studentNumber,
            firstName,
            lastName,
            sheetName,
            sex: studentRow[2]?.effectiveValue?.stringValue ?? "",
            level: studentRow[3]?.effectiveValue?.stringValue ?? "",
            grade: studentRow[4]?.effectiveValue?.stringValue ?? "",
        };
        students.push(student);

        const studentReportDataRow: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { numberValue: studentNumber } },
            { userEnteredValue: { stringValue: firstName } },
            { userEnteredValue: { stringValue: lastName } },
            { userEnteredValue: { stringValue: sheetName } },
        ];
        if (emptyBefore && !firstTime) {
            studentReportData.push([]);
            emptyBefore = false;
        }
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
): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    academicFields: AcademicField[];
} {
    const academicFields: AcademicField[] = [];

    const setupSubjectData = MappedNamedRange.getCellDataArray(setupRanges[SetupSheetSchema.sheets.groupData.ranges.subjects]);

    let currentField: AcademicField = {
        name: "",
        color: "",
        subjects: [],
    };

    for (const dataRow of setupSubjectData) {
        const iterator = dataRow[Symbol.iterator]();
        const first = iterator.next();

        if (first.done) continue;

        const fieldCell = first.value;
        const fieldName = (fieldCell?.effectiveValue?.stringValue ?? "").trim();

        if (fieldName !== "") {
            // Save previous Field
            if (currentField.name !== "" && currentField.color !== "" && currentField.subjects.length > 0) academicFields.push(currentField);

            // Start a new one
            const bgColor = fieldCell?.effectiveFormat?.backgroundColor;
            currentField = {
                name: fieldName,
                color: colorToHex(bgColor),
                subjects: [],
            };
        }
        if (currentField.name !== "") {
            let subject = "";
            for (const cellData of iterator) {
                const cellVal = (cellData.effectiveValue?.stringValue ?? "").trim();
                if (cellVal !== "") {
                    subject = cellVal;
                    break;
                }
            }
            if (subject !== "") {
                currentField.subjects.push(subject);
            }
        }
    }
    // Save last Field
    if (currentField.name !== "" && currentField.color !== "" && currentField.subjects.length > 0) academicFields.push(currentField);

    const requests = buildReportFieldsAndSubjects(reportRanges, academicFields);

    return { requests, academicFields };
}

/**
 * Creates the requests to fields and subjects into the persistent data.
 */
function buildReportFieldsAndSubjects(
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
    academicFields: AcademicField[],
): GoogleAppsScript.Sheets.Schema.Request[] {
    const subjectsData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];
    const fieldsData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    for (const field of academicFields) {
        const fieldsDataRow: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { stringValue: field.color } },
            { userEnteredValue: { stringValue: field.name } },
            { userEnteredValue: { numberValue: field.subjects.length } },
        ];
        fieldsData.push(fieldsDataRow);

        const subjectsDataRow: GoogleAppsScript.Sheets.Schema.CellData[][] = field.subjects.map((subject) => [{ userEnteredValue: { stringValue: subject } }]);
        subjectsData.push(...subjectsDataRow);
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
