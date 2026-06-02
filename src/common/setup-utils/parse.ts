import { colorToHex, getSheetsDate } from "../gas-utils";
import type { AcademicField, Student, StudentRow, WeightedSubject } from "../report-utils";
import { MS_PER_DAY, sanitizeSheetName } from "../utils";

/**
 * Parses a 2D array of Google Sheets CellData into structured AcademicFields and WeightedSubjects.
 */
export function parseAcademicFieldsAndSubjects(setupSubjectData: readonly GoogleAppsScript.Sheets.Schema.CellData[][]): {
    academicFields: AcademicField[];
    rawSubjects: WeightedSubject[];
} {
    const academicFields: AcademicField[] = [];
    const rawSubjects: WeightedSubject[] = [];

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
            // Save previous Field if it's complete.
            if (currentField.name !== "" && currentField.color !== "" && currentField.subjects > 0) {
                academicFields.push(currentField);
            }

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
                // Safely handle potentially undefined cellData via optional chaining
                const cellNumVal = cellData?.effectiveValue?.numberValue;
                if (cellNumVal !== undefined) {
                    weight = cellNumVal;
                    continue;
                }

                const cellStringVal = (cellData?.effectiveValue?.stringValue ?? "").trim();
                if (cellStringVal !== "") {
                    subject = cellStringVal;
                }
            }

            if (subject !== "") {
                currentField.subjects++;
                rawSubjects.push({ weight, subject });
            }
        }
    }

    // Save the final Field
    if (currentField.name !== "" && currentField.color !== "" && currentField.subjects > 0) {
        academicFields.push(currentField);
    }

    return { academicFields, rawSubjects };
}

/**
 * Parses a 2D array of CellData into structured StudentRows.
 */
export function parseStudentList(studentSetupData: readonly GoogleAppsScript.Sheets.Schema.CellData[][]): {
    studentReportData: GoogleAppsScript.Sheets.Schema.CellData[][];
    students: StudentRow[];
} {
    const students: StudentRow[] = [];
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
            sex: studentRow[2]?.effectiveValue?.stringValue ?? (studentRow[2]?.effectiveValue?.numberValue != null ? `${studentRow[2].effectiveValue.numberValue}` : ""),
            level:
                studentRow[3]?.effectiveValue?.stringValue ?? (studentRow[3]?.effectiveValue?.numberValue != null ? `${studentRow[3].effectiveValue.numberValue}` : ""),
            grade:
                studentRow[4]?.effectiveValue?.stringValue ?? (studentRow[4]?.effectiveValue?.numberValue != null ? `${studentRow[4].effectiveValue.numberValue}º` : ""),
            curp: studentRow[5]?.effectiveValue?.stringValue ?? (studentRow[5]?.effectiveValue?.numberValue != null ? `${studentRow[5].effectiveValue.numberValue}` : ""),
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

    return { studentReportData, students };
}

/**
 * Parse a 2D calendar into a list of dates
 */
export function parseCalendarDays(
    initialDay: number,
    calendarRawData: GoogleAppsScript.Sheets.Schema.CellData[][],
): {
    sheetDays: GoogleAppsScript.Sheets.Schema.CellData[][];
    days: number[];
} {
    if (new Date(initialDay).getUTCDay() !== 0) throw new Error("Calendario no inicia en Domingo.");

    const days: number[] = [];
    const sheetDays: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    for (const [rowIndex, row] of calendarRawData.entries()) {
        for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
            const columnIndex = dayIndex * 2 + 1;
            const cell = row[columnIndex];

            const isChecked = cell?.effectiveValue?.boolValue === true;

            if (isChecked) {
                const day = initialDay + (rowIndex * 7 + dayIndex) * MS_PER_DAY;
                days.push(day);
                sheetDays.push([
                    {
                        userEnteredValue: { numberValue: getSheetsDate(day) },
                    },
                ]);
            }
        }
    }

    return { days, sheetDays };
}
