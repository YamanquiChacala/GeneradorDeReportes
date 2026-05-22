import { colorToHex } from "../gas-utils";
import type { AcademicField, WeightedSubject } from "../report-utils";

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
