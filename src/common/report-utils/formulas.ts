import { getA1Notation, getColumnLetter, type MappedNamedRange } from "../gas-utils";
import type { AcademicField } from ".";

/**
 * Define the formala for calculating the attendance percent.
 */
export function createAttendaceFormulas(row: number, startCol: number, endCol: number): { percent: string; count: string } {
    if (startCol === -1 || startCol > endCol) return { percent: '=""', count: '=""' };

    const str = getColumnLetter(startCol);
    const end = getColumnLetter(endCol);

    const range = `$${str}${row + 1}:$${end}${row + 1}`;

    const percent = `=IF(COUNTA(${range}) >= COLUMNS(${range}) / 10, (COUNTA(${range}) - SUM(${range})) / COUNTA(${range}), "")`;
    const count = `=IF(COUNTA(${range}) >= COLUMNS(${range}) / 10, SUM(${range}), "")`;

    return { percent, count };
}

/**
 * Helper function to build the formula for shorter comments (for the SEP).
 */
export function getShortCommentFormula(source: MappedNamedRange, rowOffset: number, colOffset: number): string {
    const sourceLetter = getColumnLetter((source.range.startColumnIndex ?? 0) + colOffset);
    const sourceNumber = (source.range.startRowIndex ?? 0) + rowOffset + 1;
    return `=LET(
    raw_text, ${sourceLetter}${sourceNumber},
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
 * Generate a formula for the general absences.
 */
export function createStudentGeneralAttendanceFormula(
    period: 0 | 1 | 2,
    firstNameMappedRange: MappedNamedRange,
    lastNameMappedRange: MappedNamedRange,
    attendanceSheetName: string,
    absences = false,
): string {
    const firstNameA1 = getA1Notation({ mappedRange: firstNameMappedRange, width: 1, height: 1, lockRows: true, lockColumns: true });
    const lastNameA1 = getA1Notation({ mappedRange: lastNameMappedRange, width: 1, height: 1, lockRows: true, lockColumns: true });
    const returnColumn = getColumnLetter((absences ? 4 : 3) + period * 2);
    return `=LET(
    first_names, '${attendanceSheetName}'!B:B,
    last_names, '${attendanceSheetName}'!C:C,
    return_data, '${attendanceSheetName}'!${returnColumn}:${returnColumn},
    raw_result, XLOOKUP(${firstNameA1}&${lastNameA1}, ARRAYFORMULA(first_names&last_names), return_data),
    ARRAYFORMULA(IF(ISNUMBER(raw_result), ROUND(raw_result * { ${absences ? 1 : 10} }, 2), raw_result))
)`;
}

/**
 * Generate the student forumala to grab the absences per subject.
 */
export function createStudentPerSubjectAttendanceFormula(
    period: 0 | 1 | 2,
    subjectRange: MappedNamedRange,
    rowOffset: number,
    firstNameRange: MappedNamedRange,
    lastNameRange: MappedNamedRange,
    attendanceSheetName: string,
): string {
    const subjectA1 = getA1Notation({ mappedRange: subjectRange, width: 1, height: 1, rowOffset, lockColumns: true });
    const firstNameA1 = getA1Notation({ mappedRange: firstNameRange, width: 1, height: 1, lockRows: true, lockColumns: true });
    const lastNameA1 = getA1Notation({ mappedRange: lastNameRange, width: 1, height: 1, lockRows: true, lockColumns: true });
    const returnColum = getColumnLetter(3 + period * 2);
    return `=LET(
    start_row, MATCH(${subjectA1}, ${attendanceSheetName}!A:A, 0),
    next_subject_offset,
        IFERROR(
            MATCH(
                TRUE,
                ARRAYFORMULA(
                    ISTEXT(
                        INDEX(${attendanceSheetName}!A:A, start_row + 1):
                        INDEX(${attendanceSheetName}!A:A, ROWS(${attendanceSheetName}!A:A))
                    )
                ),
                0
            ),
            ROWS(${attendanceSheetName}!A:A) - start_row
        ),
    height, next_subject_offset - 1,
    first_names, OFFSET(${attendanceSheetName}!B$1, start_row, 0, height, 1),
    last_names, OFFSET(${attendanceSheetName}!C$1, start_row, 0, height, 1),
    return_data, OFFSET(${attendanceSheetName}!${returnColum}$1, start_row, 0, height, 2),
    raw_result, XLOOKUP(${firstNameA1} & ${lastNameA1}, ARRAYFORMULA(first_names & last_names), return_data),
    ARRAYFORMULA(IF(ISNUMBER(raw_result), ROUND(raw_result * { 10, 1 }, 2), raw_result))
)`;
}

/**
 * Formula to claculate each subject's average.
 */
export function createIndividualSubjectAverageFormula(valuesRange: MappedNamedRange, rowOffset: number, weightsRange: MappedNamedRange): string {
    const valuesA1 = getA1Notation({ mappedRange: valuesRange, rowOffset, colOffset: 1, height: 1, width: 3, lockColumns: true });
    const weightsA1 = getA1Notation({ mappedRange: weightsRange, includeSheetName: true, lockRows: true, lockColumns: true });

    return `=IFERROR(
    ROUND(
        AVERAGE.WEIGHTED(
            ${valuesA1},
            TRANSPOSE(${weightsA1})
        )
    ), ""
)`;
}

/**
 * Formula to calculate the final subject's average.
 */
export function createFinalSubjectAverageFormula(trim1: MappedNamedRange, trim2: MappedNamedRange, trim3: MappedNamedRange, rowOffset: number): string {
    const colOffset = (trim3.range.endColumnIndex ?? 0) - 2;

    const trim1A1 = getA1Notation({ mappedRange: trim1, rowOffset, colOffset, height: 1, width: 1, lockColumns: true });
    const trim2A1 = getA1Notation({ mappedRange: trim2, rowOffset, colOffset, height: 1, width: 1, lockColumns: true });
    const trim3A1 = getA1Notation({ mappedRange: trim3, rowOffset, colOffset, height: 1, width: 1, lockColumns: true });

    return `=IF(
    COUNT(${trim1A1}, ${trim2A1}, ${trim3A1})=3,
    ROUND(AVERAGE(${trim1A1}, ${trim2A1}, ${trim3A1}),1),
    ""
)`;
}

/**
 * Helper function to calculate the field value from the subjects.
 */
export function createFieldFormula(
    field: AcademicField,
    colOffset: number,
    subjectsOffset: number,
    subjectsRange: MappedNamedRange,
    weightsRange: MappedNamedRange,
): string {
    const valuesA1 = getA1Notation({
        mappedRange: subjectsRange,
        rowOffset: subjectsOffset,
        colOffset,
        width: 1,
        height: field.subjects,
        lockRows: true,
    });
    const weightsA1 = getA1Notation({
        includeSheetName: true,
        mappedRange: weightsRange,
        rowOffset: subjectsOffset,
        colOffset: 1,
        width: 1,
        height: field.subjects,
        lockRows: true,
        lockColumns: true,
    });

    return `=IFERROR(
    AVERAGE.WEIGHTED(${valuesA1}, ${weightsA1})
)`;
}

/**
 * Helper to calculate the average of the academic fields
 */
export function createFieldAverageFormula(mappedFields: MappedNamedRange, colOffset: number): string {
    const valuesA1 = getA1Notation({ mappedRange: mappedFields, colOffset, width: 1, lockRows: true });
    return `=IFERROR(ROUND(AVERAGE(${valuesA1}), 1))`;
}

/**
 * Helper to crete the average of all the subjects
 */
export function createAllSubjectsAverageFormula(mappedSubjects: MappedNamedRange, mappedWeights: MappedNamedRange, colOffset: number): string {
    const valuesA1 = getA1Notation({ mappedRange: mappedSubjects, colOffset, width: 1, lockRows: true });
    const weightsA1 = getA1Notation({ mappedRange: mappedWeights, colOffset: 1, width: 1, includeSheetName: true, lockRows: true, lockColumns: true });
    return `=IFERROR(ROUND(AVERAGE.WEIGHTED(${valuesA1}, ${weightsA1}), 1))`;
}
