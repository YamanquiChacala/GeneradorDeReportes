import { createAttendaceFormulas, type StudentRow, type TrimesterRanges } from "../report-utils";
import { TemplateSize } from "./types";

interface MonthGroupMeta {
    year: number;
    month: number;
    startCol: number;
    count: number;
    template: TemplateSize;
}

/**
 * Builds the data to form the Assistance Headers.
 */
export function calculateCalendarHeaders(
    days: number[],
    frozenCols: number,
    names1: string[],
    names2: string[],
    names5: string[],
    dayNames: string[],
): { monthGroups: MonthGroupMeta[]; row1Values: (string | null)[]; row2Values: string[]; row3Values: number[] } {
    const monthGroups: MonthGroupMeta[] = [];
    const row1Values: (string | null)[] = [];
    const row2Values: string[] = [];
    const row3Values: number[] = [];

    if (days.length === 0) {
        return { monthGroups, row1Values, row2Values, row3Values };
    }

    let currentGroup: MonthGroupMeta | null = null;

    const finalizeGroup = (group: MonthGroupMeta) => {
        const shortYear = String(group.year).slice(-2);
        const longYear = String(group.year);
        let text = "";

        if (group.count === 1) {
            group.template = TemplateSize.SMALL;
            const name = names1[group.month] ?? "";
            text = `${name}\n${shortYear}`;
        } else if (group.count >= 2 && group.count <= 4) {
            group.template = TemplateSize.MEDIUM;
            const name = names2[group.month] ?? "";
            text = `${name}\n${longYear}`;
        } else {
            group.template = TemplateSize.LARGE;
            const name = names5[group.month] ?? "";
            text = `${name}\n${longYear}`;
        }

        monthGroups.push(group);

        for (let i = 0; i < group.count; i++) {
            row1Values.push(i === 0 ? text : null);
        }
    };

    days.forEach((dayValue, i) => {
        const date = new Date(dayValue);
        const year = date.getUTCFullYear();
        const month = date.getUTCMonth();
        const dayOfWeek = date.getUTCDay();
        const dayOfMonth = date.getUTCDate();

        const targetCol = frozenCols + i;

        if (!currentGroup) {
            currentGroup = { year, month, startCol: targetCol, count: 1, template: TemplateSize.SMALL };
        } else if (currentGroup.year === year && currentGroup.month === month) {
            currentGroup.count++;
        } else {
            finalizeGroup(currentGroup);
            currentGroup = { year, month, startCol: targetCol, count: 1, template: TemplateSize.SMALL };
        }

        const dayName = dayNames[dayOfWeek] ?? "";
        row2Values.push(dayName);
        row3Values.push(dayOfMonth);
    });

    // biome-ignore lint/style/noNonNullAssertion: At lest 1 day exists, so the cycle ran at least 1.
    finalizeGroup(currentGroup!);

    return { monthGroups, row1Values, row2Values, row3Values };
}

/**
 * Builds the data for the student list in Attendance
 */
export function generateStudentGrid(students: StudentRow[], initialRow: number, trimesters: TrimesterRanges): GoogleAppsScript.Sheets.Schema.CellData[][] {
    const result: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    for (let i = 0; i < students.length; i++) {
        const studentRow = students[i];

        if (studentRow?.type === "student") {
            const trim1 = createAttendaceFormulas(initialRow + i, trimesters.trim1.start, trimesters.trim1.end);
            const trim2 = createAttendaceFormulas(initialRow + i, trimesters.trim2.start, trimesters.trim2.end);
            const trim3 = createAttendaceFormulas(initialRow + i, trimesters.trim3.start, trimesters.trim3.end);

            result.push([
                { userEnteredValue: { numberValue: studentRow.id } },
                { userEnteredValue: { stringValue: studentRow.firstName } },
                { userEnteredValue: { stringValue: studentRow.lastName } },
                { userEnteredValue: { formulaValue: trim1.percent } },
                { userEnteredValue: { formulaValue: trim1.count } },
                { userEnteredValue: { formulaValue: trim2.percent } },
                { userEnteredValue: { formulaValue: trim2.count } },
                { userEnteredValue: { formulaValue: trim3.percent } },
                { userEnteredValue: { formulaValue: trim3.count } },
            ]);
        } else {
            result.push([]); // Keep empty rows for separators
        }
    }

    return result;
}
