import { MORE_THAN_A_YEAR, MS_PER_DAY } from "../utils";
import { type CalendarDates, type CalendarGrid, type DayData, DayType, type MonthBlock, type WeekData } from "./types";

/**
 * Ensures the user given dates are in order and valid
 */
export function validateDates(dates: [number, number, number, number]): string | null {
    if (dates.length < 4 || dates.some((d) => !d)) return "Faltan fechas.";
    if (!(dates[0] < dates[1] && dates[1] < dates[2] && dates[2] < dates[3])) return "Fechas en desorden.";
    if (dates[3] - dates[0] > MORE_THAN_A_YEAR) return "Calendario demasiado largo.";
    return null; // Valid
}

/**
 * Calculates the data to create the calendar form the period dates.
 * Every date must be in Unix Epoch milliseconds.
 */
export function calculateCalendarDates(dates: [number, number, number, number]): CalendarDates {
    const dateError = validateDates(dates);
    if (dateError) {
        throw new Error(dateError);
    }

    // Snap to first Sunday
    const dateStartDayOfWeek = new Date(dates[0]).getUTCDay();
    const calStart = dates[0] - dateStartDayOfWeek * MS_PER_DAY;

    // Snap to last Saturday
    const dateEndDayOfWeek = new Date(dates[3]).getUTCDay();
    const calEnd = dates[3] + (6 - dateEndDayOfWeek) * MS_PER_DAY;

    const totalDays = (calEnd - calStart) / MS_PER_DAY + 1;
    const totalRows = 1 + Math.ceil(totalDays / 7); // Header + weeks

    return { dateStart: dates[0], dateTrimester1: dates[1], dateTrimester2: dates[2], dateEnd: dates[3], calStart, calEnd, totalDays, totalRows };
}

/**
 * Calculates the calendar grid structure without GAS dependencies.
 */
export function calculateCalendarGrid(dates: CalendarDates): CalendarGrid {
    const monthBlocks: MonthBlock[] = [];
    const weeks: WeekData[] = [];

    let currentMs = dates.calStart;
    let currentRowNumber = 1;
    let currentMonthIndex = -1;
    let currentYear = -1;
    let monthStartRow = 1;

    while (currentMs <= dates.calEnd) {
        // Wednesday determines the month
        const wednesdayMs = currentMs + 3 * MS_PER_DAY;
        const wednesdayDate = new Date(wednesdayMs);
        const monthIndex = wednesdayDate.getUTCMonth();
        const year = wednesdayDate.getUTCFullYear();

        // Month boundary logic
        if (currentMonthIndex !== monthIndex) {
            if (currentMonthIndex !== -1) {
                monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });
            }
            currentMonthIndex = monthIndex;
            currentYear = year;
            monthStartRow = currentRowNumber;
        }

        const days: DayData[] = [];

        // Process 7 days for this week/row
        for (let i = 0; i < 7; i++) {
            const dayDate = new Date(currentMs);
            const dayOfWeek = dayDate.getUTCDay();
            const isWeekday = dayOfWeek !== 0 && dayOfWeek !== 6;
            const inBounds = currentMs >= dates.dateStart && currentMs <= dates.dateEnd;

            // Determine day formatting
            let dayType = DayType.REST;
            if (isWeekday && inBounds) {
                if (currentMs > dates.dateTrimester2) dayType = DayType.TRIM3;
                else if (currentMs > dates.dateTrimester1) dayType = DayType.TRIM2;
                else dayType = DayType.TRIM1;
            }

            days.push({
                ms: currentMs,
                dateNumber: dayDate.getUTCDate(),
                isWeekday,
                inBounds,
                dayType,
            });

            currentMs += MS_PER_DAY;
        }

        weeks.push({ rowNumber: currentRowNumber, days });
        currentRowNumber++;
    }

    // Close final month block
    monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });

    return { monthBlocks, weeks };
}
