import { formatDateRange } from "../utils";
import type { ReportPersistentData } from ".";

/**
 * Returns a nice string for the asked period.
 */
export function generatePeriodString(data: ReportPersistentData, period: 0 | 1 | 2): string {
    const { configData, calendar } = data;

    const nextPeriod = [1, 2, 3] as const;

    const startBoundary = configData.dates[period] - 1;
    const endBoundary = configData.dates[nextPeriod[period]];

    // Get the indices using our single binary search function
    const startIndex = getUpperBoundIndex(calendar, startBoundary);
    const endIndex = getUpperBoundIndex(calendar, endBoundary);

    // Resolve the actual days.
    const actualStart = calendar[startIndex];
    const actualEnd = calendar[endIndex - 1];

    // Exit if dates weren't found, or if the boundaries resulted in crossed dates
    if (!actualStart || !actualEnd || actualStart > actualEnd) return "No hay fechas";

    return formatDateRange(actualStart, actualEnd);
}

/**
 * Returns the index of the first element in the calendar strictly greater than the threshold.
 */
function getUpperBoundIndex(calendar: readonly number[], threshold: number): number {
    let left = 0;
    let right = calendar.length - 1;
    let bestIndex = calendar.length;

    while (left <= right) {
        const mid = Math.floor((left + right) / 2);
        const midVal = calendar[mid];

        if (midVal === undefined) {
            break;
        }

        if (midVal > threshold) {
            bestIndex = mid;
            right = mid - 1; // It's strictly greater, but look left to find an earlier one
        } else {
            left = mid + 1; // It's <= threshold, look right
        }
    }

    return bestIndex;
}
