import { formatDateRange, getUpperBoundIndex } from "../utils";
import type { ReportPersistentData } from "./types";

/**
 * Returns a nice string for the asked period.
 */
export function generatePeriodString(data: ReportPersistentData, period: 0 | 1 | 2): string | undefined {
    const { configData, calendar } = data;

    const startBoundary = configData.dates[period] - (period === 0 ? 1 : 0);
    // biome-ignore lint/style/noNonNullAssertion: dates has 4 elements.
    const endBoundary = configData.dates[period + 1]!;

    // Get the indices using our single binary search function
    const startIndex = getUpperBoundIndex(calendar, startBoundary);
    const endIndex = getUpperBoundIndex(calendar, endBoundary);

    // Resolve the actual days.
    const actualStart = calendar[startIndex];
    const actualEnd = calendar[endIndex - 1];

    // Exit if dates weren't found, or if the boundaries resulted in crossed dates
    if (!actualStart || !actualEnd || actualStart > actualEnd) return undefined;

    return formatDateRange(actualStart, actualEnd);
}
