import { getUpperBoundIndex } from "../utils";
import type { FrozenArea, Range, ReportPersistentData, TrimesterRanges } from "./types";

/**
 * Calculate trimester column indexes.
 */
export function calculateAssistanceTrimRanges(data: ReportPersistentData, frozenArea: FrozenArea): TrimesterRanges {
    const frozenCols = frozenArea.cols;
    const format = (start: number, end: number): Range => (start <= end ? { start: start + frozenCols, end: end + frozenCols } : { start: -1, end: -1 });

    return {
        trim1: format(getUpperBoundIndex(data.calendar, data.configData.dates[0] - 1), getUpperBoundIndex(data.calendar, data.configData.dates[1]) - 1),
        trim2: format(getUpperBoundIndex(data.calendar, data.configData.dates[1]), getUpperBoundIndex(data.calendar, data.configData.dates[2]) - 1),
        trim3: format(getUpperBoundIndex(data.calendar, data.configData.dates[2]), getUpperBoundIndex(data.calendar, data.configData.dates[3]) - 1),
    };
}
