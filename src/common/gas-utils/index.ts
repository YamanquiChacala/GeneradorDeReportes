import type { MappedNamedRange } from "../utils/mapped-name-range";

export * from "./color";
export * from "./notation";
export * from "./range";
export * from "./request";
export * from "./time";

export interface OffsetGridRangeProperties {
    origin: GoogleAppsScript.Sheets.Schema.GridRange;
    rowOffset?: number;
    colOffset?: number;
    height?: number;
    width?: number;
}

export interface ResizeRangeParams {
    target: MappedNamedRange;
    targetRows?: number;
    targetCols?: number;
    rowOffset?: number;
    colOffset?: number;
}

export interface RangeOperationResult {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    rowOffset: number;
    colOffset: number;
}

// TODO: Really needed?
/**
 * Trasnsforms from `effetive...` to `userEntered...` for use in writting.
 * @param stripOthers if true, only keep `userEntered...`, remove everything elese.
 */
export function makeUserEntered(data: GoogleAppsScript.Sheets.Schema.CellData[][], stripOthers: boolean = false): GoogleAppsScript.Sheets.Schema.CellData[][] {
    return data.map((row) =>
        row.map((cell) => {
            const finalValue = cell.effectiveValue !== undefined ? cell.effectiveValue : cell.userEnteredValue;
            const finalFormat = cell.effectiveFormat !== undefined ? cell.effectiveFormat : cell.userEnteredFormat;

            const newCell: GoogleAppsScript.Sheets.Schema.CellData = {};

            if (!stripOthers) {
                Object.assign(newCell, cell);
            }

            if (finalValue !== undefined) newCell.userEnteredValue = finalValue;
            if (finalFormat !== undefined) newCell.userEnteredFormat = finalFormat;

            delete newCell.effectiveValue;
            delete newCell.effectiveFormat;

            return newCell;
        }),
    );
}
