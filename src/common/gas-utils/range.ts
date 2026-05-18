import type { OffsetGridRangeProperties } from ".";

/**
 * Creates a range with the given data.
 */
export function createRange(sheetId: number, startRowIndex: number, startColumnIndex: number, height?: number, width?: number): GoogleAppsScript.Sheets.Schema.GridRange {
    const result: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId, startRowIndex, startColumnIndex };
    if (height && height > 0) result.endRowIndex = startRowIndex + height;
    if (width && width > 0) result.endColumnIndex = startColumnIndex + width;
    return result;
}

/**
 * Creates a range with a single cell.
 */

export function createSingleCellRange(sheetId: number, startRowIndex: number, startColumnIndex: number): GoogleAppsScript.Sheets.Schema.GridRange {
    return createRange(sheetId, startRowIndex, startColumnIndex, 1, 1);
}

/**
 * Changes the sheetId of the grid. Moving a grid to another sheet.
 */

export function changeGridRangeSheet(grid: GoogleAppsScript.Sheets.Schema.GridRange, sheetId: number): GoogleAppsScript.Sheets.Schema.GridRange {
    const newGrid = { ...grid };
    newGrid.sheetId = sheetId;
    return newGrid;
}

/**
 * Offsets a `range` and optionally bounds it to a given size.
 * height or width negative unbounds the range.
 */
export function offsetGridRange({ origin, rowOffset = 0, colOffset = 0, height, width }: OffsetGridRangeProperties): GoogleAppsScript.Sheets.Schema.GridRange {
    const startRow = Math.max(0, (origin.startRowIndex ?? 0) + rowOffset);
    const startCol = Math.max(0, (origin.startColumnIndex ?? 0) + colOffset);

    const result: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: origin.sheetId };

    if (startRow > 0) {
        result.startRowIndex = startRow;
    }
    if (startCol > 0) {
        result.startColumnIndex = startCol;
    }

    if (height != null) {
        if (height > 0) result.endRowIndex = startRow + height;
    } else if (origin.endRowIndex != null) {
        result.endRowIndex = Math.max(startRow, origin.endRowIndex + rowOffset);
    }

    if (width != null) {
        if (width > 0) result.endColumnIndex = startCol + width;
    } else if (origin.endColumnIndex != null) {
        result.endColumnIndex = Math.max(startCol, origin.endColumnIndex + colOffset);
    }

    return result;
}

// TODO: Still needed?
/**
 * Calculates a new shifted range and the resulting offset for sequential ranges.
 */
export function calculateRangeShift(
    originRange: GoogleAppsScript.Sheets.Schema.GridRange,
    dataLength: number,
    currentOffset: number,
): { newRange: GoogleAppsScript.Sheets.Schema.GridRange; nextRowOffset: number } {
    const oldHeight = (originRange.endRowIndex ?? 0) - (originRange.startRowIndex ?? 0);
    const newRange = offsetGridRange({ origin: originRange, rowOffset: currentOffset, height: dataLength });

    return {
        newRange,
        nextRowOffset: currentOffset + dataLength - oldHeight,
    };
}
