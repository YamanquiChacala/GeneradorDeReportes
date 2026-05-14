// TODO: Add tests

import { Dimension } from "../gas-enums";
import type { OffsetGridRangeProperties, RangeOperationResult, ResizeRangeParams } from ".";

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

    const result: GoogleAppsScript.Sheets.Schema.GridRange = {
        sheetId: origin.sheetId,
        startRowIndex: startRow,
        startColumnIndex: startCol,
    };

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

// TODO: Tests for this
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

/**
 * Updates the size of a MappedNamedRange in Sheets and syncs its in-memory state.
 * Allows collapsing to 0 to destroy unused template ranges.
 */
export function resizeMappedRange({ target, targetRows, targetCols, rowOffset = 0, colOffset = 0 }: ResizeRangeParams): RangeOperationResult {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const sheetId = target.range.sheetId ?? 0;

    // Use the helper to determine where the range ACTUALLY is right now before resizing
    const actualRange = offsetGridRange({ origin: target.range, rowOffset, colOffset });
    const actualStartRow = actualRange.startRowIndex ?? 0;
    const actualEndRow = actualRange.endRowIndex ?? 0;
    const actualStartCol = actualRange.startColumnIndex ?? 0;
    const actualEndCol = actualRange.endColumnIndex ?? 0;

    const currentRows = actualEndRow - actualStartRow;
    const currentCols = actualEndCol - actualStartCol;

    // If undefined, default to keeping the current dimensions
    const finalTargetRows = targetRows ?? currentRows;
    const finalTargetCols = targetCols ?? currentCols;

    let newRowOffset = rowOffset;
    let newColOffset = colOffset;

    // Adjust Rows
    if (finalTargetRows > currentRows) {
        const diff = finalTargetRows - currentRows;
        requests.push({
            insertDimension: {
                range: { sheetId, dimension: Dimension.ROWS, startIndex: actualStartRow + 1, endIndex: actualStartRow + 1 + diff },
                inheritFromBefore: true,
            },
        });
        newRowOffset += diff;
    } else if (finalTargetRows < currentRows) {
        const diff = currentRows - finalTargetRows;
        requests.push({
            deleteDimension: {
                // If finalTargetRows is 0, this deletes from actualStartRow to actualEndRow (destroying it)
                range: { sheetId, dimension: Dimension.ROWS, startIndex: actualStartRow + finalTargetRows, endIndex: actualEndRow },
            },
        });
        newRowOffset -= diff;
    }

    // Adjust Columns
    if (finalTargetCols > currentCols) {
        const diff = finalTargetCols - currentCols;
        requests.push({
            insertDimension: {
                range: { sheetId, dimension: Dimension.COLUMNS, startIndex: actualStartCol + 1, endIndex: actualStartCol + 1 + diff },
                inheritFromBefore: true,
            },
        });
        newColOffset += diff;
    } else if (finalTargetCols < currentCols) {
        const diff = currentCols - finalTargetCols;
        requests.push({
            deleteDimension: {
                range: { sheetId, dimension: Dimension.COLUMNS, startIndex: actualStartCol + finalTargetCols, endIndex: actualEndCol },
            },
        });
        newColOffset -= diff;
    }

    // Mutate the in-memory object using the helper to set the final position and size
    target.range = offsetGridRange({
        origin: target.range,
        rowOffset, // Feed it the incoming offset, since origin is still the unmodified original
        colOffset,
        height: finalTargetRows,
        width: finalTargetCols,
    });

    return { requests, rowOffset: newRowOffset, colOffset: newColOffset };
}
