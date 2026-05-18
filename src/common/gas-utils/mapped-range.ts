import { Dimension } from "../constants";
import { getEpochDate, type MappedNamedRange, type RangeOperationResult, type ResizeRangeParams } from ".";
import { offsetGridRange } from "./range";

interface GetCellParams {
    mappedRange: MappedNamedRange;
    rowOffset?: number;
    columnOffset?: number;
}

export function getCellDataArray(mappedRange: MappedNamedRange, unboundRows = false, unboundColumns = false): GoogleAppsScript.Sheets.Schema.CellData[][] {
    const { range, sheet } = mappedRange;

    if (!range || !sheet) return [];

    const startRow = range.startRowIndex ?? 0;
    const startCol = range.startColumnIndex ?? 0;
    const endRow = (unboundRows ? undefined : range.endRowIndex) ?? sheet.properties?.gridProperties?.rowCount ?? startRow;
    const endCol = (unboundColumns ? undefined : range.endColumnIndex) ?? sheet.properties?.gridProperties?.columnCount ?? startCol;

    const numRows = Math.max(0, endRow - startRow);
    const numCols = Math.max(0, endCol - startCol);

    const result = Array.from({ length: numRows }, () => Array<GoogleAppsScript.Sheets.Schema.CellData>(numCols).fill({}));

    if (!sheet.data || numRows === 0 || numCols === 0) return result;

    for (const gridData of sheet.data) {
        const gridStartRow = gridData.startRow ?? 0;
        const gridStartCol = gridData.startColumn ?? 0;
        const rowDataArray = gridData.rowData ?? [];

        for (let r = 0; r < rowDataArray.length; r++) {
            const absoluteRow = gridStartRow + r;

            // Skip if this row is outside the named range bounds.
            if (absoluteRow < startRow || absoluteRow >= endRow) continue;

            const values = rowDataArray[r]?.values ?? [];
            const targetRow = result[absoluteRow - startRow];

            if (!targetRow) continue;

            for (let c = 0; c < values.length; c++) {
                const absoluteCol = gridStartCol + c;

                // Skip is this colum is outside the named range bounds.
                if (absoluteCol < startCol || absoluteCol >= endCol) continue;

                const resultColIndex = absoluteCol - startCol; // Where in the result this will end.
                const cellData = values[c];

                if (cellData != null) {
                    targetRow[resultColIndex] = cellData;
                }
            }
        }
    }
    return result;
}

export function getCellData({ mappedRange, rowOffset, columnOffset }: GetCellParams): GoogleAppsScript.Sheets.Schema.CellData {
    const absoluteRowIndex = (mappedRange.range.startRowIndex ?? 0) + (rowOffset ?? 0);
    const absoluteColIndex = (mappedRange.range.startColumnIndex ?? 0) + (columnOffset ?? 0);

    const endRow = mappedRange.range.endRowIndex;
    const endColumn = mappedRange.range.endColumnIndex;

    if ((endRow && absoluteRowIndex >= endRow) || (endColumn && absoluteColIndex >= endColumn)) return {};

    const sheetData = mappedRange.sheet.data;
    if (!sheetData) return {};

    for (const gridData of sheetData) {
        const startRow = gridData.startRow ?? 0;
        const startColumn = gridData.startColumn ?? 0;

        const relativeRow = absoluteRowIndex - startRow;
        const relativeCol = absoluteColIndex - startColumn;

        if (relativeRow >= 0 && gridData.rowData && relativeRow < gridData.rowData.length) {
            const rowData = gridData.rowData[relativeRow];

            if (relativeCol >= 0 && rowData?.values && relativeCol < rowData.values.length) {
                return rowData.values[relativeCol] ?? {};
            }
        }
    }

    return {};
}

export function getCellEffectiveValue(args: GetCellParams): GoogleAppsScript.Sheets.Schema.ExtendedValue {
    const cellData = getCellData(args);

    return cellData.effectiveValue ?? {};
}

export function getCellText(args: GetCellParams): string {
    const effectiveValue = getCellEffectiveValue(args);

    return effectiveValue.stringValue ?? "";
}

export function getCellBoolean(args: GetCellParams): boolean {
    const effectiveValue = getCellEffectiveValue(args);

    return effectiveValue.boolValue ?? false;
}

export function getCellNumber(args: GetCellParams): number {
    const effectiveValue = getCellEffectiveValue(args);

    return effectiveValue.numberValue ?? 0;
}

export function getCellUnixEpoch(args: GetCellParams): number {
    const cellNumber = getCellNumber(args);

    return getEpochDate(cellNumber);
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
