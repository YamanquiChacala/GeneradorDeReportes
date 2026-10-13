import type { OffsetGridRangeProperties } from "./types";

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
 * Offsets a `range` and optionally bounds it to a given size.
 * height or width negative unbounds the range.
 */
export function offsetGridRange({ origin, sheetId, rowOffset = 0, colOffset = 0, height, width }: OffsetGridRangeProperties): GoogleAppsScript.Sheets.Schema.GridRange {
    const result: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: origin.sheetId };

    if (sheetId != null) {
        result.sheetId = sheetId;
    }

    const startRow = Math.max(0, (origin.startRowIndex ?? 0) + rowOffset);
    const startCol = Math.max(0, (origin.startColumnIndex ?? 0) + colOffset);

    if (startRow > 0) {
        result.startRowIndex = startRow;
    }
    if (startCol > 0) {
        result.startColumnIndex = startCol;
    }

    if (height != null) {
        if (height >= 0) result.endRowIndex = startRow + height;
    } else if (origin.endRowIndex != null) {
        result.endRowIndex = Math.max(startRow, origin.endRowIndex + rowOffset);
    }

    if (width != null) {
        if (width >= 0) result.endColumnIndex = startCol + width;
    } else if (origin.endColumnIndex != null) {
        result.endColumnIndex = Math.max(startCol, origin.endColumnIndex + colOffset);
    }

    return result;
}

/**
 * Safely calculates the width of a GridRange, defaulting to 0 if bounds are missing.
 */
export function getRangeWidth(range: GoogleAppsScript.Sheets.Schema.GridRange): number {
    return Math.max(0, (range.endColumnIndex ?? 0) - (range.startColumnIndex ?? 0));
}

/**
 * Safely calculates the height of a GridRange, defaulting to 0 if bounds are missing.
 */
export function getRangeHeight(range: GoogleAppsScript.Sheets.Schema.GridRange): number {
    return Math.max(0, (range.endRowIndex ?? 0) - (range.startRowIndex ?? 0));
}

/**
 * Shrinks a range by the given number of columns.
 */
export function shrinkRangeWidth(range: GoogleAppsScript.Sheets.Schema.GridRange, colsToRemove: number): GoogleAppsScript.Sheets.Schema.GridRange {
    const currentWidth = getRangeWidth(range);
    const desiredWidth = Math.max(0, currentWidth - colsToRemove);
    return offsetGridRange({ origin: range, width: desiredWidth });
}

/**
 * Returns the cell data within a GridRange, leaving cells without returned data as empty objects.
 * Missing range end indexes default to the sheet's grid dimensions when available.
 */
export function getRangeData(sheet: GoogleAppsScript.Sheets.Schema.Sheet, range: GoogleAppsScript.Sheets.Schema.GridRange): GoogleAppsScript.Sheets.Schema.CellData[][] {
    const startRow = range.startRowIndex ?? 0;
    const startCol = range.startColumnIndex ?? 0;
    const endRow = range.endRowIndex ?? sheet.properties?.gridProperties?.rowCount ?? startRow;
    const endCol = range.endColumnIndex ?? sheet.properties?.gridProperties?.columnCount ?? startCol;

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

            // Skip if this row is outside the range bounds.
            if (absoluteRow < startRow || absoluteRow >= endRow) continue;

            const values = rowDataArray[r]?.values ?? [];
            // biome-ignore lint/style/noNonNullAssertion: Math proves this must exist.
            const targetRow = result[absoluteRow - startRow]!;

            for (let c = 0; c < values.length; c++) {
                const absoluteCol = gridStartCol + c;

                // Skip is this colum is outside the named range bounds.
                if (absoluteCol < startCol || absoluteCol >= endCol) continue;

                const resultColIndex = absoluteCol - startCol; // Where in the result this will end.
                // biome-ignore lint/style/noNonNullAssertion: `c` iterates on the lenght of values.
                const cellData = values[c]!;

                targetRow[resultColIndex] = cellData;
            }
        }
    }
    return result;
}
