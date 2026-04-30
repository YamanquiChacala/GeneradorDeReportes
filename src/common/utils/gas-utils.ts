import type { PasteType } from "../gas-enums";

/**
 * Transforms a column number into it's corresponding column letter, using 0-based index.
 */
export function getColumnLetter(column: number): string {
    let temp: number,
        letter = "";
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
}

export function createSingleCellRange(sheetId: number, startRowIndex: number, startColumnIndex: number): GoogleAppsScript.Sheets.Schema.GridRange {
    return {
        sheetId,
        startRowIndex,
        startColumnIndex,
        endRowIndex: startRowIndex + 1,
        endColumnIndex: startColumnIndex + 1,
    };
}

/**
 * Changes the sheetId of the grid. Moving a grid to another sheet.
 */
export function changeGridRangeSheet(grid: GoogleAppsScript.Sheets.Schema.GridRange, sheetId: number): GoogleAppsScript.Sheets.Schema.GridRange {
    const newGrid = { ...grid };
    newGrid.sheetId = sheetId;
    return newGrid;
}

interface OffsetGridRangeProperties {
    origin: GoogleAppsScript.Sheets.Schema.GridRange;
    rowOffset?: number;
    colOffset?: number;
    height?: number;
    width?: number;
}

/**
 * Offsets a `range` and optionally bounds it to a given size.
 */
export function offsetGridRange({ origin, rowOffset = 0, colOffset = 0, height, width }: OffsetGridRangeProperties): GoogleAppsScript.Sheets.Schema.GridRange {
    const startRow = (origin.startRowIndex ?? 0) + rowOffset;
    const startCol = (origin.startColumnIndex ?? 0) + colOffset;

    const result: GoogleAppsScript.Sheets.Schema.GridRange = {
        sheetId: origin.sheetId,
        startRowIndex: startRow,
        startColumnIndex: startCol,
    };

    if (height != null) {
        result.endRowIndex = startRow + height;
    } else if (origin.endRowIndex != null) {
        result.endRowIndex = origin.endRowIndex + rowOffset;
    }

    if (width != null) {
        result.endColumnIndex = startCol + width;
    } else if (origin.endColumnIndex != null) {
        result.endColumnIndex = origin.endColumnIndex + colOffset;
    }

    return result;
}

/** Helper function to convert Sheets API Color to a Hex string */
export function colorToHex(color?: GoogleAppsScript.Sheets.Schema.Color, fallback = "#FFFFFF"): string {
    // If there's no color provided, default to white as requested
    if (!color) return fallback;

    // Sheets API color channels default to 0 if not present in the object
    const r = Math.round((color.red || 0) * 255);
    const g = Math.round((color.green || 0) * 255);
    const b = Math.round((color.blue || 0) * 255);

    const toHex = (n: number) => n.toString(16).padStart(2, "0").toUpperCase();

    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

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

/**
 * Generates batch update `copyPaste` request to copy data from `origin` into `destination` ranges.
 */
export function buildCopyPasteRequest(
    source: GoogleAppsScript.Sheets.Schema.GridRange | undefined,
    destination: GoogleAppsScript.Sheets.Schema.GridRange | undefined,
    pasteType: PasteType,
): GoogleAppsScript.Sheets.Schema.Request | undefined {
    if (!source || !destination) return undefined;

    return {
        copyPaste: {
            source,
            destination,
            pasteType,
            pasteOrientation: "NORMAL",
        },
    };
}

/**
 * Generates batch update requests to put `data` into the range defined by `destination`.
 * @param fields Mask to see what to copy.
 * @param adaptRange If true, add/remove rows and columns to adapt the range to the size of `data`
 */
export function buildTransferRequest(
    destination: GoogleAppsScript.Sheets.Schema.GridRange,
    data: GoogleAppsScript.Sheets.Schema.CellData[][],
    fields: string,
    adaptRange: boolean = false,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const sheetId = destination.sheetId ?? 0;
    const startRow = destination.startRowIndex ?? 0;
    const endRow = destination.endRowIndex ?? 0;
    const startCol = destination.startColumnIndex ?? 0;
    const endCol = destination.endColumnIndex ?? 0;

    const destRows = endRow - startRow;
    const destCols = endCol - startCol;

    const dataRows = data.length;
    // Safely find the widest row in case of jagged arrays
    const dataCols = dataRows > 0 ? Math.max(...data.map((row) => row.length)) : 0;

    let finalRows = destRows;
    let finalCols = destCols;

    if (adaptRange) {
        finalRows = dataRows;
        finalCols = dataCols;

        // Safeguard: If data is completely empty, skip resizing to avoid destroying the named range
        if (dataRows === 0 || dataCols === 0) return requests;

        // Adjust Rows
        if (dataRows > destRows) {
            requests.push({
                insertDimension: {
                    range: { sheetId, dimension: "ROWS", startIndex: startRow + 1, endIndex: startRow + 1 + (dataRows - destRows) },
                    inheritFromBefore: true,
                },
            });
        } else if (dataRows < destRows) {
            requests.push({
                deleteDimension: {
                    range: { sheetId, dimension: "ROWS", startIndex: startRow + dataRows, endIndex: endRow },
                },
            });
        }

        // Adjust Columns
        if (dataCols > destCols) {
            requests.push({
                insertDimension: {
                    range: { sheetId, dimension: "COLUMNS", startIndex: startCol + 1, endIndex: startCol + 1 + (dataCols - destCols) },
                    inheritFromBefore: true,
                },
            });
        } else if (dataCols < destCols) {
            requests.push({
                deleteDimension: {
                    range: { sheetId, dimension: "COLUMNS", startIndex: startCol + dataCols, endIndex: endCol },
                },
            });
        }
    }

    const finalRowData: GoogleAppsScript.Sheets.Schema.RowData[] = [];
    for (let r = 0; r < finalRows; r++) {
        const rowValues: GoogleAppsScript.Sheets.Schema.CellData[] = [];
        const sourceRow = data[r] ?? [];

        for (let c = 0; c < finalCols; c++) {
            rowValues.push(sourceRow[r] ?? {});
        }

        finalRowData.push({ values: rowValues });
    }

    // Write the data
    if (finalRows && finalCols) {
        requests.push({
            updateCells: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: startRow,
                    endRowIndex: startRow + finalRows,
                    startColumnIndex: startCol,
                    endColumnIndex: startCol + finalCols,
                },
                rows: finalRowData,
                fields: fields,
            },
        });
    }

    return requests;
}
