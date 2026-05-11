import { MS_PER_DAY, SHEETS_EPOCH_OFFSET } from "../constants";
import { Dimension, type MergeType, PasteOrientation, type PasteType } from "../gas-enums";
import { type HSLColor, hslToRgb, type RGBColor, rgbToHex } from "./color-utils";
import type { ExtractRangeNames, NestedSheetSchema } from "./mapped-name-range";

/**
 * Transforms a column number into it's corresponding column letter, using 0-based index.
 */
export function getColumnLetter(column: number): string {
    let temp: number;
    let letter = "";
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
}

// TODO: Make tests for this
/**
 * Returns the A1 notation of a range.
 */
export function getA1Notation(range: GoogleAppsScript.Sheets.Schema.GridRange, first = false, lockRows = false, lockColumns = false): string {
    // 1. Define our lock prefixes
    const rowPrefix = lockRows ? "$" : "";
    const colPrefix = lockColumns ? "$" : "";

    // 2. Check if dimensions are bounded (undefined means unbounded in that dimension)
    const hasRows = range.startRowIndex !== undefined || range.endRowIndex !== undefined;
    const hasCols = range.startColumnIndex !== undefined || range.endColumnIndex !== undefined;

    // 3. Get initial coordinates (defaults to 0 for the 'first' flag fallback)
    const startRow = (range.startRowIndex ?? 0) + 1;
    const startColLetter = getColumnLetter(range.startColumnIndex ?? 0);

    if (first) return `${colPrefix}${startColLetter}${rowPrefix}${startRow}`;

    // 4. Build starting strings (omit entirely if the dimension is unbounded)
    const startRowStr = hasRows ? `${rowPrefix}${startRow}` : "";
    const startColStr = hasCols ? `${colPrefix}${startColLetter}` : "";

    // 5. Build ending strings
    // If end index is omitted on a bounded dimension, we assume it spans 1 unit (end = start + 1)
    // Because end indices are exclusive, endRowIndex exactly matches the 1-based row number.
    const endRow = range.endRowIndex !== undefined ? range.endRowIndex : startRow;
    const endColIndex = range.endColumnIndex !== undefined ? range.endColumnIndex : (range.startColumnIndex ?? 0) + 1;

    // For columns, we subtract 1 to get the actual final column included in the range.
    const endColLetter = getColumnLetter(endColIndex - 1);

    const endRowStr = hasRows ? `${rowPrefix}${endRow}` : "";
    const endColStr = hasCols ? `${colPrefix}${endColLetter}` : "";

    const startA1 = `${startColStr}${startRowStr}`;
    const endA1 = `${endColStr}${endRowStr}`;

    // 6. Handle single entities
    if (startA1 === endA1) {
        if (!hasRows) return `${startColStr}:${startColStr}`; // Unbounded single column (e.g., "$A:$A")
        if (!hasCols) return `${startRowStr}:${startRowStr}`; // Unbounded single row (e.g., "$1:$1")
        return startA1; // Single cell (e.g., "$A$1")
    }

    // 7. Return the standard range
    return `${startA1}:${endA1}`; // e.g., "$A$1:$B$2", "A$1:A", etc.
}

/**
 * Converts a Unix Epoch into a Sheets Epoch.
 */
export function getSheetsDate(epoch: number): number {
    return epoch / MS_PER_DAY + SHEETS_EPOCH_OFFSET;
}

/**
 * Converts a Sheets Epock into a Unix Epoch.
 */
export function getEpochDate(sheetsDate: number): number {
    return (sheetsDate - SHEETS_EPOCH_OFFSET) * MS_PER_DAY;
}

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

interface OffsetGridRangeProperties {
    origin: GoogleAppsScript.Sheets.Schema.GridRange;
    rowOffset?: number;
    colOffset?: number;
    height?: number;
    width?: number;
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
 * Creates a banding for alternating cell backgrounds in a given hue [0,1]
 */
export function createBanding(hue: number, header = false, footer = false): GoogleAppsScript.Sheets.Schema.BandingProperties {
    const banding: GoogleAppsScript.Sheets.Schema.BandingProperties = {
        firstBandColor: { red: 0.98, green: 0.98, blue: 0.98 },
        secondBandColor: hslToColor({ h: hue, s: 1, l: 0.95 }),
    };
    if (header) banding.headerColor = hslToColor({ h: hue, s: 0.75, l: 0.2 });
    if (footer) banding.footerColor = hslToColor({ h: hue, s: 0.5, l: 0.7 });

    return banding;
}

/** Helper function to convert Sheets API Color to a Hex string */
export function colorToHex(color?: GoogleAppsScript.Sheets.Schema.Color, fallback = "#FFFFFF"): string {
    // If there's no color provided, default to white as requested
    if (!color) return fallback;

    const rbgColor: RGBColor = {
        r: color.red ?? 0,
        g: color.green ?? 0,
        b: color.blue ?? 0,
    };

    return rgbToHex(rbgColor);
}

/** Helper function to convert a HSL color into Google Sheets color */
function hslToColor(color: HSLColor): GoogleAppsScript.Sheets.Schema.Color {
    const rgbColor = hslToRgb(color);
    return {
        red: rgbColor.r,
        green: rgbColor.g,
        blue: rgbColor.b,
    };
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
            pasteOrientation: PasteOrientation.NORMAL,
        },
    };
}

interface BuildTranferRequestParams {
    destination?: GoogleAppsScript.Sheets.Schema.GridRange;
    data: GoogleAppsScript.Sheets.Schema.CellData[][];
    fields: string;
    adaptRange?: boolean;
}

/**
 * Generates batch upate `mergeCells` request.
 */
export function buildMergeCellsRequest(
    range: GoogleAppsScript.Sheets.Schema.GridRange | undefined,
    mergeType: MergeType,
): GoogleAppsScript.Sheets.Schema.Request | undefined {
    if (!range) return undefined;
    return {
        mergeCells: {
            range,
            mergeType,
        },
    };
}

/**
 * Generates batch update `addNamedRange` request.
 */
export function buildAddNamedRangeRequest<T extends NestedSheetSchema>(
    name: ExtractRangeNames<T>,
    range?: GoogleAppsScript.Sheets.Schema.GridRange,
): GoogleAppsScript.Sheets.Schema.Request | undefined {
    if (!range) return undefined;
    return {
        addNamedRange: {
            namedRange: {
                name,
                range,
            },
        },
    };
}

/**
 * Generates batch update requests to put `data` into the range defined by `destination`.
 * @param fields Mask to see what to copy.
 * @param adaptRange If true, add/remove rows and columns to adapt the range to the size of `data`
 */
export function buildTransferRequests({ destination, data, fields, adaptRange = false }: BuildTranferRequestParams): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    if (!destination) return requests;

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
                    range: { sheetId, dimension: Dimension.ROWS, startIndex: startRow + 1, endIndex: startRow + 1 + (dataRows - destRows) },
                    inheritFromBefore: true,
                },
            });
        } else if (dataRows < destRows) {
            requests.push({
                deleteDimension: {
                    range: { sheetId, dimension: Dimension.ROWS, startIndex: startRow + dataRows, endIndex: endRow },
                },
            });
        }

        // Adjust Columns
        if (dataCols > destCols) {
            requests.push({
                insertDimension: {
                    range: { sheetId, dimension: Dimension.COLUMNS, startIndex: startCol + 1, endIndex: startCol + 1 + (dataCols - destCols) },
                    inheritFromBefore: true,
                },
            });
        } else if (dataCols < destCols) {
            requests.push({
                deleteDimension: {
                    range: { sheetId, dimension: Dimension.COLUMNS, startIndex: startCol + dataCols, endIndex: endCol },
                },
            });
        }
    }

    const finalRowData: GoogleAppsScript.Sheets.Schema.RowData[] = [];
    for (let r = 0; r < finalRows; r++) {
        const rowValues: GoogleAppsScript.Sheets.Schema.CellData[] = [];
        const sourceRow = data[r] ?? [];

        for (let c = 0; c < finalCols; c++) {
            rowValues.push(sourceRow[c] ?? {});
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
