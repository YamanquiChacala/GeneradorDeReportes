import { Dimension, type MergeType, PasteOrientation, type PasteType } from "../gas-enums";
import type { ExtractRangeNames, MappedNamedRange, NestedSheetSchema } from "../utils/mapped-name-range";
import { offsetGridRange, type RangeOperationResult, resizeMappedRange } from ".";

/**
 * Generates batch update `copyPaste` request to copy data from `origin` into `destination` ranges.
 */
export function buildCopyPasteRequest(
    source: GoogleAppsScript.Sheets.Schema.GridRange,
    destination: GoogleAppsScript.Sheets.Schema.GridRange,
    pasteType: PasteType,
): GoogleAppsScript.Sheets.Schema.Request {
    return {
        copyPaste: {
            source,
            destination,
            pasteType,
            pasteOrientation: PasteOrientation.NORMAL,
        },
    };
}

/**
 * Generates batch upate `mergeCells` request.
 */
export function buildMergeCellsRequest(range: GoogleAppsScript.Sheets.Schema.GridRange, mergeType: MergeType): GoogleAppsScript.Sheets.Schema.Request {
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
    range: GoogleAppsScript.Sheets.Schema.GridRange,
): GoogleAppsScript.Sheets.Schema.Request {
    return {
        addNamedRange: {
            namedRange: {
                name,
                range,
            },
        },
    };
}

interface BuildTransferRequestParams {
    destination?: MappedNamedRange;
    data: GoogleAppsScript.Sheets.Schema.CellData[][];
    fields: string;
    adaptRange?: boolean;
    rowOffset?: number;
    colOffset?: number;
}

export function buildTransferRequests({ destination, data, fields, adaptRange = false, rowOffset = 0, colOffset = 0 }: BuildTransferRequestParams): RangeOperationResult {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    let currentRowOffset = rowOffset;
    let currentColOffset = colOffset;

    if (!destination) {
        return { requests, rowOffset: currentRowOffset, colOffset: currentColOffset };
    }

    const dataRows = data.length;
    const dataCols = dataRows > 0 ? Math.max(...data.map((row) => row.length)) : 0;

    if (adaptRange) {
        const resizeResult = resizeMappedRange({
            target: destination,
            targetRows: dataRows,
            targetCols: dataCols,
            rowOffset: currentRowOffset,
            colOffset: currentColOffset,
        });

        requests.push(...resizeResult.requests);
        currentRowOffset = resizeResult.rowOffset;
        currentColOffset = resizeResult.colOffset;
    } else {
        // If not adapting, we still apply offsets so the destination is accurate for writing
        destination.range = offsetGridRange({
            origin: destination.range,
            rowOffset: currentRowOffset,
            colOffset: currentColOffset,
        });
    }

    // destination.range is now fully accurate
    const sheetId = destination.range.sheetId ?? 0;
    const finalStartRow = destination.range.startRowIndex ?? 0;
    const finalEndRow = destination.range.endRowIndex ?? 0;
    const finalStartCol = destination.range.startColumnIndex ?? 0;
    const finalEndCol = destination.range.endColumnIndex ?? 0;

    const finalRows = finalEndRow - finalStartRow;
    const finalCols = finalEndCol - finalStartCol;

    // Only attempt to build data and write if the range still exists (wasn't destroyed)
    if (finalRows > 0 && finalCols > 0) {
        const finalRowData: GoogleAppsScript.Sheets.Schema.RowData[] = [];
        for (let r = 0; r < finalRows; r++) {
            const rowValues: GoogleAppsScript.Sheets.Schema.CellData[] = [];
            const sourceRow = data[r] ?? [];

            for (let c = 0; c < finalCols; c++) {
                rowValues.push(sourceRow[c] ?? {});
            }

            finalRowData.push({ values: rowValues });
        }

        requests.push({
            updateCells: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: finalStartRow,
                    endRowIndex: finalEndRow,
                    startColumnIndex: finalStartCol,
                    endColumnIndex: finalEndCol,
                },
                rows: finalRowData,
                fields: fields,
            },
        });
    }

    return {
        requests,
        rowOffset: currentRowOffset,
        colOffset: currentColOffset,
    };
}

interface BuildTranferRequestParamsBackup {
    destination?: GoogleAppsScript.Sheets.Schema.GridRange;
    data: GoogleAppsScript.Sheets.Schema.CellData[][];
    fields: string;
    adaptRange?: boolean;
}

/**
 * Generates batch update requests to put `data` into the range defined by `destination`.
 * @param fields Mask to see what to copy.
 * @param adaptRange If true, add/remove rows and columns to adapt the range to the size of `data`
 */
export function buildTransferRequestsBackup({
    destination,
    data,
    fields,
    adaptRange = false,
}: BuildTranferRequestParamsBackup): GoogleAppsScript.Sheets.Schema.Request[] {
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
