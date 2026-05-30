import { type MergeType, PasteOrientation, type PasteType } from "../constants";
import type { ExtractRangeNames, MappedNamedRange, NestedSheetSchema } from ".";
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

/**
 * Generate batch update `addBanding` request.
 */
export function buildAddBandingRequest(
    range: GoogleAppsScript.Sheets.Schema.GridRange,
    bandingProperties: GoogleAppsScript.Sheets.Schema.BandingProperties,
): GoogleAppsScript.Sheets.Schema.Request {
    return {
        addBanding: {
            bandedRange: {
                range,
                rowProperties: bandingProperties,
            },
        },
    };
}

interface BuildUpdateCellsRequestParams {
    destination: GoogleAppsScript.Sheets.Schema.GridRange;
    data: GoogleAppsScript.Sheets.Schema.CellData[][];
    fields: string;
}

/**
 * Generates a batch update requests to put `data` into the range defined by `destination`.
 * @param fields Mask to see what to copy.
 */
export function buildUpdateCellsRequest({ destination, data, fields }: BuildUpdateCellsRequestParams): GoogleAppsScript.Sheets.Schema.Request | undefined {
    const startRow = destination.startRowIndex ?? 0;
    const endRow = destination.endRowIndex ?? 0;
    const startCol = destination.startColumnIndex ?? 0;
    const endCol = destination.endColumnIndex ?? 0;

    const finalRows = endRow - startRow;
    const finalCols = endCol - startCol;

    if (!finalRows || !finalCols) return undefined;
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
    return {
        updateCells: {
            range: destination,
            rows: finalRowData,
            fields: fields,
        },
    };
}

interface BuildTransferRequestParams {
    destination: MappedNamedRange;
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
        destination.namedRange.range = offsetGridRange({
            origin: destination.namedRange.range,
            rowOffset: currentRowOffset,
            colOffset: currentColOffset,
        });
    }

    // destination.range is now fully accurate
    const sheetId = destination.namedRange.range.sheetId ?? 0;
    const finalStartRow = destination.namedRange.range.startRowIndex ?? 0;
    const finalEndRow = destination.namedRange.range.endRowIndex ?? 0;
    const finalStartCol = destination.namedRange.range.startColumnIndex ?? 0;
    const finalEndCol = destination.namedRange.range.endColumnIndex ?? 0;

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
