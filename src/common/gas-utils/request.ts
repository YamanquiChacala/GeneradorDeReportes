import { type ExtractRangeNames, type MappedNamedRange, type NestedSheetSchema, RangeBehavior, type RangeOperationResult, resizeMappedRange } from ".";
import { type MergeType, PasteOrientation, type PasteType } from "./api-types";

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
    namedRangeId?: string,
): GoogleAppsScript.Sheets.Schema.Request {
    return {
        addNamedRange: {
            namedRange: {
                namedRangeId,
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

interface BuildUpdateSheetPropertiesParams {
    sheetId: number;
    index?: number;
    hidden?: boolean;
    rowCount?: number;
    columnCount?: number;
    frozenRowCount?: number;
    frozenColumnCount?: number;
    hideGridlines?: boolean;
}

/**
 * Generates batch update `updateSheetProperties`
 */
export function buildUpdateSheetPropertiesRequest({
    sheetId,
    index,
    hidden,
    rowCount,
    columnCount,
    frozenRowCount,
    frozenColumnCount,
    hideGridlines = true,
}: BuildUpdateSheetPropertiesParams): GoogleAppsScript.Sheets.Schema.Request {
    const properties: GoogleAppsScript.Sheets.Schema.SheetProperties = { sheetId };
    const fields: string[] = [];

    if (index !== undefined) {
        properties.index = index;
        fields.push("index");
    }
    if (hidden !== undefined) {
        properties.hidden = hidden;
        fields.push("hidden");
    }

    const gridProperties: GoogleAppsScript.Sheets.Schema.GridProperties = {};
    let hasGridProperties = false;

    if (rowCount !== undefined) {
        gridProperties.rowCount = rowCount;
        fields.push("gridProperties.rowCount");
        hasGridProperties = true;
    }
    if (columnCount !== undefined) {
        gridProperties.columnCount = columnCount;
        fields.push("gridProperties.columnCount");
        hasGridProperties = true;
    }
    if (frozenRowCount !== undefined) {
        gridProperties.frozenRowCount = frozenRowCount;
        fields.push("gridProperties.frozenRowCount");
        hasGridProperties = true;
    }
    if (frozenColumnCount !== undefined) {
        gridProperties.frozenColumnCount = frozenColumnCount;
        fields.push("gridProperties.frozenColumnCount");
        hasGridProperties = true;
    }
    if (hideGridlines) {
        gridProperties.hideGridlines = hideGridlines;
        fields.push("gridProperties.hideGridlines");
        hasGridProperties = true;
    }

    if (hasGridProperties) {
        properties.gridProperties = gridProperties;
    }

    return {
        updateSheetProperties: {
            properties,
            fields: fields.join(","),
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
    rowBehavior?: RangeBehavior;
    colBehavior?: RangeBehavior;
    rowOffset?: number;
    colOffset?: number;
}

export function buildTransferRequests({
    destination,
    data,
    fields,
    rowBehavior = RangeBehavior.IGNORE,
    colBehavior = RangeBehavior.IGNORE,
    rowOffset = 0,
    colOffset = 0,
}: BuildTransferRequestParams): RangeOperationResult {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const dataRows = data.length;
    const dataCols = dataRows > 0 ? Math.max(...data.map((row) => row.length)) : 0;

    const resizeResult = resizeMappedRange({
        target: destination,
        targetRows: dataRows,
        targetCols: dataCols,
        rowOffset,
        colOffset,
        rowBehavior,
        colBehavior,
    });

    requests.push(...resizeResult.requests);

    const updateRequest = buildUpdateCellsRequest({
        destination: destination.namedRange.range,
        data,
        fields,
    });

    if (updateRequest) requests.push(updateRequest);

    return {
        requests,
        rowOffset: resizeResult.rowOffset,
        colOffset: resizeResult.colOffset,
    };
}
