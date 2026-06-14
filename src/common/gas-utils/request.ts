import { getRandomId } from "../utils";
import { type MergeType, PasteOrientation, type PasteType } from "./api-types";
import { createRequiredGetter } from "./helpers";
import { resizeMappedRange } from "./mapped-range";
import { offsetGridRange } from "./range";
import {
    type ExtractDynamicRangeNames,
    type ExtractRangeNames,
    type ExtractSheetNames,
    type MappedNamedRange,
    type NestedSheetSchema,
    type ParsedSpreadsheet,
    RangeBehavior,
    type RangeOperationResult,
    type StrictNameRange,
} from "./types";

/**
 * Generates batch update `copyPaste` request to copy data from `origin` into `destination` ranges.
 */
export function buildCopyPasteRequest(
    source: Readonly<GoogleAppsScript.Sheets.Schema.GridRange>,
    destination: Readonly<GoogleAppsScript.Sheets.Schema.GridRange>,
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
export function buildMergeCellsRequest(range: Readonly<GoogleAppsScript.Sheets.Schema.GridRange>, mergeType: MergeType): GoogleAppsScript.Sheets.Schema.Request {
    return {
        mergeCells: {
            range,
            mergeType,
        },
    };
}

/**
 * Generates batch update `unmergeCells` request.
 */
export function buildUnmergeCellsRequest(range: Readonly<GoogleAppsScript.Sheets.Schema.GridRange>): GoogleAppsScript.Sheets.Schema.Request {
    return {
        unmergeCells: {
            range,
        },
    };
}

/**
 * Generate batch update `addBanding` request.
 */
export function buildAddBandingRequest(
    range: Readonly<GoogleAppsScript.Sheets.Schema.GridRange>,
    bandingProperties: Readonly<GoogleAppsScript.Sheets.Schema.BandingProperties>,
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

/**
 * Generates batch updates to either modify or create a protectedRange.
 */
export function buildProtectSheetRequest<T extends NestedSheetSchema>(
    parsedData: ParsedSpreadsheet<T>,
    sheetName: ExtractSheetNames<T>,
    unprotectedRanges?: GoogleAppsScript.Sheets.Schema.GridRange[],
): GoogleAppsScript.Sheets.Schema.Request {
    const getSheet = createRequiredGetter(parsedData.mappedSheets, "hoja para proteger");
    const sheet = getSheet(sheetName);

    return buildSingleSheetProtectRequest(sheet, parsedData.usedIds, unprotectedRanges);
}

/**
 * Generates batch updates to add or modify protectedRanges of a series of sheets.
 */
export function buildProtectExtraSheetRequests<T extends NestedSheetSchema>(
    parsedData: ParsedSpreadsheet<T>,
    unprotectedRanges?: GoogleAppsScript.Sheets.Schema.GridRange[],
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    for (const sheet of parsedData.extraSheets) {
        const sheetId = sheet.properties?.sheetId ?? 0;
        const sheetUnprotectedRanges = unprotectedRanges?.map((range) => offsetGridRange({ origin: range, sheetId }));
        requests.push(buildSingleSheetProtectRequest(sheet, parsedData.usedIds, sheetUnprotectedRanges));
    }

    return requests;
}

/**
 * Helper function to generate a protect request for a single sheet.
 * Mutates the sheet.protectedRanges property in place.
 */
function buildSingleSheetProtectRequest(
    sheet: GoogleAppsScript.Sheets.Schema.Sheet,
    usedIds: Set<number>,
    unprotectedRanges?: GoogleAppsScript.Sheets.Schema.GridRange[],
): GoogleAppsScript.Sheets.Schema.Request {
    const sheetId = sheet.properties?.sheetId;
    const sheetName = sheet.properties?.title;
    const protectedRanges = sheet.protectedRanges;

    const newProtectedRange: GoogleAppsScript.Sheets.Schema.ProtectedRange = {
        range: { sheetId },
        description: sheetName,
        warningOnly: false,
        unprotectedRanges,
    };

    let request: GoogleAppsScript.Sheets.Schema.Request;

    if (!protectedRanges || protectedRanges.length === 0) {
        newProtectedRange.protectedRangeId = getRandomId(usedIds);
        request = { addProtectedRange: { protectedRange: newProtectedRange } };
    } else if (protectedRanges.length !== 1) {
        throw new Error("Demasiados rangos protegidos en la hoja.");
    } else {
        newProtectedRange.protectedRangeId = protectedRanges[0]?.protectedRangeId;
        request = { updateProtectedRange: { protectedRange: newProtectedRange } };
    }

    // Update the sheet object
    sheet.protectedRanges = [newProtectedRange];

    return request;
}

interface BuildUpdateSheetPropertiesParams {
    readonly sheetId: number;
    readonly index?: number;
    readonly hidden?: boolean;
    readonly rowCount?: number;
    readonly columnCount?: number;
    readonly frozenRowCount?: number;
    readonly frozenColumnCount?: number;
    readonly hideGridlines?: boolean;
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
    readonly destination: GoogleAppsScript.Sheets.Schema.GridRange;
    readonly data: GoogleAppsScript.Sheets.Schema.CellData[][];
    readonly fields: string;
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
    readonly destination: MappedNamedRange;
    readonly data: GoogleAppsScript.Sheets.Schema.CellData[][];
    readonly fields: string;
    readonly rowBehavior?: RangeBehavior;
    readonly colBehavior?: RangeBehavior;
    readonly rowOffset?: number;
    readonly colOffset?: number;
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

interface BaseAddSheetParams<T extends NestedSheetSchema> {
    readonly parsedData: ParsedSpreadsheet<T>;
    readonly sourceSheetTitle: ExtractSheetNames<T>;
    readonly insertSheetIndex: number;
}

// Option A: Single sheet tied to the schema
interface AddSchemaSheetParams<T extends NestedSheetSchema> extends BaseAddSheetParams<T> {
    readonly schemaSheetName: ExtractSheetNames<T>;
    readonly multipleSheetNames?: never;
}

// Option B: Multiple non-schema sheets
interface AddExtraSheetsParams<T extends NestedSheetSchema> extends BaseAddSheetParams<T> {
    readonly schemaSheetName?: never;
    readonly multipleSheetNames: string[];
}

type AddNewSheetParams<T extends NestedSheetSchema> = AddSchemaSheetParams<T> | AddExtraSheetsParams<T>;

/**
 * Duplicates a template sheet.
 */
export function addNewSheet<T extends NestedSheetSchema>({
    parsedData,
    sourceSheetTitle,
    insertSheetIndex,
    schemaSheetName,
    multipleSheetNames,
}: AddNewSheetParams<T>): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newSheetIds: number[] } {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const newSheetIds: number[] = [];

    const getSheetNamedRanges = createRequiredGetter(parsedData.mappedSheetNamedRanges, "lista de rangos en hoja");
    const getSheet = createRequiredGetter(parsedData.mappedSheets, "hoja");

    const sourceSheetId = getSheet(sourceSheetTitle).properties?.sheetId ?? 0;

    // Temporarly remove named ranges before duplicating template
    const templateNamedRanges = getSheetNamedRanges(sourceSheetTitle);
    for (const namedRange of templateNamedRanges) {
        requests.push({ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } });
    }

    let sheetNamesToCreate: string[] = [];
    let isSchemaBound = false;

    if (schemaSheetName) {
        sheetNamesToCreate = [schemaSheetName];
        isSchemaBound = true;
    } else {
        // biome-ignore lint/style/noNonNullAssertion: if `schemaSheetName` is not defined, `multipleSheetNames` must be.
        sheetNamesToCreate = multipleSheetNames!;
    }

    // Mass duplication
    for (const [index, targetSheetName] of sheetNamesToCreate.entries()) {
        const newSheetId = getRandomId(parsedData.usedIds);
        newSheetIds.push(newSheetId);

        const targetIndex = insertSheetIndex + index;

        const newSheet: GoogleAppsScript.Sheets.Schema.Sheet = {
            properties: {
                sheetId: newSheetId,
                title: targetSheetName,
                index: targetIndex,
            },
        };

        if (isSchemaBound) {
            const targetName = targetSheetName as ExtractSheetNames<T>;
            // If the sheet already exists, delete it and recreate it.
            const oldNamedRanges = parsedData.mappedSheetNamedRanges[targetName] ?? [];
            for (const oldRange of oldNamedRanges) {
                requests.push({ deleteNamedRange: { namedRangeId: oldRange.namedRangeId } });
            }
            const oldSheet = parsedData.mappedSheets[targetName];
            if (oldSheet) requests.push({ deleteSheet: { sheetId: oldSheet.properties?.sheetId } });
            // Create new one
            parsedData.mappedSheets[targetName] = newSheet;
            parsedData.mappedSheetNamedRanges[targetName] = [];
        } else {
            parsedData.extraSheets.push(newSheet);
        }

        requests.push({
            duplicateSheet: {
                sourceSheetId,
                newSheetName: targetSheetName,
                insertSheetIndex: targetIndex,
                newSheetId,
            },
        });
    }

    // Restor named ranges to template
    for (const namedRange of templateNamedRanges) requests.push({ addNamedRange: { namedRange } });

    return { requests, newSheetIds };
}

interface BaseAddNamedRangeParams<T extends NestedSheetSchema> {
    readonly parsedData: ParsedSpreadsheet<T>;
    readonly sheetTitle: ExtractSheetNames<T>;
    readonly gridRange: GoogleAppsScript.Sheets.Schema.GridRange;
}

interface AddStaticNamedRangeParams<T extends NestedSheetSchema> extends BaseAddNamedRangeParams<T> {
    readonly staticRangeKey: ExtractRangeNames<T>;
    readonly rangeName?: never;
    readonly dynamicRangeKey?: never;
}

interface AddDynamicNamedRangeParams<T extends NestedSheetSchema> extends BaseAddNamedRangeParams<T> {
    readonly staticRangeKey?: never;
    readonly rangeName: string;
    readonly dynamicRangeKey: ExtractDynamicRangeNames<T>;
}

type AddNamedRangeParams<T extends NestedSheetSchema> = AddStaticNamedRangeParams<T> | AddDynamicNamedRangeParams<T>;

/**
 * Adds a new named range, both to the spreadsheet, and to memory.
 */
export function addNewNamedRange<T extends NestedSheetSchema>({
    parsedData,
    sheetTitle,
    gridRange,
    staticRangeKey,
    rangeName,
    dynamicRangeKey,
}: AddNamedRangeParams<T>): GoogleAppsScript.Sheets.Schema.Request {
    const getSheet = createRequiredGetter(parsedData.mappedSheets, "Adding range name to sheet");
    const sheet = getSheet(sheetTitle);

    gridRange.sheetId = sheet.properties?.sheetId;

    const finalRangeName = rangeName ?? staticRangeKey;

    const newStricNamedRange: StrictNameRange = {
        namedRangeId: Utilities.getUuid(),
        name: finalRangeName,
        range: gridRange,
    };

    const newMappedNamedRange: MappedNamedRange = {
        namedRange: newStricNamedRange,
        sheet: sheet,
    };

    if (!parsedData.mappedSheetNamedRanges[sheetTitle]) {
        parsedData.mappedSheetNamedRanges[sheetTitle] = [];
    }
    parsedData.mappedSheetNamedRanges[sheetTitle].push(newStricNamedRange);

    if (staticRangeKey) {
        const key = staticRangeKey as ExtractRangeNames<T>;
        parsedData.mappedRanges[key] = newMappedNamedRange;
    } else {
        const key = dynamicRangeKey as ExtractDynamicRangeNames<T>;
        if (!parsedData.dynamicMappedRanges[key]) {
            parsedData.dynamicMappedRanges[key] = [];
        }
        parsedData.dynamicMappedRanges[key].push(newMappedNamedRange);
    }

    return { addNamedRange: { namedRange: newStricNamedRange } };
}
