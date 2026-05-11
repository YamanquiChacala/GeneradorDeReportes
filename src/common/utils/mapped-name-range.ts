export interface MappedNamedRange {
    range: GoogleAppsScript.Sheets.Schema.GridRange;
    sheet: GoogleAppsScript.Sheets.Schema.Sheet;
}

export interface NestedSheetSchema {
    readonly sheets: Record<
        string,
        {
            readonly sheetName: string;
            readonly ranges?: Record<string, string>;
        }
    >;
}

export type ExtractSheetNames<T extends NestedSheetSchema> = T["sheets"][keyof T["sheets"]]["sheetName"];

export type ExtractRangeNames<T extends NestedSheetSchema> = {
    [K in keyof T["sheets"]]: T["sheets"][K] extends { ranges?: infer R } ? (R extends undefined ? never : R[keyof R]) : never;
}[keyof T["sheets"]];

interface GetCellParams {
    mappedRange: MappedNamedRange | undefined;
    rowOffset?: number;
    columnOffset?: number;
}

export const MappedNamedRange = {
    getCellDataArray(mappedRange: MappedNamedRange | undefined, unboundRows = false, unboundColumns = false): GoogleAppsScript.Sheets.Schema.CellData[][] {
        const { range, sheet } = mappedRange ?? {};

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
    },

    getCellData({ mappedRange, rowOffset, columnOffset }: GetCellParams): GoogleAppsScript.Sheets.Schema.CellData | undefined {
        const absoluteRowIndex = (mappedRange?.range.startRowIndex ?? 0) + (rowOffset ?? 0);
        const absoluteColIndex = (mappedRange?.range.startColumnIndex ?? 0) + (columnOffset ?? 0);

        const endRow = mappedRange?.range.endRowIndex;
        const endColumn = mappedRange?.range.endColumnIndex;

        if ((endRow && absoluteRowIndex >= endRow) || (endColumn && absoluteColIndex >= endColumn)) return undefined;

        const sheetData = mappedRange?.sheet.data;
        if (!sheetData) return undefined;

        for (const gridData of sheetData) {
            const startRow = gridData.startRow ?? 0;
            const startColumn = gridData.startColumn ?? 0;

            const relativeRow = absoluteRowIndex - startRow;
            const relativeCol = absoluteColIndex - startColumn;

            if (relativeRow >= 0 && gridData.rowData && relativeRow < gridData.rowData.length) {
                const rowData = gridData.rowData[relativeRow];

                if (relativeCol >= 0 && rowData?.values && relativeCol < rowData.values.length) {
                    return rowData.values[relativeCol];
                }
            }
        }

        return undefined;
    },

    getCellEffectiveValue(args: GetCellParams): GoogleAppsScript.Sheets.Schema.ExtendedValue | undefined {
        const cellData = MappedNamedRange.getCellData(args);

        return cellData?.effectiveValue;
    },

    getCellText(args: GetCellParams): string | undefined {
        const effectiveValue = MappedNamedRange.getCellEffectiveValue(args);

        return effectiveValue?.stringValue;
    },

    getCellNumber(args: GetCellParams): number | undefined {
        const effectiveValue = MappedNamedRange.getCellEffectiveValue(args);

        return effectiveValue?.numberValue;
    },

    getCellUnixEpoch(args: GetCellParams): number | undefined {
        const cellNumber = MappedNamedRange.getCellNumber(args);

        if (cellNumber == null) return undefined;

        const msPerDay = 24 * 60 * 60 * 1000;
        const sheetsEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
        return sheetsEpoch + cellNumber * msPerDay;
    },
} as const;

export interface ParsedSpreadsheet<T extends NestedSheetSchema> {
    sheets: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.Sheet>>;
    sheetNamedRanges: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.NamedRange[]>>;
    namedRanges: Partial<Record<ExtractRangeNames<T>, MappedNamedRange>>;
}

/**
 * Parses a Spreadsheet coming from Sheets API, so that the Sheets and named ranges are easy to find and work with.
 */
export function parseSpreadsheet<T extends NestedSheetSchema>(spreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet | undefined, schema: T): ParsedSpreadsheet<T> {
    const mappedSheets: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.Sheet>> = {};
    const mappedSheetNamedRanges: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.NamedRange[]>> = {};
    const mappedRanges: Partial<Record<ExtractRangeNames<T>, MappedNamedRange>> = {};

    if (!spreadsheet?.sheets) return { sheets: mappedSheets, sheetNamedRanges: mappedSheetNamedRanges, namedRanges: mappedRanges };

    const allowedSheetNames = new Set<string>();
    const allowedRangeNames = new Set<string>();

    for (const key of Object.keys(schema.sheets)) {
        const sheetConfig = schema.sheets[key];
        if (sheetConfig) allowedSheetNames.add(sheetConfig.sheetName);

        if (sheetConfig?.ranges) {
            for (const rangeKey of Object.keys(sheetConfig.ranges)) {
                if (sheetConfig.ranges[rangeKey]) allowedRangeNames.add(sheetConfig.ranges[rangeKey]);
            }
        }
    }

    const sheetIdLookup: Record<number, GoogleAppsScript.Sheets.Schema.Sheet> = {};

    for (const sheet of spreadsheet.sheets) {
        sheetIdLookup[sheet.properties?.sheetId ?? 0] = sheet;
        const sheetTitle = sheet.properties?.title;
        if (sheetTitle != null && allowedSheetNames.has(sheetTitle)) {
            mappedSheets[sheetTitle as ExtractSheetNames<T>] = sheet;
            mappedSheetNamedRanges[sheetTitle as ExtractSheetNames<T>] = spreadsheet.namedRanges?.filter(
                (namedRange) => namedRange.range?.sheetId === sheet.properties?.sheetId,
            );
        }
    }

    if (spreadsheet.namedRanges) {
        for (const namedRange of spreadsheet.namedRanges) {
            if (namedRange.name == null || namedRange.range == null) continue;
            const linkedSheet = sheetIdLookup[namedRange.range.sheetId ?? 0];

            if (linkedSheet && allowedRangeNames.has(namedRange.name)) {
                mappedRanges[namedRange.name as ExtractRangeNames<T>] = {
                    range: namedRange.range,
                    sheet: linkedSheet,
                };
            }
        }
    }

    return { sheets: mappedSheets, sheetNamedRanges: mappedSheetNamedRanges, namedRanges: mappedRanges };
}
