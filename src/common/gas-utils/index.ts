export * from "./color";
export * from "./gas-types";
export * from "./helpers";
export * from "./mapped-range";
export * from "./notation";
export * from "./parse";
export * from "./range";
export * from "./request";
export * from "./time";

export interface MappedNamedRange {
    range: GoogleAppsScript.Sheets.Schema.GridRange;
    sheet: GoogleAppsScript.Sheets.Schema.Sheet;
}

export interface ParsedSpreadsheet<T extends NestedSheetSchema> {
    sheets: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.Sheet>>;
    sheetNamedRanges: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.NamedRange[]>>;
    mappedRanges: Partial<Record<ExtractRangeNames<T>, MappedNamedRange>>;
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

export interface OffsetGridRangeProperties {
    origin: GoogleAppsScript.Sheets.Schema.GridRange;
    rowOffset?: number;
    colOffset?: number;
    height?: number;
    width?: number;
}

export interface ResizeRangeParams {
    target: MappedNamedRange;
    targetRows?: number;
    targetCols?: number;
    rowOffset?: number;
    colOffset?: number;
}

export interface RangeOperationResult {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    rowOffset: number;
    colOffset: number;
}
