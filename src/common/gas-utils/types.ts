export interface StrictNameRange extends GoogleAppsScript.Sheets.Schema.NamedRange {
    name: string;
    range: GoogleAppsScript.Sheets.Schema.GridRange;
}

export interface MappedNamedRange {
    namedRange: StrictNameRange;
    sheet: GoogleAppsScript.Sheets.Schema.Sheet;
}

export interface ParsedSpreadsheet<T extends NestedSheetSchema> {
    mappedSheets: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.Sheet>>;
    mappedSheetNamedRanges: Partial<Record<ExtractSheetNames<T>, StrictNameRange[]>>;
    mappedRanges: Partial<Record<ExtractRangeNames<T>, MappedNamedRange>>;
    dynamicMappedRanges: Partial<Record<ExtractDynamicRangeKeys<T>, MappedNamedRange[]>>;
}

export interface NestedSheetSchema {
    readonly sheets: Record<
        string,
        {
            readonly sheetName: string;
            readonly ranges?: Record<string, string>;
            readonly dynamicRanges?: Record<string, string>;
        }
    >;
}

export type ExtractSheetNames<T extends NestedSheetSchema> = T["sheets"][keyof T["sheets"]]["sheetName"];

export type ExtractRangeNames<T extends NestedSheetSchema> = {
    [K in keyof T["sheets"]]: T["sheets"][K] extends { ranges?: infer R } ? (R extends undefined ? never : R[keyof R]) : never;
}[keyof T["sheets"]];

export type ExtractDynamicRangeKeys<T extends NestedSheetSchema> = {
    [K in keyof T["sheets"]]: T["sheets"][K] extends { dynamicRanges?: infer D } ? (D extends undefined ? never : keyof D) : never;
}[keyof T["sheets"]];

export interface OffsetGridRangeProperties {
    readonly origin: GoogleAppsScript.Sheets.Schema.GridRange;
    readonly sheetId?: number;
    readonly rowOffset?: number;
    readonly colOffset?: number;
    readonly height?: number;
    readonly width?: number;
}

export interface ResizeRangeParams {
    target: MappedNamedRange;
    targetRows?: number;
    targetCols?: number;
    rowOffset?: number;
    colOffset?: number;
    rowBehavior?: RangeBehavior;
    colBehavior?: RangeBehavior;
}

export interface RangeOperationResult {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    rowOffset: number;
    colOffset: number;
}

export enum RangeBehavior {
    IGNORE = "IGNORE",
    INSERT_DELETE = "INSERT_DELETE",
    MODIFY_RANGE = "MODIFY_RANGE",
}

export enum FileType {
    SETUP = "MontessoriChacalaSchoolGroupSetup",
    REPORT = "MontessoriChacalaSchoolReport",
}

export enum Urls {
    MEDIA_SERVER = "https://media.githubusercontent.com/media/YamanquiChacala/GeneradorDeReportes/refs/heads/main/",
}

export enum Colors {
    LOGO_OSCURO = "#34666A",
    LOGO_CLARO = "#159A5E",
    ORANGE = "#EA4335",
    LANGUAGE = "#c9daf8",
    SCIENCE = "#fce5cd",
    NATURE = "#d9ead3",
    HUMANITIES = "#ead1dc",
}
