export interface StrictNameRange extends GoogleAppsScript.Sheets.Schema.NamedRange {
    readonly namedRangeId: string;
    readonly name: string;
    range: GoogleAppsScript.Sheets.Schema.GridRange; // Explicitly not readonly, since it's mutated.
}

export interface MappedNamedRange {
    readonly namedRange: StrictNameRange;
    readonly sheet: GoogleAppsScript.Sheets.Schema.Sheet;
}

export interface ParsedSpreadsheet<T extends NestedSheetSchema> {
    readonly mappedSheets: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.Sheet>>;
    readonly mappedSheetNamedRanges: Partial<Record<ExtractSheetNames<T>, StrictNameRange[]>>;
    readonly mappedRanges: Partial<Record<ExtractRangeNames<T>, MappedNamedRange>>;
    readonly dynamicMappedRanges: Partial<Record<ExtractDynamicRangeNames<T>, MappedNamedRange[]>>;
    readonly extraSheets: GoogleAppsScript.Sheets.Schema.Sheet[];
    readonly usedIds: Set<number>;
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

export type ExtractDynamicRangeNames<T extends NestedSheetSchema> = {
    [K in keyof T["sheets"]]: T["sheets"][K] extends { dynamicRanges?: infer R } ? (R extends undefined ? never : R[keyof R]) : never;
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
    readonly target: MappedNamedRange;
    readonly targetRows?: number;
    readonly targetCols?: number;
    readonly rowOffset?: number;
    readonly colOffset?: number;
    readonly rowBehavior?: RangeBehavior;
    readonly colBehavior?: RangeBehavior;
}

export interface RangeOperationResult {
    readonly requests: GoogleAppsScript.Sheets.Schema.Request[];
    readonly rowOffset: number;
    readonly colOffset: number;
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

export enum InputType {
    STRING = "string",
    NUMBER = "number",
    BOOLEAN = "boolean",
    DATE = "date",
    TIME = "time",
    ARRAY = "array",
}

export enum ParamType {
    STRING = "string",
    NUMBER = "number",
    BOOLEAN = "boolean",
}
