type InputType = "string" | "number" | "boolean" | "date" | "time" | "array";

export type MappedInput<T extends InputType> = T extends "string"
    ? string
    : T extends "number"
      ? number
      : T extends "boolean"
        ? boolean
        : T extends "date"
          ? number
          : T extends "time"
            ? { hours: number; minutes: number }
            : T extends "array"
              ? string[]
              : never;

type ParamType = "string" | "number" | "boolean";

type MappedParam<T extends ParamType> = T extends "string" ? string : T extends "number" ? number : T extends "boolean" ? boolean : never;

type GASFormInputs = GoogleAppsScript.Addons.CommonEventObject["formInputs"];

export enum PasteType {
    PASTE_NORMAL = "PASTE_NORMAL",
    PASTE_VALUES = "PASTE_VALUES",
    PASTE_FORMAT = "PASTE_FORMAT",
    PASTE_NO_BORDERS = "PASTE_NO_BORDERS",
    PASTE_FORMULA = "PASTE_FORMULA",
    PASTE_DATA_VALIDATION = "PASTE_DATA_VALIDATION",
    PASTE_CONDITIONAL_FORMATTING = "PASTE_CONDITIONAL_FORMATTING",
}

export interface MappedNamedRange {
    range: GoogleAppsScript.Sheets.Schema.GridRange;
    sheet: GoogleAppsScript.Sheets.Schema.Sheet;
}

interface GetCellParams {
    mappedRange: MappedNamedRange | undefined;
    row?: number;
    column?: number;
}

interface CopyPasteParams {
    mappedRange: MappedNamedRange | undefined;
    destinationSheetId: number;
    destinationStartRow: number;
    destinationStartColumn: number;
    pasteType: PasteType;
    offsetRow?: number;
    offsetColumn?: number;
    height?: number;
    width?: number;
}

export const MappedNamedRange = {
    getCellData({ mappedRange, row, column }: GetCellParams): GoogleAppsScript.Sheets.Schema.CellData | undefined {
        const rowIndex = (mappedRange?.range.startRowIndex ?? 0) + (row ?? 0);
        const colIndex = (mappedRange?.range.startColumnIndex ?? 0) + (column ?? 0);

        const endRow = mappedRange?.range.endRowIndex;
        const endColumn = mappedRange?.range.endColumnIndex;

        if ((endRow && rowIndex >= endRow) || (endColumn && colIndex >= endColumn)) return undefined;

        return mappedRange?.sheet.data?.[0]?.rowData?.[rowIndex]?.values?.[colIndex];
    },

    getCellDisplay({ mappedRange, row, column }: GetCellParams): string | undefined {
        const cellData = MappedNamedRange.getCellData({ mappedRange, row, column });

        return cellData?.formattedValue;
    },

    getCellNumber({ mappedRange, row, column }: GetCellParams): number | undefined {
        const cellData = MappedNamedRange.getCellData({ mappedRange, row, column });

        return cellData?.effectiveValue?.numberValue;
    },

    getCellUnixEpoch({ mappedRange, row, column }: GetCellParams): number | undefined {
        const cellNumber = MappedNamedRange.getCellNumber({ mappedRange, row, column });

        if (cellNumber == null) return undefined;

        const msPerDay = 24 * 60 * 60 * 1000;
        const sheetsEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
        return sheetsEpoch + cellNumber * msPerDay;
    },

    buildCopyPasteRequest({
        mappedRange,
        destinationSheetId,
        destinationStartRow,
        destinationStartColumn,
        pasteType,
        offsetRow,
        offsetColumn,
        height,
        width,
    }: CopyPasteParams): GoogleAppsScript.Sheets.Schema.Request | undefined {
        if (!mappedRange) return undefined;
        const srcStartRow = (mappedRange.range.startRowIndex ?? 0) + (offsetRow ?? 0);
        const srcStartColumn = (mappedRange.range.startColumnIndex ?? 0) + (offsetColumn ?? 0);

        const endRow = mappedRange.range.endRowIndex;
        const endColumn = mappedRange.range.endColumnIndex;

        const finalHeight = height ?? (endRow != null ? endRow - srcStartRow : undefined);
        const finalWidth = width ?? (endColumn != null ? endColumn - srcStartColumn : undefined);

        if (!finalHeight || !finalWidth) return undefined;

        if ((endRow != null && endRow < srcStartRow + finalHeight) || (endColumn != null && endColumn < srcStartColumn + finalWidth)) return undefined;

        return {
            copyPaste: {
                source: {
                    sheetId: mappedRange.range.sheetId ?? 0,
                    startRowIndex: srcStartRow,
                    endRowIndex: srcStartRow + finalHeight,
                    startColumnIndex: srcStartColumn,
                    endColumnIndex: srcStartColumn + finalWidth,
                },
                destination: {
                    sheetId: destinationSheetId,
                    startRowIndex: destinationStartRow,
                    endRowIndex: destinationStartRow + finalHeight,
                    startColumnIndex: destinationStartColumn,
                    endColumnIndex: destinationStartColumn + finalWidth,
                },
                pasteType: pasteType,
                pasteOrientation: "NORMAL",
            },
        };
    },
} as const;

/**
 * Helper function to define a type configuration linking a string to a {string, type}
 */
export function defineRangesDataConfig<T extends Record<string, { range: string; type: InputType }>>(fields: T) {
    return fields;
}

/**
 * Helper function to creating Form Input schemas for Cards and Callbacks.
 */
export function defineInputsSchema<T extends Record<string, InputType>>(schema: T) {
    return {
        schema,
        // This function does nothing at runtime but return the string you pass it.
        // But at compile time, TS forces you to only pass valid keys from your schema!
        nameOf: (key: keyof T & string) => key,
    };
}

/**
 * Helper funciton to create Action parameters for Cards and Callbacks.
 */
export function defineActionParameters<T extends Record<string, ParamType>>(schema: T) {
    return {
        build: (params: { [K in keyof T]: MappedParam<T[K]> }): Record<string, string> => {
            const result: Record<string, string> = {};
            for (const [key, val] of Object.entries(params)) {
                result[key] = String(val);
            }
            return result;
        },
        parse: (rawParams: Record<string, string> | undefined): Partial<{ [K in keyof T]: MappedParam<T[K]> }> => {
            const result: Record<string, unknown> = {};
            if (!rawParams) return result as Partial<{ [K in keyof T]: MappedParam<T[K]> }>;
            for (const [key, type] of Object.entries(schema)) {
                const val = rawParams[key];

                if (val === undefined) continue;

                switch (type) {
                    case "string":
                        result[key] = val;
                        break;
                    case "number": {
                        const num = Number(val);
                        if (!Number.isNaN(num)) result[key] = num;
                        break;
                    }
                    case "boolean":
                        result[key] = val === "true";
                        break;
                }
            }
            return result as Partial<{ [K in keyof T]: MappedParam<T[K]> }>;
        },
    };
}

/**
 * Extract the inputs from a Card callback. The schema should define what to expect from the inputs.
 */
export function getInputs<T extends Record<string, InputType>>(formInputs: GASFormInputs, schema: T): Partial<{ [K in keyof T]: MappedInput<T[K]> }> {
    const result: Record<string, unknown> = {};

    // if (!formInputs) return result as Partial<{ [K in keyof T]: MappedInput<T[K]> }>;

    for (const [key, expectedType] of Object.entries(schema)) {
        const rawField = formInputs[key];

        // 1. Boolean (Checkboxe, Switch). Anything except "false" or and empty string is true
        if (expectedType === "boolean") {
            const firstStringValue = rawField?.stringInputs?.value[0];
            result[key] = rawField !== undefined && firstStringValue !== "false" && firstStringValue !== "";
            continue;
        }

        // If rawField doesn't exist and isn't a boolean, it's safely undefined
        if (!rawField) continue;

        // 2. Dates (DatePickers, DateTimePicker). Can return NaN on malformed Text Input
        if (expectedType === "date") {
            const epoch = rawField.dateInput?.msSinceEpoch ?? rawField.dateTimeInput?.msSinceEpoch;
            if (epoch) {
                result[key] = Number(epoch);
                continue;
            }
            // Fallback: try to parse string input
            const fallbackStr = rawField.stringInputs?.value[0];
            if (fallbackStr) {
                const num = Number(fallbackStr);
                const num2 = Number.isNaN(num) ? new Date(fallbackStr).getTime() : num;
                if (Number.isNaN(num2)) continue;
                result[key] = num2;
            }
            // Neither DatePicker, DateTimePicker or Text Input, so we ignore it.
            continue;
        }

        // 3. Time (TimePicker)
        if (expectedType === "time") {
            if (rawField.timeInput) {
                result[key] = {
                    hours: rawField.timeInput.hours ?? 0,
                    minutes: rawField.timeInput.minutes ?? 0,
                };
            }
            // No TimePicker, so we ignore.
            continue;
        }

        // 4. Handle standard String Inputs (Text, Selects, Radios)
        const stringVals = rawField.stringInputs?.value;
        if (!stringVals || stringVals.length === 0) continue;

        switch (expectedType) {
            case "string":
                result[key] = stringVals[0] ?? "";
                break;
            case "number": {
                const parsedNum = Number(stringVals[0]);
                if (!Number.isNaN(parsedNum)) result[key] = parsedNum;
                break;
            }
            case "array":
                result[key] = stringVals ?? [];
                break;
        }
    }
    return result as Partial<{ [K in keyof T]: MappedInput<T[K]> }>;
}

/**
 * Parses a Spreadsheet coming from Sheets API, so that the Sheets and named ranges are easy to find and work with.
 */
export function parseSpreadsheet<S extends Record<string, string>, R extends Record<string, string>>(
    spreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet | undefined,
    schema: { readonly sheetNames: S; readonly namedRanges: R },
) {
    const mappedSheets: Partial<Record<S[keyof S], GoogleAppsScript.Sheets.Schema.Sheet>> = {};
    const mappedRanges: Partial<Record<R[keyof R], MappedNamedRange>> = {};

    if (!spreadsheet?.sheets) return { sheets: mappedSheets, namedRanges: mappedRanges };

    const allowedSheetNames = new Set<string>(Object.values(schema.sheetNames));
    const allowedRangeNames = new Set<string>(Object.values(schema.namedRanges));

    const sheetIdLookup: Record<number, GoogleAppsScript.Sheets.Schema.Sheet> = {};

    for (const sheet of spreadsheet.sheets) {
        sheetIdLookup[sheet.properties?.sheetId ?? 0] = sheet;
        const sheetTitle = sheet.properties?.title;
        if (sheetTitle != null && allowedSheetNames.has(sheetTitle)) {
            mappedSheets[sheetTitle as S[keyof S]] = sheet;
        }
    }

    if (spreadsheet.namedRanges) {
        for (const namedRange of spreadsheet.namedRanges) {
            if (namedRange.name == null || namedRange.range == null) continue;
            const linkedSheet = sheetIdLookup[namedRange.range.sheetId ?? 0];

            if (linkedSheet && allowedRangeNames.has(namedRange.name)) {
                mappedRanges[namedRange.name as R[keyof R]] = {
                    range: namedRange.range,
                    sheet: linkedSheet,
                };
            }
        }
    }

    return { sheets: mappedSheets, namedRanges: mappedRanges };
}
