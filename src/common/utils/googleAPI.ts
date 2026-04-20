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

export interface MappedNamedRange {
    range: GoogleAppsScript.Sheets.Schema.GridRange;
    sheet: GoogleAppsScript.Sheets.Schema.Sheet;
}

export const MappedNamedRange = {
    getFirstCellData(mapped?: MappedNamedRange): GoogleAppsScript.Sheets.Schema.CellData | undefined {
        if (mapped?.range.startRowIndex == null || mapped.range.startColumnIndex == null) return undefined;

        return mapped.sheet.data?.[0]?.rowData?.[mapped.range.startRowIndex]?.values?.[mapped.range.startColumnIndex];
    },

    getFirstCellNumber(mapped?: MappedNamedRange): number | undefined {
        const cell = MappedNamedRange.getFirstCellData(mapped);

        return cell?.effectiveValue?.numberValue;
    },

    getFirstCellUnixEpoch(mapped?: MappedNamedRange): number | undefined {
        const GASDate = MappedNamedRange.getFirstCellNumber(mapped);
        if (GASDate == null) return undefined;
        const msPerDay = 24 * 60 * 60 * 1000;
        const sheetsEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
        return sheetsEpoch + GASDate * msPerDay;
    },
} as const;

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
                result[key] = Number.isNaN(num) ? new Date(fallbackStr).getTime() : num;
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
                result[key] = stringVals[0];
                break;
            case "number": {
                const parsedNum = Number(stringVals[0]);
                if (!Number.isNaN(parsedNum)) result[key] = parsedNum;
                break;
            }
            case "array":
                result[key] = stringVals;
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
        if (sheet.properties?.sheetId != null) {
            sheetIdLookup[sheet.properties.sheetId] = sheet;
        }
        const sheetTitle = sheet.properties?.title;
        if (sheetTitle != null && allowedSheetNames.has(sheetTitle)) {
            mappedSheets[sheetTitle as S[keyof S]] = sheet;
        }
    }

    if (spreadsheet.namedRanges) {
        for (const namedRange of spreadsheet.namedRanges) {
            if (namedRange.name == null || namedRange.range?.sheetId == null) continue;
            const linkedSheet = sheetIdLookup[namedRange.range.sheetId];

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
