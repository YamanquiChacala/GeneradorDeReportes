type InputType = "string" | "number" | "boolean" | "date" | "time" | "array";

type MappedInput<T extends InputType> = T extends "string"
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

interface MappedNamedRange {
    range: GoogleAppsScript.Sheets.Schema.GridRange;
    sheet: GoogleAppsScript.Sheets.Schema.Sheet;
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
 *
 */
export function getDateMs(range?: GoogleAppsScript.Sheets.Schema.GridRange, sheet?: GoogleAppsScript.Sheets.Schema.Sheet): number | null {
    if (range?.startRowIndex == null || range?.startColumnIndex == null) return null;

    const cell = sheet?.data?.[0]?.rowData?.[range.startRowIndex]?.values?.[range.startColumnIndex];
    const rawNumber = cell?.effectiveValue?.numberValue;

    if (rawNumber == null) return null;

    const msPerDay = 24 * 60 * 60 * 1000;
    const sheetsEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
    return sheetsEpoch + rawNumber * msPerDay;
}

/**
 * Maps the namedRanges of a Spreadsheet to a dictionary with the sheet data and the range for ease of use.
 */
export function mapNamedRages<T extends Record<string, string>>(
    spreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet,
    allowedRangesConfig: T,
): Partial<Record<T[keyof T], MappedNamedRange>> {
    const dictionary: Partial<Record<T[keyof T], MappedNamedRange>> = {};

    if (!spreadsheet.namedRanges || !spreadsheet.sheets) return dictionary;

    const sheetLookup: Record<number, GoogleAppsScript.Sheets.Schema.Sheet> = {};
    for (const sheet of spreadsheet.sheets) {
        if (sheet.properties?.sheetId != null) {
            sheetLookup[sheet.properties.sheetId] = sheet;
        }
    }

    const allowedValues = new Set<string>(Object.values(allowedRangesConfig));

    for (const namedRange of spreadsheet.namedRanges) {
        if (!namedRange.name || namedRange.range?.sheetId == null) continue;
        const linkedSheet = sheetLookup[namedRange.range.sheetId];
        if (linkedSheet && allowedValues.has(namedRange.name)) {
            const key = namedRange.name as T[keyof T];
            dictionary[key] = {
                range: namedRange.range,
                sheet: linkedSheet,
            };
        }
    }

    return dictionary;
}
