import { type CheckboxGroup, InputType, ParamType } from "./types";

type SchemaValue = InputType | CheckboxGroup;

type MappedInput<T extends SchemaValue> =
    T extends CheckboxGroup<infer K>
        ? Record<K, boolean>
        : T extends InputType.STRING
          ? string
          : T extends InputType.NUMBER
            ? number
            : T extends InputType.BOOLEAN
              ? boolean
              : T extends InputType.DATE
                ? number
                : T extends InputType.TIME
                  ? { hours: number; minutes: number }
                  : T extends InputType.ARRAY
                    ? string[]
                    : never;

type MappedParam<T extends ParamType> = T extends ParamType.STRING ? string : T extends ParamType.NUMBER ? number : T extends ParamType.BOOLEAN ? boolean : never;

/** The part of the event object dealing with form inputs */
type GASFormInputs = GoogleAppsScript.Addons.CommonEventObject["formInputs"];

type CheckboxFields<T> = { [F in keyof T]: T[F] extends CheckboxGroup ? F : never }[keyof T] & string;

type OptionOf<G> = G extends CheckboxGroup<infer O> ? O : never;

/**
 * Helper function to create a input CheckboxGroup
 *
 * Usage:
 * checkboxGroup(["a","b","c"])
 */
export function checkboxGroup<const K extends string>(options: readonly K[]): CheckboxGroup<K> {
    return { kind: "checkboxGroup", options };
}

/**
 * Helper function to creating Form Input schemas for Cards and Callbacks.
 */
export function defineInputsSchema<T extends Record<string, SchemaValue>>(schema: T) {
    return {
        schema,
        fieldName: (key: keyof T & string) => key,
        option: <F extends CheckboxFields<T>>(_field: F, value: OptionOf<T[F]>): OptionOf<T[F]> => value,
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
                    case ParamType.STRING:
                        result[key] = val;
                        break;
                    case ParamType.NUMBER: {
                        const num = Number(val);
                        if (!Number.isNaN(num)) result[key] = num;
                        break;
                    }
                    case ParamType.BOOLEAN:
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
export function getInputs<T extends Record<string, SchemaValue>>(formInputs: GASFormInputs, schema: T): Partial<{ [K in keyof T]: MappedInput<T[K]> }> {
    const result: Record<string, unknown> = {};

    for (const [key, expectedType] of Object.entries(schema)) {
        const rawField = formInputs?.[key];

        // Checkbox group: absence of the field means "nothing checked", which is
        // a valid state, so we always produce a full record of booleans.
        if (typeof expectedType === "object" && expectedType.kind === "checkboxGroup") {
            const selected = new Set(rawField?.stringInputs?.value ?? []);
            result[key] = Object.fromEntries(expectedType.options.map((opt) => [opt, selected.has(opt)]));
            continue;
        }

        // Boolean (Checkboxe, Switch). Anything except "false" or and empty string is true
        if (expectedType === InputType.BOOLEAN) {
            const firstStringValue = rawField?.stringInputs?.value[0];
            result[key] = rawField !== undefined && firstStringValue !== "false" && firstStringValue !== "";
            continue;
        }

        // If rawField doesn't exist and isn't a boolean, it's safely undefined
        if (!rawField) continue;

        // Dates (DatePickers, DateTimePicker).
        if (expectedType === InputType.DATE) {
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
            // Malformed input, we return nothing.
            continue;
        }

        // Time (TimePicker)
        if (expectedType === InputType.TIME) {
            if (rawField.timeInput) {
                result[key] = {
                    hours: rawField.timeInput.hours,
                    minutes: rawField.timeInput.minutes,
                };
            }
            // No TimePicker, so we ignore.
            continue;
        }

        // Handle standard String Inputs (Text, Selects, Radios)
        const stringVals = rawField.stringInputs?.value;
        if (!stringVals || stringVals.length === 0) continue;

        switch (expectedType) {
            case InputType.STRING:
                result[key] = stringVals[0] ?? "";
                break;
            case InputType.NUMBER: {
                const parsedNum = Number(stringVals[0]);
                if (!Number.isNaN(parsedNum)) result[key] = parsedNum;
                break;
            }
            case InputType.ARRAY:
                result[key] = stringVals;
                break;
        }
    }
    return result as Partial<{ [K in keyof T]: MappedInput<T[K]> }>;
}
