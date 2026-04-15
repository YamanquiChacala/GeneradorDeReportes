/**
 * Define the exact shapes we expect from Google Apps Script payloads
 */
interface GasInput {
    stringInputs?: { value: string[] };
    dateInput?: { msSinceEpoch: string | number };
    dateTimeInput?: { msSinceEpoch: string | number };
    [key: string]: unknown;
}

/**
 * Flattens the nested GAS formInputs object into a clean dictionary.
 * Uses a generic <T> so you can strongly type the returned object!
 */
export function flattenFormInputs<T = Record<string, unknown>>(formInputs: Record<string, GasInput> | undefined): Partial<T> {
    const flat: Record<string, unknown> = {};

    if (!formInputs) return flat as Partial<T>;

    for (const [key, obj] of Object.entries(formInputs)) {
        if (!obj) continue;

        if (obj.stringInputs?.value !== undefined) {
            const values = obj.stringInputs.value;
            if (values.length === 1) {
                const val = values[0];
                flat[key] = val === "true" ? true : val;
            } else {
                flat[key] = values;
            }
        } else if (obj.dateInput?.msSinceEpoch !== undefined) {
            flat[key] = Number(obj.dateInput.msSinceEpoch);
        } else if (obj.dateTimeInput?.msSinceEpoch !== undefined) {
            flat[key] = Number(obj.dateTimeInput.msSinceEpoch);
        } else {
            flat[key] = obj;
        }
    }

    return flat as Partial<T>;
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
 * Fetches multiple named ranges in a single Advanced Sheets API call.
 * Auto-flattens 1x1 ranges into primitives.
 */
export function batchFetchNamedRangesValues<T extends Record<string, string>>(spreadsheetId: string, schemaGroup: T): Partial<Record<keyof T, unknown>> {
    const keys = Object.keys(schemaGroup) as (keyof T)[];
    const rangeNames = Object.values(schemaGroup) as string[];

    const response = Sheets?.Spreadsheets.Values.batchGet(spreadsheetId, {
        ranges: rangeNames,
        majorDimension: "ROWS",
        valueRenderOption: "UNFORMATTED_VALUE",
    });

    const result: Partial<Record<keyof T, unknown>> = {};
    const valueRanges = response?.valueRanges || [];

    keys.forEach((key, index) => {
        const values = valueRanges[index]?.values;

        if (!values || values.length === 0) {
            result[key] = undefined;
            return;
        }
        const firstRow = values[0];

        if (values.length === 1 && Array.isArray(firstRow) && firstRow.length === 1) {
            result[key] = firstRow[0];
        } else {
            result[key] = values;
        }
    });

    return result;
}
