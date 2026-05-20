/**
 * Trasnsforms from `effetive...` to `userEntered...` for use in writting.
 * @param stripOthers if true, only keep `userEntered...`, remove everything elese.
 */
export function makeUserEntered(data: GoogleAppsScript.Sheets.Schema.CellData[][], stripOthers: boolean = false): GoogleAppsScript.Sheets.Schema.CellData[][] {
    return data.map((row) =>
        row.map((cell) => {
            const finalValue = cell.effectiveValue !== undefined ? cell.effectiveValue : cell.userEnteredValue;
            const finalFormat = cell.effectiveFormat !== undefined ? cell.effectiveFormat : cell.userEnteredFormat;

            const newCell: GoogleAppsScript.Sheets.Schema.CellData = {};

            if (!stripOthers) {
                Object.assign(newCell, cell);
            }

            if (finalValue !== undefined) newCell.userEnteredValue = finalValue;
            if (finalFormat !== undefined) newCell.userEnteredFormat = finalFormat;

            delete newCell.effectiveValue;
            delete newCell.effectiveFormat;

            return newCell;
        }),
    );
}

/**
 * Factory function to create an easy way to get partial properties or throw an error.
 */
export function createRequiredGetter<K extends PropertyKey, V>(obj: Partial<Record<K, V>>, context?: string) {
    return (key: K): V => {
        const value = obj[key];

        if (value === undefined) {
            throw new Error(context ? `Falta ${context}: ${String(key)}` : `Falta propidad: ${String(key)}`);
        }

        return value;
    };
}

/**
 * Recursively unwraps Arrays.
 * E.g., Sheet[] becomes Sheet. Sheet[][] becomes Sheet.
 */
type UnwrapArray<T> = T extends Array<infer U> ? UnwrapArray<U> : T;

/**
 * A hardcoded tuple to limit TypeScript's recursion depth.
 * 5 levels deep is perfect for Sheets API (e.g., sheets.data.rowData.values.effectiveValue.numberValue).
 */
type Decrement = [never, 0, 1, 2, 3, 4, 5];

/**
 * The magic type. It extracts all keys, unwraps arrays, ignores optional (?) markers,
 * and recursively builds a union of all valid dot-notation paths.
 */
type FieldPaths<T, Depth extends Decrement[number] = 5> = [Depth] extends [never]
    ? never
    : T extends object
      ? {
            [K in keyof T]-?: K extends string
                ? UnwrapArray<NonNullable<T[K]>> extends object
                    ? // If it's an object, we keep the parent key OR recurse deeper
                      K | `${K}.${FieldPaths<UnwrapArray<NonNullable<T[K]>>, Decrement[Depth]>}`
                    : // If it's a primitive (string, number), we just return the key
                      K
                : never;
        }[keyof T]
      : never;

/**
 * Takes an array of type-checked dot-notation strings and joins them into a mask.
 * You pass the Google Apps Script Schema interface as the generic <T>.
 */
export function buildFieldsMask<T>(...paths: FieldPaths<T>[]): string {
    return paths.join(",");
}
