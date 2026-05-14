/**
 * Transforms a column number into it's corresponding column letter, using 0-based index.
 */
export function getColumnLetter(column: number): string {
    let temp: number;
    let letter = "";
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
}

// TODO: Make tests for this
/**
 * Returns the A1 notation of a range.
 */

export function getA1Notation(range: GoogleAppsScript.Sheets.Schema.GridRange, first = false, lockRows = false, lockColumns = false): string {
    // 1. Define our lock prefixes
    const rowPrefix = lockRows ? "$" : "";
    const colPrefix = lockColumns ? "$" : "";

    // 2. Check if dimensions are bounded (undefined means unbounded in that dimension)
    const hasRows = range.startRowIndex !== undefined || range.endRowIndex !== undefined;
    const hasCols = range.startColumnIndex !== undefined || range.endColumnIndex !== undefined;

    // 3. Get initial coordinates (defaults to 0 for the 'first' flag fallback)
    const startRow = (range.startRowIndex ?? 0) + 1;
    const startColLetter = getColumnLetter(range.startColumnIndex ?? 0);

    if (first) return `${colPrefix}${startColLetter}${rowPrefix}${startRow}`;

    // 4. Build starting strings (omit entirely if the dimension is unbounded)
    const startRowStr = hasRows ? `${rowPrefix}${startRow}` : "";
    const startColStr = hasCols ? `${colPrefix}${startColLetter}` : "";

    // 5. Build ending strings
    // If end index is omitted on a bounded dimension, we assume it spans 1 unit (end = start + 1)
    // Because end indices are exclusive, endRowIndex exactly matches the 1-based row number.
    const endRow = range.endRowIndex !== undefined ? range.endRowIndex : startRow;
    const endColIndex = range.endColumnIndex !== undefined ? range.endColumnIndex : (range.startColumnIndex ?? 0) + 1;

    // For columns, we subtract 1 to get the actual final column included in the range.
    const endColLetter = getColumnLetter(endColIndex - 1);

    const endRowStr = hasRows ? `${rowPrefix}${endRow}` : "";
    const endColStr = hasCols ? `${colPrefix}${endColLetter}` : "";

    const startA1 = `${startColStr}${startRowStr}`;
    const endA1 = `${endColStr}${endRowStr}`;

    // 6. Handle single entities
    if (startA1 === endA1) {
        if (!hasRows) return `${startColStr}:${startColStr}`; // Unbounded single column (e.g., "$A:$A")
        if (!hasCols) return `${startRowStr}:${startRowStr}`; // Unbounded single row (e.g., "$1:$1")
        return startA1; // Single cell (e.g., "$A$1")
    }

    // 7. Return the standard range
    return `${startA1}:${endA1}`; // e.g., "$A$1:$B$2", "A$1:A", etc.
}
