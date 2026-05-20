import type { MappedNamedRange } from ".";
import { offsetGridRange } from ".";

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

interface A1NotationParams {
    mappedRange: MappedNamedRange;
    includeSheetName?: boolean;
    lockRows?: boolean;
    lockColumns?: boolean;
    rowOffset?: number;
    colOffset?: number;
    height?: number;
    width?: number;
}

/**
 * Returns the A1 notation of a range.
 */
export function getA1Notation({
    mappedRange,
    includeSheetName = false,
    lockRows = false,
    lockColumns = false,
    rowOffset = 0,
    colOffset = 0,
    height,
    width,
}: A1NotationParams): string {
    const adjustedRange = offsetGridRange({ origin: mappedRange.range, rowOffset, colOffset, height, width });

    // Fetch sheet dimensions to cap unbounded ends
    const gridProps = mappedRange.sheet.properties?.gridProperties;
    const maxRows = gridProps?.rowCount ?? 1000;
    const maxCols = gridProps?.columnCount ?? 1000;

    // Identify bounded vs unbounded states
    const startRowDefined = adjustedRange.startRowIndex !== undefined;
    const endRowDefined = adjustedRange.endRowIndex !== undefined;
    const startColDefined = adjustedRange.startColumnIndex !== undefined;
    const endColDefined = adjustedRange.endColumnIndex !== undefined;

    // A dimension is fully unbounded only if BOTH start and end are omitted
    const isRowFullyUnbounded = !startRowDefined && !endRowDefined;
    const isColFullyUnbounded = !startColDefined && !endColDefined;

    // Resolve indices (0-based) to 1-based coordinates
    // Start defaults to 0 (row 1 or col A)
    const startRow = (adjustedRange.startRowIndex ?? 0) + 1;
    const startColLetter = getColumnLetter(adjustedRange.startColumnIndex ?? 0);

    // End defaults to the maximum bound of the sheet
    const endRow = adjustedRange.endRowIndex ?? maxRows;
    const endColIndex = adjustedRange.endColumnIndex ?? maxCols;
    const endColLetter = getColumnLetter(endColIndex - 1); // Subtract 1 because endIndex is exclusive

    // Build lock prefixes
    const rowPrefix = lockRows ? "$" : "";
    const colPrefix = lockColumns ? "$" : "";

    // Build coordinate strings (omit completely if fully unbounded to support A:A or 1:1)
    const startRowStr = isRowFullyUnbounded ? "" : `${rowPrefix}${startRow}`;
    const endRowStr = isRowFullyUnbounded ? "" : `${rowPrefix}${endRow}`;

    const startColStr = isColFullyUnbounded ? "" : `${colPrefix}${startColLetter}`;
    const endColStr = isColFullyUnbounded ? "" : `${colPrefix}${endColLetter}`;

    const startA1 = `${startColStr}${startRowStr}`;
    const endA1 = `${endColStr}${endRowStr}`;

    // Formulate the final A1 notation
    let a1Notation = "";

    if (isRowFullyUnbounded && isColFullyUnbounded) {
        // Fallback for an entirely unbounded range (the whole sheet)
        a1Notation = `${colPrefix}A${rowPrefix}1:${colPrefix}${endColLetter}${rowPrefix}${maxRows}`;
    } else if (startA1 === endA1) {
        // If start and end strings are identical, it's a single entity
        if (isRowFullyUnbounded || isColFullyUnbounded) {
            a1Notation = `${startA1}:${startA1}`; // Single entire row (1:1) or col (A:A)
        } else {
            a1Notation = startA1; // Single cell (A1)
        }
    } else {
        // Standard range (A1:B2, A:B, 1:2, or C4:Z1000)
        a1Notation = `${startA1}:${endA1}`;
    }

    // Append sheet name if requested
    if (includeSheetName && mappedRange.sheet.properties?.title) {
        const escapedTitle = mappedRange.sheet.properties.title.replace(/'/g, "''");
        return `'${escapedTitle}'!${a1Notation}`;
    }

    return a1Notation;
}
