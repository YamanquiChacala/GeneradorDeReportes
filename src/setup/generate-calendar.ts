import { MORE_THAN_A_YEAR, MS_PER_DAY } from "../common/constants";
import { ConditionType, MergeType, PasteType } from "../common/gas-enums";
import { SetupSheetSchema } from "../common/sheet-schema";
import { buildFieldsMask } from "../common/utils/gas-types";
import { buildCopyPasteRequest, createSingleCellRange, offsetGridRange } from "../common/utils/gas-utils";
import { type ExtractRangeNames, type ExtractSheetNames, MappedNamedRange, parseSpreadsheet } from "../common/utils/mapped-name-range";

interface CalendarDates {
    dateStart: number;
    dateTrimester1: number;
    dateTrimester2: number;
    dateEnd: number;
    calStart: number;
    calEnd: number;
    totalDays: number;
    totalRows: number;
}

interface MonthBlock {
    startRow: number;
    endRow: number;
    monthIndex: number;
    year: number;
}

/**
 * Generates a Calendar in the given Spreadsheet
 * The spreadsheet must be a Group Setup file.
 */
export function generateCalendar(setupFileId: string) {
    // Fetch and Parse Initial Data
    const fieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>(
        "sheets.properties.sheetId",
        "sheets.properties.title",
        "sheets.data.rowData.values.effectiveValue",
        "namedRanges",
    );
    const SetupSpreadsheet = Sheets?.Spreadsheets.get(setupFileId, { fields: fieldsMask });
    const { sheets, sheetNamedRanges, namedRanges } = parseSpreadsheet(SetupSpreadsheet, SetupSheetSchema);

    // Extract and Validate Dates
    const dates = calculateCalendarDates(namedRanges);
    const calendarSheetId = Math.floor(Math.random() * (2 ** 31 - 1));

    // Generate the Day-by-Day data and formats
    const { rowDataArray, monthBlocks, requests: dayRequests } = buildDayDataAndFormats(dates, namedRanges, calendarSheetId);

    // Compile all API Requests
    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [
        ...buildSheetSetupRequests(sheets, sheetNamedRanges, calendarSheetId, dates.totalRows),
        ...dayRequests,
        ...buildMonthLabelRequests(monthBlocks, rowDataArray, namedRanges, calendarSheetId),
        ...buildFinalizationRequests(rowDataArray, dates, calendarSheetId),
    ];

    // Execute Batch Update
    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, setupFileId);
}

/**
 * Extracts dates from named ranges, validates them, and calculates grid boundaries.
 */
function calculateCalendarDates(namedRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>): CalendarDates {
    const dateStart = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.sheets.groupData.ranges.dateStart] });
    const dateTrimester1 = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.sheets.groupData.ranges.dateTrim1] });
    const dateTrimester2 = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.sheets.groupData.ranges.dateTrim2] });
    const dateEnd = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.sheets.groupData.ranges.dateEnd] });

    if (!dateStart || !dateTrimester1 || !dateTrimester2 || !dateEnd) throw new Error("Faltan las fechas.");
    if (dateStart >= dateTrimester1 || dateTrimester1 >= dateTrimester2 || dateTrimester2 >= dateEnd) throw new Error("Fechas en desorden.");
    if (dateEnd - dateStart > MORE_THAN_A_YEAR) throw new Error("Calendario demasiado grande.");

    // Snap to first Sunday
    const dateStartDayOfWeek = new Date(dateStart).getUTCDay();
    const calStart = dateStart - dateStartDayOfWeek * MS_PER_DAY;

    // Snap to last Saturday
    const dateEndDayOfWeek = new Date(dateEnd).getUTCDay();
    const calEnd = dateEnd + (6 - dateEndDayOfWeek) * MS_PER_DAY;

    const totalDays = (calEnd - calStart) / MS_PER_DAY + 1;
    const totalRows = 1 + Math.ceil(totalDays / 7);

    return { dateStart, dateTrimester1, dateTrimester2, dateEnd, calStart, calEnd, totalDays, totalRows };
}

/**
 * Prepares the structural requests: deleting old sheets, duplicating the template, and resizing.
 */
function buildSheetSetupRequests(
    sheets: Partial<Record<ExtractSheetNames<typeof SetupSheetSchema>, GoogleAppsScript.Sheets.Schema.Sheet>>,
    sheetNamedRanges: Partial<Record<ExtractSheetNames<typeof SetupSheetSchema>, GoogleAppsScript.Sheets.Schema.NamedRange[]>>,
    calendarSheetId: number,
    totalRows: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const calendarTemplateSheet = sheets[SetupSheetSchema.sheets.calendarTemplate.sheetName];
    if (!calendarTemplateSheet) throw new Error("Faltan hojas o formato.");

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // Clean up old calendar named ranges and sheet
    const calendarNamedRanges = sheetNamedRanges[SetupSheetSchema.sheets.calendar.sheetName] ?? [];
    for (const namedRange of calendarNamedRanges) {
        apiRequests.push({ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } });
    }

    const calendarSheet = sheets[SetupSheetSchema.sheets.calendar.sheetName];
    if (calendarSheet) {
        apiRequests.push({ deleteSheet: { sheetId: calendarSheet.properties?.sheetId } });
    }

    // Temporarily remove template named ranges so they don't duplicate, then restore them
    const templateNamedRanges = sheetNamedRanges[SetupSheetSchema.sheets.calendarTemplate.sheetName] ?? [];
    for (const namedRange of templateNamedRanges) {
        apiRequests.push({ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } });
    }

    apiRequests.push({
        duplicateSheet: {
            sourceSheetId: calendarTemplateSheet.properties?.sheetId,
            insertSheetIndex: 2,
            newSheetId: calendarSheetId,
            newSheetName: SetupSheetSchema.sheets.calendar.sheetName,
        },
    });

    for (const namedRange of templateNamedRanges) {
        apiRequests.push({ addNamedRange: { namedRange: namedRange } });
    }

    // Adjust sheet size and visibility
    apiRequests.push({
        updateSheetProperties: {
            properties: {
                sheetId: calendarSheetId,
                hidden: false,
                gridProperties: { rowCount: totalRows, frozenRowCount: 1, frozenColumnCount: 1 },
            },
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.SheetProperties>(
                "hidden",
                "gridProperties.rowCount",
                "gridProperties.frozenRowCount",
                "gridProperties.frozenColumnCount",
            ),
        },
    });

    apiRequests.push({
        updateSheetProperties: {
            properties: { sheetId: calendarTemplateSheet.properties?.sheetId, hidden: true },
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.SheetProperties>("hidden"),
        },
    });

    return apiRequests;
}

/**
 * Iterates through every day, building the cell data and formatting requests, while tracking month boundaries.
 */
function buildDayDataAndFormats(
    dates: CalendarDates,
    namedRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>,
    calendarSheetId: number,
) {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const rowDataArray: GoogleAppsScript.Sheets.Schema.RowData[] = [];
    const monthBlocks: MonthBlock[] = [];

    let currentMs = dates.calStart;
    let currentRowNumber = 1;
    let currentMonthIndex = -1;
    let currentYear = -1;
    let monthStartRow = 1;

    while (currentMs <= dates.calEnd) {
        const wednesdayMs = currentMs + 3 * MS_PER_DAY;
        const wednesdayDate = new Date(wednesdayMs);
        const monthIndex = wednesdayDate.getUTCMonth();
        const year = wednesdayDate.getUTCFullYear();

        if (currentMonthIndex !== monthIndex) {
            if (currentMonthIndex !== -1) {
                monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });
            }
            currentMonthIndex = monthIndex;
            currentYear = year;
            monthStartRow = currentRowNumber;
        }

        const rowCells: GoogleAppsScript.Sheets.Schema.CellData[] = [{}]; // Empty Col A for month label

        for (let i = 0; i < 7; i++) {
            const dayDate = new Date(currentMs);
            const dayOfWeek = dayDate.getUTCDay();
            const isWeekday = dayOfWeek !== 0 && dayOfWeek !== 6;
            const inBounds = currentMs >= dates.dateStart && currentMs <= dates.dateEnd;

            rowCells.push({ userEnteredValue: { numberValue: dayDate.getUTCDate() } });
            rowCells.push({
                userEnteredValue: { boolValue: inBounds && isWeekday },
                dataValidation: { condition: { type: ConditionType.BOOLEAN }, strict: true, showCustomUi: true },
            });

            // Format routing
            let formatSource: MappedNamedRange | undefined;
            if (isWeekday && inBounds) {
                formatSource = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.trimester1Day];
                if (currentMs > dates.dateTrimester2) formatSource = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.trimester3Day];
                else if (currentMs > dates.dateTrimester1) formatSource = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.trimester2Day];
            } else {
                formatSource = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.restDay];
            }

            const destinationRange = createSingleCellRange(calendarSheetId, currentRowNumber, 2 * i + 1);
            const formatRequest = buildCopyPasteRequest(formatSource?.range, destinationRange, PasteType.PASTE_FORMAT);
            if (formatRequest) requests.push(formatRequest);

            currentMs += MS_PER_DAY;
        }
        rowDataArray.push({ values: rowCells });
        currentRowNumber++;
    }

    // Close final month block
    monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });

    return { rowDataArray, monthBlocks, requests };
}

/**
 * Handles merging and applying month names/formatting to Column A.
 */
function buildMonthLabelRequests(
    monthBlocks: MonthBlock[],
    rowDataArray: GoogleAppsScript.Sheets.Schema.RowData[],
    namedRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>,
    calendarSheetId: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    monthBlocks.forEach((block) => {
        const rowSpan = block.endRow - block.startRow;
        let monthLabelRange = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.monthNames3];

        if (rowSpan === 2) monthLabelRange = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.monthNames2];
        else if (rowSpan === 1) monthLabelRange = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.monthNames1];

        if (!monthLabelRange?.range) throw new Error("Faltan nombres de meses.");

        const monthName = MappedNamedRange.getCellText({ mappedRange: monthLabelRange, rowOffset: block.monthIndex });
        const monthYearText = `${monthName ?? ""}\n${block.year}`;

        const sourceMonthNameRange = offsetGridRange({ origin: monthLabelRange.range, rowOffset: block.monthIndex, height: 1, width: 1 });
        const destinationMonthNameRange = createSingleCellRange(calendarSheetId, block.startRow, 0);

        const formatRequest = buildCopyPasteRequest(sourceMonthNameRange, destinationMonthNameRange, PasteType.PASTE_FORMAT);
        if (formatRequest) requests.push(formatRequest);

        requests.push({
            mergeCells: {
                range: { sheetId: calendarSheetId, startRowIndex: block.startRow, endRowIndex: block.endRow, startColumnIndex: 0, endColumnIndex: 1 },
                mergeType: MergeType.MERGE_ALL,
            },
        });

        // Inject text into the pre-existing row data array
        const targetRowData = rowDataArray[block.startRow - 1];
        if (targetRowData?.values) {
            targetRowData.values[0] = { userEnteredValue: { stringValue: monthYearText } };
        }
    });

    return requests;
}

/**
 * Applies the generated values to the grid, creates named ranges, and protects the sheet.
 */
function buildFinalizationRequests(
    rowDataArray: GoogleAppsScript.Sheets.Schema.RowData[],
    dates: CalendarDates,
    calendarSheetId: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    requests.push({
        updateCells: {
            rows: rowDataArray,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue", "dataValidation"),
            start: { sheetId: calendarSheetId, rowIndex: 1, columnIndex: 0 },
        },
    });

    const calendarFirstCell = createSingleCellRange(calendarSheetId, 0, 0);

    requests.push({
        repeatCell: {
            cell: { userEnteredValue: { numberValue: dates.calStart } },
            range: calendarFirstCell,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.numberValue"),
        },
    });

    requests.push({
        addNamedRange: {
            namedRange: {
                name: SetupSheetSchema.sheets.calendar.ranges.start,
                range: calendarFirstCell,
            },
        },
    });

    requests.push({
        addNamedRange: {
            namedRange: {
                name: SetupSheetSchema.sheets.calendar.ranges.calendar,
                range: { sheetId: calendarSheetId, startRowIndex: 1, endRowIndex: dates.totalRows, startColumnIndex: 1, endColumnIndex: 15 },
            },
        },
    });

    requests.push({
        addProtectedRange: {
            protectedRange: {
                range: { sheetId: calendarSheetId },
                description: SetupSheetSchema.sheets.calendar.sheetName,
                warningOnly: false,
            },
        },
    });

    return requests;
}
