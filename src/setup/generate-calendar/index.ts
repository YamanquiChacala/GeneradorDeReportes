import { ConditionType, MergeType, MORE_THAN_A_YEAR, MS_PER_DAY, PasteType } from "../../common/constants";
import { SetupSheetSchema } from "../../common/gas-parts";
import {
    buildAddNamedRangeRequest,
    buildCopyPasteRequest,
    buildFieldsMask,
    buildMergeCellsRequest,
    createRange,
    createRequiredGetter,
    createSingleCellRange,
    type ExtractRangeNames,
    type ExtractSheetNames,
    getCellText,
    getCellUnixEpoch,
    type MappedNamedRange,
    offsetGridRange,
    parseSpreadsheet,
} from "../../common/gas-utils";

type SheetName = ExtractSheetNames<typeof SetupSheetSchema>;
type RangeName = ExtractRangeNames<typeof SetupSheetSchema>;

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
    const { sheets, sheetNamedRanges, mappedRanges } = parseSpreadsheet(SetupSpreadsheet, SetupSheetSchema);

    // Extract and Validate Dates
    const dates = calculateCalendarDates(mappedRanges);
    const calendarSheetId = Math.floor(Math.random() * (2 ** 31 - 1));

    // Generate the Day-by-Day data and formats
    const { rowDataArray, monthBlocks, requests: dayRequests } = buildDayDataAndFormats(dates, mappedRanges, calendarSheetId);

    // Compile all API Requests
    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [
        ...buildSheetSetupRequests(sheets, sheetNamedRanges, calendarSheetId, dates.totalRows),
        ...dayRequests,
        ...buildMonthLabelRequests(monthBlocks, rowDataArray, mappedRanges, calendarSheetId),
        ...buildFinalizationRequests(rowDataArray, dates, calendarSheetId),
    ];

    // Execute Batch Update
    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, setupFileId);
}

/**
 * Extracts dates from named ranges, validates them, and calculates grid boundaries.
 */
function calculateCalendarDates(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>): CalendarDates {
    const rangeNames = SetupSheetSchema.sheets.groupData.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango de modelo de calendario");

    const dates = getMappedRange(rangeNames.dates);

    // Get user provided dates.
    const date0 = getCellUnixEpoch({ mappedRange: dates, rowOffset: 0 });
    const date1 = getCellUnixEpoch({ mappedRange: dates, rowOffset: 1 });
    const date2 = getCellUnixEpoch({ mappedRange: dates, rowOffset: 2 });
    const date3 = getCellUnixEpoch({ mappedRange: dates, rowOffset: 3 });

    // Validate user provided dates.
    if (!date0 || !date1 || !date2 || !date3) throw new Error("Faltan las fechas.");
    if (date0 >= date1 || date1 >= date2 || date2 >= date3) throw new Error("Fechas en desorden.");
    if (date3 - date0 > MORE_THAN_A_YEAR) throw new Error("Calendario demasiado grande.");

    // Snap to first Sunday
    const dateStartDayOfWeek = new Date(date0).getUTCDay();
    const calStart = date0 - dateStartDayOfWeek * MS_PER_DAY;

    // Snap to last Saturday
    const dateEndDayOfWeek = new Date(date3).getUTCDay();
    const calEnd = date3 + (6 - dateEndDayOfWeek) * MS_PER_DAY;

    const totalDays = (calEnd - calStart) / MS_PER_DAY + 1;
    const totalRows = 1 + Math.ceil(totalDays / 7);

    return { dateStart: date0, dateTrimester1: date1, dateTrimester2: date2, dateEnd: date3, calStart, calEnd, totalDays, totalRows };
}

/**
 * Iterates through every day, building the cell data and formatting requests, while tracking month boundaries.
 */
function buildDayDataAndFormats(
    dates: CalendarDates,
    mappedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    calendarSheetId: number,
): {
    rowDataArray: GoogleAppsScript.Sheets.Schema.RowData[];
    monthBlocks: MonthBlock[];
    requests: GoogleAppsScript.Sheets.Schema.Request[];
} {
    const rangeNames = SetupSheetSchema.sheets.calendarTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango del calendario");

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const rowDataArray: GoogleAppsScript.Sheets.Schema.RowData[] = [];
    const monthBlocks: MonthBlock[] = [];

    let currentMs = dates.calStart; // The day we're processing (we advance by week rows, so 7 days)
    let currentRowNumber = 1; // 0 Is headers.
    let currentMonthIndex = -1; // undefined, 0-11 for month.
    let currentYear = -1; // undefined
    let monthStartRow = 1; // undefined. What row did this month started on, used to merge the full month label.

    while (currentMs <= dates.calEnd) {
        // Wednesday (day 3) defines what month the week "belongs" to.
        const wednesdayMs = currentMs + 3 * MS_PER_DAY;
        const wednesdayDate = new Date(wednesdayMs);
        const monthIndex = wednesdayDate.getUTCMonth();
        const year = wednesdayDate.getUTCFullYear();

        // If this week is the start of a new month.
        if (currentMonthIndex !== monthIndex) {
            if (currentMonthIndex !== -1) {
                // Don't close a month before even starting.
                monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });
            }
            currentMonthIndex = monthIndex;
            currentYear = year;
            monthStartRow = currentRowNumber;
        }

        const rowCells: GoogleAppsScript.Sheets.Schema.CellData[] = [{}]; // Empty Col A for month label

        // Process each day of the week
        for (let i = 0; i < 7; i++) {
            const dayDate = new Date(currentMs);
            const dayOfWeek = dayDate.getUTCDay();
            const isWeekday = dayOfWeek !== 0 && dayOfWeek !== 6;
            const inBounds = currentMs >= dates.dateStart && currentMs <= dates.dateEnd;

            // Day number at the left
            rowCells.push({ userEnteredValue: { numberValue: dayDate.getUTCDate() } });
            // Checkmark at the rigth
            rowCells.push({
                userEnteredValue: { boolValue: inBounds && isWeekday },
                dataValidation: { condition: { type: ConditionType.BOOLEAN }, strict: true, showCustomUi: true },
            });

            // Format of our number and checkmark
            let formatSource: MappedNamedRange;
            if (isWeekday && inBounds) {
                formatSource = getMappedRange(rangeNames.trimester1Day);
                if (currentMs > dates.dateTrimester2) formatSource = getMappedRange(rangeNames.trimester3Day);
                else if (currentMs > dates.dateTrimester1) formatSource = getMappedRange(rangeNames.trimester2Day);
            } else {
                formatSource = getMappedRange(rangeNames.restDay);
            }

            const destinationRange = createSingleCellRange(calendarSheetId, currentRowNumber, 2 * i + 1);
            requests.push(buildCopyPasteRequest(formatSource.range, destinationRange, PasteType.PASTE_FORMAT));

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
 * Prepares the structural requests: deleting old sheets, duplicating the template, and resizing.
 */
function buildSheetSetupRequests(
    sheets: Partial<Record<SheetName, GoogleAppsScript.Sheets.Schema.Sheet>>,
    sheetNamedRanges: Partial<Record<SheetName, GoogleAppsScript.Sheets.Schema.NamedRange[]>>,
    calendarSheetId: number,
    totalRows: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const getSheet = createRequiredGetter(sheets, "hoja para calendario");

    const calendarTemplateSheet = getSheet(SetupSheetSchema.sheets.calendarTemplate.sheetName);

    // Delete old calendar named ranges
    const calendarNamedRanges = sheetNamedRanges[SetupSheetSchema.sheets.calendar.sheetName] ?? [];
    for (const namedRange of calendarNamedRanges) {
        apiRequests.push({ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } });
    }

    // Delete old calendar sheet if it exists.
    const calendarSheet = sheets[SetupSheetSchema.sheets.calendar.sheetName];
    if (calendarSheet) {
        apiRequests.push({ deleteSheet: { sheetId: calendarSheet.properties?.sheetId } });
    }

    // Temporarily remove template named ranges so they don't duplicate, then restore them
    const templateNamedRanges = sheetNamedRanges[SetupSheetSchema.sheets.calendarTemplate.sheetName] ?? [];
    for (const namedRange of templateNamedRanges) {
        apiRequests.push({ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } });
    }

    // Create a brand new calendar sheet
    apiRequests.push({
        duplicateSheet: {
            sourceSheetId: calendarTemplateSheet.properties?.sheetId,
            insertSheetIndex: 2,
            newSheetId: calendarSheetId,
            newSheetName: SetupSheetSchema.sheets.calendar.sheetName,
        },
    });

    // Restore template named ranges.
    for (const namedRange of templateNamedRanges) {
        apiRequests.push({ addNamedRange: { namedRange: namedRange } });
    }

    // Adjust calendar sheet size and visibility
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

    // Hide calendar template to be used when changing the calendar.
    apiRequests.push({
        updateSheetProperties: {
            properties: { sheetId: calendarTemplateSheet.properties?.sheetId, hidden: true },
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.SheetProperties>("hidden"),
        },
    });

    return apiRequests;
}

/**
 * Handles merging and applying month names/formatting to Column A.
 */
function buildMonthLabelRequests(
    monthBlocks: MonthBlock[],
    rowDataArray: GoogleAppsScript.Sheets.Schema.RowData[],
    mappedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    calendarSheetId: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const rangeNames = SetupSheetSchema.sheets.calendarTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango del modelo de calendario");

    monthBlocks.forEach((block) => {
        // How many weeks (rows) this month spans
        const rowSpan = block.endRow - block.startRow;

        // Grab the name month names for the correct span.
        let monthLabelRange = getMappedRange(rangeNames.monthNames3);
        if (rowSpan === 2) monthLabelRange = getMappedRange(rangeNames.monthNames2);
        else if (rowSpan === 1) monthLabelRange = getMappedRange(rangeNames.monthNames1);

        const monthName = getCellText({ mappedRange: monthLabelRange, rowOffset: block.monthIndex }) ?? "";
        const monthYearText = `${monthName}\n${block.year}`;

        const sourceMonthNameRange = offsetGridRange({ origin: monthLabelRange.range, rowOffset: block.monthIndex, height: 1, width: 1 });
        const destinationMonthNameRange = createSingleCellRange(calendarSheetId, block.startRow, 0);
        const monthMergeRange = createRange(calendarSheetId, block.startRow, 0, rowSpan, 1);

        // Make the requests
        requests.push(buildCopyPasteRequest(sourceMonthNameRange, destinationMonthNameRange, PasteType.PASTE_FORMAT));
        requests.push(buildMergeCellsRequest(monthMergeRange, MergeType.MERGE_ALL));

        // Inject text into the pre-existing row data array
        const targetRowData = rowDataArray[block.startRow - 1]; // Array starts on row 1, blocks on 0, so we have to substract the header.
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

    // Put all the months, days and checkmarks for the calendar.
    requests.push({
        updateCells: {
            rows: rowDataArray,
            start: { sheetId: calendarSheetId, rowIndex: 1, columnIndex: 0 },
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue", "dataValidation"),
        },
    });

    // Put the ms of the calendar start in A1, for ease of calculation
    const calendarFirstCell = createSingleCellRange(calendarSheetId, 0, 0);
    requests.push({
        repeatCell: {
            cell: { userEnteredValue: { numberValue: dates.calStart } },
            range: calendarFirstCell,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.numberValue"),
        },
    });

    // Create named range for the start ms in A1
    requests.push(buildAddNamedRangeRequest(SetupSheetSchema.sheets.calendar.ranges.start, calendarFirstCell));

    // Create named range for the actual calendar.
    requests.push(buildAddNamedRangeRequest(SetupSheetSchema.sheets.calendar.ranges.calendar, createRange(calendarSheetId, 1, 1, dates.totalRows - 1, 14)));

    // Protect the sheet
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
