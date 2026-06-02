import { SetupSheetSchema } from "../../common/gas-parts";
import {
    buildAddNamedRangeRequest,
    buildCopyPasteRequest,
    buildFieldsMask,
    buildMergeCellsRequest,
    buildUpdateSheetPropertiesRequest,
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
import { ConditionType, MergeType, PasteType } from "../../common/gas-utils/api-types";
import { type CalendarDates, calculateCalendarDates, calculateCalendarGrid, type MonthBlock } from "../../common/setup-utils";

type SheetName = ExtractSheetNames<typeof SetupSheetSchema>;
type RangeName = ExtractRangeNames<typeof SetupSheetSchema>;

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
    const { mappedSheets: sheets, mappedSheetNamedRanges: sheetNamedRanges, mappedRanges } = parseSpreadsheet(SetupSpreadsheet, SetupSheetSchema);

    // Extract and Validate Dates
    const dates = getCalendarDates(mappedRanges);
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
function getCalendarDates(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>): CalendarDates {
    const rangeNames = SetupSheetSchema.sheets.groupData.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango de modelo de calendario");

    const datesRange = getMappedRange(rangeNames.dates);

    const dates: [number, number, number, number] = [
        getCellUnixEpoch({ mappedRange: datesRange, rowOffset: 0 }),
        getCellUnixEpoch({ mappedRange: datesRange, rowOffset: 1 }),
        getCellUnixEpoch({ mappedRange: datesRange, rowOffset: 2 }),
        getCellUnixEpoch({ mappedRange: datesRange, rowOffset: 3 }),
    ];

    return calculateCalendarDates(dates);
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

    const grid = calculateCalendarGrid(dates);

    grid.weeks.forEach((week) => {
        const rowCells: GoogleAppsScript.Sheets.Schema.CellData[] = [{}]; // Empty Col A for month label

        week.days.forEach((day, i) => {
            // Day number
            rowCells.push({ userEnteredValue: { numberValue: day.dateNumber } });

            // Checkmark
            rowCells.push({
                userEnteredValue: { boolValue: day.inBounds && day.isWeekday },
                dataValidation: { condition: { type: ConditionType.BOOLEAN }, strict: true, showCustomUi: true },
            });

            // Format
            let formatSource: MappedNamedRange;
            switch (day.dayType) {
                case "trimester1":
                    formatSource = getMappedRange(rangeNames.trimester1Day);
                    break;
                case "trimester2":
                    formatSource = getMappedRange(rangeNames.trimester2Day);
                    break;
                case "trimester3":
                    formatSource = getMappedRange(rangeNames.trimester3Day);
                    break;
                case "rest":
                    formatSource = getMappedRange(rangeNames.restDay);
                    break;
            }

            const destinationRange = createSingleCellRange(calendarSheetId, week.rowNumber, 2 * i + 1);
            requests.push(buildCopyPasteRequest(formatSource.namedRange.range, destinationRange, PasteType.PASTE_FORMAT));
        });

        rowDataArray.push({ values: rowCells });
    });

    return { rowDataArray, monthBlocks: grid.monthBlocks, requests };
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
    apiRequests.push(
        buildUpdateSheetPropertiesRequest({
            sheetId: calendarSheetId,
            hidden: false,
            rowCount: totalRows,
            frozenRowCount: 1,
            frozenColumnCount: 1,
        }),
    );

    // Hide calendar template to be used when changing the calendar.
    apiRequests.push(
        buildUpdateSheetPropertiesRequest({
            sheetId: calendarTemplateSheet.properties?.sheetId ?? 0,
            hidden: true,
        }),
    );

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

        const sourceMonthNameRange = offsetGridRange({ origin: monthLabelRange.namedRange.range, rowOffset: block.monthIndex, height: 1, width: 1 });
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
