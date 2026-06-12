import { SetupSheetSchema } from "../../common/gas-parts";
import {
    addNewNamedRange,
    addNewSheet,
    buildCopyPasteRequest,
    buildFieldsMask,
    buildMergeCellsRequest,
    buildProtectSheetRequest,
    buildUpdateCellsRequest,
    buildUpdateSheetPropertiesRequest,
    createRange,
    createRequiredGetter,
    createSingleCellRange,
    type ExtractRangeNames,
    getCellText,
    getCellUnixEpoch,
    type MappedNamedRange,
    offsetGridRange,
    type ParsedSpreadsheet,
    parseSpreadsheet,
} from "../../common/gas-utils";
import { ConditionType, MergeType, PasteType } from "../../common/gas-utils/api-types";
import { type CalendarDates, calculateCalendarDates, calculateCalendarGrid, DayType, type MonthBlock } from "../../common/setup-utils";

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
        "sheets.protectedRanges",
        "sheets.data.rowData.values.effectiveValue",
        "namedRanges",
    );
    const setupSpreadsheet = Sheets?.Spreadsheets.get(setupFileId, { fields: fieldsMask });
    const parsedSetup = parseSpreadsheet(setupSpreadsheet, SetupSheetSchema);

    // Extract and Validate Dates
    const dates = getCalendarDates(parsedSetup.mappedRanges);

    // Create Calendar sheet and give it the correct size.
    const { requests: setupRequests, newSheetId: calendarSheetId } = buildSheetSetupRequests(parsedSetup, dates.totalRows);

    // Generate the Day-by-Day data and formats
    const { calendarData, monthBlocks, requests: dayRequests } = buildDayDataAndFormats(dates, parsedSetup.mappedRanges, calendarSheetId);

    // Compile all API Requests
    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [
        ...setupRequests,
        ...dayRequests,
        ...buildMonthLabelRequests(monthBlocks, calendarData, parsedSetup.mappedRanges, calendarSheetId),
        ...buildFinalizationRequests(parsedSetup, calendarData, dates, calendarSheetId),
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
 * Prepares the structural requests: deleting old sheets, duplicating the template, and resizing.
 */
function buildSheetSetupRequests(
    parsedSetup: ParsedSpreadsheet<typeof SetupSheetSchema>,
    totalRows: number,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; newSheetId: number } {
    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const getMappedSheet = createRequiredGetter(parsedSetup.mappedSheets, "hoja de registro");
    const calendarTemplateId = getMappedSheet(SetupSheetSchema.sheets.calendarTemplate.sheetName).properties?.sheetId ?? 0;

    const { requests, newSheetIds } = addNewSheet({
        parsedData: parsedSetup,
        sourceSheetTitle: SetupSheetSchema.sheets.calendarTemplate.sheetName,
        insertSheetIndex: 2,
        schemaSheetName: SetupSheetSchema.sheets.calendar.sheetName,
    });

    apiRequests.push(...requests);
    // biome-ignore lint/style/noNonNullAssertion: `addNewSheet` always returns at least one id.
    const newSheetId = newSheetIds[0]!;

    // Adjust calendar sheet size and visibility
    apiRequests.push(
        buildUpdateSheetPropertiesRequest({
            sheetId: newSheetId,
            hidden: false,
            rowCount: totalRows,
            frozenRowCount: 1,
            frozenColumnCount: 1,
        }),
    );

    // Hide calendar template to be used when changing the calendar.
    apiRequests.push(
        buildUpdateSheetPropertiesRequest({
            sheetId: calendarTemplateId,
            hidden: true,
        }),
    );

    return { requests: apiRequests, newSheetId };
}

/**
 * Iterates through every day, building the cell data and formatting requests, while tracking month boundaries.
 */
function buildDayDataAndFormats(
    dates: CalendarDates,
    mappedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    calendarSheetId: number,
): {
    calendarData: GoogleAppsScript.Sheets.Schema.CellData[][];
    monthBlocks: MonthBlock[];
    requests: GoogleAppsScript.Sheets.Schema.Request[];
} {
    const rangeNames = SetupSheetSchema.sheets.calendarTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango del calendario");

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const calendarData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

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
                case DayType.TRIM1:
                    formatSource = getMappedRange(rangeNames.trimester1Day);
                    break;
                case DayType.TRIM2:
                    formatSource = getMappedRange(rangeNames.trimester2Day);
                    break;
                case DayType.TRIM3:
                    formatSource = getMappedRange(rangeNames.trimester3Day);
                    break;
                case DayType.REST:
                    formatSource = getMappedRange(rangeNames.restDay);
                    break;
            }

            const destinationRange = createSingleCellRange(calendarSheetId, week.rowNumber, 2 * i + 1);
            requests.push(buildCopyPasteRequest(formatSource.namedRange.range, destinationRange, PasteType.PASTE_FORMAT));
        });

        calendarData.push(rowCells);
    });

    return { calendarData, monthBlocks: grid.monthBlocks, requests };
}

/**
 * Handles merging and applying month names/formatting to Column A.
 */
function buildMonthLabelRequests(
    monthBlocks: MonthBlock[],
    calendarData: GoogleAppsScript.Sheets.Schema.CellData[][],
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
        const targetRowData = calendarData[block.startRow - 1]; // Array starts on row 1, blocks on 0, so we have to substract the header.
        if (targetRowData) {
            targetRowData[0] = { userEnteredValue: { stringValue: monthYearText } };
        }
    });

    return requests;
}

/**
 * Applies the generated values to the grid, creates named ranges, and protects the sheet.
 */
function buildFinalizationRequests(
    parsedData: ParsedSpreadsheet<typeof SetupSheetSchema>,
    calendarData: GoogleAppsScript.Sheets.Schema.CellData[][],
    dates: CalendarDates,
    calendarSheetId: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const calendarRange = createRange(calendarSheetId, 1, 0, dates.totalRows - 1, 15);

    const fillCalendarRequest = buildUpdateCellsRequest({
        destination: calendarRange,
        data: calendarData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue", "dataValidation"),
    });
    if (fillCalendarRequest) requests.push(fillCalendarRequest);

    // Put the ms of the calendar start in A1, for ease of calculation
    const calendarFirstCell = createSingleCellRange(calendarSheetId, 0, 0);
    const putFirstDateRequest = buildUpdateCellsRequest({
        destination: calendarFirstCell,
        data: [[{ userEnteredValue: { numberValue: dates.calStart } }]],
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.numberValue"),
    });
    if (putFirstDateRequest) requests.push(putFirstDateRequest);

    // Create named range for the start ms in A1
    requests.push(
        addNewNamedRange({
            parsedData,
            sheetTitle: SetupSheetSchema.sheets.calendar.sheetName,
            gridRange: calendarFirstCell,
            staticRangeKey: SetupSheetSchema.sheets.calendar.ranges.start,
        }),
    );

    // Create named range for the actual calendar.
    requests.push(
        addNewNamedRange({
            parsedData,
            sheetTitle: SetupSheetSchema.sheets.calendar.sheetName,
            gridRange: createRange(calendarSheetId, 1, 1, dates.totalRows - 1, 14),
            staticRangeKey: SetupSheetSchema.sheets.calendar.ranges.calendar,
        }),
    );

    // Protect the sheet
    requests.push(buildProtectSheetRequest(parsedData, SetupSheetSchema.sheets.calendar.sheetName));

    return requests;
}
