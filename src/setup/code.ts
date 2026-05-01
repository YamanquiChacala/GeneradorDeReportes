import { MORE_THAN_A_YEAR } from "../common/constants";
import { FileType } from "../common/enums";
import { PasteType } from "../common/gas-enums";
import { ReportSheetSchema, SetupSheetSchema } from "../common/sheet-schema";
import { key as FILE_VALIDATION_KEY } from "../common/utils/file-validation";
import { buildFieldsMask, defineRangesDataConfig, type MappedInput } from "../common/utils/gas-types";
import { buildCopyPasteRequest, colorToHex, createSingleCellRange, offsetGridRange } from "../common/utils/gas-utils";
import { MappedNamedRange, parseSpreadsheet } from "../common/utils/mapped-name-range";
import { sanitizeFileName } from "../common/utils/text";

const SetupFileDataConfig = defineRangesDataConfig({
    groupName: { range: SetupSheetSchema.sheets.groupData.ranges.groupName, type: "string" },
    attendancePerClass: { range: SetupSheetSchema.sheets.groupData.ranges.attendancePerClass, type: "boolean" },
    averagePerField: { range: SetupSheetSchema.sheets.groupData.ranges.averagePerField, type: "boolean" },
    dateStart: { range: SetupSheetSchema.sheets.groupData.ranges.dateStart, type: "date" },
    dateEndTrimester1: { range: SetupSheetSchema.sheets.groupData.ranges.dateTrim1, type: "date" },
    dateEndTrimester2: { range: SetupSheetSchema.sheets.groupData.ranges.dateTrim2, type: "date" },
    dateEnd: { range: SetupSheetSchema.sheets.groupData.ranges.dateEnd, type: "date" },
} as const);

export type SetupFileData = { folderId: string } & {
    [K in keyof typeof SetupFileDataConfig]: MappedInput<(typeof SetupFileDataConfig)[K]["type"]>;
};

/**
 * Creates a new Group Initialization file with the given data.
 */
export function createSetupFile(initData: SetupFileData) {
    // ========== Create File ============

    const fileName = `__Registro Inicial - ${initData.groupName}`;

    const newFile = Drive?.Files.copy(
        {
            name: fileName,
            parents: [initData.folderId],
            appProperties: {
                [FILE_VALIDATION_KEY]: FileType.SETUP,
            },
        },
        SetupSheetSchema.templateId,
        {
            supportsAllDrives: true,
        },
    );

    const newFileId = newFile?.id;

    if (!newFileId) throw new Error("Error al crear copia del Regirsto Inicial");

    // ======== Update namedRanges ==========

    const updateData: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

    for (const [key, config] of Object.entries(SetupFileDataConfig)) {
        const rawValue = initData[key as keyof typeof SetupFileDataConfig];

        if (rawValue == null) continue;

        let cellValue = rawValue;

        if (config.type === "date") {
            const date = new Date(rawValue as number);
            cellValue = Utilities.formatDate(date, "UTC", "yyyy-MM-dd");
        }
        updateData.push({
            range: config.range,
            values: [[cellValue]],
        });
    }

    Sheets?.Spreadsheets.Values.batchUpdate(
        {
            valueInputOption: "USER_ENTERED",
            data: updateData,
        },
        newFileId,
    );

    // ========= Create Calendar ==========
    generateCalendar(newFileId);
}

/**
 * Generates a Calendar in the given Spreadsheet
 * The spreadsheet must be a Setup Group.
 */
export function generateCalendar(setupFileId: string) {
    const fieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>(
        "sheets.properties.sheetId",
        "sheets.properties.title",
        "sheets.data.rowData.values.effectiveValue",
        "namedRanges",
    );
    const SetupSpreadsheet = Sheets?.Spreadsheets.get(setupFileId, { fields: fieldsMask });

    const { sheets, sheetNamedRanges, namedRanges } = parseSpreadsheet(SetupSpreadsheet, SetupSheetSchema);

    const calendarTemplateSheet = sheets[SetupSheetSchema.sheets.calendarTemplate.sheetName];

    if (!calendarTemplateSheet) throw new Error("Faltan hojas o formato.");

    const dateStart = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.sheets.groupData.ranges.dateStart] });
    const dateTrimester1 = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.sheets.groupData.ranges.dateTrim1] });
    const dateTrimester2 = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.sheets.groupData.ranges.dateTrim2] });
    const dateEnd = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.sheets.groupData.ranges.dateEnd] });

    if (!dateStart || !dateTrimester1 || !dateTrimester2 || !dateEnd) throw new Error("Faltan las fechas.");
    if (dateStart >= dateTrimester1 || dateTrimester1 >= dateTrimester2 || dateTrimester2 >= dateEnd) throw new Error("Fechas en desorden.");
    if (dateEnd - dateStart > MORE_THAN_A_YEAR) throw new Error("Calendario demasiado grande.");

    // ========= Date calculations ===========
    const msPerDay = 24 * 60 * 60 * 1000;

    // Snap to first Sunday
    const dateStartDayOfWeek = new Date(dateStart).getUTCDay();
    const calStart = dateStart - dateStartDayOfWeek * msPerDay;

    // Snap to last Saturday
    const dateEndDayOfWeek = new Date(dateEnd).getUTCDay();
    const calEnd = dateEnd + (6 - dateEndDayOfWeek) * msPerDay;

    const totalDays = (calEnd - calStart) / msPerDay + 1;
    const totalRows = 1 + Math.ceil(totalDays / 7);

    // ========= Prepara the new sheet =========
    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // Remove the old calendar, if it exists.

    const calendarNamedRanges = sheetNamedRanges[SetupSheetSchema.sheets.calendar.sheetName] ?? [];

    calendarNamedRanges.forEach((namedRange) => {
        apiRequests.push({
            deleteNamedRange: {
                namedRangeId: namedRange.namedRangeId,
            },
        });
    });

    const calendarSheet = sheets[SetupSheetSchema.sheets.calendar.sheetName];
    if (calendarSheet) {
        apiRequests.push({
            deleteSheet: {
                sheetId: calendarSheet.properties?.sheetId,
            },
        });
    }

    // Tempraraty remove named ranges so they don't get duplicated
    const templateNamedRanges = sheetNamedRanges[SetupSheetSchema.sheets.calendarTemplate.sheetName] ?? [];
    // const templateNamedRanges = (SetupSpreadsheet?.namedRanges ?? []).filter((namedRange) => namedRange.range?.sheetId === calendarTemplateSheet.properties?.sheetId);

    templateNamedRanges.forEach((namedRange) => {
        apiRequests.push({
            deleteNamedRange: {
                namedRangeId: namedRange.namedRangeId,
            },
        });
    });

    // Duplicate the template
    const calendarSheetId = Math.floor(Math.random() * (2 ** 31 - 1));

    apiRequests.push({
        duplicateSheet: {
            sourceSheetId: calendarTemplateSheet.properties?.sheetId,
            insertSheetIndex: 2,
            newSheetId: calendarSheetId,
            newSheetName: SetupSheetSchema.sheets.calendar.sheetName,
        },
    });

    // Add the named ranges to the template again.
    templateNamedRanges.forEach((namedRange) => {
        apiRequests.push({
            addNamedRange: {
                namedRange: namedRange,
            },
        });
    });

    // Adjust sheet size
    apiRequests.push({
        updateSheetProperties: {
            properties: {
                sheetId: calendarSheetId,
                hidden: false,
                // tabColorStyle: {rgbColor: {red: 1.0, green: 1.0, blue: 1.0}}, // TODO: Add to "@types/google-apps-script
                gridProperties: {
                    rowCount: totalRows,
                    frozenRowCount: 1,
                    frozenColumnCount: 1,
                },
            },
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.SheetProperties>(
                "hidden",
                "gridProperties.rowCount",
                "gridProperties.frozenRowCount",
                "gridProperties.frozenColumnCount",
            ),
        },
    });

    // Hide template
    apiRequests.push({
        updateSheetProperties: {
            properties: {
                sheetId: calendarTemplateSheet.properties?.sheetId,
                hidden: true,
            },
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.SheetProperties>("hidden"),
        },
    });

    // ========== Build each day =============
    let currentMs = calStart;
    let currentRowNumber = 1;

    let currentMonthIndex = -1;
    let currentYear = -1;
    let monthStartRow = 1;

    const monthBlocks: { startRow: number; endRow: number; monthIndex: number; year: number }[] = [];

    const rowDataArray: GoogleAppsScript.Sheets.Schema.RowData[] = [];

    while (currentMs <= calEnd) {
        // Handle change of month, this happens on wednesday.
        const wednesdayMs = currentMs + 3 * msPerDay;
        const wednesdaydate = new Date(wednesdayMs);
        const monthIndex = wednesdaydate.getUTCMonth();
        const year = wednesdaydate.getUTCFullYear();

        if (currentMonthIndex !== monthIndex) {
            if (currentMonthIndex !== -1) {
                monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });
            }
            currentMonthIndex = monthIndex;
            currentYear = year;
            monthStartRow = currentRowNumber;
        }

        const rowCells: GoogleAppsScript.Sheets.Schema.CellData[] = [{}]; // Column A is empty for now, later we'll inject the month.

        // Handle each day of the week
        for (let i = 0; i < 7; i++) {
            const dayDate = new Date(currentMs);
            const dayOfWeek = dayDate.getUTCDay();
            const isWeekday = dayOfWeek !== 0 && dayOfWeek !== 6;
            const inBounds = currentMs >= dateStart && currentMs <= dateEnd;

            rowCells.push({ userEnteredValue: { numberValue: dayDate.getUTCDate() } });
            rowCells.push({
                userEnteredValue: { boolValue: inBounds && isWeekday },
                dataValidation: { condition: { type: "BOOLEAN" }, strict: true, showCustomUi: true },
            });

            // Format the cells
            let formatSource: MappedNamedRange | undefined;
            if (isWeekday && inBounds) {
                formatSource = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.trimester1Day]; // Default to trimester 1;
                if (currentMs > dateTrimester2) formatSource = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.trimester3Day];
                else if (currentMs > dateTrimester1) formatSource = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.trimester2Day];
            } else {
                formatSource = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.restDay];
            }

            const destinationRange = createSingleCellRange(calendarSheetId, currentRowNumber, 2 * i + 1);

            const formatRequest = buildCopyPasteRequest(formatSource?.range, destinationRange, PasteType.PASTE_FORMAT);

            if (formatRequest) apiRequests.push(formatRequest);

            currentMs += msPerDay;
        }
        rowDataArray.push({ values: rowCells });
        currentRowNumber++;
    }
    // ============ Handle month and year labels ============

    // Close last month block.
    monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });

    monthBlocks.forEach((block) => {
        const rowSpan = block.endRow - block.startRow;

        let monthLabelRange = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.monthNames3];

        if (rowSpan === 2) {
            monthLabelRange = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.monthNames2];
        } else if (rowSpan === 1) {
            monthLabelRange = namedRanges[SetupSheetSchema.sheets.calendarTemplate.ranges.monthNames1];
        }

        const monthName = MappedNamedRange.getCellText({ mappedRange: monthLabelRange, rowOffset: block.monthIndex });
        const monthYearText = `${monthName ?? ""}\n${block.year}`;

        if (!monthLabelRange?.range) throw new Error("Missing month names.");

        const sourceMonthNameRange = offsetGridRange({ origin: monthLabelRange.range, rowOffset: block.monthIndex, height: 1, width: 1 });
        const destinationMonthNameRange = createSingleCellRange(calendarSheetId, block.startRow, 0);

        const formatRequest = buildCopyPasteRequest(sourceMonthNameRange, destinationMonthNameRange, PasteType.PASTE_FORMAT);
        if (formatRequest) apiRequests.push(formatRequest);

        // Merge
        apiRequests.push({
            mergeCells: {
                range: {
                    sheetId: calendarSheetId,
                    startRowIndex: block.startRow,
                    endRowIndex: block.endRow,
                    startColumnIndex: 0,
                    endColumnIndex: 1,
                },
                mergeType: "MERGE_ALL",
            },
        });

        //Insert text in the data array.
        const targetRowData = rowDataArray[block.startRow - 1];
        if (targetRowData?.values) {
            targetRowData.values[0] = { userEnteredValue: { stringValue: monthYearText } };
        }
    });

    // ============= Update all the cell values. =============
    apiRequests.push({
        updateCells: {
            rows: rowDataArray,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue", "dataValidation"),
            start: { sheetId: calendarSheetId, rowIndex: 1, columnIndex: 0 },
        },
    });

    apiRequests.push({
        repeatCell: {
            cell: { userEnteredValue: { numberValue: calStart } },
            range: {
                sheetId: calendarSheetId,
                startRowIndex: 0,
                endRowIndex: 1,
                startColumnIndex: 0,
                endColumnIndex: 1,
            },
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        },
    });

    // ============= Calendar named ranges =============

    apiRequests.push({
        addNamedRange: {
            namedRange: {
                name: SetupSheetSchema.sheets.calendar.ranges.start,
                range: {
                    sheetId: calendarSheetId,
                    startRowIndex: 0,
                    endRowIndex: 1,
                    startColumnIndex: 0,
                    endColumnIndex: 1,
                },
            },
        },
    });

    apiRequests.push({
        addNamedRange: {
            namedRange: {
                name: SetupSheetSchema.sheets.calendar.ranges.calendar,
                range: {
                    sheetId: calendarSheetId,
                    startRowIndex: 1,
                    endRowIndex: totalRows,
                    startColumnIndex: 1,
                    endColumnIndex: 15,
                },
            },
        },
    });

    // ============= Protect Calendar ==============

    apiRequests.push({
        addProtectedRange: {
            protectedRange: {
                range: { sheetId: calendarSheetId },
                description: "Calendario",
                warningOnly: false,
            },
        },
    });

    // Execute them all!
    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, setupFileId);
}

/**
 * Copies the SetupFile `fileId` changing the group name and saves it into `folderId`.
 */
export function copySetupFile(setupFileId: string, folderId: string | undefined, groupName: string) {
    const filename = `__Registro ${groupName}`;

    const fileRequest: GoogleAppsScript.Drive_v3.Drive.V3.Schema.File = {
        name: filename,
        appProperties: {
            [FILE_VALIDATION_KEY]: FileType.SETUP,
        },
    };
    if (folderId) {
        fileRequest.parents = [folderId];
    }

    const newFile = Drive?.Files.copy(fileRequest, setupFileId, {
        supportsAllDrives: true,
    });

    const newFileId = newFile?.id;

    if (!newFileId) throw new Error("Error al crear copia del Regirsto Inicial");

    Sheets?.Spreadsheets.Values.update({ values: [[groupName]] }, newFileId, SetupSheetSchema.sheets.groupData.ranges.groupName, { valueInputOption: "USER_ENTERED" });
}

/**
 * Initializes a Report spreadsheet in the same folder as the setupFile, with the information from the Setup file.
 */
export function initializeReport(setupFileId: string, parentId: string) {
    // ========= Data Gathering ===========
    const fieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>(
        "sheets.properties.sheetId",
        "sheets.properties.title",
        "sheets.data.rowData.values.effectiveValue",
        "namedRanges",
    );
    const SetupSpreadsheet = Sheets?.Spreadsheets.get(setupFileId, { fields: fieldsMask });

    const { sheets: setupSheets, namedRanges: setupRanges } = parseSpreadsheet(SetupSpreadsheet, SetupSheetSchema);

    const setupSheetCalendar = setupSheets[SetupSheetSchema.sheets.calendar.sheetName];
    const setupSheetData = setupSheets[SetupSheetSchema.sheets.groupData.sheetName];

    const groupName = setupRanges[SetupSheetSchema.sheets.groupData.ranges.groupName];

    if (!setupSheetCalendar || !setupSheetData || !groupName) throw new Error("Registro Inicial de grupo incompleto.");

    const sanitizedGroupName = sanitizeFileName(`_${MappedNamedRange.getCellText({ mappedRange: groupName })}`);

    // Copy Report Template
    const reportFile = Drive?.Files.copy(
        {
            name: sanitizedGroupName,
            parents: [parentId],
            appProperties: {
                [FILE_VALIDATION_KEY]: FileType.REPORT,
            },
        },
        ReportSheetSchema.templateId,
        {
            supportsAllDrives: true,
        },
    );

    const reportFileId = reportFile?.id;

    if (!reportFileId) throw new Error("Error al copiar template de Reporte");

    const reportSpreadsheet = Sheets?.Spreadsheets.get(reportFileId, { fields: fieldsMask });

    const { sheets: reportSheets, namedRanges: reportRanges } = parseSpreadsheet(reportSpreadsheet, ReportSheetSchema);

    const reportSheetData = reportSheets[ReportSheetSchema.sheets.persistentData.sheetName];
    const reportAssistanceTemplate = reportSheets[ReportSheetSchema.sheets.attendanceTemplate.sheetName];

    if (!reportSheetData || !reportAssistanceTemplate) throw new Error("Archivo de reportes incompleto.");

    // ============ Batch Changes ==============

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // ============ Single cell data copies ================

    const cellMapping = [
        {
            source: SetupSheetSchema.sheets.groupData.ranges.attendancePerClass,
            dest: ReportSheetSchema.sheets.persistentData.ranges.attendancePerClass,
        },
        {
            source: SetupSheetSchema.sheets.groupData.ranges.averagePerField,
            dest: ReportSheetSchema.sheets.persistentData.ranges.averagePerField,
        },
        {
            source: SetupSheetSchema.sheets.groupData.ranges.dateStart,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateStart,
        },
        {
            source: SetupSheetSchema.sheets.groupData.ranges.dateTrim1,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateTrim1,
        },
        {
            source: SetupSheetSchema.sheets.groupData.ranges.dateTrim2,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateTrim2,
        },
        {
            source: SetupSheetSchema.sheets.groupData.ranges.dateEnd,
            dest: ReportSheetSchema.sheets.persistentData.ranges.dateEnd,
        },
    ] as const;

    const singleCellRequests: GoogleAppsScript.Sheets.Schema.Request[] = cellMapping.map(({ source, dest }) => ({
        repeatCell: {
            cell: { userEnteredValue: MappedNamedRange.getCellEffectiveValue({ mappedRange: setupRanges[source] }) },
            range: reportRanges[dest]?.range,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        },
    }));

    apiRequests.push(...singleCellRequests);

    // ============ Execute Batch update ===========

    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, reportFileId);
}

/**
 * Helper function to extract the fields and subjects from the setup page.
 */
function extractFieldsAndSubjects(rawData: GoogleAppsScript.Sheets.Schema.CellData[][]): {
    fields: GoogleAppsScript.Sheets.Schema.CellData[][];
    subjects: GoogleAppsScript.Sheets.Schema.CellData[][];
} {
    const fields: GoogleAppsScript.Sheets.Schema.CellData[][] = [];
    const subjects: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    // Track the active group as we iterate through the rows
    let currentGroupInfo: { name: string; color: string; count: number } | null = null;

    for (let r = 0; r < rawData.length; r++) {
        const row = rawData[r];
        if (!row || row.length === 0) continue;

        // Check for a Group in the first column
        const groupCell = row[0];
        const possibleGroupVal = (groupCell?.effectiveValue?.stringValue || "").trim();

        if (possibleGroupVal !== "") {
            // Before starting a new group, save the previous one IF it had elements
            if (currentGroupInfo && currentGroupInfo.count > 0) {
                fields.push([
                    { userEnteredValue: { stringValue: currentGroupInfo.color } },
                    { userEnteredValue: { stringValue: currentGroupInfo.name } },
                    { userEnteredValue: { numberValue: currentGroupInfo.count } },
                ]);
            }

            // Initialize the new group
            const bgColor = groupCell?.effectiveFormat?.backgroundColor;
            currentGroupInfo = {
                name: possibleGroupVal,
                color: colorToHex(bgColor),
                count: 0,
            };
        }
        if (currentGroupInfo) {
            // Check for an Element in the remaining columns
            let foundElementVal = "";
            for (let c = 1; c < row.length; c++) {
                const cellVal = (row[c]?.effectiveValue?.stringValue || "").trim();
                if (cellVal !== "") {
                    foundElementVal = cellVal;
                    break; // Found the element, stop scanning this row
                }
            }

            // If an element exists, add it to subjects and increment the group count
            if (foundElementVal !== "") {
                subjects.push([{ userEnteredValue: { stringValue: foundElementVal } }]);
                currentGroupInfo.count++;
            }
        }
    }

    // 4. End of data: Make sure to push the final group if it had elements
    if (currentGroupInfo && currentGroupInfo.count > 0) {
        fields.push([
            { userEnteredValue: { stringValue: currentGroupInfo.color } },
            { userEnteredValue: { stringValue: currentGroupInfo.name } },
            { userEnteredValue: { numberValue: currentGroupInfo.count } },
        ]);
    }

    return { fields, subjects };
}
