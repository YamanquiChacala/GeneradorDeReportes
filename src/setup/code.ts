import { FileType, Numbers } from "../common/enums";
import { SetupSheetSchema } from "../common/sheetSchema";
import { key as FILE_VALIDATION_KEY } from "../common/utils/fileValidation";
import { defineRangesDataConfig, type MappedInput, MappedNamedRange, PasteType, parseSpreadsheet } from "../common/utils/googleAPI";

const SetupFileDataConfig = defineRangesDataConfig({
    groupName: { range: SetupSheetSchema.namedRanges.groupName, type: "string" },
    attendancePerClass: { range: SetupSheetSchema.namedRanges.attendancePerClass, type: "boolean" },
    averagePerField: { range: SetupSheetSchema.namedRanges.averagePerField, type: "boolean" },
    dateStart: { range: SetupSheetSchema.namedRanges.dateStart, type: "date" },
    dateEndTrimester1: { range: SetupSheetSchema.namedRanges.dateTrimester1, type: "date" },
    dateEndTrimester2: { range: SetupSheetSchema.namedRanges.dateTrimester2, type: "date" },
    dateEnd: { range: SetupSheetSchema.namedRanges.dateEnd, type: "date" },
} as const);

export type SetupFileData = { folderId: string } & {
    [K in keyof typeof SetupFileDataConfig]: MappedInput<(typeof SetupFileDataConfig)[K]["type"]>;
};

/**
 * Creates a new Group Initialization file with the given data.
 */
export function createSetupFile(initData: SetupFileData) {
    // ========== Create File ============

    const fileName = `__Registro ${initData.groupName}`;

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
export function generateCalendar(fileId: string) {
    const SetupSpreadsheet = Sheets?.Spreadsheets.get(fileId, {
        fields: "sheets(properties(sheetId,title),data(rowData/values(formattedValue,effectiveValue/numberValue))),namedRanges",
    });

    const { sheets, namedRanges } = parseSpreadsheet(SetupSpreadsheet, SetupSheetSchema);

    const calendarTemplateSheet = sheets[SetupSheetSchema.sheetNames.calendarTemplate];
    const calendarSheet = sheets[SetupSheetSchema.sheetNames.calendar];

    if (!calendarTemplateSheet) throw new Error("Faltan hojas o formato.");

    const dateStart = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.namedRanges.dateStart] });
    const dateTrimester1 = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.namedRanges.dateTrimester1] });
    const dateTrimester2 = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.namedRanges.dateTrimester2] });
    const dateEnd = MappedNamedRange.getCellUnixEpoch({ mappedRange: namedRanges[SetupSheetSchema.namedRanges.dateEnd] });

    if (!dateStart || !dateTrimester1 || !dateTrimester2 || !dateEnd) throw new Error("Faltan las fechas.");
    if (dateStart >= dateTrimester1 || dateTrimester1 >= dateTrimester2 || dateTrimester2 >= dateEnd) throw new Error("Fechas en desorden.");
    if (dateEnd - dateStart > Numbers.MORE_THAN_A_YEAR) throw new Error("Calendario demasiado grande.");

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
    if (calendarSheet) {
        apiRequests.push({
            deleteSheet: {
                sheetId: calendarSheet.properties?.sheetId,
            },
        });
    }

    // Tempraraty remove named ranges so they don't get duplicated
    const templateNamedRanges = (SetupSpreadsheet?.namedRanges ?? []).filter((namedRange) => namedRange.range?.sheetId === calendarTemplateSheet.properties?.sheetId);

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
            newSheetName: SetupSheetSchema.sheetNames.calendar,
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
            fields: "hidden,gridProperties(rowCount,frozenRowCount,frozenColumnCount)",
        },
    });

    // Hide template
    apiRequests.push({
        updateSheetProperties: {
            properties: {
                sheetId: calendarTemplateSheet.properties?.sheetId,
                hidden: true,
            },
            fields: "hidden",
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
                formatSource = namedRanges[SetupSheetSchema.namedRanges.trimester1Day]; // Default to trimester 1;
                if (currentMs > dateTrimester2) formatSource = namedRanges[SetupSheetSchema.namedRanges.trimester3Day];
                else if (currentMs > dateTrimester1) formatSource = namedRanges[SetupSheetSchema.namedRanges.trimester2Day];
            } else {
                formatSource = namedRanges[SetupSheetSchema.namedRanges.restDay];
            }

            const formatRequest = MappedNamedRange.buildCopyPasteRequest({
                mappedRange: formatSource,
                destinationSheetId: calendarSheetId,
                destinationStartRow: currentRowNumber,
                destinationStartColumn: 2 * i + 1,
                pasteType: PasteType.PASTE_FORMAT,
            });
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

        let monthLabelRange = namedRanges[SetupSheetSchema.namedRanges.monthNames3];

        if (rowSpan === 2) {
            monthLabelRange = namedRanges[SetupSheetSchema.namedRanges.monthNames2];
        } else if (rowSpan === 1) {
            monthLabelRange = namedRanges[SetupSheetSchema.namedRanges.monthNames1];
        }

        const monthName = MappedNamedRange.getCellDisplay({ mappedRange: monthLabelRange, row: block.monthIndex });
        const monthYearText = `${monthName ?? ""}\n${block.year}`;

        const formatRequest = MappedNamedRange.buildCopyPasteRequest({
            mappedRange: monthLabelRange,
            destinationSheetId: calendarSheetId,
            destinationStartRow: block.startRow,
            destinationStartColumn: 0,
            offsetRow: block.monthIndex,
            height: 1,
            width: 1,
            pasteType: PasteType.PASTE_FORMAT,
        });
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
            fields: "userEnteredValue,dataValidation",
            start: { sheetId: calendarSheetId, rowIndex: 1, columnIndex: 0 },
        },
    });

    // Execute them all!
    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, fileId);
}
