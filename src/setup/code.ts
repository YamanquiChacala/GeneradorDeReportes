import { FileType, Numbers } from "../common/enums";
import { SetupSheet } from "../common/sheetSchema";
import { key as FILE_VALIDATION_KEY } from "../common/utils/fileValidation";
import { getDateMs } from "../common/utils/googleAPI";

export interface InitFileData {
    folderId: string;
    [SetupSheet.namedRanges.groupName]: string;
    [SetupSheet.namedRanges.attendancePerClass]: boolean;
    [SetupSheet.namedRanges.averagePerField]: boolean;
    [SetupSheet.namedRanges.dateStart]: number;
    [SetupSheet.namedRanges.dateEndTrimester1]: number;
    [SetupSheet.namedRanges.dateEndTrimester2]: number;
    [SetupSheet.namedRanges.dateEnd]: number;
}

/**
 * Creates a new Group Initialization file with the given data.
 */
export function createInitializationFile(initData: InitFileData) {
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
        SetupSheet.templateId,
        {
            supportsAllDrives: true,
        },
    );

    const newFileId = newFile?.id;

    if (!newFileId) throw new Error("Error al crear copia del Regirsto Inicial");

    // ======== Update namedRanges ==========

    const namedRanges: Array<keyof InitFileData> = [
        SetupSheet.namedRanges.groupName,
        SetupSheet.namedRanges.attendancePerClass,
        SetupSheet.namedRanges.averagePerField,
        SetupSheet.namedRanges.dateStart,
        SetupSheet.namedRanges.dateEndTrimester1,
        SetupSheet.namedRanges.dateEndTrimester2,
        SetupSheet.namedRanges.dateEnd,
    ];

    const updateData: GoogleAppsScript.Sheets.Schema.ValueRange[] = namedRanges.map((name) => {
        let cellValue = initData[name];
        if (name.startsWith("date")) {
            const date = new Date(cellValue as number);
            cellValue = Utilities.formatDate(date, "UTC", "yyyy-MM-dd");
        }
        return {
            range: name,
            values: [[cellValue]],
        };
    });

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
    const registroSpreadsheet = Sheets?.Spreadsheets.get(fileId, {
        fields: "sheets(properties(sheetId,title),data(rowMetadata/pixelSize,columnMetadata/pixelSize,rowData/values(formattedValue,effectiveValue/numberValue))),namedRanges(name,range)",
    });

    const dataSheet = registroSpreadsheet?.sheets?.find((s) => s.properties?.title === SetupSheet.sheetNames.groupData);
    const calendarTemplateSheet = registroSpreadsheet?.sheets?.find((s) => s.properties?.title === SetupSheet.sheetNames.calendarTemplate);
    const calendarSheet = registroSpreadsheet?.sheets?.find((s) => s.properties?.title === SetupSheet.sheetNames.calendar);

    const templateSheetId = calendarTemplateSheet?.properties?.sheetId;

    if (!dataSheet || !calendarTemplateSheet || !registroSpreadsheet?.namedRanges) throw new Error("Faltan hojas de datos o formato.");

    const namedRanges = registroSpreadsheet.namedRanges.reduce(
        (acc, namedRange) => {
            if (namedRange.name) {
                acc[namedRange.name] = namedRange;
            }
            return acc;
        },
        {} as Record<string, GoogleAppsScript.Sheets.Schema.NamedRange>,
    );

    const dateStart = getDateMs(namedRanges[SetupSheet.namedRanges.dateStart]?.range, dataSheet);
    const dateEndTrimester1 = getDateMs(namedRanges[SetupSheet.namedRanges.dateEndTrimester1]?.range, dataSheet);
    const dateEndTrimester2 = getDateMs(namedRanges[SetupSheet.namedRanges.dateEndTrimester2]?.range, dataSheet);
    const dateEnd = getDateMs(namedRanges[SetupSheet.namedRanges.dateEnd]?.range, dataSheet);

    if (!dateStart || !dateEndTrimester1 || !dateEndTrimester2 || !dateEnd) throw new Error("Faltan las fechas.");
    if (dateStart >= dateEndTrimester1 || dateEndTrimester1 >= dateEndTrimester2 || dateEndTrimester2 >= dateEnd) throw new Error("Fechas en desorden.");
    if (dateEnd - dateStart > Numbers.MORE_THAN_A_YEAR) throw new Error("Calendario demasiado grande.");

    const columnWidths = {
        month: calendarTemplateSheet.data?.[0]?.columnMetadata?.[0]?.pixelSize ?? 100,
        number: calendarTemplateSheet.data?.[0]?.columnMetadata?.[1]?.pixelSize ?? 100,
        checkmark: calendarTemplateSheet.data?.[0]?.columnMetadata?.[2]?.pixelSize ?? 100,
    };

    const rowHights = {
        header: calendarTemplateSheet.data?.[0]?.rowMetadata?.[0]?.pixelSize ?? 21,
        day: calendarTemplateSheet.data?.[0]?.rowMetadata?.[1]?.pixelSize ?? 21,
    };

    const monthNamesRowStart = namedRanges[SetupSheet.namedRanges.monthNames3]?.range?.startRowIndex ?? 0;
    const monthNamesRowEnd = namedRanges[SetupSheet.namedRanges.monthNames3]?.range?.endRowIndex ?? 0;
    const monthNamesColHigh3 = namedRanges[SetupSheet.namedRanges.monthNames3]?.range?.startColumnIndex ?? 0;
    const monthNamesColHigh2 = namedRanges[SetupSheet.namedRanges.monthNames2]?.range?.startColumnIndex ?? 0;
    const monthNamesColHigh1 = namedRanges[SetupSheet.namedRanges.monthNames1]?.range?.startColumnIndex ?? 0;

    const monthNamesHigh3: string[] = [];
    const monthNamesHigh2: string[] = [];
    const monthNamesHigh1: string[] = [];
    for (let i = monthNamesRowStart; i <= monthNamesRowEnd; i++) {
        const row = calendarTemplateSheet.data?.[0]?.rowData?.[i];
        const cellHigh3 = row?.values?.[monthNamesColHigh3];
        const cellHigh2 = row?.values?.[monthNamesColHigh2];
        const cellHigh1 = row?.values?.[monthNamesColHigh1];
        monthNamesHigh3.push(cellHigh3?.formattedValue ?? "");
        monthNamesHigh2.push(cellHigh2?.formattedValue ?? "");
        monthNamesHigh1.push(cellHigh1?.formattedValue ?? "");
    }

    // Date calculations
    const msPerDay = 24 * 60 * 60 * 1000;

    // Snap to first Sunday
    const dateStartDayOfWeek = new Date(dateStart).getUTCDay();
    const calStart = dateStart - dateStartDayOfWeek * msPerDay;

    // Snap to last Saturday
    const dateEndDayOfWeek = new Date(dateEnd).getUTCDay();
    const calEnd = dateEnd + (6 - dateEndDayOfWeek) * msPerDay;

    const totalDays = (calEnd - calStart) / msPerDay + 1;
    const totalRows = 1 + Math.ceil(totalDays / 7);

    // Create new sheet
    const calendarSheetId = Math.floor(Math.random() * (2 ** 31 - 1));

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    if (calendarSheet) {
        apiRequests.push({
            deleteSheet: {
                sheetId: calendarSheet.properties?.sheetId,
            },
        });
    }

    apiRequests.push({
        addSheet: {
            properties: {
                sheetId: calendarSheetId,
                title: SetupSheet.sheetNames.calendar,
                gridProperties: {
                    rowCount: totalRows,
                    columnCount: 15,
                    frozenRowCount: 1,
                    frozenColumnCount: 1,
                    hideGridlines: true,
                },
            },
        },
    });

    // Add column sizes:
    apiRequests.push({
        updateDimensionProperties: {
            range: { sheetId: calendarSheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 1 },
            properties: { pixelSize: columnWidths.month },
            fields: "pixelSize",
        },
    });
    for (let i = 1; i < 8; i++) {
        apiRequests.push({
            updateDimensionProperties: {
                range: { sheetId: calendarSheetId, dimension: "COLUMNS", startIndex: 2 * i - 1, endIndex: 2 * i },
                properties: { pixelSize: columnWidths.number },
                fields: "pixelSize",
            },
        });
        apiRequests.push({
            updateDimensionProperties: {
                range: { sheetId: calendarSheetId, dimension: "COLUMNS", startIndex: 2 * i, endIndex: 2 * i + 1 },
                properties: { pixelSize: columnWidths.checkmark },
                fields: "pixelSize",
            },
        });
    }

    // Add row sizes:
    apiRequests.push({
        updateDimensionProperties: {
            range: { sheetId: calendarSheetId, dimension: "ROWS", startIndex: 0, endIndex: 1 },
            properties: { pixelSize: rowHights.header },
            fields: "pixelSize",
        },
    });

    apiRequests.push({
        updateDimensionProperties: {
            range: { sheetId: calendarSheetId, dimension: "ROWS", startIndex: 1, endIndex: totalRows + 1 },
            properties: { pixelSize: rowHights.day },
            fields: "pixelSize",
        },
    });

    // Do every day of the calendar

    const formatTrimester1Row = namedRanges[SetupSheet.namedRanges.trimester1Day]?.range?.startRowIndex ?? 0;
    const formatTrimester2Row = namedRanges[SetupSheet.namedRanges.trimester2Day]?.range?.startRowIndex ?? 0;
    const formatTrimester3Row = namedRanges[SetupSheet.namedRanges.trimester3Day]?.range?.startRowIndex ?? 0;

    const formatTrimesterCol = namedRanges[SetupSheet.namedRanges.trimester1Day]?.range?.startColumnIndex ?? 0;
    const formatFreeDayCol = namedRanges[SetupSheet.namedRanges.restDay]?.range?.startColumnIndex ?? 0;

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
            let formatRowIndex = formatTrimester1Row; // Default to trimester 1;
            if (currentMs > dateEndTrimester2) formatRowIndex = formatTrimester3Row;
            else if (currentMs > dateEndTrimester1) formatRowIndex = formatTrimester2Row;

            const formatColumnIndex = isWeekday && inBounds ? formatTrimesterCol : formatFreeDayCol; // Column D and H

            apiRequests.push({
                copyPaste: {
                    source: {
                        sheetId: templateSheetId,
                        startRowIndex: formatRowIndex,
                        endRowIndex: formatRowIndex + 1,
                        startColumnIndex: formatColumnIndex,
                        endColumnIndex: formatColumnIndex + 2,
                    },
                    destination: {
                        sheetId: calendarSheetId,
                        startRowIndex: currentRowNumber,
                        endRowIndex: currentRowNumber + 1,
                        startColumnIndex: i * 2 + 1,
                        endColumnIndex: i * 2 + 3,
                    },
                    pasteType: "PASTE_FORMAT",
                },
            });
            currentMs += msPerDay;
        }
        rowDataArray.push({ values: rowCells });
        currentRowNumber++;
    }

    monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });

    // Handle month and year labels
    monthBlocks.forEach((block) => {
        const rowSpan = block.endRow - block.startRow;
        let monthLabelColumn = monthNamesColHigh3;

        let monthName = monthNamesHigh3[block.monthIndex];

        if (rowSpan === 2) {
            monthLabelColumn = monthNamesColHigh2;
            monthName = monthNamesHigh2[block.monthIndex];
        } else if (rowSpan === 1) {
            monthLabelColumn = monthNamesColHigh1;
            monthName = monthNamesHigh1[block.monthIndex];
        }

        const monthYearText = `${monthName}\n${block.year}`;

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
        //Format
        apiRequests.push({
            copyPaste: {
                source: {
                    sheetId: templateSheetId,
                    startRowIndex: monthNamesRowStart + block.monthIndex,
                    endRowIndex: monthNamesRowStart + block.monthIndex + 1,
                    startColumnIndex: monthLabelColumn,
                    endColumnIndex: monthLabelColumn + 1,
                },
                destination: {
                    sheetId: calendarSheetId,
                    startRowIndex: block.startRow,
                    endRowIndex: block.endRow,
                    startColumnIndex: 0,
                    endColumnIndex: 1,
                },
                pasteType: "PASTE_FORMAT",
            },
        });
        //Insert text in the data array.
        const targetRowData = rowDataArray[block.startRow - 1];
        if (targetRowData?.values) {
            targetRowData.values[0] = { userEnteredValue: { stringValue: monthYearText } };
        }
    });

    // Update all the cell values.
    apiRequests.push({
        updateCells: {
            rows: rowDataArray,
            fields: "userEnteredValue,dataValidation",
            start: { sheetId: calendarSheetId, rowIndex: 1, columnIndex: 0 },
        },
    });

    // Copy and paste header (At the end to overwrite borders of adjacent cells.)
    apiRequests.push({
        copyPaste: {
            source: { sheetId: templateSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 15 },
            destination: { sheetId: calendarSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 15 },
            pasteType: "PASTE_NORMAL",
        },
    });

    // Store the start of the calendar in the metadata
    apiRequests.push({
        createDeveloperMetadata: {
            developerMetadata: {
                metadataKey: "calStart",
                metadataValue: calStart.toString(),
                location: {
                    sheetId: calendarSheetId,
                },
                visibility: "PROJECT",
            },
        },
    });

    // Protect the Calendar Sheet
    const unprotectedRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];
    for (let i = 2; i < 15; i += 2) {
        unprotectedRanges.push({
            sheetId: calendarSheetId,
            startRowIndex: 1,
            startColumnIndex: i,
            endColumnIndex: i + 1,
        });
    }
    apiRequests.push({
        addProtectedRange: {
            protectedRange: {
                range: { sheetId: calendarSheetId },
                description: "Calendario",
                warningOnly: true,
                unprotectedRanges: unprotectedRanges,
            },
        },
    });

    // Clean up
    apiRequests.push({
        updateSheetProperties: {
            properties: {
                sheetId: templateSheetId,
                hidden: true,
            },
            fields: "hidden",
        },
    });

    // Execute them all!
    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, fileId);
}
