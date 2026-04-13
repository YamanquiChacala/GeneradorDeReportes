const Initialization = {
    /**
     * Creates a new Group Initialization file with the given data.
     * @param {InitFileData} initData
     */
    createInitializationFile(initData) {

        // ========== Create File ============

        const fileName = `__Registro ${initData.groupName}`;

        const newFile = Drive.Files.copy({
            name: fileName,
            parents: [initData.folderId],
            appProperties: {
                [FileValidation.KEY]: FileType.setup
            }
        }, DriveFiles.initializationTemplateId, {
            supportsAllDrives: true
        });

        const newFileId = newFile.id;

        if (!newFileId) throw new Error("Error al copiar el Regirsto Inicial");


        // ======== Update namedRanges ==========

        /** @type {Array<keyof InitFileData>} */
        const namedRanges = ["groupName", "attendancePerClass", "averagePerField", "dateStart", "dateEndTrimester1", "dateEndTrimester2", "dateEnd"];

        /** @type {GoogleAppsScript.Sheets.Schema.ValueRange[]} */
        const updateData = namedRanges.map(name => {
            let cellValue = initData[name];
            if (name.startsWith("date")) {
                const date = new Date(/** @type {number} */(cellValue));
                cellValue = Utilities.formatDate(date, "UTC", "yyyy-MM-dd");
            }
            return {
                range: name,
                values: [[cellValue]],
            }
        });

        Sheets.Spreadsheets.Values.batchUpdate({
            valueInputOption: "USER_ENTERED",
            data: updateData,
        }, newFileId)

        // ========= Create Calendar ==========
        Initialization.generateCalendar(newFileId);

    },

    /**
     * 
     * @param {GoogleAppsScript.Sheets.Schema.GridRange} [range] 
     * @param {GoogleAppsScript.Sheets.Schema.Sheet} [sheet] 
     * @returns {number | null}
     */
    getDateMs(range, sheet) {
        if (!range?.startRowIndex || !range.startColumnIndex) return null;

        const cell = sheet?.data?.[0]?.rowData?.[range.startRowIndex]?.values?.[range.startColumnIndex];
        const rawNumber = cell?.effectiveValue?.numberValue;

        if (!rawNumber) return null;

        const msPerDay = 24 * 60 * 60 * 1000;
        const sheetsEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
        return sheetsEpoch + (rawNumber * msPerDay);
    },

    /**
     * Generates a Calendar in the given Spreadsheet
     * The spreadsheet must be a Setup Group.
     * @param {string} fileId 
     */
    generateCalendar(fileId) {
        const registroSpreadsheet = Sheets.Spreadsheets.get(fileId, {
            fields: "sheets(properties(sheetId,title),data(rowMetadata/pixelSize,columnMetadata/pixelSize,rowData/values(formattedValue,effectiveValue/numberValue))),namedRanges(name,range)"
        });

        const dataSheet = registroSpreadsheet.sheets?.find(s => s.properties?.title === "Registro Inicial");
        const calendarTemplateSheet = registroSpreadsheet.sheets?.find(s => s.properties?.title === "CalendarioTemplate");
        const calendarSheet = registroSpreadsheet.sheets?.find(s => s.properties?.title === "Calendario");

        const templateSheetId = calendarTemplateSheet?.properties?.sheetId;

        if (!dataSheet || !calendarTemplateSheet || !registroSpreadsheet.namedRanges) throw new Error("Error al crear calendario: Faltan hojas de datos o formato.");

        const namedRanges = registroSpreadsheet.namedRanges.reduce((acc, namedRange) => {
            if (namedRange.name) {
                acc[namedRange.name] = namedRange;
            }
            return acc;
        },/** @type {Record<string, GoogleAppsScript.Sheets.Schema.NamedRange>} */({}));

        const dateStart = Initialization.getDateMs(namedRanges["dateStart"]?.range, dataSheet);
        const dateEndTrimester1 = Initialization.getDateMs(namedRanges["dateEndTrimester1"]?.range, dataSheet);
        const dateEndTrimester2 = Initialization.getDateMs(namedRanges["dateEndTrimester2"]?.range, dataSheet);
        const dateEnd = Initialization.getDateMs(namedRanges["dateEnd"]?.range, dataSheet);

        if (!dateStart || !dateEndTrimester1 || !dateEndTrimester2 || !dateEnd) return new Error("Error al crear calendario: Faltan las fechas.")

        const columnWidths = {
            month: calendarTemplateSheet.data?.[0]?.columnMetadata?.[0]?.pixelSize ?? 100,
            number: calendarTemplateSheet.data?.[0]?.columnMetadata?.[1]?.pixelSize ?? 100,
            checkmark: calendarTemplateSheet.data?.[0]?.columnMetadata?.[2]?.pixelSize ?? 100,
        }

        const rowHights = {
            header: calendarTemplateSheet.data?.[0]?.rowMetadata?.[0]?.pixelSize ?? 21,
            day: calendarTemplateSheet.data?.[0]?.rowMetadata?.[1]?.pixelSize ?? 21,
        }

        /** @type {string[]} */
        const monthNamesHigh3 = [];
        /** @type {string[]} */
        const monthNamesHigh2 = [];
        /** @type {string[]} */
        const monthNamesHigh1 = [];
        for (let i = 7; i < 19; i++) {
            const row = calendarTemplateSheet.data?.[0]?.rowData?.[i];
            const cellHigh3 = row?.values?.[0];
            const cellHigh2 = row?.values?.[2];
            const cellHigh1 = row?.values?.[4];
            monthNamesHigh3.push(cellHigh3?.formattedValue ?? "")
            monthNamesHigh2.push(cellHigh2?.formattedValue ?? "")
            monthNamesHigh1.push(cellHigh1?.formattedValue ?? "")
        }

        // Date calculations
        const msPerDay = 24 * 60 * 60 * 1000;

        // Snap to first Sunday
        const dateStartDayOfWeek = new Date(dateStart).getUTCDay();
        const calStart = dateStart - (dateStartDayOfWeek * msPerDay);

        // Snap to last Saturday
        const dateEndDayOfWeek = new Date(dateEnd).getUTCDay();
        const calEnd = dateEnd + ((6 - dateEndDayOfWeek) * msPerDay);

        const totalDays = ((calEnd - calStart) / msPerDay) + 1;
        const totalRows = 1 + Math.ceil(totalDays / 7);

        // Create new sheet
        const calendarSheetId = Math.floor(Math.random() * (Math.pow(2, 31) - 1));

        /** @type {GoogleAppsScript.Sheets.Schema.Request[]} */
        const apiRequests = [];

        if (calendarSheet) {
            apiRequests.push({
                deleteSheet: {
                    sheetId: calendarSheet.properties?.sheetId,
                }
            });
        }

        apiRequests.push({
            addSheet: {
                properties: {
                    sheetId: calendarSheetId,
                    title: "Calendario",
                    gridProperties: {
                        rowCount: totalRows,
                        columnCount: 15,
                        frozenRowCount: 1,
                        frozenColumnCount: 1,
                        hideGridlines: true,
                    }
                }
            }
        });

        // Add column sizes:
        apiRequests.push({
            updateDimensionProperties: {
                range: { sheetId: calendarSheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 1 },
                properties: { pixelSize: columnWidths.month },
                fields: "pixelSize",
            }
        });
        for (let i = 1; i < 8; i++) {
            apiRequests.push({
                updateDimensionProperties: {
                    range: { sheetId: calendarSheetId, dimension: "COLUMNS", startIndex: 2 * i - 1, endIndex: 2 * i },
                    properties: { pixelSize: columnWidths.number },
                    fields: "pixelSize",
                }
            });
            apiRequests.push({
                updateDimensionProperties: {
                    range: { sheetId: calendarSheetId, dimension: "COLUMNS", startIndex: 2 * i, endIndex: 2 * i + 1 },
                    properties: { pixelSize: columnWidths.checkmark },
                    fields: "pixelSize",
                }
            });
        }

        // Add row sizes:
        apiRequests.push({
            updateDimensionProperties: {
                range: { sheetId: calendarSheetId, dimension: "ROWS", startIndex: 0, endIndex: 1 },
                properties: { pixelSize: rowHights.header },
                fields: "pixelSize",
            }
        });

        apiRequests.push({
            updateDimensionProperties: {
                range: { sheetId: calendarSheetId, dimension: "ROWS", startIndex: 1, endIndex: totalRows + 1 },
                properties: { pixelSize: rowHights.day },
                fields: "pixelSize",
            }
        });

        // Do every day of the calendar

        let currentMs = calStart;
        let currentRowNumber = 1;

        let currentMonthIndex = -1;
        let currentYear = -1;
        let monthStartRow = 1;

        /** @type {{startRow: number, endRow: number, monthIndex: number, year: number}[]} */
        const monthBlocks = [];

        /** @type {GoogleAppsScript.Sheets.Schema.RowData[]} */
        const rowDataArray = [];

        while (currentMs <= calEnd) {
            // Handle change of month, this happens on wednesday.
            const wednesdayMs = currentMs + (3 * msPerDay);
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

            /** @type {GoogleAppsScript.Sheets.Schema.CellData[]} */
            const rowCells = [{}]; // Column A is empty for now.

            // Handle each day of the week
            for (let i = 0; i < 7; i++) {
                const dayDate = new Date(currentMs);
                const dayOfWeek = dayDate.getUTCDay();
                const isWeekday = dayOfWeek !== 0 && dayOfWeek !== 6;
                const inBounds = currentMs >= dateStart && currentMs <= dateEnd;

                rowCells.push({ userEnteredValue: { numberValue: dayDate.getUTCDate() } });
                rowCells.push({
                    userEnteredValue: { boolValue: (inBounds && isWeekday) },
                    dataValidation: { condition: { type: "BOOLEAN" }, strict: true, showCustomUi: true },
                });

                // TODO: Some magic numbers here, it would be best to have them elsewhere.
                // Format the cells
                let formatRowIndex = 2; // Default to trimester 1;
                if (currentMs > dateEndTrimester2) formatRowIndex = 4;
                else if (currentMs > dateEndTrimester1) formatRowIndex = 3;

                const formatColumnIndex = isWeekday && inBounds ? 3 : 7; // Column D and H

                apiRequests.push({
                    copyPaste: {
                        source: {
                            sheetId: templateSheetId,
                            startRowIndex: formatRowIndex, endRowIndex: formatRowIndex + 1,
                            startColumnIndex: formatColumnIndex, endColumnIndex: formatColumnIndex + 2
                        },
                        destination: {
                            sheetId: calendarSheetId,
                            startRowIndex: currentRowNumber, endRowIndex: currentRowNumber + 1,
                            startColumnIndex: i * 2 + 1, endColumnIndex: i * 2 + 3
                        },
                        pasteType: "PASTE_FORMAT"
                    }
                });
                currentMs += msPerDay;
            }
            rowDataArray.push({ values: rowCells });
            currentRowNumber++;
        }

        monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex: currentMonthIndex, year: currentYear });

        // Handle month and year labels
        monthBlocks.forEach(block => {
            const rowSpan = block.endRow - block.startRow;
            let monthLabelRow = 7;
            let monthLabelColumn = 0;

            let monthName = monthNamesHigh3[block.monthIndex];

            if (rowSpan === 2) {
                monthLabelColumn = 2;
                monthName = monthNamesHigh2[block.monthIndex];
            } else if (rowSpan === 1) {
                monthLabelColumn = 4;
                monthName = monthNamesHigh1[block.monthIndex];
            }

            const monthYearText = `${monthName}\n${block.year}`

            // Merge
            apiRequests.push({
                mergeCells: {
                    range: {
                        sheetId: calendarSheetId,
                        startRowIndex: block.startRow, endRowIndex: block.endRow,
                        startColumnIndex: 0, endColumnIndex: 1,
                    },
                    mergeType: "MERGE_ALL",
                }
            });
            //Format
            apiRequests.push({
                copyPaste: {
                    source: {
                        sheetId: templateSheetId,
                        startRowIndex: monthLabelRow + block.monthIndex, endRowIndex: monthLabelRow + block.monthIndex + 1,
                        startColumnIndex: monthLabelColumn, endColumnIndex: monthLabelColumn + 1,
                    },
                    destination: {
                        sheetId: calendarSheetId,
                        startRowIndex: block.startRow, endRowIndex: block.endRow,
                        startColumnIndex: 0, endColumnIndex: 1,
                    },
                    pasteType: "PASTE_FORMAT"
                },
            });
            //Insert text in the data array.
            const targetRowData = rowDataArray[block.startRow - 1];
            if (targetRowData.values) {
                targetRowData.values[0] = { userEnteredValue: { stringValue: monthYearText } };
            }
        });

        // Update all the cell values.
        apiRequests.push({
            updateCells: {
                rows: rowDataArray,
                fields: "userEnteredValue,dataValidation",
                start: { sheetId: calendarSheetId, rowIndex: 1, columnIndex: 0 }
            }
        });

        // Copy and paste header (At the end to overwrite borders of adjacent cells.)
        apiRequests.push({
            copyPaste: {
                source: { sheetId: templateSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 15 },
                destination: { sheetId: calendarSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 15 },
                pasteType: "PASTE_NORMAL"
            }
        });

        // Protect the Calendar Sheet
        /** @type {GoogleAppsScript.Sheets.Schema.GridRange[]} */
        const unprotectedRanges = []
        for (let i = 2; i < 15; i += 2) {
            unprotectedRanges.push({
                sheetId: calendarSheetId,
                startRowIndex: 1,
                startColumnIndex: i, endColumnIndex: i + 1
            });
        }
        apiRequests.push({
            addProtectedRange: {
                protectedRange: {
                    range: { sheetId: calendarSheetId },
                    description: "Calendario",
                    warningOnly: true,
                    unprotectedRanges: unprotectedRanges,
                }
            }
        });

        // Clean up
        apiRequests.push({
            updateSheetProperties: {
                properties: {
                    sheetId: templateSheetId,
                    hidden: true,
                },
                fields: "hidden"
            }
        });

        // Execute them all!
        Sheets.Spreadsheets.batchUpdate({ requests: apiRequests }, fileId);
    }
}