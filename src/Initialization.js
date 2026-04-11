/**
 * Callback to the button to create a new Initialization Group File.
 * @param {GoogleAppsScript.Addons.EventObject} e 
 * @returns {GoogleAppsScript.Card_Service.ActionResponse}
 */
function onCreateInitializationFile(e) {
    const formInput = e.commonEventObject.formInputs;
    const folderId = e.commonEventObject.parameters.destinationFolder;

    const groupName = Utils.sanitizeFileName(formInput.groupName.stringInputs?.value[0]);

    const attendancePerClass = formInput.attendancePerClass ? true : false;
    const averagePerField = formInput.averagePerField ? true : false;

    const dateStart = parseInt(formInput.dateStart.dateInput?.msSinceEpoch ?? "");
    const dateEndTrimester1 = parseInt(formInput.dateEndTrimester1.dateInput?.msSinceEpoch ?? "");
    const dateEndTrimester2 = parseInt(formInput.dateEndTrimester2.dateInput?.msSinceEpoch ?? "");
    const dateEnd = parseInt(formInput.dateEnd.dateInput?.msSinceEpoch ?? "");

    if (!(dateStart < dateEndTrimester1 && dateEndTrimester1 < dateEndTrimester2 && dateEndTrimester2 < dateEnd)) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
                .setText("❌ Las fechas deben estar en orden ascendente."))
            .build();
    }

    const moreThanAYear = 400 * 24 * 60 * 60 * 1000;
    if (dateEnd - dateStart > moreThanAYear) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
                .setText("❌ Periodo demasiado largo."))
            .build();
    }

    /** @type {InitFileData} */
    const initData = {
        folderId,
        groupName,
        attendancePerClass,
        averagePerField,
        dateStart,
        dateEndTrimester1,
        dateEndTrimester2,
        dateEnd,
    }

    // TODO: Surrond by try-catch, show a card with the error if error.
    Initialization.createInitializationFile(initData);
    /*
    const trigger = ScriptApp.newTrigger(fireCreateInitializationFile.name)
        .timeBased()
        .after(1)
        .create();

    const triggerId = trigger.getUniqueId();
    PropertiesService.getUserProperties().setProperty(triggerId, JSON.stringify(initData));
    */

    // TODO: Extract this into a generic function buildParagraphCard(header, htmlText)
    const successCard = CardService.newCardBuilder()
        .setHeader(CardParts.headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", image: "school" }))
        .addSection(CardService.newCardSection()
            .addWidget(CardService.newTextParagraph()
                .setText(`✅ Creación del Registro para "<b>${groupName}</b>" en proceso.<br><br>El archivo aparecerá la carpeta de Drive en un momento.<br><br>(Por favor espera al menos 1 minúto antes de intentarlo de nuevo.)`))
            .addWidget(CardService.newTextButton()
                .setText("Regresar al inicio")
                .setOnClickAction(CardService.newAction().setFunctionName(onPopCardStack.name))))
        .build();

    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().pushCard(successCard))
        .build();
}

// TODO: Move to file with generic reusable functions.
/**
 * Pops the Card Stack to the root.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse}
 */
function onPopCardStack() {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popToRoot()).build();
}

/**
 * Trigger callback to create a new Initialization Group File.
 * @param {GoogleAppsScript.Events.TimeDriven} e 
 */
function fireCreateInitializationFile(e) {
    const triggerId = e.triggerUid;
    if (!triggerId) return;

    const userProperties = PropertiesService.getUserProperties();
    const payloadString = userProperties.getProperty(triggerId);

    userProperties.deleteProperty(triggerId);
    ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getUniqueId() === triggerId) {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    if (!payloadString) return;
    /** @type {InitFileData} */
    const initData = JSON.parse(payloadString);

    try {
        Initialization.createInitializationFile(initData);
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error(`Error creando Registro para "${initData.groupName}": ${errorMessage}`);
    }
}


const Initialization = {
    /**
     * Presents the user with a form to fill and a button to create the initialization file.
     * 
     * @param {string} folderId The Drive folder where the Initialization file will be created.
     * @returns {GoogleAppsScript.Card_Service.Card}
     */
    buildCreateInitializationCard(folderId) {
        const card = CardService.newCardBuilder()
            .setHeader(CardParts.headerImage({ title: "Registro Inicial de Grupos", subtitle: "Montessory Chacala", image: "school" }));

        const section = CardService.newCardSection().setHeader("Información del Grupo:");

        section.addWidget(CardService.newTextInput()
            .setValue("Secundaria")
            .setFieldName("groupName")
            .setTitle("Nombre del Grupo")
            .setHint("Ejemplo: 5to y 6to"));

        section.addWidget(CardService.newDecoratedText()
            .setText("Asistencia individual por materia")
            .setSwitchControl(CardService.newSwitch()
                .setFieldName("attendancePerClass")
                .setValue("attendancePerClass")
                .setSelected(false)));

        section.addWidget(CardService.newDecoratedText()
            .setText("Promedios por Campo Formativo")
            .setSwitchControl(CardService.newSwitch()
                .setFieldName("averagePerField")
                .setValue("averagePerField")
                .setSelected(false)));

        section.addWidget(CardService.newDatePicker()
            .setValueInMsSinceEpoch(1788220800000)
            .setFieldName("dateStart")
            .setTitle("Primer dia de clases"));

        section.addWidget(CardService.newDatePicker()
            .setValueInMsSinceEpoch(1793491200000)
            .setFieldName("dateEndTrimester1")
            .setTitle("Último día del primer trimestre"));

        section.addWidget(CardService.newDatePicker()
            .setValueInMsSinceEpoch(1797292800000)
            .setFieldName("dateEndTrimester2")
            .setTitle("Último día del segundo trimestre"));

        section.addWidget(CardService.newDatePicker()
            .setValueInMsSinceEpoch(1812326400000)
            .setFieldName("dateEnd")
            .setTitle("Último día de clases"));

        const createAction = CardService.newAction()
            .setFunctionName(onCreateInitializationFile.name)
            .setParameters({ destinationFolder: folderId })
            .addRequiredWidget("groupName")
            .addRequiredWidget("dateStart")
            .addRequiredWidget("dateEndTrimester1")
            .addRequiredWidget("dateEndTrimester2")
            .addRequiredWidget("dateEnd");

        section.addWidget(CardService.newTextButton()
            .setText("📋 Crear Registro Inicial del Grupo")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(createAction));

        return card.addSection(section).build();
    },

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

        const calendarFormatResponse = Sheets.Spreadsheets.get(newFileId, {
            ranges: ["CalendarioTemplate!A1:O19"],
            includeGridData: true
        })

        const calendarFormatData = calendarFormatResponse.sheets?.[0]?.data?.[0];
        if (!calendarFormatData) throw new Error("Error al crear calendario: No se encuentra la hoja \"CalendarioTemplate\"");

        const templateSheetId = calendarFormatResponse.sheets?.[0].properties?.sheetId ?? 0;

        const columnWidths = {
            month: calendarFormatData.columnMetadata?.[0].pixelSize ?? 100,
            number: calendarFormatData.columnMetadata?.[1].pixelSize ?? 100,
            checkmark: calendarFormatData.columnMetadata?.[2].pixelSize ?? 100,
        }

        const rowHights = {
            header: calendarFormatData.rowMetadata?.[0].pixelSize ?? 21,
            day: calendarFormatData.rowMetadata?.[1].pixelSize ?? 21,
        }

        // TODO: Do I need this? Couldn't I do a NORMAL_COPY of the full header, to get everything, including format.
        const calendarHeaders = calendarFormatData.rowData?.[0].values?.map(cellData => cellData.formattedValue ?? "");

        /** @type {string[]} */
        const monthNamesHigh3 = [];
        /** @type {string[]} */
        const monthNamesHigh2 = [];
        /** @type {string[]} */
        const monthNamesHigh1 = [];
        for (let i = 7; i < 19; i++) {
            const row = calendarFormatData.rowData?.[i];
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
        const dateStartDayOfWeek = new Date(initData.dateStart).getUTCDay();
        const calStart = initData.dateStart - (dateStartDayOfWeek * msPerDay);

        // Snap to last Saturday
        const dateEndDayOfWeek = new Date(initData.dateEnd).getUTCDay();
        const calEnd = initData.dateEnd + ((6 - dateEndDayOfWeek) * msPerDay);

        const totalDays = ((calEnd - calStart) / msPerDay) + 1;
        const totalRows = 1 + Math.ceil(totalDays / 7);

        // Create new sheet
        const calendarSheetId = Math.floor(Math.random() * (Math.pow(2, 31) - 1));

        /** @type {GoogleAppsScript.Sheets.Schema.Request[]} */
        const apiRequests = [];

        apiRequests.push({
            addSheet: {
                properties: {
                    sheetId: calendarSheetId,
                    title: "Calendario",
                    gridProperties: {
                        rowCount: totalRows,
                        columnCount: 15,
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
                    monthBlocks.push({ startRow: monthStartRow, endRow: currentRowNumber, monthIndex, year });
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
                const inBounds = currentMs >= initData.dateStart && currentMs <= initData.dateEnd;

                rowCells.push({ userEnteredValue: { numberValue: dayDate.getUTCDate() } });
                rowCells.push({ userEnteredValue: { boolValue: (inBounds && isWeekday) }, dataValidation: { condition: { type: "BOOLEAN" }, strict: true, showCustomUi: true } });

                // TODO: Some magic numbers here, it would be best to have them elsewhere.
                // Format the cells
                let formatRowIndex = 2; // Default to trimester 1;
                if (currentMs > initData.dateEndTrimester2) formatRowIndex = 4;
                else if (currentMs > initData.dateEndTrimester1) formatRowIndex = 3;

                const formatColumnIndex = isWeekday ? 3 : 7; // Column D and H

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

        Sheets.Spreadsheets.batchUpdate({ requests: apiRequests }, newFileId);
    }
}
