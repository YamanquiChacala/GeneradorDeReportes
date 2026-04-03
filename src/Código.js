/**
 * This function builds the homepage card for the Add-on.
 * It will run whenever the Add-on is opened in Drive or Sheets.
 * @param {GoogleAppsScript.Addons.EventObject} e 
 * @returns {GoogleAppsScript.Card_Service.Card}
 */
function buildHelloWorldCard(e) {
    // 1. Create a text widget
    var textParagraph = CardService.newTextParagraph()
        .setText("Hello World! This is my first Workspace Add-on.");

    // 2. Create a section and add the widget to it
    var section = CardService.newCardSection()
        .addWidget(textParagraph);

    // 3. Build the card, add a header, and attach the section
    var card = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Report Card Automator"))
        .addSection(section)
        .build();

    // 4. Return the card so Google Workspace knows what to display
    return card;
}

/**
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e
 * @returns {GoogleAppsScript.Card_Service.Card} 
 */
function buildDriveCard(e) {
    const builder = CardService.newCardBuilder();
    const section = CardService.newCardSection();

    const generateAction = CardService.newAction().setFunctionName('handleSavePdfClick');

    const button = CardService.newTextButton()
        .setText('Generate & Save PDF')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(generateAction);

    section.addWidget(button);
    builder.addSection(section);

    return builder.build();
}

/**
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} 
 */
function handleSavePdfClick(e) {
    const pdfBlob = createPdf();

    let targetFolder = DriveApp.getRootFolder();

    if (e.drive && e.drive.selectedItems && e.drive.selectedItems.length > 0) {
        const selectedItem = e.drive.selectedItems[0];

        if (selectedItem.mimeType === MimeType.FOLDER) {
            targetFolder = DriveApp.getFolderById(selectedItem.id);
        } else {
            const file = DriveApp.getFileById(selectedItem.id);
            const parents = file.getParents();
            if (parents.hasNext()) {
                targetFolder = parents.next();
            }
        }
    }

    //const newFile = targetFolder.createFile('Reporte_debug.html', pdfBlob, MimeType.HTML);
    const newFile = targetFolder.createFile(pdfBlob);

    return CardService.newActionResponseBuilder()
        .setNotification(
            CardService.newNotification()
                .setText("PDF saved successfully to: " + targetFolder.getName())
        )
        .build();
}


/**
 * @returns {GoogleAppsScript.Base.Blob}
 */
function createPdf() {
    const htmlTemplate = HtmlService.createTemplateFromFile('src/template');

    htmlTemplate.data = {
        start_year: "2025",
        end_year: "2026",
        period: "3er trimestre",
        date: "Del 18 de abril del 2026 al 30 de Julio del 2026",
        first_names: "Yamanqui",
        last_names: "García Rosales",
        id: "GARY801114MDFRG09",
        grade: "3º",
        level: "Secundaria",
        absences: 4,
        p1_average: 9.75,
        p2_average: 8.5,
        p3_average: 5.2,
        pf_average: 7.7,
        groups: [
            {
                name: "Lenguajes",
                color: "#c9daf8",
                courses: [
                    { name: "Español", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Inglés", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Arte", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" }
                ]
            },
            {
                name: "Pensamiento científico",
                color: "#fce5cd",
                courses: [
                    { name: "Matemáticas", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Física", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Tecnología", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" }
                ]
            },
            {
                name: "Ética, naturaleza y sociedad",
                color: "#d9ead3",
                courses: [
                    { name: "Historia", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Desarrollo Socioemocional", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" }
                ]
            },
            {
                name: "Humano y Comunitario",
                color: "#ead1dc",
                courses: [
                    { name: "Proyecto comunitario", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Deportes", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" }
                ]
            }
        ],
        courses: [
            { name: "Español", absences: 0, comment: "Muy buen alumno" },
            { name: "Inglés", absences: 1, comment: "Muy buen alumno" },
            { name: "Arte", absences: 2, comment: "Muy buen alumno" },
            { name: "Matemáticas", absences: 0, comment: "Muy buen alumno" },
            { name: "Física", absences: 3, comment: "Muy buen alumno" },
            { name: "Tecnología", absences: 0, comment: "Muy buen alumno" },
            { name: "Historia", absences: 4, comment: "Muy buen alumno" },
            { name: "Desarrollo Socioemocional", absences: 0, comment: "Muy buen alumno" },
            { name: "Proyecto comunitario", absences: 0, comment: "Muy buen alumno" },
            { name: "Deportes", absences: 5, comment: "Muy buen alumno" },
        ],
        images: {
            sep: BASE64_IMAGES.sep,
            school: BASE64_IMAGES.school,
            signature: BASE64_IMAGES.signature,
        },
        fonts: {
            regular: BASE64_FONTS.montserratRegular,
            bold: BASE64_FONTS.montserratBold,
            italic: BASE64_FONTS.montserratItalic,
        }
    };

    const htmlOutput = htmlTemplate.evaluate();
    //return htmlOutput.getContent();

    return htmlOutput.getAs(MimeType.PDF).setName('Reporte_' + htmlTemplate.data.first_names + '.pdf');
}