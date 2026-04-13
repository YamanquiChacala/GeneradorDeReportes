import { Base64Fonts, Base64Images } from "../common/base64Constants";

interface StudentData {
    start_year: string;
    end_year: string;
    period: string;
    date: string;
    first_names: string;
    last_names: string;
    id: string;
    grade: string;
    level: string;
    absences?: number;
    p1_average?: number;
    p2_average?: number;
    p3_average?: number;
    pf_average?: number;
    groups: CourseGroup[];
    courses: CourseComment[];
    images: {
        sep: string;
        school: string;
        signature: string;
    };
    fonts: {
        regular: string;
        bold: string;
        italic: string;
    };
}

interface CourseGroup {
    name: string;
    color: string;
    courses: CourseGrades[];
}

interface CourseGrades {
    name: string;
    p1?: number;
    p2?: number;
    p3?: number;
    final?: number;
    h1: string;
    h2: string;
    h3: string;
    h4: string;
}

interface CourseComment {
    name: string;
    absences?: number;
    comment: string;
}

interface MyTemplate extends GoogleAppsScript.HTML.HtmlTemplate {
    data: StudentData;
}

export function buildDriveTestPdfCreationCard(): GoogleAppsScript.Card_Service.Card {
    const builder = CardService.newCardBuilder();
    const section = CardService.newCardSection();

    const generateAction = CardService.newAction().setFunctionName(onTestSavePdf.name);

    const button = CardService.newTextButton().setText("Generate & Save PDF").setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(generateAction);

    section.addWidget(button);
    builder.addSection(section);

    return builder.build();
}

function onTestSavePdf(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    let targetFolder = DriveApp.getRootFolder();

    const selectedItem = e?.drive?.selectedItems[0];

    if (selectedItem) {
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

    const pdfBlob = createPdf();

    targetFolder.createFile("Reporte_debug.html", pdfBlob[0], MimeType.HTML);
    targetFolder.createFile(pdfBlob[1]);

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`PDF saved successfully to: ${targetFolder.getName()}`))
        .build();
}

function createPdf(): [string, GoogleAppsScript.Base.Blob] {
    const htmlTemplate = HtmlService.createTemplateFromFile("PdfPrintTemplate") as MyTemplate;

    const data: StudentData = {
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
                    { name: "Arte", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                ],
            },
            {
                name: "Pensamiento científico",
                color: "#fce5cd",
                courses: [
                    { name: "Matemáticas", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Física", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Tecnología", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                ],
            },
            {
                name: "Ética, naturaleza y sociedad",
                color: "#d9ead3",
                courses: [
                    { name: "Historia", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Desarrollo Socioemocional", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                ],
            },
            {
                name: "Humano y Comunitario",
                color: "#ead1dc",
                courses: [
                    { name: "Proyecto comunitario", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Deportes", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                ],
            },
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
            sep: Base64Images.SEP,
            school: Base64Images.SCHOOL,
            signature: Base64Images.SIGNATURE,
        },
        fonts: {
            regular: Base64Fonts.MONTSERRAT_REGULAR,
            bold: Base64Fonts.MONTSERRAT_BOLD,
            italic: Base64Fonts.MONTSERRAT_ITALIC,
        },
    };

    htmlTemplate.data = data;

    const htmlOutput = htmlTemplate.evaluate();

    return [htmlOutput.getContent(), htmlOutput.getAs(MimeType.PDF).setName(`Reporte_${htmlTemplate.data.first_names}.pdf`)];

    //return htmlOutput.getAs(MimeType.PDF).setName(`Reporte_${htmlTemplate.data.first_names}.pdf`);
}
