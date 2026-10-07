import { Colors } from "../common/gas-utils";
import { getHTMLreport, type StudentData } from "./pdf-creation";

const data: StudentData = {
    start_year: 2025,
    end_year: 2026,
    period: "3er trimestre",
    date: "Del 18 de abril del 2026 al 30 de Julio del 2026",
    first_names: "Yamanqui",
    last_names: "García Rosales",
    id: "GARY801114MDFRG09",
    grade: "3º",
    level: "Secundaria",
    absences: 4,
    averages: [9.3, 4.5, 7.8, 5.23],
    subjectSections: [
        {
            name: "Lenguajes",
            color: Colors.LANGUAGE,
            subjects: [
                { name: "Español", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
                { name: "Inglés", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
                { name: "Arte", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
            ],
        },
        {
            name: "Pensamiento científico",
            color: Colors.SCIENCE,
            subjects: [
                { name: "Matemáticas", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
                { name: "Física", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
                { name: "Tecnología", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
            ],
        },
        {
            name: "Ética, naturaleza y sociedad",
            color: Colors.NATURE,
            subjects: [
                { name: "Historia", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
                { name: "Desarrollo Socioemocional", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
            ],
        },
        {
            name: "Humano y Comunitario",
            color: Colors.HUMANITIES,
            subjects: [
                { name: "Proyecto comunitario", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
                { name: "Deportes", grades: [10, 8, 5, 8], habilities: ["E", "B", "S", "R"] },
            ],
        },
    ],
    comments: [
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
};

let parent: GoogleAppsScript.Drive.Folder; // Get a Drive folder

const htmlString = getHTMLreport(data);

const pdfBlob = HtmlService.createHtmlOutput(htmlString).getAs(MimeType.PDF);

parent.createFile(pdfBlob.setName(`Reporte_${data.first_names}.pdf`));
