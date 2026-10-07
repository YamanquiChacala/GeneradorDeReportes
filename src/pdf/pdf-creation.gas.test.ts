import { Colors } from "../common/gas-utils";
import { GasTestRunner } from "../testing/gas-test-runner";
import { getHTMLreport, type StudentData } from "./pdf-creation";

export function testPDFReportGeneration() {
    const runner = new GasTestRunner();
    const { describe, test, beforeAll, expect } = runner;

    const TEST_FOLDER_NAME = "PDF_Tests";

    let testFolder: GoogleAppsScript.Drive.Folder;

    describe("Student Report PDF Generation", () => {
        beforeAll(() => {
            const scriptId = ScriptApp.getScriptId();
            const scriptFile = DriveApp.getFileById(scriptId);
            const parents = scriptFile.getParents();

            const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

            const folders = parentFolder.getFoldersByName(TEST_FOLDER_NAME);
            if (folders.hasNext()) {
                testFolder = folders.next();
            } else {
                testFolder = parentFolder.createFolder(TEST_FOLDER_NAME);
            }
        });

        test("Big group, individual assistance", () => {
            const prefix = "sec_big";
            const tiemzone = Session.getScriptTimeZone();
            const timestamp = Utilities.formatDate(new Date(), tiemzone, "yyyyMMdd_HHmmss");
            const expectedFileName = `${prefix}_${timestamp}`;

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
                // absences: 4,
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
                    {
                        name: "Español",
                        absences: 0,
                        comment:
                            "Proin quis eros malesuada, fermentum ante sit amet, pretium nulla.\n\nPellentesque fermentum risus eu velit vehicula, a posuere urna tincidunt.\nNulla cursus, tellus eu varius ornare, neque dui ullamcorper mi, nec pharetra odio dolor non nisl. Aliquam eu congue leo. Nulla non ipsum vitae ligula posuere feugiat. Ut dignissim nunc nec luctus commodo. Phasellus bibendum mauris aliquet magna pellentesque, et ultrices sem interdum. Maecenas et condimentum dui, vel feugiat ligula. Suspendisse ut ipsum id dolor ullamcorper porttitor. Phasellus ultricies, metus ut aliquam lobortis, dui enim dictum felis, eget consectetur leo ante at neque.",
                    },
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

            const htmlString = getHTMLreport(data);
            const savedHtmlFile = testFolder.createFile(`${expectedFileName}.html`, htmlString, MimeType.HTML);

            const pdfBlob = HtmlService.createHtmlOutput(htmlString).getAs(MimeType.PDF);
            const savedFile = testFolder.createFile(pdfBlob.setName(`${expectedFileName}.pdf`));

            expect(savedHtmlFile).toBeTruthy();

            expect(savedFile).toBeTruthy();
            expect(savedFile.getName()).toBe(`${expectedFileName}.pdf`);
            expect(savedFile.getMimeType()).toBe(MimeType.PDF);
            if (savedFile.getSize() === 0) throw new Error("Generated PDF is empty");
        });
    });

    runner.execute();
}
