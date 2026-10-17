import { GasTestRunner } from "../../testing/gas-test-runner";
import { Colors } from "../gas-utils";
import { pickRandom, randomInt } from "../utils";
import { type Grades, getHTMLreport, type StudentData } from "./pdf-creation";

const LOREM_SENTENCES = [
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
    "Proin quis eros malesuada, fermentum ante sit amet, pretium nulla.",
    "Pellentesque fermentum risus eu velit vehicula, a posuere urna tincidunt.",
    "Nulla cursus, tellus eu varius ornare, neque dui ullamcorper mi, nec pharetra odio dolor non nisl.",
    "Aliquam eu congue leo.",
    "Nulla non ipsum vitae ligula posuere feugiat.",
    "Ut dignissim nunc nec luctus commodo.",
    "Phasellus bibendum mauris aliquet magna pellentesque, et ultrices sem interdum.",
    "Maecenas et condimentum dui, vel feugiat ligula.",
    "Suspendisse ut ipsum id dolor ullamcorper porttitor.",
    "Phasellus ultricies, metus ut aliquam lobortis, dui enim dictum felis, eget consectetur leo ante at neque.",
    "Donec vitae sapien ut libero venenatis faucibus.",
    "Curabitur ligula sapien, tincidunt non, euismod vitae, posuere imperdiet, leo.",
    "Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae.",
    "Mauris placerat eleifend leo.",
    "Quisque sit amet est et sapien ullamcorper pharetra.",
    "Sed lectus.",
    "Etiam rhoncus, maecenas tempus, tellus eget condimentum rhoncus, sem quam semper libero.",
];

function randomComment(maxSentences = 12, newlineChance = 0.25): string {
    const count = randomInt(0, maxSentences);
    let result = "";

    for (let i = 0; i < count; i++) {
        result += pickRandom(LOREM_SENTENCES, 1);

        let usingNewLine = false;
        let rand = Math.random();

        while (rand < newlineChance) {
            usingNewLine = true;
            result += "\n";
            rand = Math.random();
        }
        if (!usingNewLine) {
            result += " ";
        }
    }
    return result;
}

interface Comment {
    name: string;
    absences?: string;
    comment: string;
}

function createComment(name: string, absences: boolean) {
    const result: Comment = {
        name,
        comment: randomComment(),
    };
    if (absences) {
        result.absences = randomInt(0, 20).toFixed(0);
    }
    return result;
}

function createGrades(decimalPlaces: number, period: 0 | 1 | 2): Grades {
    if (!Number.isInteger(decimalPlaces) || decimalPlaces < 0) {
        throw new RangeError("decimalPlaces must be a non-negative integer");
    }

    const scale = 10 ** decimalPlaces;

    // Integers in [0, 10 * scale], so 0 and 10 are both possible
    const a = randomInt(5 * scale, 10 * scale) / scale;
    const b = randomInt(5 * scale, 10 * scale) / scale;
    const c = randomInt(5 * scale, 10 * scale) / scale;

    // Average expressed in units of 1/avgScale, rounded to a whole unit
    const average = (a + b + c) / 3;

    return [
        a.toFixed(decimalPlaces),
        period >= 1 ? b.toFixed(decimalPlaces) : "",
        period >= 2 ? c.toFixed(decimalPlaces) : "",
        period >= 2 ? average.toFixed(decimalPlaces + 1) : "",
    ];
}

function createSubject(name: string, period: 0 | 1 | 2) {
    const habilityOptions = ["E", "B", "S", "R"];
    return {
        name,
        grades: createGrades(0, period),
        habilities: pickRandom(habilityOptions, 4) as [string, string, string, string],
    };
}

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
            const prefix = "secundaria";
            const tiemzone = Session.getScriptTimeZone();
            const timestamp = Utilities.formatDate(new Date(), tiemzone, "yyyyMMdd_HHmmss");
            const expectedFileName = `${prefix}_${timestamp}`;

            const data: StudentData = {
                startYear: "2000",
                endYear: "2134",
                period: "N-ésimo periodo",
                date: "Del 35 de martius del 2000 al 42 de quintilis del 2134",
                firstName: "Alejandrino Güillermina",
                lastName: "de las Heras y Ponce de León del Olmo y Fernández de Córdoba",
                id: "HEOA001231HNYRLL37",
                grade: "4Xº",
                level: "Secundaria",
                // absences: 4,
                averages: createGrades(1, 2),
                subjectSections: [
                    {
                        name: "Lenguajes",
                        color: Colors.LANGUAGE,
                        subjects: [createSubject("Español", 2), createSubject("Inglés", 2), createSubject("Arte y Creatividad", 2)],
                    },
                    {
                        name: "Saberes y Pensamiento Científico",
                        color: Colors.SCIENCE,
                        subjects: [
                            createSubject("Matemáticas", 2),
                            createSubject("Física", 2),
                            createSubject("Química", 2),
                            createSubject("Biología", 2),
                            createSubject("Tecnología", 2),
                        ],
                    },
                    {
                        name: "Ética, Naturaleza y Sociedad",
                        color: Colors.NATURE,
                        subjects: [createSubject("Historia", 2), createSubject("Geografía", 2), createSubject("Ética", 2)],
                    },
                    {
                        name: "De lo Humano y Comunitario",
                        color: Colors.HUMANITIES,
                        subjects: [createSubject("Educación Socioemocional", 2), createSubject("Desarrollo Personal", 2), createSubject("Deportes", 2)],
                    },
                ],
                comments: [
                    createComment("Español", true),
                    createComment("Inglés", true),
                    createComment("Arte y Creatividad", true),
                    createComment("Matemáticas", true),
                    createComment("Física", true),
                    createComment("Química", true),
                    createComment("Biología", true),
                    createComment("Tecnología", true),
                    createComment("Historia", true),
                    createComment("Geografía", true),
                    createComment("Ética", true),
                    createComment("Educación Socioemocional", true),
                    createComment("Desarrollo Personal", true),
                    createComment("Deportes", true),
                ],
            };

            const htmlString = getHTMLreport(data);
            // const savedHtmlFile = testFolder.createFile(`${expectedFileName}.html`, htmlString, MimeType.HTML);

            const pdfBlob = HtmlService.createHtmlOutput(htmlString).getAs(MimeType.PDF);
            const savedFile = testFolder.createFile(pdfBlob.setName(`${expectedFileName}.pdf`));

            // expect(savedHtmlFile).toBeTruthy();

            expect(savedFile).toBeTruthy();
            expect(savedFile.getName()).toBe(`${expectedFileName}.pdf`);
            expect(savedFile.getMimeType()).toBe(MimeType.PDF);
            if (savedFile.getSize() === 0) throw new Error("Generated PDF is empty");
        });

        test("Small group, shared assistance", () => {
            const prefix = "primaria";
            const tiemzone = Session.getScriptTimeZone();
            const timestamp = Utilities.formatDate(new Date(), tiemzone, "yyyyMMdd_HHmmss");
            const expectedFileName = `${prefix}_${timestamp}`;

            const data: StudentData = {
                startYear: "2000",
                endYear: "2134",
                period: "N-ésimo periodo",
                date: "Del 35 de martius del 2000 al 42 de quintilis del 2134",
                firstName: "Anoñio",
                lastName: "d'or 田中",
                id: "שֵׁם מִשׁפָּחָה",
                grade: "IXª",
                level: "Primaria",
                absences: "4",
                averages: createGrades(1, 1),
                subjectSections: [
                    {
                        name: "Lenguajes",
                        color: Colors.LANGUAGE,
                        subjects: [createSubject("Español", 1), createSubject("Inglés", 1), createSubject("Arte y Creatividad", 1)],
                    },
                    {
                        name: "Saberes y Pensamiento Científico",
                        color: Colors.SCIENCE,
                        subjects: [createSubject("Matemáticas", 1), createSubject("Ciencias Naturales", 1)],
                    },
                    {
                        name: "Ética, Naturaleza y Sociedad",
                        color: Colors.NATURE,
                        subjects: [createSubject("Historia y Geografía", 1)],
                    },
                    {
                        name: "De lo Humano y Comunitario",
                        color: Colors.HUMANITIES,
                        subjects: [createSubject("Deportes", 1)],
                    },
                ],
                comments: [
                    createComment("Español", false),
                    createComment("Inglés", false),
                    createComment("Arte y Creatividad", false),
                    createComment("Matemáticas", false),
                    createComment("Ciencias Naturales", false),
                    createComment("Historia y Geografía", false),
                    createComment("Deportes", false),
                ],
            };

            const htmlString = getHTMLreport(data);
            // const savedHtmlFile = testFolder.createFile(`${expectedFileName}.html`, htmlString, MimeType.HTML);

            const pdfBlob = HtmlService.createHtmlOutput(htmlString).getAs(MimeType.PDF);
            const savedFile = testFolder.createFile(pdfBlob.setName(`${expectedFileName}.pdf`));

            // expect(savedHtmlFile).toBeTruthy();

            expect(savedFile).toBeTruthy();
            expect(savedFile.getName()).toBe(`${expectedFileName}.pdf`);
            expect(savedFile.getMimeType()).toBe(MimeType.PDF);
            if (savedFile.getSize() === 0) throw new Error("Generated PDF is empty");
        });

        test("Small group, first period", () => {
            const prefix = "preescolar";
            const tiemzone = Session.getScriptTimeZone();
            const timestamp = Utilities.formatDate(new Date(), tiemzone, "yyyyMMdd_HHmmss");
            const expectedFileName = `${prefix}_${timestamp}`;

            const data: StudentData = {
                startYear: "2000",
                endYear: "2134",
                period: "N-ésimo periodo",
                date: "Del 35 de martius del 2000 al 42 de quintilis del 2134",
                firstName: "Úngüla",
                lastName: '"La destructora" Hernández',
                id: "C0NN4N",
                grade: "五十六",
                level: "Preescolar",
                absences: "25",
                averages: createGrades(1, 0),
                subjectSections: [
                    {
                        name: "Lenguajes",
                        color: Colors.LANGUAGE,
                        subjects: [createSubject("Español", 0), createSubject("Inglés", 0), createSubject("Arte y Creatividad", 0)],
                    },
                    {
                        name: "Saberes y Pensamiento Científico",
                        color: Colors.SCIENCE,
                        subjects: [createSubject("Matemáticas", 0)],
                    },
                    {
                        name: "Ética, Naturaleza y Sociedad",
                        color: Colors.NATURE,
                        subjects: [createSubject("Historia", 0)],
                    },
                    {
                        name: "De lo Humano y Comunitario",
                        color: Colors.HUMANITIES,
                        subjects: [createSubject("Deportes", 0)],
                    },
                ],
                comments: [
                    createComment("Español", false),
                    createComment("Inglés", false),
                    createComment("Arte y Creatividad", false),
                    createComment("Matemáticas", false),
                    createComment("Historia", false),
                    createComment("Deportes", false),
                ],
            };

            const htmlString = getHTMLreport(data);
            // const savedHtmlFile = testFolder.createFile(`${expectedFileName}.html`, htmlString, MimeType.HTML);

            const pdfBlob = HtmlService.createHtmlOutput(htmlString).getAs(MimeType.PDF);
            const savedFile = testFolder.createFile(pdfBlob.setName(`${expectedFileName}.pdf`));

            // expect(savedHtmlFile).toBeTruthy();

            expect(savedFile).toBeTruthy();
            expect(savedFile.getName()).toBe(`${expectedFileName}.pdf`);
            expect(savedFile.getMimeType()).toBe(MimeType.PDF);
            if (savedFile.getSize() === 0) throw new Error("Generated PDF is empty");
        });
    });

    runner.execute();
}
