import { Colors } from "../common/gas-utils";
import { pickRandom, randomInt } from "../common/utils";
import { GasTestRunner } from "../testing/gas-test-runner";
import { getHTMLreport, type StudentData } from "./pdf-creation";

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
    absences?: number;
    comment: string;
}

function createComment(name: string, absences: boolean) {
    const result: Comment = {
        name,
        comment: randomComment(),
    };
    if (absences) {
        result.absences = randomInt(0, 20);
    }
    return result;
}

type Grades = [number, number, number, number];

function createGrades(decimalPlaces: number): Grades {
    if (!Number.isInteger(decimalPlaces) || decimalPlaces < 0) {
        throw new RangeError("decimalPlaces must be a non-negative integer");
    }

    const scale = 10 ** decimalPlaces;
    const avgScale = scale * 10; // one extra decimal place

    // Integers in [0, 10 * scale], so 0 and 10 are both possible
    const a = randomInt(0, 10 * scale);
    const b = randomInt(0, 10 * scale);
    const c = randomInt(0, 10 * scale);

    // Average expressed in units of 1/avgScale, rounded to a whole unit
    const scaledAvg = Math.round(((a + b + c) * 10) / 3);

    return [a / scale, b / scale, c / scale, scaledAvg / avgScale];
}

function createSubject(name: string) {
    const habilityOptions = ["E", "B", "S", "R"];
    return {
        name,
        grades: createGrades(0),
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
                start_year: 2000,
                end_year: 2134,
                period: "N-ésimo periodo",
                date: "Del 35 de martius del 2000 al 42 de quintilis del 2134",
                first_names: "Alejandrino Güillermina",
                last_names: "de las Heras y Ponce de León del Olmo y Fernández de Córdoba",
                id: "HEOA001231HNYRLL37",
                grade: "4Xº",
                level: "Secundaria",
                // absences: 4,
                averages: createGrades(1),
                subjectSections: [
                    {
                        name: "Lenguajes",
                        color: Colors.LANGUAGE,
                        subjects: [createSubject("Español"), createSubject("Inglés"), createSubject("Arte y Creatividad")],
                    },
                    {
                        name: "Saberes y Pensamiento Científico",
                        color: Colors.SCIENCE,
                        subjects: [
                            createSubject("Matemáticas"),
                            createSubject("Física"),
                            createSubject("Química"),
                            createSubject("Biología"),
                            createSubject("Tecnología"),
                        ],
                    },
                    {
                        name: "Ética, Naturaleza y Sociedad",
                        color: Colors.NATURE,
                        subjects: [createSubject("Historia"), createSubject("Geografía"), createSubject("Ética")],
                    },
                    {
                        name: "De lo Humano y Comunitario",
                        color: Colors.HUMANITIES,
                        subjects: [createSubject("Educación Socioemocional"), createSubject("Desarrollo Personal"), createSubject("Deportes")],
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
                start_year: 2000,
                end_year: 2134,
                period: "N-ésimo periodo",
                date: "Del 35 de martius del 2000 al 42 de quintilis del 2134",
                first_names: "Anoñio",
                last_names: "d'or 田中",
                id: "שֵׁם מִשׁפָּחָה",
                grade: "IXª",
                level: "Primaria",
                absences: 4,
                averages: createGrades(1),
                subjectSections: [
                    {
                        name: "Lenguajes",
                        color: Colors.LANGUAGE,
                        subjects: [createSubject("Español"), createSubject("Inglés"), createSubject("Arte y Creatividad")],
                    },
                    {
                        name: "Saberes y Pensamiento Científico",
                        color: Colors.SCIENCE,
                        subjects: [createSubject("Matemáticas"), createSubject("Ciencias Naturales")],
                    },
                    {
                        name: "Ética, Naturaleza y Sociedad",
                        color: Colors.NATURE,
                        subjects: [createSubject("Historia y Geografía")],
                    },
                    {
                        name: "De lo Humano y Comunitario",
                        color: Colors.HUMANITIES,
                        subjects: [createSubject("Deportes")],
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
    });

    runner.execute();
}
