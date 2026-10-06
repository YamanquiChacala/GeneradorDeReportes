import { Colors } from "../common/gas-utils";
import type { AcademicField } from "../common/report-utils";
import { Base64Fonts, Base64Images } from "../common/utils/base64-constants";
import { REPORT_CSS_STYLE } from "../templates/report.css";
import { Templates } from "./types";

type Grades = [number | null, number | null, number | null, number | null];

interface StudentData {
    start_year: number;
    end_year: number;
    period: string;
    date: string;
    first_names: string;
    last_names: string;
    id: string;
    grade: string;
    level: string;
    absences?: number;
    averages: Grades;
    academicFields: AcademicField[];
    subjects: subjectGrades[];
    comments: CourseComment[];
}

interface subjectGrades {
    name: string;
    grades: Grades;
    habilities: [string | null, string | null, string | null];
}

interface CourseComment {
    name: string;
    absences?: number;
    comment: string;
}

export function getHTMLreport(data: StudentData): string {
    const out = (
        <html lang="es">
            <DocHead />
            <body>
                <Header start={data.start_year} end={data.end_year} />
                <StudentInfo firstName={data.first_names} lastName={data.last_names} id={data.id} grade={data.grade} level={data.level} />
                <div class="grades-body"></div>
            </body>
        </html>
    );
    if (typeof out !== "string") throw new Error("Async components aren't supported");
    return out;
}

function DocHead(): JSX.Element {
    return (
        <head>
            <title>Reporte de calificaciones</title>
            <base target="_top" />
            <style>{REPORT_CSS_STYLE}</style>
        </head>
    );
}

function Header(prop: { start: number; end: number }): JSX.Element {
    return (
        <div class="header">
            <div class="header-images">
                <img src={Base64Images.SEP} alt="Secretaría de Educación" />
                <img src={Base64Images.SCHOOL} alt="Montessori Chacala" />
            </div>
            <div class="header-title">
                Reporte de evaluación
                <br />
                <div class="header-dates">
                    Ciclo escolar {prop.start} - {prop.end}
                </div>
            </div>
        </div>
    );
}

function StudentInfo(props: { firstName: string; lastName: string; id: string; grade: string; level: string }): JSX.Element {
    return (
        <div class="student-info big">
            <div id="nombre">
                <b>Nombre(s) y apellidos: </b>{" "}
                <span safe>
                    {props.firstName} {props.lastName}
                </span>
            </div>
            <div id="id">
                <b>CURP: </b> <span safe>{props.id}</span>
            </div>
            <div id="grado">
                <b>Grado: </b>{" "}
                <span safe>
                    {props.grade} {props.level}
                </span>
            </div>
        </div>
    );
}

function GradesLeft(props: {fields: AcademicField[]}): JSX.Element{
    return (
        <div class="left-column">
        <table>
          <thead>
            <tr>
              <th colspan="6" class="black" style="height: 2rem;" >
                Programa académico
              </th>
              <th rowspan="2" class="text-vertical" style="width: 8%;" >
                Actitud
              </th>
              <th
                rowspan="2"
                class="text-vertical"
                style="width: 8%;"
              >Hábitos de<br/>estudio / trabajo</th>
              <th
                rowspan="2"
                class="text-vertical"
                style="width: 8%;"
              >Pensamiento<br/>crítico</th>
              <th
                rowspan="2"
                class="text-vertical"
                style="width: 8%;"
              >Desarrollo<br/>socioemocional</th>
            </tr>
            <tr>
              <th style="width: 16%;">Campos<br/>formativos</th>
              <th style="width: 16%;">Período</th>
              <th style="width: 8%;">1º</th>
              <th style="width: 8%;">2º</th>
              <th style="width: 8%;">3º</th>
              <th style="width: 8%;" class="average">Media<br/>final</th>
            </tr>
          </thead>
          <tbody>
            {props.fields.map((field) => (

            ))}
            <? data.groups.forEach(function(group) { ?>

            <? group.courses.forEach(function(course, index) { ?>
            <tr>

              <? if (index === 0) { ?>
              <th rowspan="<?= group.courses.length ?>" style="background-color: <?= group.color ?>;">
                <?= group.name ?>
              </th>
              <? } ?>

              <td class="course">
                <?= course.name ?>
              </td>

              <td class="<?= course.p1 < 6 ? 'failed' : '' ?>">
                <?= course.p1 ?>
              </td>
              <td class="<?= course.p2 < 6 ? 'failed' : '' ?>">
                <?= course.p2 ?>
              </td>
              <td class="<?= course.p3 < 6 ? 'failed' : '' ?>">
                <?= course.p3 ?>
              </td>
              <td class="average <?= course.final < 6 ? 'failed' : '' ?>">
                <?= course.final ?>
              </td>

              <td class="<?= course.h1 === 'R' ? 'failed' : '' ?>">
                <?= course.h1 ?>
              </td>
              <td class="<?= course.h2 === 'R' ? 'failed' : '' ?>">
                <?= course.h2 ?>
              </td>
              <td class="<?= course.h3 === 'R' ? 'failed' : '' ?>">
                <?= course.h3 ?>
              </td>
              <td class="<?= course.h4 === 'R' ? 'failed' : '' ?>">
                <?= course.h4 ?>
              </td>

            </tr>
            <? }); ?>
            <? }); ?>
          </tbody>
          <tfoot>
            <tr>
              <th colspan="2" class="black">
                Promedio del periodo
              </th>
              <th class="average <?= data.p1_average < 6 ? 'failed' : '' ?>">
                <?= data.p1_average ?>
              </th>
              <th class="average <?= data.p2_average < 6 ? 'failed' : '' ?>">
                <?= data.p2_average ?>
              </th>

              <th class="average <?= data.p3_average < 6 ? 'failed' : '' ?>">
                <?= data.p3_average ?>
              </th>

              <th class="average <?= data.pf_average < 6 ? 'failed' : '' ?>">
                <?= data.pf_average ?>
              </th>
            </tr>
          </tfoot>
        </table>
      </div>
    )
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
                color: Colors.LANGUAGE,
                courses: [
                    { name: "Español", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Inglés", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Arte", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                ],
            },
            {
                name: "Pensamiento científico",
                color: Colors.SCIENCE,
                courses: [
                    { name: "Matemáticas", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Física", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Tecnología", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                ],
            },
            {
                name: "Ética, naturaleza y sociedad",
                color: Colors.NATURE,
                courses: [
                    { name: "Historia", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                    { name: "Desarrollo Socioemocional", p1: 10, p2: 8, p3: 5, final: 8, h1: "E", h2: "B", h3: "S", h4: "R" },
                ],
            },
            {
                name: "Humano y Comunitario",
                color: Colors.HUMANITIES,
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

    const pdfBlob = createPdf(data);

    targetFolder.createFile(`Reporte_${data.first_names}.html`, pdfBlob[0], MimeType.HTML);
    targetFolder.createFile(pdfBlob[1].setName(`Reporte_${data.first_names}.pdf`));

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`PDF saved successfully to: ${targetFolder.getName()}`))
        .build();
}

/**
 * Creates a PDF of the given data based on the given html format;
 */
export function createPdf(data: StudentData): [string, GoogleAppsScript.Base.Blob] {
    const htmlTemplate = HtmlService.createTemplateFromFile(Templates.HTML_TO_PDF_TEMPLATE) as MyTemplate;

    htmlTemplate.data = data;

    const htmlOutput = htmlTemplate.evaluate();

    return [htmlOutput.getContent(), htmlOutput.getAs(MimeType.PDF)];

    //return htmlOutput.getAs(MimeType.PDF).setName(`Reporte_${htmlTemplate.data.first_names}.pdf`);
}
