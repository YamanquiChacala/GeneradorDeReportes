import { Base64Images } from "../common/utils/base64-constants";
import { REPORT_CSS_STYLE } from "../templates/report.css";

type Grades = [number | null, number | null, number | null, number | null];

export interface StudentData {
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
    subjectSections: SubjectSection[];
    comments: SubjectComment[];
}

interface SubjectSection {
    name: string;
    color: string;
    subjects: SubjectGrades[];
}

interface SubjectGrades {
    name: string;
    grades: Grades;
    habilities: [string | null, string | null, string | null, string | null];
}

interface SubjectComment {
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
                <div class="grades-body">
                    <GradesLeft averages={data.averages} sections={data.subjectSections} />
                    <NoticeRight />
                </div>
                <Footer />
                <Comments absences={data.absences} comments={data.comments} date={data.date} start={data.start_year} end={data.end_year} period={data.period} />
            </body>
        </html>
    );
    if (typeof out !== "string") throw new Error("Async components aren't supported");
    return `<!DOCTYPE html>${out}`;
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

function GradesLeft(props: { sections: SubjectSection[]; averages: Grades }): JSX.Element {
    return (
        <div class="left-column">
            <table>
                <thead>
                    <tr>
                        <th colspan="6" class="black" style={{ height: "1rem" }}>
                            Programa académico
                        </th>
                        <th rowspan="2" class="text-vertical" style={{ width: "8%" }}>
                            Actitud
                        </th>
                        <th rowspan="2" class="text-vertical" style={{ width: "8%" }}>
                            Hábitos de
                            <br />
                            estudio
                        </th>
                        <th rowspan="2" class="text-vertical" style={{ width: "8%" }}>
                            Pensamiento
                            <br />
                            crítico
                        </th>
                        <th rowspan="2" class="text-vertical" style={{ width: "8%" }}>
                            Desarrollo
                            <br />
                            emocional
                        </th>
                    </tr>
                    <tr>
                        <th style={{ width: "16%" }}>
                            Campos
                            <br />
                            formativos
                        </th>
                        <th style={{ width: "16%" }}>Período</th>
                        <th style={{ width: "8%" }}>1º</th>
                        <th style={{ width: "8%" }}>2º</th>
                        <th style={{ width: "8%" }}>3º</th>
                        <th style={{ width: "8%" }} class="average">
                            Final
                        </th>
                    </tr>
                </thead>
                <tbody>
                    {props.sections.map((field) =>
                        field.subjects.map((subject, index) => (
                            <tr>
                                {index === 0 && (
                                    <th rowspan={field.subjects.length} style={{ backgroundColor: field.color }} safe>
                                        {field.name}
                                    </th>
                                )}
                                <td class="course" safe>
                                    {subject.name}
                                </td>
                                {subject.grades.map((grade, gradeIndex) => (
                                    <td class={`${grade !== null && grade < 6 ? "failed" : ""} ${gradeIndex === 3 ? "average" : ""}`}>{grade}</td>
                                ))}
                                {subject.habilities.map((mark) => (
                                    <td class={mark === "R" ? "failed" : ""} safe>
                                        {mark}
                                    </td>
                                ))}
                            </tr>
                        )),
                    )}
                </tbody>
                <tfoot>
                    <tr>
                        <th colspan="2" class="black">
                            Promedio del periodo
                        </th>
                        {props.averages.map((grade) => (
                            <th class={`${grade !== null && grade <= 5 ? "failed " : ""}average`}>{grade}</th>
                        ))}
                    </tr>
                </tfoot>
            </table>
        </div>
    );
}

function NoticeRight(): JSX.Element {
    return (
        <div class="right-column">
            <table>
                <thead>
                    <tr>
                        <th class="black">Habilidades para el aprendizaje</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="text-left">
                            <b>Actitud:</b> La manera en que un estudiante se acerca al aprendizaje, su nivel de interés, motivación, compromiso, y su disposición a
                            participar de manera activa y positiva.
                        </td>
                    </tr>
                    <tr>
                        <td class="text-left">
                            <b>Hábitos de estudio y trabjo:</b> Nivel de autonomía en el proceso de aprendizaje, capacidad de investigación, organización personal,
                            gestión del tiempo, dedicación y calidad en la elaboración de proyectos, tareas o actividades.
                        </td>
                    </tr>
                    <tr>
                        <td class="text-left">
                            <b>Pensamiento crítico:</b> Habilidad para analizar, evaluar y entender de manera reflexiva la información, ideas, conceptos y argumentos.
                            Implica pensar de manera independiente, cuestionar, formar juicios, resolver problemas y tomar decisiones informadas.
                        </td>
                    </tr>
                    <tr>
                        <td class="text-left">
                            <b>Desarrollo socioemocional:</b> Habilidades de comunicación, escucha activa, colaboración, trabajo en equipo, empatía, autorregulación
                            emocional, autoestima y confianza en sí mismo(a).
                        </td>
                    </tr>
                    <tr></tr>
                </tbody>
                <tfoot>
                    <tr>
                        <td class="yellow table-title">
                            <b>E</b>=Excelente - <b>B</b>=Bueno - <b>S</b>=Suficiente - <b>R</b>=Requiere apoyo
                        </td>
                    </tr>
                </tfoot>
            </table>

            <table>
                <thead>
                    <tr>
                        <th class="table-title" colspan="2">
                            Criterio de calificación
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="text-left" style={{ width: "90%" }}>
                            Indica <b>dominio sobresaliente</b> de los aprendizajes esperados. La/el estudiante ha demostrado los conocimientos, habilidades, actitudes y
                            valores requeridos con un alto grado de efectividad.
                        </td>
                        <td>
                            <b>10</b>
                        </td>
                    </tr>
                    <tr>
                        <td class="text-left">
                            Indica <b>dominio satisfactorio</b> de los aprendizajes esperados. La/el estudiante ha demostrado los conocimientos, habilidades, actitudes y
                            valores requeridos con efectividad.
                        </td>
                        <td>
                            <b>8 y 9</b>
                        </td>
                    </tr>
                    <tr>
                        <td class="text-left">
                            Indica <b>dominio básico</b> de los aprendizajes esperados. La/el estudiante tiene dificultades para demostrar los conocimientos, habilidades,
                            actitudes y valores requeridos.
                        </td>
                        <td>
                            <b>6 y 7</b>
                        </td>
                    </tr>
                    <tr>
                        <td class="text-left">
                            Indica <b>dominio insuficiente</b> de los aprendizajes esperados. La/el estudiante tiene carencias fundamentales en los conocimientos,
                            habilidades, actitudes y valores requeridos.{" "}
                        </td>
                        <td>
                            <b>5</b>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
    );
}

function Footer(): JSX.Element {
    return (
        <div class="footer big">
            <div class="address">
                <b>Montessori Chacala</b>
                C.C.T. Preescolar: 18PJN0157K
                <br />
                C.C.T. Primaria: 18PPR0093G
                <br />
                C.C.T. Secundaria: 18PES0053Z
                <br />
                T: (327) 219-4011 / info@chacala.school
            </div>
            <div class="signature">
                <div class="sig-line">
                    <img src={Base64Images.SIGNATURE} alt="Firma" class="sig-image" />
                    <b>Direción y control escolar</b>
                </div>
            </div>
            <div class="signature">
                <div class="sig-line">
                    <b>Nombre y firma de la madre o padre o tutor(a)</b>
                </div>
            </div>
        </div>
    );
}

function Comments(props: { start: number; end: number; date: string; period: string; absences?: number; comments: SubjectComment[] }): JSX.Element {
    return (
        <div class="page-break">
            <table>
                <thead>
                    <tr>
                        <th colspan="3" class="seaparation no-padding">
                            <div class="header">
                                <div class="header-images">
                                    <img src={Base64Images.SEP} alt="Secretaría de Educación" />
                                    <img src={Base64Images.SCHOOL} alt="Montessori Chacala" />
                                </div>
                                <div class="header-title">
                                    Reporte de evaluación
                                    <br />
                                    <div class="header-dates">
                                        Ciclo escolar {props.start} - {props.end}
                                    </div>
                                </div>
                            </div>
                        </th>
                    </tr>
                    <tr>
                        <th class="table-title" style={{ minWidth: "7em" }} safe>
                            {props.period}
                        </th>
                        {props.absences == null ? (
                            <th colspan="2" safe>
                                {props.date}
                            </th>
                        ) : (
                            <>
                                <th style={{ minWidth: "4.4em" }}>Faltas: {props.absences}</th>
                                <th safe>{props.date}</th>
                            </>
                        )}
                    </tr>
                    <tr>
                        <th colspan="3" class="seaparation"></th>
                    </tr>
                    <tr>
                        {props.absences == null ? (
                            <>
                                <th style={{ width: "20%" }}>Asignatura</th>
                                <th style={{ width: "5%" }}>Faltas</th>
                            </>
                        ) : (
                            <th colspan="2" style={{ width: "20%" }}>
                                Asignatura
                            </th>
                        )}
                        <th>Comentarios del Docente</th>
                    </tr>
                </thead>

                <tbody>
                    {props.comments.map((comm) => (
                        <tr>
                            {comm.absences == null ? (
                                <td colspan="2">
                                    <b safe>{comm.name}</b>
                                </td>
                            ) : (
                                <>
                                    <td>
                                        <b safe>{comm.name}</b>
                                    </td>
                                    <td>
                                        <b>{comm.absences}</b>
                                    </td>
                                </>
                            )}
                            <td class="text-left">
                                {comm.comment
                                    .split("\n")
                                    .filter((line) => line.trim() !== "")
                                    .map((line) => (
                                        <p safe>{line}</p>
                                    ))}
                            </td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
}
