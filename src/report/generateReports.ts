import { type Grades, getHTMLreport, type StudentData, type SubjectComment, type SubjectGrades, type SubjectSection } from "../common/gas-parts/pdf-creation";
import { DEFAULT_COMMENT, type ReportPersistentData, type Student, type StudentRow, StudentRowType } from "../common/report-utils";
import { loadReport, loadStudentData, type StudentRangeKey } from "./load";

export interface HtmlContent {
    html: string;
    name: string;
}

const isStudent = (row: StudentRow): row is Student => row.type === StudentRowType.STUDENT;

/**
 * Generate HTML reports for every student row in a report spreadsheet.
 */
export function generateAllReports(reportFileId: string): HtmlContent[] {
    const { parsedReport, persistentData } = loadReport(reportFileId);

    const students = persistentData.students.filter(isStudent);

    return loadStudentData(reportFileId, parsedReport, students).map(({ student, ranges }) => ({
        name: student.sheetName,
        html: getHTMLreport(getStudentDataForPdf(persistentData, ranges, student.firstName, student.lastName)),
    }));
}

/**
 * Generate the HTML report for one student.
 */
export function generateSingleReport(reportFileId: string, studentRowIndex: number): HtmlContent {
    const { parsedReport, persistentData } = loadReport(reportFileId);

    const student = persistentData.students[studentRowIndex];

    if (student == null || student.type === StudentRowType.SEPARATOR) throw new Error("Alumno no válido.");

    const studentData = loadStudentData(reportFileId, parsedReport, [student])[0];

    if (studentData == null) throw new Error("Fallo al extraer detalles de alumno.");

    const pdfParams = getStudentDataForPdf(persistentData, studentData.ranges, student.firstName, student.lastName);

    return { html: getHTMLreport(pdfParams), name: student.sheetName };
}

/**
 * Assemble the loaded spreadsheet values and persistent settings into the data required by the PDF template.
 */
function getStudentDataForPdf(persistenData: ReportPersistentData, rawData: Record<StudentRangeKey, string[][]>, firstName: string, lastName: string): StudentData {
    const startDate = persistenData.configData.dates[0];
    const endDate = persistenData.configData.dates[3];

    const startYear = new Date(startDate).getUTCFullYear().toString();
    const endYear = new Date(endDate).getUTCFullYear().toString();

    const id = rawData.generalInfo[0]?.[0] ?? "";
    const grade = rawData.generalInfo[1]?.[0] ?? "";
    const level = rawData.generalInfo[2]?.[0] ?? "";
    const period = rawData.generalInfo[3]?.[0] ?? "";
    const date = rawData.generalInfo[4]?.[0] ?? "";
    const absences = rawData.generalInfo[5]?.[0];

    const allGrades: Grades[] = [];
    const subjectAbsences: string[] = [];

    for (let i = 0; i < persistenData.subjects.length; i++) {
        allGrades.push([rawData.trim1[i]?.[1] ?? "", rawData.trim2[i]?.[1] ?? "", rawData.trim3[i]?.[1] ?? "", rawData.trim3[i]?.[2] ?? ""]);

        if (persistenData.configData.attendancePerClass) {
            const currentAbsences = [rawData.trim1[i]?.[0] ?? "", rawData.trim2[i]?.[0] ?? "", rawData.trim3[i]?.[0] ?? ""];
            subjectAbsences.push(currentAbsences.filter(Boolean).join(" / "));
        }
    }

    const averages: Grades = [rawData.trim1Totals[0]?.[0] ?? "", rawData.trim2Totals[0]?.[0] ?? "", rawData.trim3Totals[0]?.[0] ?? "", rawData.trim3Totals[0]?.[1] ?? ""];

    let currentIndex = 0;

    const subjectSections: SubjectSection[] = persistenData.academicFields.map((field) => {
        const sectionSubjects: SubjectGrades[] = [];

        for (let i = 0; i < field.subjects; i++) {
            sectionSubjects.push({
                // biome-ignore lint/style/noNonNullAssertion: Subjects must have all the subjects.
                name: persistenData.subjects[currentIndex]!.subject,
                // biome-ignore lint/style/noNonNullAssertion: Subjects must have all the subjects.
                grades: allGrades[currentIndex]!,
                habilities: [
                    rawData.abilities[currentIndex]?.[0] ?? "",
                    rawData.abilities[currentIndex]?.[1] ?? "",
                    rawData.abilities[currentIndex]?.[2] ?? "",
                    rawData.abilities[currentIndex]?.[3] ?? "",
                ],
            });

            currentIndex++;
        }

        return {
            name: field.name,
            color: field.color,
            subjects: sectionSubjects,
        };
    });

    // Omit a comment when it contains only the configured default prompt.
    const cleanComment = (raw: string): string => {
        const trimmed = raw.trim();
        return DEFAULT_COMMENT.startsWith(trimmed) ? "" : raw;
    };

    const comments: SubjectComment[] = [];

    persistenData.subjects.forEach((subject, index) => {
        comments.push({
            name: subject.subject,
            absences: subjectAbsences[index],
            comment: cleanComment(rawData.comments[index]?.[0] ?? ""),
        });
    });

    return {
        startYear,
        endYear,
        firstName,
        lastName,
        id,
        grade,
        level,
        period,
        date,
        absences,
        averages,
        subjectSections,
        comments,
    };
}
