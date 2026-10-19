import { ReportSheetSchema } from "../common/gas-parts";
import { type Grades, getHTMLreport, type StudentData, type SubjectComment, type SubjectGrades, type SubjectSection } from "../common/gas-parts/pdf-creation";
import {
    createRequiredGetter,
    type ExtractRangeNames,
    getStringRangeData,
    type MappedNamedRange,
    type OffsetGridRangeProperties,
    offsetGridRange,
} from "../common/gas-utils";
import { DEFAULT_COMMENT, type ReportPersistentData, StudentRowType } from "../common/report-utils";
import { type LoadedStudentData, loadReport, loadStudentData, type StudentRangeKey } from "./load";

export function generateSingleReport(reportFileId: string, studentRowIndex: number): string {
    const { parsedReport, persistentData } = loadReport(reportFileId);

    const student = persistentData.students[studentRowIndex];

    if (student == null || student.type === StudentRowType.SEPARATOR) throw new Error("Alumno no válido.");

    const sheetName = student.sheetName;

    const studentData = loadStudentData(reportFileId, parsedReport, new Set([sheetName]))[0];

    if (studentData == null) throw new Error("Fallo al extraer detalles de alumno.");

    const pdfParams = getStudentDataForPdf(persistentData, studentData.ranges);

    return getHTMLreport(pdfParams);
}

function getStudentDataForPdf(persistenData: ReportPersistentData, rawData: Record<StudentRangeKey, string[][]>): StudentData {
    const startDate = persistenData.configData.dates[0];
    const endDate = persistenData.configData.dates[3];

    const startYear = new Date(startDate).getUTCFullYear().toString();
    const endYear = new Date(endDate).getUTCFullYear().toString();

    const firstName = rawData.firstName[];
}

function getStudentDataForPDF2(
    mappedRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
    studentSheet: GoogleAppsScript.Sheets.Schema.Sheet,
    persistenData: ReportPersistentData,
): StudentData {
    const getMappedRange = createRequiredGetter(mappedRanges, "rango de reporte");
    const getRangeData = (rangeName: ExtractRangeNames<typeof ReportSheetSchema>, options: StudentRangeOptions = {}) =>
        getStudentRangeData(getMappedRange(rangeName), studentSheet, options);
    const getLastColumnsData = (rangeName: ExtractRangeNames<typeof ReportSheetSchema>, width: number) =>
        getStudentLastColumnsData(getMappedRange(rangeName), studentSheet, width);

    const startDate = persistenData.configData.dates[0];
    const endDate = persistenData.configData.dates[3];

    const startYear = new Date(startDate).getUTCFullYear().toString();
    const endYear = new Date(endDate).getUTCFullYear().toString();

    const firstNameRange = getRangeData(ReportSheetSchema.sheets.studentTemplate.ranges.firstName, { width: 1 });
    // biome-ignore lint/style/noNonNullAssertion: Guaranteed strings.
    const firstName = firstNameRange[0]![0]!;

    const lastNameRange = getRangeData(ReportSheetSchema.sheets.studentTemplate.ranges.lastName, { width: 1 });
    // biome-ignore lint/style/noNonNullAssertion: Guaranteed strings.
    const lastName = lastNameRange[0]![0]!;

    const generalInfo = getRangeData(ReportSheetSchema.sheets.studentTemplate.ranges.generalInfo);

    const id = generalInfo[0]?.[0] ?? "";
    const grade = generalInfo[1]?.[0] ?? "";
    const level = generalInfo[2]?.[0] ?? "";
    const period = generalInfo[3]?.[0] ?? "";
    const date = generalInfo[4]?.[0] ?? "";
    const absences = generalInfo[5]?.[0];

    const trim1 = getLastColumnsData(ReportSheetSchema.sheets.studentTemplate.ranges.trim1Subjects, 2);
    const trim1Totals = getLastColumnsData(ReportSheetSchema.sheets.studentTemplate.ranges.trim1Totals, 1);
    const trim2 = getLastColumnsData(ReportSheetSchema.sheets.studentTemplate.ranges.trim2Subjects, 2);
    const trim2Totals = getLastColumnsData(ReportSheetSchema.sheets.studentTemplate.ranges.trim2Totals, 1);
    const trim3 = getLastColumnsData(ReportSheetSchema.sheets.studentTemplate.ranges.trim3Subjects, 3);
    const trim3Totals = getLastColumnsData(ReportSheetSchema.sheets.studentTemplate.ranges.trim3Totals, 2);

    const allGrades: Grades[] = [];
    const subjectAbsences: string[] = [];

    for (let i = 0; i < persistenData.subjects.length; i++) {
        // biome-ignore lint/style/noNonNullAssertion: All arrays are of known size
        allGrades.push([trim1[i]![1]!, trim2[i]![1]!, trim3[i]![1]!, trim3[i]![2]!]);

        if (persistenData.configData.attendancePerClass) {
            // biome-ignore lint/style/noNonNullAssertion: All arrays are of known size
            const currentAbsences = [trim1[i]![0]!, trim2[i]![0]!, trim3[i]![0]!];
            subjectAbsences.push(currentAbsences.filter(Boolean).join(" / "));
        }
    }

    // biome-ignore lint/style/noNonNullAssertion: All arrays are guaranteed to be 1 row in height.
    const averages = [...trim1Totals[0]!, ...trim2Totals[0]!, ...trim3Totals[0]!] as Grades;

    const habilities = getRangeData(ReportSheetSchema.sheets.studentTemplate.ranges.abilities, { colOffset: 1, width: 4 });

    let currentIndex = 0;

    const subjectSections: SubjectSection[] = persistenData.academicFields.map((field) => {
        const sectionSubjects: SubjectGrades[] = [];

        for (let i = 0; i < field.subjects; i++) {
            sectionSubjects.push({
                // biome-ignore lint/style/noNonNullAssertion: Subjects must have all the subjects.
                name: persistenData.subjects[currentIndex]!.subject,
                // biome-ignore lint/style/noNonNullAssertion: Subjects must have all the subjects.
                grades: allGrades[currentIndex]!,
                habilities: habilities[currentIndex] as Grades,
            });

            currentIndex++;
        }

        return {
            name: field.name,
            color: field.color,
            subjects: sectionSubjects,
        };
    });

    const commentsRaw = getRangeData(ReportSheetSchema.sheets.studentTemplate.ranges.comments, { colOffset: 1, width: 1 });

    const cleanComment = (raw: string): string => {
        const trimmed = raw.trim();
        return DEFAULT_COMMENT.startsWith(trimmed) ? "" : raw;
    };

    const comments: SubjectComment[] = [];

    persistenData.subjects.forEach((subject, index) => {
        comments.push({
            name: subject.subject,
            absences: subjectAbsences[index],
            comment: cleanComment(commentsRaw[index]?.[0] ?? ""),
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

type StudentRangeOptions = Pick<OffsetGridRangeProperties, "colOffset" | "width">;

function getStudentRangeData(mappedRange: MappedNamedRange, studentSheet: GoogleAppsScript.Sheets.Schema.Sheet, options: StudentRangeOptions = {}): string[][] {
    const range = offsetGridRange({
        origin: mappedRange.namedRange.range,
        sheetId: studentSheet.properties?.sheetId,
        ...options,
    });
    return getStringRangeData(studentSheet, range);
}

function getStudentLastColumnsData(mappedRange: MappedNamedRange, studentSheet: GoogleAppsScript.Sheets.Schema.Sheet, width: number): string[][] {
    const { endColumnIndex = /* istanbul ignore next */ 0, startColumnIndex = /* istanbul ignore next */ 0 } = mappedRange.namedRange.range;
    return getStudentRangeData(mappedRange, studentSheet, { colOffset: endColumnIndex - startColumnIndex - width, width });
}
