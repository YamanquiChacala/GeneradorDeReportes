import { REPORT_CSS_STYLE } from "./report.css";

interface Grade {
    subject: string;
    score: number;
}
export interface ReportData {
    studentName: string;
    photoBase64?: string;
    grades: Grade[];
    comments?: string;
}

const GradesTable = ({ grades }: { grades: Grade[] }) => (
    <table>
        <tbody>
            {grades.map((g) => (
                <tr>
                    <td safe>{g.subject}</td>
                    <td>{g.score}</td>
                </tr>
            ))}
        </tbody>
    </table>
);

const Report = ({ data }: { data: ReportData }) => (
    <html lang="es">
        <head>
            <style>{REPORT_CSS_STYLE}</style>
        </head>
        <body>
            <h1 safe>{data.studentName}</h1>
            {data.photoBase64 && <img src={`${data.photoBase64}`} alt="nada" />}
            <GradesTable grades={data.grades} />
            {data.comments && <p safe>{data.comments}</p>}
        </body>
    </html>
);

export function renderReport(data: ReportData): string {
    const out = <Report data={data} />;
    if (typeof out !== "string") throw new Error("Async components aren't supported");
    return out;
}
