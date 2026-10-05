// import { getEpochDate } from "../gas-utils";
// import type { ReportPersistentData } from "./types";

// interface ParsePersistentDataParams {
//     // Config Data
//     attendancePerClass: GoogleAppsScript.Sheets.Schema.CellData[][];
//     averagePerField: GoogleAppsScript.Sheets.Schema.CellData[][];
//     dates: GoogleAppsScript.Sheets.Schema.CellData[][];
//     subjectGradingWeights: GoogleAppsScript.Sheets.Schema.CellData[][];
//     // Protected sections
//     protect_habilities: GoogleAppsScript.Sheets.Schema.CellData[][];
//     protect_comments: GoogleAppsScript.Sheets.Schema.CellData[][];
//     protect_trimesters: GoogleAppsScript.Sheets.Schema.CellData[][];
//     // Other stuff
//     academicFields: GoogleAppsScript.Sheets.Schema.CellData[][];
//     subjects: GoogleAppsScript.Sheets.Schema.CellData[][];
//     students: GoogleAppsScript.Sheets.Schema.CellData[][];
//     calendar: GoogleAppsScript.Sheets.Schema.CellData[][];
// }

// export function parsePersistentData(params: ParsePersistentDataParams): ReportPersistentData {
//     const attendancePerClass = throwIfNull(params.attendancePerClass[0]?.[0]?.effectiveValue?.boolValue);
//     const averagePerField = throwIfNull(params.averagePerField[0]?.[0]?.effectiveValue?.boolValue);

//     const dates: [number, number, number, number] = [
//         getEpochDate(throwIfNull(params.dates[0]?.[0]?.effectiveValue?.numberValue)),
//         getEpochDate(throwIfNull(params.dates[1]?.[0]?.effectiveValue?.numberValue)),
//         getEpochDate(throwIfNull(params.dates[2]?.[0]?.effectiveValue?.numberValue)),
//         getEpochDate(throwIfNull(params.dates[3]?.[0]?.effectiveValue?.numberValue)),
//     ];

//     const subjectGradingWeights: [number, number, number] = [
//         throwIfNull(params.subjectGradingWeights[0]?.[0]?.effectiveValue?.numberValue),
//         throwIfNull(params.subjectGradingWeights[1]?.[0]?.effectiveValue?.numberValue),
//         throwIfNull(params.subjectGradingWeights[2]?.[0]?.effectiveValue?.numberValue),
//     ];
// }

// function throwIfNull<T>(value: T): NonNullable<T> {
//     if (value == null) {
//         throw new Error("Falta dato de hoja '_Persistente'");
//     }
//     return value;
// }
