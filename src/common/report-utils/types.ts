export interface ReportPersistentData {
    configData: ConfigData;
    protectedSections: ProtectedSections;
    academicFields: AcademicField[];
    subjects: WeightedSubject[];
    students: StudentRow[];
    calendar: number[];
}

export interface ConfigData {
    attendancePerClass: boolean;
    averagePerField: boolean;
    dates: [number, number, number, number];
    subjectGradingWeights: [number, number, number];
}

interface ProtectedSections {
    habilities: boolean;
    comments: boolean;
    trimesters: [boolean, boolean, boolean];
}

export interface AcademicField {
    name: string;
    color: string;
    subjects: number;
}

export interface WeightedSubject {
    subject: string;
    weight: number;
}

export type StudentRow = Student | StudentSpace;

export enum StudentRowType {
    STUDENT = "STUDENT",
    SEPARATOR = "SEPARATOR",
}

export interface Student {
    type: StudentRowType.STUDENT;
    id: number;
    firstName: string;
    lastName: string;
    sheetName: string;
    sex: string;
    level: string;
    grade: string;
    curp: string;
}

interface StudentSpace {
    type: StudentRowType.SEPARATOR;
}

export enum Period {
    FIRST = 0,
    SECOND = 1,
    THIRD = 2,
}

export interface FrozenArea {
    rows: number;
    cols: number;
}

export interface Range {
    readonly start: number;
    readonly end: number;
}

export interface TrimesterRanges {
    readonly trim1: Range;
    readonly trim2: Range;
    readonly trim3: Range;
}
