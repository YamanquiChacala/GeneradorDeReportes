export * from "./data";
export * from "./dates";
export * from "./math";
export * from "./parse";
export * from "./requests";

export type DayType = "trimester1" | "trimester2" | "trimester3" | "rest";

export interface Range {
    start: number;
    end: number;
}

export interface Trimesters {
    trim1Range: Range;
    trim2Range: Range;
    trim3Range: Range;
}

export interface SubjectBlockLayout {
    subjectIndex: number;
    titleFormatStartRow: number;
    studentStartRow: number;
    bandingStartRow: number;
    bandingNumRows: number;
}

export interface CalendarGrid {
    monthBlocks: MonthBlock[];
    weeks: WeekData[];
}

export interface CalendarDates {
    dateStart: number;
    dateTrimester1: number;
    dateTrimester2: number;
    dateEnd: number;
    calStart: number;
    calEnd: number;
    totalDays: number;
    totalRows: number;
}

export interface MonthBlock {
    startRow: number;
    endRow: number;
    monthIndex: number;
    year: number;
}

export interface DayData {
    ms: number;
    dateNumber: number;
    isWeekday: boolean;
    inBounds: boolean;
    dayType: DayType;
}

export interface WeekData {
    rowNumber: number;
    days: DayData[];
}

export interface SetupFileData {
    folderId: string;
    groupName: string;
    attendancePerClass: boolean;
    averagePerField: boolean;
    dateStart: number;
    dateEndTrimester1: number;
    dateEndTrimester2: number;
    dateEnd: number;
}
