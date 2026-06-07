export interface Range {
    readonly start: number;
    readonly end: number;
}

export interface Trimesters {
    readonly trim1Range: Range;
    readonly trim3Range: Range;
    readonly trim2Range: Range;
}

export interface SubjectBlockLayout {
    readonly subjectIndex: number;
    readonly titleFormatStartRow: number;
    readonly studentStartRow: number;
    readonly bandingStartRow: number;
    readonly bandingNumRows: number;
}

export interface CalendarGrid {
    readonly monthBlocks: MonthBlock[];
    readonly weeks: WeekData[];
}

export interface CalendarDates {
    readonly dateStart: number;
    readonly dateTrimester1: number;
    readonly dateTrimester2: number;
    readonly dateEnd: number;
    readonly calStart: number;
    readonly calEnd: number;
    readonly totalDays: number;
    readonly totalRows: number;
}

export interface MonthBlock {
    readonly startRow: number;
    readonly endRow: number;
    readonly monthIndex: number;
    readonly year: number;
}

export interface DayData {
    readonly ms: number;
    readonly dateNumber: number;
    readonly isWeekday: boolean;
    readonly inBounds: boolean;
    readonly dayType: DayType;
}

export interface WeekData {
    readonly rowNumber: number;
    readonly days: DayData[];
}

export interface SetupFileData {
    readonly folderId: string;
    readonly groupName: string;
    readonly attendancePerClass: boolean;
    readonly averagePerField: boolean;
    readonly dateStart: number;
    readonly dateEndTrimester1: number;
    readonly dateEndTrimester2: number;
    readonly dateEnd: number;
}

export enum TemplateSize {
    SMALL = 1,
    MEDIUM = 2,
    LARGE = 5,
}

export enum DayType {
    TRIM1 = "trimester 1 day",
    TRIM2 = "trimester 2 day",
    TRIM3 = "trimester 3 day",
    REST = "rest day",
}
