export enum FileType {
    SETUP = "MontessoriChacalaSchoolGroupSetup",
    REPORT = "MontessoriChacalaSchoolReport",
}

export enum Icon {
    WARNING = "material-symbols/warning-rounded",
    FOLDER_QUESTION = "mdi/folder-question",
    CLIPBOARD = "noto/clipboard",
    CHART = "noto/bar-chart",
}

export enum DriveFiles {
    INITIALIZATION_TEMPLATE_ID = "19WMef0XLfSNkK48IQDQ1WUxeLa4Ir1LfvHZ4W5xEiIk",
}

export enum Urls {
    MEDIA_SERVER = "https://media.githubusercontent.com/media/YamanquiChacala/GeneradorDeReportes/refs/heads/main/",
}

export enum Tempates {
    HTML_TO_PDF_TEMPLATE = "PdfPrintTemplate",
}

export enum UserRoles {
    ADMIN,
}

export const Users: Record<UserRoles, readonly string[]> = {
    [UserRoles.ADMIN]: ["info@chacala.school"],
};
