export * from "./base64-constants";
export * from "./enums";
export * from "./gas-enums";

export const MS_PER_DAY = 24 * 60 * 60 * 1000;
export const MORE_THAN_A_YEAR = 400 * MS_PER_DAY;
export const SHEETS_EPOCH_OFFSET = 25569; // Days from Dec 30, 1899 (Sheets) to Jan 1, 1970 (Unix)

export const SETUP_FILE_PREFIX = "__Registro Inicial - ";

export const BASE_ASSISTANCE_PROTECTED_RANGE = "asist_per";

export const MONTH_NAMES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"] as const;
export const TRIMESTER_NAMES = ["1er Momento", "2do Momento", "3er Momento"] as const;
export const DEFAULT_COMMENT =
    "Escribe aquí tus comentarios, comienza con un comentario positivo, seguido de tus observaciones. " +
    "Concluye con un comentario alentador, felicitaciones y/o recomendaciones. " +
    "¡No te olvides de revisar la ortografía y redacción!\n" +
    "=> Para la SEP, sólo se enviará el primer párrafo.\n" +
    "FORTALEZAS:\n" +
    "ÁREAS DE OPORTUNIDAD:\n" +
    "SUGERENCIAS";

// Banding coloring
export const LIGHT_GREY_COLOR: GoogleAppsScript.Sheets.Schema.Color = { red: 0.98, green: 0.98, blue: 0.98 };
export const BAND_LIGHT = 0.95;
export const HEADER_SAT = 0.75;
export const HEADER_LIGH = 0.2;
export const FOOTER_SAT = 0.5;
export const FOOTER_LIGH = 0.7;
