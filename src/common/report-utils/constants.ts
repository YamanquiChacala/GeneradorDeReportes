export const TRIMESTER_NAMES = ["1er Momento", "2do Momento", "3er Momento"] as const;

export const DEFAULT_COMMENT =
    "Escribe aquí tus comentarios, comienza con un comentario positivo, seguido de tus observaciones. " +
    "Concluye con un comentario alentador, felicitaciones y/o recomendaciones.\n\n" +
    "¡No te olvides de revisar la ortografía y redacción!";

export const DEFAULT_SEP_STRENGHT = [
    "[Escribe el logro del alumno de forma directa. Usa este espacio como límite visual]",
    "[Escribe el logro del alumno directamente]",
    "[Escribe un logro directo.]",
] as const;

export const DEFAULT_SEP_WEAKNESS = [
    "[Indica directamente el área a mejorar. Este texto te sirve como límite de tamaño]",
    "[Indica directo el tema a fortalecer aquí]",
    "[Tema directo a fortalecer.]",
] as const;

export const DEFAULT_SEP_SUGGESTION = [
    "[Escribe una recomendación directa. Este texto muestra el límite máximo de letras]",
    "[Anota una acción directa para su mejora.]",
    "[Sugiere una acción directa]",
] as const;
