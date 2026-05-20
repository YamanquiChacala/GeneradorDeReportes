/**
 * Replacement for Ojbect.entries that preserves types (as much as possible).
 */
export function typedEntries<T extends object>(
    obj: T,
): {
    [K in keyof T]: [K, T[K]];
}[keyof T][] {
    return Object.entries(obj) as {
        [K in keyof T]: [K, T[K]];
    }[keyof T][];
}
