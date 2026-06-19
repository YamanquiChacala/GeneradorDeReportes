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

/**
 * Simple Zip function as a generator
 */
export function* zip<A, B>(arr1: A[], arr2: B[]): Generator<[A, B]> {
    const lenght = Math.min(arr1.length, arr2.length);

    for (let i = 0; i < lenght; i++) {
        const item1 = arr1[i];
        const item2 = arr2[i];

        if (item1 != null && item2 != null) {
            yield [item1, item2];
        }
    }
}
