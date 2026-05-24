/**
 * Calculates how much to move in a circle of `n` elements to go over every element without repeating, while jumping as much as possible.
 */
export function getStep(n: number) {
    const target = Math.floor(n / 2);
    for (let delta = 0; delta <= target; delta++) {
        const k1 = target - delta;
        const k2 = target + delta;

        if (k1 > 0 && gcd(n, k1) === 1) {
            return k1;
        }
        if (k2 < n && gcd(n, k2) === 1) {
            return k2;
        }
    }
    return 1;
}

/**
 * Calculates the grates common denominator.
 */
export function gcd(a: number, b: number) {
    while (b !== 0) {
        const t = b;
        b = a % b;
        a = t;
    }
    return a;
}

/**
 * Binary search
 * Returns the index of the first element in the array strictly greater than the threshold.
 * If no element is greater than threshold, returns the lenght of the array.
 */
export function getUpperBoundIndex(orderedArray: readonly number[], threshold: number): number {
    let left = 0;
    let right = orderedArray.length - 1;
    let bestIndex = orderedArray.length;

    while (left <= right) {
        const mid = Math.floor((left + right) / 2);
        // biome-ignore lint/style/noNonNullAssertion: By definition in the middle of the array.
        const midVal = orderedArray[mid]!;

        if (midVal > threshold) {
            bestIndex = mid;
            right = mid - 1; // It's strictly greater, but look left to find an earlier one
        } else {
            left = mid + 1; // It's <= threshold, look right
        }
    }

    return bestIndex;
}
