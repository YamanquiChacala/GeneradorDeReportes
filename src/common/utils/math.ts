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
