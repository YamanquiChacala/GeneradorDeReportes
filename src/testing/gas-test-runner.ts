type TestCallback = () => void;

interface TestSuite {
    name: string;
    tests: { name: string; fn: TestCallback }[];
    beforeAlls: TestCallback[];
    afterAlls: TestCallback[];
}

/**
 * Usage:
 * const runner = new GasTestRunner();
 * const { describe, test, beforeAll, afterAll, expect } = runner;
 * ...
 * runner.execute();
 */
export class GasTestRunner {
    private logs: string[] = [];
    private passed = 0;
    private failed = 0;

    private suites: TestSuite[] = [];

    private currentSuite: TestSuite | null = null;

    private log(msg: string) {
        this.logs.push(msg);
    }

    // --- JEST API ---

    describe = (name: string, fn: TestCallback) => {
        const suite: TestSuite = { name, tests: [], beforeAlls: [], afterAlls: [] };
        this.suites.push(suite);
        this.currentSuite = suite;
        fn(); // Execute to register tests and hooks
        this.currentSuite = null;
    };

    test = (name: string, fn: TestCallback) => {
        if (!this.currentSuite) {
            // Fallback for tests outside a describe block
            this.describe("Global", () => {
                this.currentSuite?.tests.push({ name, fn });
            });
            return;
        }
        this.currentSuite.tests.push({ name, fn });
    };

    beforeAll = (fn: TestCallback) => {
        if (this.currentSuite) this.currentSuite.beforeAlls.push(fn);
    };

    afterAll = (fn: TestCallback) => {
        if (this.currentSuite) this.currentSuite.afterAlls.push(fn);
    };

    expect = (actual: unknown) => {
        return {
            toBe: (expected: unknown) => {
                if (actual !== expected) {
                    // Convert to strings for safer error logging when dealing with unknown types
                    throw new Error(`Expected ${String(expected)}, but got ${String(actual)}`);
                }
            },
            toEqual: (expected: unknown) => {
                if (JSON.stringify(actual) !== JSON.stringify(expected)) {
                    throw new Error(`Expected ${JSON.stringify(expected)}, but got ${JSON.stringify(actual)}`);
                }
            },
            toBeUndefined: () => {
                if (actual !== undefined) throw new Error(`Expected undefined, but got ${String(actual)}`);
            },
            toBeTruthy: () => {
                if (!actual) throw new Error(`Expected truthy, but got ${String(actual)}`);
            },
            toBeFalsy: () => {
                if (actual) throw new Error(`Expected falsy, but got ${String(actual)}`);
            },
        };
    };

    // --- RUNNER ---

    execute = () => {
        this.log("--- Starting Test Execution ---");

        for (const suite of this.suites) {
            this.log(`\n📦 Suite: ${suite.name}`);
            try {
                // Setup
                for (const fn of suite.beforeAlls) {
                    fn();
                }

                // Execute Tests
                for (const t of suite.tests) {
                    try {
                        t.fn();
                        this.passed++;
                        this.log(`  ✅ PASSED: ${t.name}`);
                    } catch (e) {
                        this.failed++;
                        const errorMessage = e instanceof Error ? e.message : String(e);
                        this.log(`  ❌ FAILED: ${t.name}\n     -> ${errorMessage}`);
                    }
                }
            } catch (e) {
                const errorMessage = e instanceof Error ? e.message : String(e);
                this.log(`🚨 SUITE FATAL ERROR (Setup): ${errorMessage}`);
                this.failed++;
            } finally {
                // Teardown
                try {
                    for (const fn of suite.afterAlls) {
                        fn();
                    }
                } catch (e) {
                    const errorMessage = e instanceof Error ? e.message : String(e);
                    this.log(`🚨 FAILED TO CLEAN UP: ${errorMessage}`);
                }
            }
        }

        this.log(`\n--- Test Complete: ${this.passed} Passed, ${this.failed} Failed ---`);
        const finalOutput = this.logs.join("\n");
        console.log(finalOutput);

        return {
            success: this.failed === 0,
            passed: this.passed,
            failed: this.failed,
            logs: this.logs,
        };
    };
}
