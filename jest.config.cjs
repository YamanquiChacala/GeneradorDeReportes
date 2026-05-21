const { createDefaultPreset } = require("ts-jest");

const tsJestTransformCfg = createDefaultPreset().transform;

/** @type {import("jest").Config} **/
module.exports = {
    testEnvironment: "node",
    transform: {
        ...tsJestTransformCfg,
    },
    testPathIgnorePatterns: [
        "/node_modules/",
        "\\.gas\\.test\\.ts$", // Tells Jest to completely ignore gas tests
    ],

    collectCoverageFrom: [
        "src/common/**/*.ts", // ONLY track coverage for files in src/common
        "!src/common/**/*.test.ts", // Exclude standard Jest test files from the report
        "!src/common/**/*.gas.test.ts", // Exclude GAS test files from the report
        "!src/common/**/*.d.ts", // Exclude TypeScript definition files
        "!src/common/gas-parts/*",
    ],
    // 'text' gives the terminal output. 'html' gives the browser report.
    // 'lcov' creates the lcov.info file needed for VSCode Coverage Gutter.
    coverageReporters: ["text", "html", "lcov"],
};
