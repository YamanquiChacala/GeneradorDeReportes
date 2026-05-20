const fs = require("node:fs");
const path = require("node:path");

try {
    // Read your dark mode css
    const darkCss = fs.readFileSync(path.resolve(__dirname, "dark-coverage.css"));

    // Append it to Jest's base.css
    fs.appendFileSync(path.resolve(__dirname, "coverage/base.css"), `\n${darkCss}`);

    console.log("Dark mode coverage styles applied!");
} catch (error) {
    console.error("Failed to append dark mode styles:", error.message);
}
