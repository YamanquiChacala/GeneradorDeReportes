import { build } from "esbuild";
import gasPlugin from "@gas-plugin/unplugin/esbuild";

build({
    entryPoints: ["src/main.ts"],
    bundle: true,
    outfile: "dist/Code.js",
    format: "esm",
    target: "es2022",
    plugins: [
        gasPlugin({
            manifest: "src/appsscript.json",
            include: ["src/**/*.html"],
            autoGlobals: true
        })
    ],
}).catch(() => process.exit(1));