import gasPlugin from "@gas-plugin/unplugin/esbuild";
import { build } from "esbuild";

build({
    entryPoints: ["src/main.ts"],
    bundle: true,
    outfile: "dist/Code.js",
    format: "esm",
    target: "es2019",
    plugins: [
        gasPlugin({
            manifest: "src/appsscript.json",
            include: ["src/**/*.html"],
            autoGlobals: false,
        }),
    ],
}).catch(() => process.exit(1));
