import { mkdir } from "node:fs/promises";
import path from "node:path";
import { build } from "esbuild";

const rootDir = process.cwd();
const outDir = path.resolve(rootDir, "dist");
const pathShim = path.resolve(rootDir, "src", "shims", "node-path-browser.js");

const nodeBuiltinAliasPlugin = {
  name: "node-builtin-alias",
  setup(buildContext) {
    buildContext.onResolve({ filter: /^node:path$/ }, () => ({ path: pathShim }));
  }
};

const commonOptions = {
  entryPoints: [path.resolve(rootDir, "src", "browser.js")],
  bundle: true,
  format: "iife",
  globalName: "pTxTq",
  platform: "browser",
  target: ["es2020"],
  charset: "utf8",
  sourcemap: true,
  plugins: [nodeBuiltinAliasPlugin],
  banner: {
    js: "/* pTxTq browser bundle */"
  }
};

await mkdir(outDir, { recursive: true });

await build({
  ...commonOptions,
  outfile: path.join(outDir, "pTxTq.js")
});

await build({
  ...commonOptions,
  minify: true,
  outfile: path.join(outDir, "pTxTq.min.js")
});

await build({
  ...commonOptions,
  minify: true,
  sourcemap: false,
  outfile: path.join(rootDir, "ptxtq.min.js")
});

console.log(`Built ${path.join(outDir, "pTxTq.js")}`);
console.log(`Built ${path.join(outDir, "pTxTq.min.js")}`);
console.log(`Built ${path.join(rootDir, "ptxtq.min.js")}`);
