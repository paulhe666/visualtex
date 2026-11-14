// Node scripts (tests, generators) import shared code from apps/shared, which has no node_modules of its own:
// packages and "@platform/<module>" resolve from the app the script runs in (the current directory).
//   tsx --import ../shared/visualtex-app/testing/node-loader.mjs scripts/<script>.mts
import { register } from "node:module";
import { pathToFileURL } from "node:url";

register("./node-loader-hooks.mjs", import.meta.url, {
  data: { appUrl: pathToFileURL(`${process.cwd()}/`).href },
});
