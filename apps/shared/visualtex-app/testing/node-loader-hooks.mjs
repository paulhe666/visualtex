// Module hooks of node-loader.mjs.
let appUrl;

export function initialize(data) {
  appUrl = data.appUrl;
}

const fromShared = (url) => typeof url === "string" && url.includes("/apps/shared/");
const bare = (specifier) => !/^(?:\.{0,2}\/|[a-z][a-z0-9+.-]*:)/i.test(specifier);

export async function resolve(specifier, context, nextResolve) {
  if (specifier.startsWith("@platform/")) {
    const name = specifier.slice("@platform/".length).replace(/\.tsx?$/, "");
    for (const extension of [".ts", ".tsx"]) {
      try {
        return await nextResolve(new URL(`src/platform/${name}${extension}`, appUrl).href, context);
      } catch {
        // try the next extension
      }
    }
  }
  try {
    return await nextResolve(specifier, context);
  } catch (error) {
    if (!bare(specifier) || !fromShared(context.parentURL)) throw error;
    return nextResolve(specifier, { ...context, parentURL: new URL("package.json", appUrl).href });
  }
}
