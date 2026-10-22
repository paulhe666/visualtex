// Node module hooks for scripts/kernel-node.mts: `mathlive` and `mathlive/ssr`
// resolve to the editing kernel's Node build.
let kernelUrl;

export function initialize(data) {
  kernelUrl = data.kernelUrl;
}

export async function resolve(specifier, context, nextResolve) {
  if (specifier === "mathlive" || specifier === "mathlive/ssr") {
    return { url: kernelUrl, shortCircuit: true };
  }
  return nextResolve(specifier, context);
}
