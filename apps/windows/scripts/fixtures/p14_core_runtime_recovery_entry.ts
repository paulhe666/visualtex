import {
  createHeadlessCoreFormulaOwner,
  getCoreEngineRuntime,
  isCoreRuntimeFault,
} from "../../src/engine/coreEngineRuntime";

interface RecoveryReport {
  status: "PASS";
  businessErrorKeptRuntime: boolean;
  fatalErrorRecognized: boolean;
  checkpointSource: string;
  recoveredSource: string;
  runtimeRebuilt: boolean;
  recoveredLayoutSvg: boolean;
  recoveredMathMl: boolean;
  oldOwnerStillReadable: boolean;
}

declare global {
  interface Window {
    p14RuntimeRecovery: { run(): Promise<RecoveryReport> };
  }
}

window.p14RuntimeRecovery = {
  async run() {
    const initialContext = await getCoreEngineRuntime();

    const unsupported = await createHeadlessCoreFormulaOwner(
      "p14-business-error",
      String.raw`\visualtexDefinitelyUnsupported{x}`,
    );
    let businessFailed = false;
    try {
      unsupported.owner.layout(initialContext.baseLayout, false);
    } catch (error) {
      businessFailed = !isCoreRuntimeFault(error);
    } finally {
      try { unsupported.owner.destroy(); } catch {}
    }
    if (!businessFailed) throw new Error("Unsupported formula did not fail as a normal business error.");
    const afterBusiness = await getCoreEngineRuntime();
    if (afterBusiness.runtime !== initialContext.runtime) {
      throw new Error("A normal unsupported formula incorrectly invalidated the shared core runtime.");
    }

    const working = await createHeadlessCoreFormulaOwner("p14-fatal-checkpoint", "x+1");
    working.owner.replaceSource("x+1+q");
    const checkpoint = working.owner.inspect();
    const checkpointSource = checkpoint.document.source;
    const activeBeforeFault = initialContext.runtime.activeSessions;
    if (activeBeforeFault < 1) throw new Error("Expected an active session before fault injection.");

    initialContext.runtime.free();
    let fatalError: unknown = null;
    try {
      working.owner.presentationMathMl(false);
    } catch (error) {
      fatalError = error;
    }
    if (!isCoreRuntimeFault(fatalError)) {
      throw new Error(`Fatal runtime invalidation was not classified correctly: ${String(fatalError)}`);
    }
    const oldOwnerStillReadable = working.owner.inspect().document.source === checkpointSource;
    if (!oldOwnerStillReadable) throw new Error("Old owner lost its cached checkpoint after runtime failure.");
    try { working.owner.destroy(); } catch {}

    const recoveredContext = await getCoreEngineRuntime();
    if (recoveredContext.runtime === initialContext.runtime) {
      throw new Error("Core runtime was not rebuilt after fatal invalidation.");
    }
    const recovered = await createHeadlessCoreFormulaOwner("p14-fatal-checkpoint-recovered", checkpointSource);
    let recoveredLayoutSvg = false;
    let recoveredMathMl = false;
    let recoveredSource = "";
    try {
      const layout = recovered.owner.layout(recoveredContext.baseLayout, true);
      recoveredLayoutSvg = Boolean(layout.svg?.includes("<svg"));
      recoveredMathMl = recovered.owner.presentationMathMl(false).mathMl.startsWith("<math");
      recoveredSource = recovered.owner.inspect().document.source;
    } finally {
      recovered.owner.destroy();
    }
    if (!recoveredLayoutSvg || !recoveredMathMl || recoveredSource !== checkpointSource) {
      throw new Error("Recovered runtime did not reconstruct the exact last committed source and exports.");
    }

    return {
      status: "PASS",
      businessErrorKeptRuntime: true,
      fatalErrorRecognized: true,
      checkpointSource,
      recoveredSource,
      runtimeRebuilt: true,
      recoveredLayoutSvg,
      recoveredMathMl,
      oldOwnerStillReadable,
    };
  },
};
