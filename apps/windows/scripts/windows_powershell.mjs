import path from "node:path";

export function windowsPowerShellPath(env = process.env) {
  const windowsRoot = env.SystemRoot ?? env.WINDIR;
  if (!windowsRoot) {
    throw new Error("SystemRoot/WINDIR is unavailable; cannot locate Windows PowerShell");
  }
  return path.join(
    windowsRoot,
    "System32",
    "WindowsPowerShell",
    "v1.0",
    "powershell.exe",
  );
}

export function windowsPowerShellEnvironment(env = process.env) {
  const result = { ...env };
  for (const key of Object.keys(result)) {
    if (key.toLowerCase() === "psmodulepath") delete result[key];
  }
  return result;
}
