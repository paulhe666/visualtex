import type { MacOfflineOfficeStatus } from "./macOfficeStatusValidation";

export function resolveMacOfficeOnboarding(
  status: MacOfflineOfficeStatus,
  previouslyConfigured: boolean,
): { mode: "setup" | "update" | "repair"; powerpointRegistrationRequired: boolean } | null {
  if (!status.compiledArtifactsAvailable) return null;

  const hosts = [status.word, status.powerpoint];
  const outdatedHosts = hosts.filter(
    (host) => host.applicationInstalled && !host.filesInstalled,
  );
  if (outdatedHosts.length === 0) return null;

  const mode = outdatedHosts.some((host) => host.filesPresent)
    ? "update"
    : previouslyConfigured || hosts.some((host) => host.filesPresent)
      ? "repair"
      : "setup";

  return {
    mode,
    powerpointRegistrationRequired:
      status.powerpoint.applicationInstalled && !status.powerpoint.filesPresent,
  };
}
