import { useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { ExternalLink } from "lucide-react";
import { macOfficeFullDiskAccessSteps } from "./macOfficeOnboarding";

/** The Full Disk Access steps for this macOS release, with Open Settings. */
export function MacOfficeFullDiskAccessHelp({
  isEn,
  updating,
}: {
  isEn: boolean;
  updating: boolean;
}) {
  const [macosMajor, setMacosMajor] = useState<number | null>(null);
  const [openError, setOpenError] = useState("");

  useEffect(() => {
    let cancelled = false;
    void invoke<number | null>("get_macos_major_version")
      .then((major) => {
        if (!cancelled) setMacosMajor(typeof major === "number" ? major : null);
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <div className="office-full-disk-access-help">
      <ol>
        {macOfficeFullDiskAccessSteps(isEn, macosMajor, updating).map((step) => (
          <li key={step}>{step}</li>
        ))}
      </ol>
      <button
        type="button"
        className="secondary-button"
        onClick={() => {
          setOpenError("");
          void invoke("open_macos_full_disk_access_settings").catch((error: unknown) => {
            setOpenError(error instanceof Error ? error.message : String(error));
          });
        }}
      >
        <ExternalLink size={14} />
        {isEn ? "Open Settings" : "打开设置"}
      </button>
      {openError && <p className="is-warning">{openError}</p>}
    </div>
  );
}
