import { useCallback, useEffect, useMemo, useState } from "react";
import { errorMessage } from "../runtime/errorMessage";
import {
  getOcrProviderConfiguration,
  isOfficeCompanionEnvironment,
  isTauriEnvironment,
  listenOcrProviderConfigurationChanged,
  saveOcrProviderConfiguration,
  type OcrModelName,
  type OcrProviderConfiguration,
} from "./ocrService";
import {
  activeOcrQuickSelectionId,
  buildOcrQuickSelectionOptions,
  createOcrProviderQuickUpdate,
  LOCAL_OCR_QUICK_CONFIGURATION,
  parseOcrQuickSelection,
} from "./ocrQuickSelection";

interface UseOcrQuickSelectionOptions {
  model: OcrModelName;
  busy: boolean;
  isEn: boolean;
  onModelChange: (model: OcrModelName) => void;
  onError: (message: string) => void;
}

export function useOcrQuickSelection({
  model,
  busy,
  isEn,
  onModelChange,
  onError,
}: UseOcrQuickSelectionOptions) {
  const [configuration, setConfiguration] =
    useState<OcrProviderConfiguration | null>(null);
  const [changing, setChanging] = useState(false);

  useEffect(() => {
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) return;
    let disposed = false;
    let unlisten: (() => void) | undefined;

    void listenOcrProviderConfigurationChanged((nextConfiguration) => {
      if (!disposed) setConfiguration(nextConfiguration);
    })
      .then((stopListening) => {
        if (disposed) stopListening();
        else unlisten = stopListening;
      })
      .catch(() => undefined);

    void getOcrProviderConfiguration()
      .then((nextConfiguration) => {
        if (!disposed) setConfiguration(nextConfiguration);
      })
      .catch((reason) => {
        if (disposed) return;
        setConfiguration(LOCAL_OCR_QUICK_CONFIGURATION);
        onError(
          errorMessage(
            reason,
            isEn
              ? "Unable to load the OCR provider"
              : "无法读取 OCR 提供器设置",
          ),
        );
      });

    return () => {
      disposed = true;
      unlisten?.();
    };
  }, [isEn, onError]);

  const options = useMemo(
    () =>
      configuration ? buildOcrQuickSelectionOptions(configuration) : [],
    [configuration],
  );
  const selection = configuration
    ? activeOcrQuickSelectionId(configuration, model)
    : "";
  const activeOption = options.find((option) => option.id === selection) ?? null;

  const handleSelectionChange = useCallback(
    async (value: string) => {
      const parsed = parseOcrQuickSelection(value);
      if (!configuration || !parsed || busy || changing || value === selection) {
        return;
      }

      setChanging(true);
      try {
        if (parsed.kind === "local") {
          if (configuration.activeProvider !== "local") {
            const saved = await saveOcrProviderConfiguration(
              createOcrProviderQuickUpdate(configuration, null),
            );
            setConfiguration(saved);
          }
          onModelChange(parsed.model);
          return;
        }

        const saved = await saveOcrProviderConfiguration(
          createOcrProviderQuickUpdate(configuration, parsed),
        );
        setConfiguration(saved);
      } catch (reason) {
        onError(
          errorMessage(
            reason,
            isEn
              ? "Unable to switch the OCR provider"
              : "无法切换 OCR 提供器",
          ),
        );
      } finally {
        setChanging(false);
      }
    }, [busy, changing, configuration, isEn, onError, onModelChange, selection]);

  return {
    selection,
    options,
    activeOption,
    busy: busy || changing,
    handleSelectionChange,
  };
}
