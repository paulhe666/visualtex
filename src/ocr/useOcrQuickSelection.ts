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
  enabled?: boolean;
  onModelChange: (model: OcrModelName) => void;
  onError: (message: string) => void;
}

export function useOcrQuickSelection({
  model,
  busy,
  isEn,
  enabled = true,
  onModelChange,
  onError,
}: UseOcrQuickSelectionOptions) {
  const [configuration, setConfiguration] =
    useState<OcrProviderConfiguration>(LOCAL_OCR_QUICK_CONFIGURATION);
  const [configurationLoaded, setConfigurationLoaded] = useState(false);
  const [changing, setChanging] = useState(false);

  const loadConfiguration = useCallback(async () => {
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) {
      return configuration;
    }
    try {
      const nextConfiguration = await getOcrProviderConfiguration();
      setConfiguration(nextConfiguration);
      setConfigurationLoaded(true);
      return nextConfiguration;
    } catch (reason) {
      setConfiguration(LOCAL_OCR_QUICK_CONFIGURATION);
      onError(
        errorMessage(
          reason,
          isEn
            ? "Unable to load the OCR provider"
            : "无法读取 OCR 提供器设置",
        ),
      );
      return LOCAL_OCR_QUICK_CONFIGURATION;
    }
  }, [configuration, isEn, onError]);

  useEffect(() => {
    if (!enabled) return;
    if (!isTauriEnvironment() && !isOfficeCompanionEnvironment()) return;
    let disposed = false;
    let unlisten: (() => void) | undefined;

    void listenOcrProviderConfigurationChanged((nextConfiguration) => {
      if (!disposed) {
        setConfiguration(nextConfiguration);
        setConfigurationLoaded(true);
      }
    })
      .then((stopListening) => {
        if (disposed) stopListening();
        else unlisten = stopListening;
      })
      .catch(() => undefined);

    void loadConfiguration();

    return () => {
      disposed = true;
      unlisten?.();
    };
  }, [enabled, loadConfiguration]);

  const options = useMemo(
    () => buildOcrQuickSelectionOptions(configuration),
    [configuration],
  );
  const selection = activeOcrQuickSelectionId(configuration, model);
  const activeOption = options.find((option) => option.id === selection) ?? null;

  const handleSelectionChange = useCallback(
    async (value: string) => {
      const parsed = parseOcrQuickSelection(value);
      if (!parsed || busy || changing || value === selection) {
        return;
      }

      setChanging(true);
      try {
        const currentConfiguration = configurationLoaded
          ? configuration
          : await loadConfiguration();
        if (parsed.kind === "local") {
          if (currentConfiguration.activeProvider !== "local") {
            const saved = await saveOcrProviderConfiguration(
              createOcrProviderQuickUpdate(currentConfiguration, null),
            );
            setConfiguration(saved);
          }
          onModelChange(parsed.model);
          return;
        }

        const saved = await saveOcrProviderConfiguration(
          createOcrProviderQuickUpdate(currentConfiguration, parsed),
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
    }, [
      busy,
      changing,
      configuration,
      configurationLoaded,
      loadConfiguration,
      onModelChange,
      selection,
    ]);

  return {
    selection,
    options,
    activeOption,
    busy: busy || changing,
    loadConfiguration,
    handleSelectionChange,
  };
}
