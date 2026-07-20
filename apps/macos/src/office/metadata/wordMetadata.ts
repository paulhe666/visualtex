import {
  formulaMetadataFromXml,
  formulaMetadataToXml,
  VISUALTEX_FORMULA_XML_NAMESPACE,
  type VisualTeXFormulaMetadata,
} from "./formulaMetadata";

export type WordMetadataStorage = "custom-xml" | "picture-alt-text";

function customXmlPartsAvailable() {
  try {
    return (
      Office.context.requirements.isSetSupported("CustomXmlParts", "1.1") &&
      Boolean(Office.context.document.customXmlParts)
    );
  } catch {
    return false;
  }
}

function officeAsync<T>(
  invoke: (callback: (result: Office.AsyncResult<T>) => void) => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    invoke((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(
          new Error(
            result.error?.message ?? "Office Custom XML operation failed.",
          ),
        );
      }
    });
  });
}

function getPartsByNamespace() {
  return officeAsync<Office.CustomXmlPart[]>((callback) =>
    Office.context.document.customXmlParts.getByNamespaceAsync(
      VISUALTEX_FORMULA_XML_NAMESPACE,
      callback,
    ),
  );
}

function getPartXml(part: Office.CustomXmlPart) {
  return officeAsync<string>((callback) => part.getXmlAsync(callback));
}

function deletePart(part: Office.CustomXmlPart) {
  return officeAsync<void>((callback) => part.deleteAsync(callback));
}

function addPart(xml: string) {
  return officeAsync<Office.CustomXmlPart>((callback) =>
    Office.context.document.customXmlParts.addAsync(xml, callback),
  );
}

export async function readWordDocumentMetadata(formulaId: string) {
  if (!customXmlPartsAvailable()) return null;
  try {
    const parts = await getPartsByNamespace();
    for (const part of parts) {
      const metadata = formulaMetadataFromXml(await getPartXml(part));
      if (metadata?.formulaId === formulaId) return metadata;
    }
  } catch {
    return null;
  }
  return null;
}

export async function writeWordDocumentMetadata(
  metadata: VisualTeXFormulaMetadata,
): Promise<WordMetadataStorage> {
  if (!customXmlPartsAvailable()) return "picture-alt-text";
  try {
    const parts = await getPartsByNamespace();
    for (const part of parts) {
      const current = formulaMetadataFromXml(await getPartXml(part));
      if (current?.formulaId === metadata.formulaId) {
        await deletePart(part);
      }
    }
    await addPart(formulaMetadataToXml(metadata));
    return "custom-xml";
  } catch {
    return "picture-alt-text";
  }
}
