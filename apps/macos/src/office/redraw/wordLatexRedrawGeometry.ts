import type { VisualTeXFormulaMetadata } from "../shared/formulaMetadata";

const WORD_IMAGE_VISUAL_SCALE = 1.1;

export type ReusableWordImageGeometry = {
  width: number;
  height: number;
  baseline: number;
  renderWidthPx: number;
  renderHeightPx: number;
  referenceWidthPt: number;
  referenceHeightPt: number;
  referenceBaselinePt: number;
  inkCenterYRatio?: number;
};

export function reusableImageGeometry(
  metadata: VisualTeXFormulaMetadata | undefined,
): ReusableWordImageGeometry | null {
  if (!metadata) return null;
  const referenceWidthPt = metadata.referenceWidthPt;
  const referenceHeightPt = metadata.referenceHeightPt;
  const referenceBaselinePt = metadata.referenceBaselinePt;
  if (
    !referenceWidthPt ||
    !Number.isFinite(referenceWidthPt) ||
    !referenceHeightPt ||
    !Number.isFinite(referenceHeightPt) ||
    referenceBaselinePt === undefined ||
    !Number.isFinite(referenceBaselinePt)
  ) {
    return null;
  }

  const renderWidthPx =
    metadata.renderWidthPx && Number.isFinite(metadata.renderWidthPx)
      ? metadata.renderWidthPx
      : referenceWidthPt / (0.75 * WORD_IMAGE_VISUAL_SCALE);
  const renderHeightPx =
    metadata.renderHeightPx && Number.isFinite(metadata.renderHeightPx)
      ? metadata.renderHeightPx
      : referenceHeightPt / (0.75 * WORD_IMAGE_VISUAL_SCALE);
  const baseline =
    renderHeightPx * (1 + referenceBaselinePt / referenceHeightPt);
  if (
    renderWidthPx <= 0 ||
    renderHeightPx <= 0 ||
    !Number.isFinite(baseline) ||
    baseline < 0 ||
    baseline > renderHeightPx
  ) {
    return null;
  }
  return {
    width: renderWidthPx,
    height: renderHeightPx,
    baseline,
    renderWidthPx,
    renderHeightPx,
    referenceWidthPt,
    referenceHeightPt,
    referenceBaselinePt,
    inkCenterYRatio: metadata.imageInkCenterYRatio,
  };
}
