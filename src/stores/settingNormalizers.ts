export const MIN_SUGGESTION_COUNT = 3;
export const MAX_SUGGESTION_COUNT = 10;
export const DEFAULT_SUGGESTION_COUNT = 6;

export function normalizeSuggestionCount(value: unknown) {
  const count =
    typeof value === "number" && Number.isFinite(value)
      ? Math.round(value)
      : DEFAULT_SUGGESTION_COUNT;
  return Math.min(MAX_SUGGESTION_COUNT, Math.max(MIN_SUGGESTION_COUNT, count));
}
