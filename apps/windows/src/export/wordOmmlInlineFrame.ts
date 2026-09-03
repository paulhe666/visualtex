/**
 * Minimum inline math frame used by Word's bundled Latin Modern Math face.
 *
 * These are font-wide typographic metrics, not formula-specific offsets. The
 * exporter combines them with every expression's own rendered ascent and
 * descent, so ordinary symbols share one stable line box while fractions,
 * radicals, large operators and matrices remain free to grow naturally.
 */
export const WORD_OMML_INLINE_MINIMUM_ASCENT_EM = 0.806;
export const WORD_OMML_INLINE_MINIMUM_DESCENT_EM = 0.194;
