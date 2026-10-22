import { _Model } from './model-private';
import { CompositionAtom } from '../atoms/composition';
import { visualTexNormalizeTextBoundaryMode } from '../editor-mathfield/keyboard-input';

/**
 * Create, remove or update a composition atom at the current location
 */
export function updateComposition(model: _Model, s: string): void {
  // VisualTeX: IME composition must inherit the editing context, not merely
  // the mode of the atom immediately to the left of the caret. At the visual
  // right edge of a \text{} run MathLive's offset still points at the final
  // TextAtom, which otherwise re-enters text mode and makes all subsequent
  // input stick to that run.
  if (model.mathfield) visualTexNormalizeTextBoundaryMode(model.mathfield);
  const cursor = model.at(model.position);

  // We're creating or updating a composition
  if (cursor.type === 'composition') {
    // Composition already in progress, update it
    cursor.value = s;
  } else {
    // No composition yet, create one

    // Remove previous caret
    const { caret } = cursor;
    cursor.caret = undefined;

    // Create 'composition' atom, with caret
    const compositionMode = model.mode === 'latex' ? 'math' : model.mode;
    const atom = new CompositionAtom(s, { mode: compositionMode });
    atom.caret = caret;
    cursor.parent!.addChildAfter(atom, cursor);

    // Move cursor one past the composition zone
    model.position += 1;
  }
}

/**
 * Remove the composition zone
 */
export function removeComposition(model: _Model): void {
  const cursor = model.at(model.position);
  if (cursor.type === 'composition') {
    cursor.parent!.removeChild(cursor);
    model.position -= 1;
  }
}
