import { register } from '../editor/commands';
import { skip } from './commands';
import { visualTexEditingCore } from '../editor-mathfield/visualtex-hooks';
import { _MathEnvironment } from '../core/math-environment';

register(
  {
    // VisualTeX editing core (selection.ts): one selection rule for keys and pointer
    selectGroup: (model) => visualTexEditingCore.selectGroup(model),

    selectAll: (model) => model.setSelection(0, model.lastOffset),
    extendSelectionForward: (model) => visualTexEditingCore.extend(model, 1),
    extendSelectionBackward: (model) => visualTexEditingCore.extend(model, -1),
    extendToNextWord: (model) => skip(model, 'forward', { extend: true }),
    extendToPreviousWord: (model) => skip(model, 'backward', { extend: true }),
    extendSelectionUpward: (model) => visualTexEditingCore.vertical(model, -1, _MathEnvironment.fractionNavigationOrder, true),
    extendSelectionDownward: (model) => visualTexEditingCore.vertical(model, 1, _MathEnvironment.fractionNavigationOrder, true),
    /**
     * Extend the selection until the next boundary is reached. A boundary
     * is defined by an atom of a different type (mbin, mord, etc...)
     * than the current focus. For example, in "1234+x=y", if the focus is between
     * "1" and "2", invoking `extendToNextBoundary_` would extend the selection
     * to "234".
     */
    extendToNextBoundary: (model) => skip(model, 'forward', { extend: true }),
    /**
     * Extend the selection until the previous boundary is reached. A boundary
     * is defined by an atom of a different type (mbin, mord, etc...)
     * than the current focus. For example, in "1+23456", if the focus is between
     * "5" and "6", invoking `extendToPreviousBoundary` would extend the selection
     * to "2345".
     */
    extendToPreviousBoundary: (model) =>
      skip(model, 'backward', { extend: true }),
    extendToGroupStart: (model) => visualTexEditingCore.extendTo(model, model.offsetOf(model.at(model.position).firstSibling)),
    extendToGroupEnd: (model) => visualTexEditingCore.extendTo(model, model.offsetOf(model.at(model.position).lastSibling)),
    extendToMathFieldStart: (model) => visualTexEditingCore.extendTo(model, 0),
    extendToMathFieldEnd: (model) => visualTexEditingCore.extendTo(model, model.lastOffset),
  },
  { target: 'model', changeSelection: true }
);
