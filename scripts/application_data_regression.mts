import assert from "node:assert/strict";
import {
  useEditorStore,
  type EditorLayout,
} from "../src/stores/editorStore.ts";
import type { InputBehaviorSettingKey, Theme } from "../src/types/formula.ts";

const store = useEditorStore;

function state() {
  return store.getState();
}

const expectedInputBehavior = {
  autoEscapeShortcuts: false,
  autoExitSuperscript: false,
  autoExitSubscript: true,
  autoExitAccent: false,
  autoExitWrapperCommand: false,
  showStructuredCommandSuggestions: true,
  showOtherCommandSuggestions: true,
} as const;

const configure = () => {
  state().setTitle("应用数据往返");
  state().setTheme("raycast" satisfies Theme);
  state().setLanguage("en");
  state().setEditorLayout("classic" satisfies EditorLayout);
  state().setZoom(1.25);
  state().setSourceOpen(true);
  state().setAutoPairDelimiters(false);
  state().setShowLineNumbers(true);
  state().setHighlightActiveLine(false);
  state().setFormulaInsetLeft(17);
  state().setFormulaInsetRight(29);
  state().setFormulaToolButtonSize(61);
  state().setFormulaToolButtonPadding(7);
  state().setFormulaRowVerticalInset(13);
  state().setPngExportBackground("#123456");
  state().setFormulaLetterFont("palatino");
  state().setFormulaChineseFont("songti");
  state().setPersonalize(false);
  state().setSuggestionCount(9);
  state().setCheckUpdatesOnStartup(false);
  state().setPowerPointDefaultFontSizePt(23.5);
  state().setClassicTileWidth(486);
  state().setClassicDockHeight(372);
  for (const [key, enabled] of Object.entries(expectedInputBehavior)) {
    state().setInputBehavior(key as InputBehaviorSettingKey, enabled);
  }
};

configure();
const saved = state().toDocument();
const serialized = JSON.parse(JSON.stringify(saved));

assert.equal(serialized.settings.theme, "raycast");
assert.equal(serialized.settings.editorLayout, "classic");
assert.equal(serialized.settings.language, "en");
assert.equal(serialized.settings.zoom, 1.25);
assert.equal(serialized.settings.sourceOpen, true);
assert.equal(serialized.settings.autoPairDelimiters, false);
assert.equal(serialized.settings.showLineNumbers, true);
assert.equal(serialized.settings.highlightActiveLine, false);
assert.equal(serialized.settings.formulaInsetLeft, 17);
assert.equal(serialized.settings.formulaInsetRight, 29);
assert.equal(serialized.settings.formulaToolButtonSize, 61);
assert.equal(serialized.settings.formulaToolButtonPadding, 7);
assert.equal(serialized.settings.formulaRowVerticalInset, 13);
assert.equal(serialized.settings.pngExportBackground, "#123456");
assert.equal(serialized.settings.formulaLetterFont, "palatino");
assert.equal(serialized.settings.formulaChineseFont, "songti");
assert.deepEqual(serialized.settings.inputBehavior, expectedInputBehavior);
assert.equal(serialized.settings.personalize, false);
assert.equal(serialized.settings.suggestionCount, 9);
assert.equal(serialized.settings.checkUpdatesOnStartup, false);
assert.equal(serialized.settings.powerPointDefaultFontSizePt, 23.5);
assert.equal(serialized.settings.classicTileWidth, 486);
assert.equal(serialized.settings.classicDockHeight, 372);

state().setTheme("light");
state().setLanguage("cn");
state().setEditorLayout("standard");
state().setZoom(0.6);
state().setSourceOpen(false);
state().setAutoPairDelimiters(true);
state().setShowLineNumbers(false);
state().setHighlightActiveLine(true);
state().setFormulaInsetLeft(34);
state().setFormulaInsetRight(34);
state().setFormulaToolButtonSize(52);
state().setFormulaToolButtonPadding(2);
state().setFormulaRowVerticalInset(5);
state().setPngExportBackground("transparent");
state().setFormulaLetterFont("katex");
state().setFormulaChineseFont("system");
state().setPersonalize(true);
state().setSuggestionCount(6);
state().setCheckUpdatesOnStartup(true);
state().setPowerPointDefaultFontSizePt(20);
state().setClassicTileWidth(300);
state().setClassicDockHeight(240);

state().loadDocument(serialized);
const restored = state();

assert.equal(restored.title, "应用数据往返");
assert.equal(restored.theme, "raycast");
assert.equal(restored.editorLayout, "classic");
assert.equal(restored.language, "en");
assert.equal(restored.zoom, 1.25);
assert.equal(restored.sourceOpen, true);
assert.equal(restored.autoPairDelimiters, false);
assert.equal(restored.showLineNumbers, true);
assert.equal(restored.highlightActiveLine, false);
assert.equal(restored.formulaInsetLeft, 17);
assert.equal(restored.formulaInsetRight, 29);
assert.equal(restored.formulaToolButtonSize, 61);
assert.equal(restored.formulaToolButtonPadding, 7);
assert.equal(restored.formulaRowVerticalInset, 13);
assert.equal(restored.pngExportBackground, "#123456");
assert.equal(restored.formulaLetterFont, "palatino");
assert.equal(restored.formulaChineseFont, "songti");
assert.deepEqual(restored.inputBehavior, expectedInputBehavior);
assert.equal(restored.personalize, false);
assert.equal(restored.suggestionCount, 9);
assert.equal(restored.checkUpdatesOnStartup, false);
assert.equal(restored.powerPointDefaultFontSizePt, 23.5);
assert.equal(restored.classicTileWidth, 486);
assert.equal(restored.classicDockHeight, 372);

console.log("VisualTeX application data regression: PASS");
