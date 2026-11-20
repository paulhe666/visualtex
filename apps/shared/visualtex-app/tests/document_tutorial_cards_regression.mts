// Document mode's tutorial cards: both platforms show the same cards in the same order, every card with an animation
// has one in both languages (the shared tutorial/<zh|en>/<id>.webp, an animated WebP), and no animation is left
// over without a card.
import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { TUTORIAL_CARDS as WINDOWS } from "../../../windows/src/platform/tutorials.ts";
import { TUTORIAL_CARDS as MACOS } from "../../../macos/src/platform/tutorials.ts";

const root = join(dirname(fileURLToPath(import.meta.url)), "../src/document/tutorial");

assert.deepEqual(
  WINDOWS.map((card) => [card.id, card.animation]),
  MACOS.map((card) => [card.id, card.animation]),
  "Windows and macOS list the same cards",
);
assert.ok(WINDOWS.length >= 3 && WINDOWS.length <= 7, "a few cards");

for (const cards of [WINDOWS, MACOS]) {
  for (const card of cards) {
    for (const text of [...card.title, ...card.text]) {
      assert.ok(text.trim(), `${card.id}: text in both languages`);
      assert.equal(text.split("`").length % 2, 1, `${card.id}: backticks in pairs: ${text}`);
    }
    assert.ok(card.animation || card.path, `${card.id}: an animation or a file path to show`);
  }
}

for (const lang of ["zh", "en"]) {
  for (const card of WINDOWS.filter((item) => item.animation)) {
    const file = join(root, lang, `${card.id}.webp`);
    assert.ok(existsSync(file), `missing ${lang}/${card.id}.webp`);
    const bytes = readFileSync(file);
    assert.equal(bytes.toString("latin1", 0, 4), "RIFF", `${lang}/${card.id}.webp is a RIFF file`);
    assert.equal(bytes.toString("latin1", 8, 12), "WEBP", `${lang}/${card.id}.webp is a WebP`);
    assert.ok(bytes.includes(Buffer.from("ANIM")), `${lang}/${card.id}.webp is animated`);
    assert.ok(bytes.length < 600 * 1024, `${lang}/${card.id}.webp stays small (${bytes.length} bytes)`);
  }
  const ids = new Set(WINDOWS.filter((item) => item.animation).map((item) => `${item.id}.webp`));
  for (const name of readdirSync(join(root, lang))) assert.ok(ids.has(name), `${lang}/${name} belongs to no card`);
}

console.log(`Document tutorial cards passed: ${WINDOWS.length} cards, ${WINDOWS.filter((card) => card.animation).length} animations in zh and en`);
