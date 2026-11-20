// Document mode's tutorial cards of the app this runs in (each app runs it: "@platform/tutorials"): every card with an
// animation has one in both languages (the shared tutorial/<zh|en>/<id>.webp, an animated WebP), and no animation is
// left over without a card, so both apps' animated cards are the same ones.
import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { TUTORIAL_CARDS } from "@platform/tutorials";

const app = basename(process.cwd());
const root = join(dirname(fileURLToPath(import.meta.url)), "../src/document/tutorial");

assert.ok(TUTORIAL_CARDS.length >= 3 && TUTORIAL_CARDS.length <= 7, `${app}: a few cards`);
assert.equal(new Set(TUTORIAL_CARDS.map((card) => card.id)).size, TUTORIAL_CARDS.length, `${app}: one card per id`);

for (const card of TUTORIAL_CARDS) {
  for (const text of [...card.title, ...card.text]) {
    assert.ok(text.trim(), `${app} ${card.id}: text in both languages`);
    assert.equal(text.split("`").length % 2, 1, `${app} ${card.id}: backticks in pairs: ${text}`);
  }
  assert.ok(card.animation || card.path, `${app} ${card.id}: an animation or a file path to show`);
}

const animated = TUTORIAL_CARDS.filter((card) => card.animation);
for (const lang of ["zh", "en"]) {
  for (const card of animated) {
    const file = join(root, lang, `${card.id}.webp`);
    assert.ok(existsSync(file), `${app}: missing ${lang}/${card.id}.webp`);
    const bytes = readFileSync(file);
    assert.equal(bytes.toString("latin1", 0, 4), "RIFF", `${lang}/${card.id}.webp is a RIFF file`);
    assert.equal(bytes.toString("latin1", 8, 12), "WEBP", `${lang}/${card.id}.webp is a WebP`);
    assert.ok(bytes.includes(Buffer.from("ANIM")), `${lang}/${card.id}.webp is animated`);
    assert.ok(bytes.length < 600 * 1024, `${lang}/${card.id}.webp stays small (${bytes.length} bytes)`);
  }
  const ids = new Set(animated.map((card) => `${card.id}.webp`));
  for (const name of readdirSync(join(root, lang))) assert.ok(ids.has(name), `${app}: ${lang}/${name} belongs to no card`);
}

console.log(`Document tutorial cards passed (${app}): ${TUTORIAL_CARDS.length} cards, ${animated.length} animations in zh and en`);
