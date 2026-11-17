// The congruence sign of Chinese geometry (≌, U+224C, \backcong) is the character the Chinese fonts carry, set in the formula's
// Chinese font like Chinese text, and it is a relation: thick space on both sides, as for "=". The math fonts (STIX, Latin
// Modern, Cambria, Apple Symbols) draw U+224C as a tilde with a gentle wave; only the Chinese fonts draw the S with curled ends
// that the textbooks print, so the layout must not set it from a TeX or math font, and the toolbar's relation group offers it.
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

const ROWS = ["a\\backcong b", "a=b", "a≌b", "a\\cong b"];
const SIGN = "≌";

await withAppBrowser(async ({ baseUrl, client, evaluate }) => {
  await client.send("Page.navigate", { url: baseUrl });
  await sleep(1200);
  await evaluate(`(() => {
    for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
    localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
    const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
    persisted.state = { ...(persisted.state || {}), theme: "light", language: "zh", lines: ${JSON.stringify(ROWS)}.map((latex, i) => ({ id: "line-" + (i + 1), latex })), activeLineId: "line-1" };
    localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "formula" }, version: 0 }));
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(3500);

  // every drawn character of each row: its text, left edge (em) and font
  const rows = await evaluate(`(() => {
    const fields = [...document.querySelectorAll("math-field.visual-mathfield")];
    const canvas = document.createElement("canvas").getContext("2d");
    return fields.map((field) => {
      const root = field.shadowRoot.querySelector(".ML__vtx");
      const glyphs = [...root.children].filter((s) => /font:/.test(s.getAttribute("style") || "") && s.textContent);
      return glyphs.map((s) => {
        const style = s.getAttribute("style");
        const font = /font:([^;]*)/.exec(style)[1].replace(/^[0-9.]+em\\s+/, "");
        canvas.font = "200px " + font;
        return { text: s.textContent, left: parseFloat(/left:(-?[0-9.]+)em/.exec(style)[1]), font, width: canvas.measureText(s.textContent).width / 200 };
      });
    });
  })()`);
  assert(rows.length === ROWS.length, `${rows.length} rows rendered, expected ${ROWS.length}`);

  const signOf = (row) => row.find((g) => g.text === SIGN);
  const [command, equals, typed, cong] = rows;
  for (const [name, row] of [["\\backcong", command], ["the typed ≌", typed]]) {
    const sign = signOf(row);
    assert(sign, `${name}: no ≌ character drawn (${JSON.stringify(row.map((g) => g.text))})`);
    assert(!/VTX|STIX|Math/i.test(sign.font), `${name} is drawn from a TeX or math font (${sign.font})`);
    assert(/YaHei|PingFang|Hiragino|SimSun|Songti|SimHei|Heiti|KaiTi|Kaiti|Noto Sans CJK|JhengHei/.test(sign.font), `${name} is not set in the formula's Chinese font (${sign.font})`);
  }

  // a relation: the same thick space before it as before "=", and the same after it
  const equalsGlyph = equals.find((g) => g.text === "\uE03D"); // cmr10 "=": U+E000 + 61 (the TeX fonts' characters are in the private use area)
  assert(equalsGlyph, `no = in the reference row (${JSON.stringify(equals.map((g) => g.text.codePointAt(0).toString(16)))})`);
  const equalsLast = equals[equals.length - 1];
  const equalsGap = equalsLast.left - (equalsGlyph.left + 0.77778);
  for (const [name, row] of [["\\backcong", command], ["the typed ≌", typed]]) {
    const sign = signOf(row);
    const last = row[row.length - 1];
    assert(Math.abs(sign.left - equalsGlyph.left) < 0.003, `${name}: space before is ${(sign.left).toFixed(4)}em, = has ${(equalsGlyph.left).toFixed(4)}em`);
    const gap = last.left - (sign.left + sign.width);
    assert(Math.abs(gap - equalsGap) < 0.006, `${name}: space after is ${gap.toFixed(4)}em, = has ${equalsGap.toFixed(4)}em`);
  }
  // the same sign through both spellings
  assert(Math.abs(signOf(command).left - signOf(typed).left) < 0.0005 && Math.abs(signOf(command).width - signOf(typed).width) < 0.0005, "the command and the typed character lay out differently");
  // ≅ stays the TeX font's (AMS \cong): not a character in the Chinese font
  assert(!signOf(cong), "\\cong is drawn as the ≌ character");

  // the toolbar's relation group: the card is there
  const click = (selector) =>
    evaluate(`(() => {
      const el = [...document.querySelectorAll(${JSON.stringify(selector)})].find((e) => e.getBoundingClientRect().width > 0);
      if (!el) return false;
      el.click();
      return true;
    })()`);
  assert(await click("[data-toolbar-edit-toggle]"), "the toolbar's edit button is not on the page");
  await sleep(500);
  assert(await click("[data-toolbar-add-item]"), "no + card in the toolbar's edit mode");
  await sleep(600);
  assert(await click('[data-toolbar-picker-source="relation"]'), "the picker has no relation category");
  await sleep(400);
  const cards = await evaluate(`[...document.querySelectorAll("[data-toolbar-item-picker] [data-toolbar-picker-item]")].map((c) => c.getAttribute("title"))`);
  assert(cards.includes("全等（≌）"), `the relation category has no 全等（≌） card (${cards.slice(0, 12).join(" | ")} ...)`);
  assert(cards.indexOf("全等（≌）") > cards.indexOf("全等"), "the ≌ card does not follow the ≅ card (全等)");

  console.log("congruence sign ≌: a relation set in the Chinese font", "sign font:", signOf(command).font.slice(0, 40), "width:", signOf(command).width.toFixed(3) + "em");
  console.log("Chinese-font relation (≌ drawn from the Chinese font as a relation, offered in the relation group): passed");
}, { windowSize: "1280,900" });
