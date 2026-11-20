// The dialog the toolbar's + card opens (Add to a group of the formula toolbar): a dense field of cards that show the character
// and nothing else (the name is the tooltip), and nothing of the formula tiles' library (its own dialog's): no card labels, no
// section headings, no "annotations" in the navigation, no library group among the sources, none among the search results.
import { withAppBrowser } from "../testing/browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

await withAppBrowser(async ({ baseUrl, client, evaluate }) => {
  await client.send("Page.navigate", { url: baseUrl });
  await sleep(1200);
  await evaluate(`(() => {
    for (const k of ["visualtex.onboarding.v3.completed", "visualtex.office.macos.first-run.v1.completed", "visualtex.onboarding.macos.desktop.v1.2.0.completed", "visualtex.office.macos.native-first-run.v1.2.0.completed"]) localStorage.setItem(k, "true");
    localStorage.setItem("visualtex.document.tutorial-shown.v1", "1");
    const persisted = JSON.parse(localStorage.getItem("visualtex-editor") || "{}");
    persisted.state = { ...(persisted.state || {}), theme: "light", language: "zh" };
    localStorage.setItem("visualtex-editor", JSON.stringify(persisted));
    localStorage.setItem("visualtex.workspace-content.v1", JSON.stringify({ state: { content: "formula" }, version: 0 }));
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(2800);

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

  const dialog = () =>
    evaluate(`(() => {
      const root = document.querySelector("[data-toolbar-item-picker]");
      if (!root) return null;
      const rect = (el) => { const r = el.getBoundingClientRect(); return { w: r.width, h: r.height, left: r.left, top: r.top }; };
      const content = root.querySelector(".formula-tile-library-content");
      const cards = [...root.querySelectorAll("[data-toolbar-picker-item]")];
      const box = content.getBoundingClientRect();
      const visible = cards.filter((c) => { const r = c.getBoundingClientRect(); return r.bottom <= box.bottom + 1 && r.top >= box.top - 1; });
      return {
        title: root.querySelector("h2")?.textContent ?? "",
        cards: cards.length,
        visible: visible.length,
        sizes: cards.slice(0, 40).map((c) => rect(c)),
        labels: root.querySelectorAll(".formula-tile-library-item-label").length,
        headings: root.querySelectorAll("h3, .toolbar-picker-nav-heading, .eyebrow").length,
        navTexts: [...root.querySelectorAll(".formula-tile-library-groups button")].map((b) => b.textContent.trim()),
        sources: [...root.querySelectorAll("[data-toolbar-picker-source]")].map((b) => b.getAttribute("data-toolbar-picker-source")),
        library: root.querySelectorAll('[data-toolbar-picker-item^="library:"]').length,
        tooltips: cards.slice(0, 12).map((c) => c.getAttribute("title")),
        search: root.querySelector("input[type=search]")?.getAttribute("placeholder") ?? "",
        dialogWidth: root.getBoundingClientRect().width,
      };
    })()`);

  const first = await dialog();
  assert(first, "the picker did not open");
  assert(first.cards >= 20, `the first category holds ${first.cards} cards`);
  assert(first.labels === 0, `the cards carry ${first.labels} text labels`);
  assert(first.headings === 0, `the dialog has ${first.headings} headings or eyebrows`);
  assert(first.library === 0, `${first.library} cards of the formula library`);
  const widest = Math.max(...first.sizes.filter((s) => s.w < 200).map((s) => s.w));
  const tall = Math.max(...first.sizes.map((s) => s.h));
  assert(widest <= 80, `a character card is ${widest.toFixed(0)}px wide (dense: at most 80)`);
  assert(tall <= 56, `a card is ${tall.toFixed(0)}px tall (dense: at most 56)`);
  assert(first.tooltips.every((t) => t && t.length > 0), "every card keeps its name as the tooltip");
  // the navigation lists the toolbar's categories (and custom symbols), not the formula library's groups
  // (the library's groups are math-algebra, physics-mechanics, chemistry-basics ...; the toolbar's categories have one word)
  assert(!first.sources.some((id) => id.includes("-")), `a library group is among the sources: ${first.sources}`);
  assert(first.navTexts.every((t) => !/中学|大学|高中|初中|School|University/.test(t)), `a level annotation in the navigation: ${first.navTexts}`);
  console.log("cards:", first.cards, "visible:", first.visible, "card size:", first.sizes[0].w.toFixed(0), "x", first.sizes[0].h.toFixed(0), "sources:", first.sources.join(","), "search:", first.search);

  // every category is a field of cards too, with the same density
  for (const id of first.sources) {
    await evaluate(`document.querySelector('[data-toolbar-picker-source="${id}"]').click()`);
    await sleep(250);
    const now = await dialog();
    assert(now.labels === 0 && now.headings === 0 && now.library === 0, `category ${id}: labels ${now.labels}, headings ${now.headings}, library ${now.library}`);
    if (now.cards > 0) assert(Math.max(...now.sizes.filter((s) => s.w < 200).map((s) => s.w)) <= 80, `category ${id}: cards wider than 80px`);
  }

  // the search: commands and custom symbols together in one field; none of the library's
  await evaluate(`(() => {
    const input = document.querySelector("[data-toolbar-item-picker] input[type=search]");
    const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set;
    set.call(input, "frac");
    input.dispatchEvent(new Event("input", { bubbles: true }));
  })()`);
  await sleep(400);
  const found = await dialog();
  assert(found.cards > 0, "the search for frac found nothing");
  assert(found.library === 0, "the search offers cards of the formula library");
  assert(found.headings === 0 && found.labels === 0, "the search results have headings or labels");

  console.log("toolbar item picker (a dense field of character cards, no annotations, no formula library): passed");
}, { windowSize: "1280,900" });
