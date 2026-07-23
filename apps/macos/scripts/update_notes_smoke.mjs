import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { localizeReleaseNotes } from "../src/update/releaseNotes.ts";

const bilingualNotes = `
## 中文

### 新增功能
- 新增自动更新提醒。
- 支持按当前语言显示更新说明。

### 问题修复
- 修复度数符号删除后光标卡住的问题。

## English

### New features
- Added automatic update notifications.
- Added localized release notes.

### Bug fixes
- Fixed the degree-symbol caret getting stuck after deletion.
`;

assert.deepEqual(localizeReleaseNotes(bilingualNotes, "cn"), {
  features: ["新增自动更新提醒。", "支持按当前语言显示更新说明。"],
  fixes: ["修复度数符号删除后光标卡住的问题。"],
  other: [],
});

assert.deepEqual(localizeReleaseNotes(bilingualNotes, "en"), {
  features: [
    "Added automatic update notifications.",
    "Added localized release notes.",
  ],
  fixes: ["Fixed the degree-symbol caret getting stuck after deletion."],
  other: [],
});

const legacyNotes = `VisualTeX improves editing stability.\n\n- Existing release note.`;
assert.deepEqual(localizeReleaseNotes(legacyNotes, "cn"), {
  features: [],
  fixes: [],
  other: ["VisualTeX improves editing stability.", "Existing release note."],
});

const updateDialogSource = await readFile("src/components/UpdateDialog.tsx", "utf8");
const qqGroupCard = await readFile("public/qq-group-card.svg", "utf8");
assert(updateDialogSource.includes('const QQ_GROUP_NUMBER = "1045801770"'));
assert(updateDialogSource.includes('const QQ_GROUP_IMAGE_URL = "/qq-group-card.svg"'));
assert(updateDialogSource.includes('className="update-community-card"'));
assert(updateDialogSource.includes("加入 VisualTeX QQ 交流群"));
assert(!updateDialogSource.includes("Join the VisualTeX QQ community"));
assert(qqGroupCard.includes("https://qm.qq.com/q/TppXdoOO8Q") === false);
assert(qqGroupCard.includes("1045801770"));
assert(qqGroupCard.includes("VisualTeX 交流群"));

console.log("Localized update notes and QQ community card smoke test passed");
