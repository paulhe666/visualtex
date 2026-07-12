import assert from "node:assert/strict";
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

console.log("Localized update notes smoke test passed");
