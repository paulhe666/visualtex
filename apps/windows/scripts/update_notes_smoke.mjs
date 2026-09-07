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
const updateDialogStyles = await readFile("src/styles.css", "utf8");
assert(updateDialogSource.includes("docs/images/wechat-pay.jpg"));
assert(updateDialogSource.includes("docs/images/alipay.jpg"));
assert(updateDialogSource.includes('className="update-community-qr-row"'));
assert(updateDialogSource.includes("支持与交流"));
assert(updateDialogSource.includes("微信打赏"));
assert(updateDialogSource.includes("支付宝打赏"));
assert(updateDialogSource.includes("自愿打赏通道"));
assert(updateDialogSource.includes("是否打赏完全不影响 VisualTeX 的任何功能和正常使用"));
assert(updateDialogSource.includes("QQ群"));
assert(updateDialogStyles.includes(".update-community-qr-row"));
assert(updateDialogStyles.includes("grid-template-columns: repeat(3, minmax(0, 1fr))"));

console.log("Localized update notes and support QR smoke test passed");
