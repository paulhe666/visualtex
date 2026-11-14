import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  localizeReleaseNotes,
  stripPersonalAuthorNames,
} from "../../shared/visualtex-app/src/update/releaseNotes.ts";

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

const privateChineseName = String.fromCodePoint(24278, 29632, 20581);
const privateEnglishName = ["Liao", "Pojian"].join(" ");
const privateIdentityNotes =
  `## 中文\n- 作者：${privateChineseName}（paulhe666）\n` +
  `## English\n- Author: ${privateEnglishName} (paulhe666)`;
for (const language of ["cn", "en"]) {
  const localized = localizeReleaseNotes(privateIdentityNotes, language);
  const visible = [
    ...localized.features,
    ...localized.fixes,
    ...localized.other,
  ].join(" ");
  assert(!visible.includes(privateChineseName));
  assert(!visible.toLocaleLowerCase().includes(privateEnglishName.toLocaleLowerCase()));
}
assert.equal(
  stripPersonalAuthorNames(`Release by ${privateChineseName}`),
  "Release by VisualTeX",
);

const updateDialogSource = await readFile("../shared/visualtex-app/src/components/UpdateDialog.tsx", "utf8");
const updateDialogStyles = await readFile("../shared/visualtex-app/src/styles/app.css", "utf8");
const qqGroupSource = await readFile("../shared/visualtex-app/src/assets/visualtexQqGroup.ts", "utf8");
assert(updateDialogSource.includes("../assets/visualtexQqGroup"));
assert(updateDialogSource.includes("VISUALTEX_QQ_GROUP_NUMBER"));
assert(updateDialogSource.includes("VISUALTEX_QQ_GROUP_QR_DATA_URL"));
assert(updateDialogSource.includes("docs/images/wechat-pay.jpg"));
assert(updateDialogSource.includes("docs/images/alipay.jpg"));
assert(updateDialogSource.includes('className="update-community-card"'));
assert(updateDialogSource.includes('className="update-community-qr-row"'));
assert(updateDialogSource.includes("支持与交流"));
assert(updateDialogSource.includes("微信打赏"));
assert(updateDialogSource.includes("支付宝打赏"));
assert(updateDialogSource.includes("打赏完全自愿，不影响任何功能"));
assert(updateDialogSource.includes("loading=\"eager\""));
assert(updateDialogSource.includes("wechatPayImageUrl"));
assert(updateDialogSource.includes("alipayImageUrl"));
assert(updateDialogSource.includes("QQ群"));
assert(updateDialogStyles.includes(".update-community-qr-row"));
assert(updateDialogStyles.includes(".update-community-qr-card"));
assert(updateDialogStyles.includes("grid-template-columns: repeat(3, minmax(0, 1fr))"));
assert(!updateDialogSource.includes(privateChineseName));
assert(
  !updateDialogSource
    .toLocaleLowerCase()
    .includes(privateEnglishName.toLocaleLowerCase()),
);
assert(!updateDialogSource.includes("Join the VisualTeX QQ community"));
assert(qqGroupSource.includes('VISUALTEX_QQ_GROUP_NUMBER = "1045801770"'));
assert(qqGroupSource.includes("VISUALTEX_QQ_GROUP_QR_DATA_URL"));
assert(qqGroupSource.includes("data:image/png;base64"));

console.log("Localized update notes and QQ community card smoke test passed");
