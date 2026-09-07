import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  localizeReleaseNotes,
  stripPersonalAuthorNames,
} from "../src/update/releaseNotes.ts";

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

const appSource = await readFile("src/App.tsx", "utf8");
const updateDialogSource = await readFile("src/components/UpdateDialog.tsx", "utf8");
const releaseWelcomeSource = await readFile("src/update/releaseWelcome.ts", "utf8");
const updateDialogStyles = await readFile("src/styles.css", "utf8");
const qqGroupCard = await readFile("public/qq-group-card.svg", "utf8");
assert(updateDialogSource.includes('const QQ_GROUP_NUMBER = "1045801770"'));
assert(updateDialogSource.includes('const QQ_GROUP_IMAGE_URL = "/qq-group-card.svg"'));
assert(updateDialogSource.includes("docs/images/wechat-pay.jpg"));
assert(updateDialogSource.includes("docs/images/alipay.jpg"));
assert(updateDialogSource.includes('className="update-community-card"'));
assert(updateDialogSource.includes('className="update-community-qr-row"'));
assert(updateDialogSource.includes("支持与交流"));
assert(updateDialogSource.includes("微信打赏"));
assert(updateDialogSource.includes("支付宝打赏"));
assert(updateDialogSource.includes("有经济能力并且觉得产品不错的可以支持一下作者呀！"));
assert(updateDialogSource.includes("打赏完全自愿，不影响 VisualTeX 的任何功能和正常使用"));
assert(updateDialogSource.includes("QQ群二维码用于加入交流群"));
assert(updateDialogSource.includes("loading=\"eager\""));
assert(updateDialogSource.includes("wechatPayImageUrl"));
assert(updateDialogSource.includes("alipayImageUrl"));
assert(updateDialogSource.includes("QQ群"));
assert(updateDialogSource.includes("updateAvailable || releaseWelcome"));
assert(updateDialogSource.includes("releaseWelcome ?"));
assert(releaseWelcomeSource.includes('RELEASE_WELCOME_VERSION = "1.2.6"'));
assert(releaseWelcomeSource.includes("visualtex.release-welcome.${RELEASE_WELCOME_VERSION}.seen"));
assert(releaseWelcomeSource.includes("VisualTeX ${RELEASE_WELCOME_VERSION}"));
assert(updateDialogSource.includes("更新内容"));
assert(appSource.includes("releaseWelcomePending"));
assert(appSource.includes("releaseWelcomeOpen"));
assert(appSource.includes("RELEASE_WELCOME_STORAGE_KEY"));
assert(appSource.includes("writeLocalStorage(RELEASE_WELCOME_STORAGE_KEY, \"true\")"));
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
assert(qqGroupCard.includes("https://qm.qq.com/q/TppXdoOO8Q") === false);
assert(qqGroupCard.includes("1045801770"));
assert(qqGroupCard.includes("VisualTeX 交流群"));

console.log("Localized update notes and QQ community card smoke test passed");
