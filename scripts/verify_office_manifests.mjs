import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { DOMParser } from "@xmldom/xmldom";

const root = resolve(import.meta.dirname, "..");
const origin = "https://127.0.0.1:43127";
const officeAppNamespace = "http://schemas.microsoft.com/office/appforoffice/1.1";
const basicTypesNamespace = "http://schemas.microsoft.com/office/officeappbasictypes/1.0";
const versionOverridesNamespace = "http://schemas.microsoft.com/office/taskpaneappversionoverrides";
const packageJson = JSON.parse(await readFile(resolve(root, "package.json"), "utf8"));
const manifests = [
  {
    platform: "macos",
    host: "word",
    path: "office/macos/manifests/VisualTeX.Word.xml",
    id: "d6fcb260-4c37-4f73-a173-cf24674f81f2",
    baseHost: "Document",
    overrideHost: "Document",
    apiSet: "WordApi",
    version: `${packageJson.version}.0`,
    commands: [
      "VisualTeX.NewFormula",
      "VisualTeX.EditSelectedFormula",
      "VisualTeX.OpenDesktopApp",
    ],
  },
  {
    platform: "macos",
    host: "powerpoint",
    path: "office/macos/manifests/VisualTeX.PowerPoint.xml",
    id: "a6d13cf2-54e8-4dfa-a20c-15de864ab3c5",
    baseHost: "Presentation",
    overrideHost: "Presentation",
    apiSet: "PowerPointApi",
    version: `${packageJson.version}.0`,
    commands: [
      "VisualTeX.NewFormula",
      "VisualTeX.EditSelectedFormula",
      "VisualTeX.OpenDesktopApp",
    ],
  },
  {
    platform: "windows-ole",
    host: "word",
    path: "office/windows/ole/manifests/VisualTeX.Word.xml",
    id: "7c7d3b35-56b2-4c40-88d9-c9eb836d6021",
    baseHost: "Document",
    overrideHost: "Document",
    apiSet: "AddinCommands",
    version: `${packageJson.version}.3`,
    metadata: {
      displayName: ["VisualTeX Windows OLE for Word", "VisualTeX Windows OLE（Word）"],
      description: [
        "VisualTeX Windows Word OLE formula integration.",
        "VisualTeX Windows Word OLE 公式集成。",
      ],
    },
    groups: [
      {
        id: "VisualTeX.Ole.Formulas",
        commands: ["newFormula", "editFormula", "openDesktop"],
      },
      { id: "VisualTeX.Ole.Numbering", commands: ["updateEquationNumbers"] },
    ],
    localizedStrings: {
      "Tab.Label": ["VisualTeX", "VisualTeX"],
      "FormulaGroup.Label": ["Formulas", "公式"],
      "NumberingGroup.Label": ["Equation Numbers", "公式编号"],
      "New.Label": ["New Formula", "新建公式"],
      "Edit.Label": ["Edit Selected Formula", "编辑所选公式"],
      "Numbering.Label": ["Update Equation Numbers", "更新公式编号"],
      "Open.Label": ["Open VisualTeX", "打开 VisualTeX"],
      "New.Desc": [
        "Insert a VisualTeX formula through the Windows OLE Bridge.",
        "通过 Windows OLE Bridge 插入 VisualTeX 公式。",
      ],
      "Edit.Desc": ["Edit the selected VisualTeX formula.", "编辑当前选中的 VisualTeX 公式。"],
      "Numbering.Desc": [
        "Renumber all numbered VisualTeX display formulas in document order.",
        "按文档顺序重新编号所有已启用编号的 VisualTeX 行间公式。",
      ],
      "Open.Desc": ["Open the VisualTeX desktop application.", "打开 VisualTeX 桌面应用。"],
    },
    commands: [
      "newFormula",
      "editFormula",
      "updateEquationNumbers",
      "openDesktop",
    ],
  },
  {
    platform: "windows-ole",
    host: "powerpoint",
    path: "office/windows/ole/manifests/VisualTeX.PowerPoint.xml",
    id: "fdc8d615-7e60-4586-bff4-5a1d728f9f6c",
    baseHost: "Presentation",
    overrideHost: "Presentation",
    apiSet: "AddinCommands",
    version: `${packageJson.version}.3`,
    metadata: {
      displayName: [
        "VisualTeX Windows OLE for PowerPoint",
        "VisualTeX Windows OLE（PowerPoint）",
      ],
      description: [
        "VisualTeX Windows PowerPoint OLE formula integration.",
        "VisualTeX Windows PowerPoint OLE 公式集成。",
      ],
    },
    groups: [
      {
        id: "VisualTeX.Ole.Formulas",
        commands: ["newFormula", "editFormula", "openDesktop"],
      },
    ],
    localizedStrings: {
      "Tab.Label": ["VisualTeX", "VisualTeX"],
      "FormulaGroup.Label": ["Formulas", "公式"],
      "New.Label": ["New Formula", "新建公式"],
      "Edit.Label": ["Edit Selected Formula", "编辑所选公式"],
      "Open.Label": ["Open VisualTeX", "打开 VisualTeX"],
      "New.Desc": [
        "Insert a VisualTeX formula on the current slide through the Windows OLE Bridge.",
        "通过 Windows OLE Bridge 在当前幻灯片中插入 VisualTeX 公式。",
      ],
      "Edit.Desc": ["Edit the selected VisualTeX formula.", "编辑当前选中的 VisualTeX 公式。"],
      "Open.Desc": ["Open the VisualTeX desktop application.", "打开 VisualTeX 桌面应用。"],
    },
    commands: [
      "newFormula",
      "editFormula",
      "openDesktop",
    ],
  },
];

function extract(xml, tag) {
  return xml.match(new RegExp(`<${tag}>([^<]+)</${tag}>`))?.[1] ?? "";
}

function assertLocalizedElement(element, namespace, expected, context) {
  assert.ok(element, `${context} is missing`);
  assert.equal(element.getAttribute("DefaultValue"), expected[0], `${context} English value`);
  const overrides = Array.from(element.getElementsByTagNameNS(namespace, "Override"));
  assert.equal(overrides.length, 1, `${context} must have exactly one locale override`);
  assert.equal(overrides[0].getAttribute("Locale"), "zh-CN", `${context} override locale`);
  assert.equal(overrides[0].getAttribute("Value"), expected[1], `${context} Chinese value`);
}

for (const manifest of manifests) {
  const absolute = resolve(root, manifest.path);
  const xml = await readFile(absolute, "utf8");
  const document = new DOMParser().parseFromString(xml, "application/xml");
  const parserErrors = document.getElementsByTagName("parsererror");
  assert.equal(parserErrors.length, 0, `${manifest.path} is not valid XML`);
  assert.equal(extract(xml, "Id"), manifest.id);
  assert.equal(extract(xml, "Version"), manifest.version);
  assert.match(xml, new RegExp(`<Host Name="${manifest.baseHost}"\\s*/>`));
  assert.match(xml, new RegExp(`<Host xsi:type="${manifest.overrideHost}">`));
  assert.match(xml, /<Permissions>ReadWriteDocument<\/Permissions>/);
  if (manifest.apiSet) {
    assert.match(xml, new RegExp(`<bt:Set Name="${manifest.apiSet}" MinVersion="1\\.1"\\s*/>`));
  } else {
    assert.ok(xml.includes('xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"'));
  }
  if (manifest.platform === "macos" && manifest.host === "powerpoint") {
    assert.match(xml, /<bt:Set Name="ImageCoercion" MinVersion="1\.1"\s*\/>/);
  }
  if (manifest.platform === "windows-ole") {
    assert.equal(extract(xml, "DefaultLocale"), "en-US");
    assert.match(xml, /<CustomTab id="VisualTeX\.WindowsOle\.Tab">/);
    assert.match(xml, /<Label resid="Tab\.Label"\s*\/><\/CustomTab>/);
    assert.doesNotMatch(xml, /<OfficeTab id="TabHome">/);
    assert.equal(
      [...xml.matchAll(/<CustomTab\b/g)].length,
      1,
      `${manifest.path} must define exactly one VisualTeX custom tab`,
    );
    assertLocalizedElement(
      document.getElementsByTagNameNS(officeAppNamespace, "DisplayName")[0],
      officeAppNamespace,
      manifest.metadata.displayName,
      `${manifest.path} DisplayName`,
    );
    assertLocalizedElement(
      document.getElementsByTagNameNS(officeAppNamespace, "Description")[0],
      officeAppNamespace,
      manifest.metadata.description,
      `${manifest.path} Description`,
    );

    const resourceStrings = Array.from(
      document.getElementsByTagNameNS(basicTypesNamespace, "String"),
    );
    for (const [id, expected] of Object.entries(manifest.localizedStrings)) {
      const resource = resourceStrings.find((item) => item.getAttribute("id") === id);
      assertLocalizedElement(
        resource,
        basicTypesNamespace,
        expected,
        `${manifest.path} resource ${id}`,
      );
    }
    const visibleResourceIds = new Set(
      [...xml.matchAll(/<(?:Label|Title|Description)\s+resid="([^"]+)"/g)].map(
        (match) => match[1],
      ),
    );
    assert.deepEqual(
      [...visibleResourceIds].sort(),
      Object.keys(manifest.localizedStrings).sort(),
      `${manifest.path} must localize every visible label and description`,
    );

    const groups = Array.from(
      document.getElementsByTagNameNS(versionOverridesNamespace, "Group"),
    );
    assert.equal(groups.length, manifest.groups.length, `${manifest.path} ribbon group count`);
    for (const [index, expectedGroup] of manifest.groups.entries()) {
      const group = groups[index];
      assert.equal(group.getAttribute("id"), expectedGroup.id, `${manifest.path} group ${index}`);
      const groupCommands = Array.from(
        group.getElementsByTagNameNS(versionOverridesNamespace, "FunctionName"),
      ).map((item) => item.textContent);
      assert.deepEqual(
        groupCommands,
        expectedGroup.commands,
        `${manifest.path} group ${expectedGroup.id} command layout`,
      );
    }
  }
  if (manifest.platform === "windows-ole" && manifest.host === "word") {
    assert.match(xml, /<Control xsi:type="Button" id="VisualTeX\.Ole\.UpdateNumbers">/);
    assert.match(xml, /<Label resid="Numbering\.Label"\s*\/>/);
    assert.match(xml, /<Description resid="Numbering\.Desc"\s*\/>/);
  }
  for (const command of manifest.commands) {
    assert.ok(xml.includes(`<FunctionName>${command}</FunctionName>`));
    assert.ok(command.length <= 32, `Office command function name is too long: ${command}`);
  }
  for (const icon of ["icon-16.png", "icon-32.png", "icon-80.png"]) {
    assert.ok(xml.includes(`${origin}/icons/${icon}`));
  }
  for (const match of xml.matchAll(/\b(?:id|resid)="([^"]+)"/g)) {
    assert.ok(match[1].length <= 32, `Office manifest id is too long: ${match[1]}`);
  }
  assert.ok(!xml.includes("{{"));
  assert.ok(!xml.includes("localhost"));
}

const ids = manifests.map((manifest) => manifest.id);
assert.equal(new Set(ids).size, ids.length, "Every platform/host manifest must have a distinct GUID");
const macSource = await readFile(resolve(root, "src-tauri/src/office/manifest.rs"), "utf8");
assert.ok(macSource.includes("office/macos/manifests"));
assert.ok(!macSource.includes("office/windows/ole/manifests"));
const windowsEntry = await readFile(resolve(root, "src/office/windows-ole/main.ts"), "utf8");
assert.ok(!windowsEntry.includes("office/macos"));
console.log("Independent macOS and Windows OLE manifest verification passed");
