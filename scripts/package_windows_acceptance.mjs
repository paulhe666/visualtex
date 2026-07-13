import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import {
  copyFileSync,
  cpSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = dirname(fileURLToPath(import.meta.url));
const projectRoot = resolve(scriptDirectory, "..");
const artifactRoot = join(projectRoot, "artifacts", "windows-acceptance");
const dateStamp = new Date().toISOString().slice(0, 10).replaceAll("-", "");
const baseCommit = execFileSync("git", ["rev-parse", "HEAD"], {
  cwd: projectRoot,
  encoding: "utf8",
}).trim();
const shortCommit = baseCommit.slice(0, 8);
const packageName = `VisualTeX-Windows-Acceptance-dev-${shortCommit}-${dateStamp}`;
const packageDirectory = join(artifactRoot, packageName);
const zipPath = join(artifactRoot, `${packageName}.zip`);

const excludedDirectoryNames = new Set([
  ".git",
  "node_modules",
  "target",
  "bin",
  "obj",
  "artifacts",
  "dist-office-macos",
]);

function shouldCopy(source) {
  const relativePath = relative(projectRoot, source);
  if (!relativePath || relativePath === ".") return true;
  const parts = relativePath.split(/[\\/]/);
  if (parts.some((part) => excludedDirectoryNames.has(part))) return false;
  if (parts.some((part) => part === ".DS_Store")) return false;
  return true;
}

function copyEntry(entry) {
  const source = join(projectRoot, entry);
  if (!existsSync(source)) return;
  const destination = join(packageDirectory, entry);
  mkdirSync(dirname(destination), { recursive: true });
  cpSync(source, destination, {
    recursive: true,
    dereference: false,
    filter: shouldCopy,
  });
}

function sha256(filePath) {
  return createHash("sha256").update(readFileSync(filePath)).digest("hex");
}

function walkFiles(directory) {
  const output = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const fullPath = join(directory, entry.name);
    if (entry.isDirectory()) output.push(...walkFiles(fullPath));
    else if (entry.isFile()) output.push(fullPath);
  }
  return output;
}

rmSync(artifactRoot, { recursive: true, force: true });
mkdirSync(packageDirectory, { recursive: true });

const entries = [
  ".github/workflows/windows-office-acceptance.yml",
  "LICENSE",
  "README.md",
  "docs/WINDOWS_OFFICE_ARCHITECTURE.md",
  "dist-office-windows-ole",
  "index.html",
  "office/windows",
  "office-bridge.html",
  "office-dialog.html",
  "office-windows-ole-bridge.html",
  "package-lock.json",
  "package.json",
  "scripts",
  "src",
  "src-tauri",
  "src-windows",
  "tsconfig.json",
  "tsconfig.node.json",
  "tsconfig.office.json",
  "vite.config.ts",
  "vite.office.config.ts",
  "vite.office.windows-ole.config.ts",
];
for (const entry of entries) copyEntry(entry);

const worktreeStatus = execFileSync("git", ["status", "--short"], {
  cwd: projectRoot,
  encoding: "utf8",
});

const readme = `# VisualTeX Windows Office 验收包

生成日期：${new Date().toISOString()}
基础提交：${baseCommit}
分支：dev

本包包含当前本地工作区的 Windows Office 集成源码、Windows Office.js 构建产物、安装脚本、单元测试和真实 Word/PowerPoint 验收脚本。当前修改尚未提交或推送。

## Windows 环境要求

- Windows 10/11 x64
- 桌面版 Microsoft Word 与 PowerPoint
- Node.js 20 或更高版本
- .NET 8 SDK
- Visual Studio 2022 Build Tools，包含 MSBuild 和 .NET Framework 4.8 Developer Pack
- 构建完整 VisualTeX Windows 桌面程序时，还需要 Rust stable、Microsoft C++ Build Tools 和 WebView2

## 推荐执行顺序

在 PowerShell 中进入解压后的目录：

\`\`\`powershell
Set-ExecutionPolicy -Scope Process Bypass
.\\BUILD-WINDOWS-ACCEPTANCE.ps1
\`\`\`

该脚本会：

1. 安装 npm 依赖；
2. 重新构建 Windows Office.js；
3. 编译并测试 C# 协议层与 OLE Bridge；
4. 生成自包含单文件 OLE Bridge EXE；
5. 编译 Word/PowerPoint 原生 COM 加载项；
6. 生成 VSTO/原生加载项 MSI；
7. 把最终产物复制到 \`WINDOWS-BUILD-OUTPUT\`。

## OLE 真实 Office 验收

先关闭不需要参与测试的 Word/PowerPoint 文档，然后执行：

\`\`\`powershell
.\\RUN-OLE-ACCEPTANCE.ps1
\`\`\`

默认会测试 Word 和 PowerPoint 中 20 个公式的插入、更新、删除、快速连续更新、UUID 定位、失败回滚、撤销/重做、只读文档、跨文档/跨幻灯片定位以及 Bridge 重启。

保留测试文档：

\`\`\`powershell
.\\RUN-OLE-ACCEPTANCE.ps1 -KeepDocuments
\`\`\`

同时测试 OLE/VSTO 模式切换：

\`\`\`powershell
.\\RUN-OLE-ACCEPTANCE.ps1 -TestModeSwitch
\`\`\`

## 安装方式

OLE 模式：

\`\`\`powershell
.\\scripts\\install_windows_ole.ps1 -EnableBackgroundStart
\`\`\`

原生加载项模式：

\`\`\`powershell
.\\scripts\\install_windows_vsto.ps1 -MsiPath .\\WINDOWS-BUILD-OUTPUT\\VisualTeX-WindowsOffice-VSTO.msi
\`\`\`

安装或切换模式前建议退出 Word 和 PowerPoint。OLE 与原生加载项不会同时启用：OLE 使用 Trusted Catalog，原生模式使用 Word/PowerPoint Addins 注册表项的 LoadBehavior。

## 日志和临时文件

- 日志：\`%LOCALAPPDATA%\\VisualTeX\\office\\logs\`
- OLE 清单目录：\`%LOCALAPPDATA%\\VisualTeX\\OfficeCatalog\`
- 临时 PNG：\`%LOCALAPPDATA%\\VisualTeX\\office\\temp\`

## 说明

本 ZIP 在 macOS 工作区中整理。Windows Office.js 已在打包前构建成功；C# EXE、MSI 和真实 Office COM 测试应在 Windows 本机生成和执行，以使用真实 MSBuild、Office COM 与 Windows SDK。
`;
writeFileSync(join(packageDirectory, "README-WINDOWS-ACCEPTANCE.md"), readme, "utf8");
writeFileSync(join(packageDirectory, "BASE-COMMIT.txt"), `${baseCommit}\n`, "utf8");
writeFileSync(join(packageDirectory, "WORKTREE-STATUS.txt"), worktreeStatus, "utf8");

const buildScript = `[CmdletBinding()]
param([switch]$SkipTests)
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Require-Command([string]$Name, [string]$Hint) {
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw "$Name is required. $Hint"
    }
}

Require-Command node "Install Node.js 20 or newer."
Require-Command npm "Install Node.js 20 or newer."
Require-Command dotnet "Install the .NET 8 SDK."

if (-not (Get-Command msbuild.exe -ErrorAction SilentlyContinue)) {
    $vswhere = Join-Path \${env:ProgramFiles(x86)} "Microsoft Visual Studio\\Installer\\vswhere.exe"
    if (Test-Path $vswhere) {
        $msbuild = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\\**\\Bin\\MSBuild.exe | Select-Object -First 1
        if ($msbuild) { $env:PATH = "$(Split-Path -Parent $msbuild);$env:PATH" }
    }
}
Require-Command msbuild.exe "Install Visual Studio 2022 Build Tools with MSBuild and the .NET Framework 4.8 Developer Pack."

npm ci
if ($LASTEXITCODE -ne 0) { throw "npm ci failed." }
npm run build:office:windows-ole
if ($LASTEXITCODE -ne 0) { throw "Windows Office.js build failed." }

if (-not $SkipTests) {
    npm run test:windows-office-architecture
    if ($LASTEXITCODE -ne 0) { throw "Architecture smoke test failed." }
}

& .\\scripts\\build_windows_office.ps1 -Configuration Release -SkipTests:$SkipTests

$output = Join-Path $root "WINDOWS-BUILD-OUTPUT"
Remove-Item $output -Recurse -Force -ErrorAction SilentlyContinue
New-Item $output -ItemType Directory -Force | Out-Null
Copy-Item .\\src-tauri\\binaries\\visualtex-windows-office-bridge-x86_64-pc-windows-msvc.exe $output -Force
Copy-Item .\\src-tauri\\resources\\windows-office\\VisualTeX-WindowsOffice-VSTO.msi $output -Force
Copy-Item .\\office\\windows\\ole\\manifests (Join-Path $output "office-manifests") -Recurse -Force
Copy-Item .\\dist-office-windows-ole (Join-Path $output "office-web") -Recurse -Force
Get-ChildItem $output -File -Recurse | Get-FileHash -Algorithm SHA256 |
    ForEach-Object { "$($_.Hash)  $($_.Path.Substring($output.Length + 1))" } |
    Set-Content (Join-Path $output "SHA256SUMS.txt") -Encoding UTF8
Write-Host "Windows build output: $output"
`;
writeFileSync(join(packageDirectory, "BUILD-WINDOWS-ACCEPTANCE.ps1"), buildScript, "utf8");

const runScript = `[CmdletBinding()]
param(
    [switch]$KeepDocuments,
    [switch]$TestModeSwitch,
    [ValidateRange(20, 200)][int]$FormulaCount = 20
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$sidecar = Join-Path $root "src-tauri\\binaries\\visualtex-windows-office-bridge-x86_64-pc-windows-msvc.exe"
$msi = Join-Path $root "src-tauri\\resources\\windows-office\\VisualTeX-WindowsOffice-VSTO.msi"
if (-not (Test-Path $sidecar)) {
    throw "The OLE Bridge has not been built. Run .\\BUILD-WINDOWS-ACCEPTANCE.ps1 first."
}
$params = @{
    SidecarPath = $sidecar
    FormulaCount = $FormulaCount
    KeepDocuments = $KeepDocuments
    TestModeSwitch = $TestModeSwitch
}
if (Test-Path $msi) { $params.VstoMsiPath = $msi }
& (Join-Path $root "scripts\\run_windows_office_acceptance.ps1") @params
`;
writeFileSync(join(packageDirectory, "RUN-OLE-ACCEPTANCE.ps1"), runScript, "utf8");

const files = walkFiles(packageDirectory).sort();
const manifestLines = files.map((filePath) => {
  const relativePath = relative(packageDirectory, filePath).replaceAll("\\", "/");
  return `${sha256(filePath)}  ${relativePath}`;
});
writeFileSync(join(packageDirectory, "PACKAGE-SHA256SUMS.txt"), `${manifestLines.join("\n")}\n`, "utf8");

rmSync(zipPath, { force: true });
execFileSync(
  "/usr/bin/ditto",
  ["-c", "-k", "--norsrc", "--keepParent", packageDirectory, zipPath],
  {
  cwd: artifactRoot,
  stdio: "inherit",
});

const zipHash = sha256(zipPath);
const checksumPath = `${zipPath}.sha256`;
writeFileSync(checksumPath, `${zipHash}  ${packageName}.zip\n`, "utf8");
const downloadsDirectory = join(process.env.HOME ?? projectRoot, "Downloads");
let downloadsZipPath = null;
if (existsSync(downloadsDirectory)) {
  downloadsZipPath = join(downloadsDirectory, `${packageName}.zip`);
  copyFileSync(zipPath, downloadsZipPath);
  copyFileSync(checksumPath, `${downloadsZipPath}.sha256`);
}
const zipSize = statSync(zipPath).size;
console.log(`Windows acceptance package: ${zipPath}`);
if (downloadsZipPath) console.log(`Downloads copy: ${downloadsZipPath}`);
console.log(`Size: ${zipSize} bytes`);
console.log(`SHA-256: ${zipHash}`);
