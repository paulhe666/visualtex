# VisualTeX Web Deployment / 网页版部署

## 中文

`web` 分支是 VisualTeX 的浏览器版本，当前不包含本地 OCR，也不会调用 Tauri/Rust/Python 运行时。

### Cloudflare 构建设置

- 生产分支：`web`
- 根目录：仓库根目录
- 构建命令：`npm run build`
- 部署命令：`npx wrangler deploy --assets ./dist`

仓库已经包含 `wrangler.jsonc`，因此部署命令也可以使用：

```bash
npx wrangler deploy
```

### 本地检查

```bash
npm ci
npm run build
npm run preview
```

Vite 的生产输出位于 `dist/`。

### 当前网页功能

- MathLive 可视化公式输入；
- 多行公式编辑；
- 公式结构与符号工具栏；
- LaTeX 命令候选与模糊搜索；
- CodeMirror LaTeX 源码同步；
- 复制多种 LaTeX 格式；
- 浏览器本地历史、设置和自动保存；
- `.visualtex.json` 文档导入与导出；
- 中英文界面、深浅主题和新手教程；
- macOS、Windows、Linux 浏览器快捷键。

### 当前限制

- 网页版暂不提供 OCR；
- 数据保存在当前浏览器的 `localStorage` 中，不会自动跨设备同步；
- 清理浏览器网站数据会删除未导出的本地历史和当前文档；
- 剪贴板写入需要 HTTPS 和浏览器权限。

## English

The `web` branch is the browser edition of VisualTeX. It currently excludes local OCR and does not invoke the Tauri, Rust, or Python runtime.

### Cloudflare build settings

- Production branch: `web`
- Root directory: repository root
- Build command: `npm run build`
- Deploy command: `npx wrangler deploy --assets ./dist`

Because the repository includes `wrangler.jsonc`, the deploy command can also be:

```bash
npx wrangler deploy
```

### Local verification

```bash
npm ci
npm run build
npm run preview
```

The Vite production output is generated in `dist/`.

### Available web features

- MathLive visual formula editing;
- Multi-line formula documents;
- Formula structure and symbol toolbar;
- LaTeX command suggestions and fuzzy search;
- Synchronized CodeMirror LaTeX source;
- Multiple LaTeX copy formats;
- Browser-local history, settings, and autosave;
- `.visualtex.json` document import and export;
- Chinese/English UI, light/dark themes, and onboarding;
- Browser shortcuts on macOS, Windows, and Linux.

### Current limitations

- OCR is not available in the web edition yet;
- Data is stored in the current browser through `localStorage` and is not synchronized across devices;
- Clearing site data removes local history and unsaved browser state;
- Clipboard writing requires HTTPS and browser permission.
