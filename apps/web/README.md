# VisualTeX web edition / 网页版

https://visualtex.pauljianliao.com — the landing page (`/`), the tutorial (`/tutorial`) and the editor in a browser
(`/editor`). / 首页（`/`）、教程（`/tutorial`）和浏览器中的编辑器（`/editor`）。

## How it is built / 结构

The editor is the shared frontend (`apps/shared/visualtex-app`) with this app's platform modules (`src/platform`),
exactly as the macOS and Windows apps build it; nothing of the shared code is copied or patched. In a browser there
is no Office, no system-wide input, no screen capture and no local OCR (`APP_FEATURES` in `src/platform/appInfo.ts`),
so the shared code leaves those parts out. Web-only parts:

- `src/main.tsx`: the routes; `src/web/EditorRoot.tsx`: the editor, the landing showcase and the tutorial's practice
  editor (both sandboxes keep their storage in memory, `src/runtime/landingPreview.ts`).
- `src/ocr/webOcrTransport.ts` + `webOcrService.ts`: OCR through online services (OpenAI-compatible APIs, Mathpix,
  PaddleOCR, SimpleTex); `worker/index.js` relays the services that refuse browser requests (`/api/ocr/*`).
- `src/landing`, `src/tutorial`: the two pages; `scripts/landing`: the landing page's brush lettering
  (`npm run build:landing-art` after editing `scripts/landing/copy.mjs`).
- `docs/help/VisualTeX-Web-Help.md`: the help manual of the web edition.

Tauri imports in shared code resolve to `src/web/tauriShim.ts`. Documents are kept in the browser (the browser side
of `documentLibraryApi`); a document's PDF goes through the browser's print dialog.

## Commands / 命令

```bash
npm ci
npm run dev            # Vite dev server
npm run build          # type check, then the site into the repository's dist/
npm run preview
npm run test:landing   # landing page source checks
npm run test:web-ocr   # OCR relay worker
```

## Deploying / 部署

Cloudflare builds the `web` branch: `npm run build` at the repository root (it installs this app's packages and
builds into `dist/`), then `npx wrangler deploy` with the root `wrangler.jsonc` (the site's files and
`worker/index.js`). Publish with `npm run deploy:web` at the root after pushing `main`: the `web` branch becomes
GitHub's `main` as a new fast-forward commit (`tools/deploy_web.mjs`).

Cloudflare 构建 `web` 分支（仓库根目录 `npm run build`，再用根目录的 `wrangler.jsonc` 部署）。推送 `main` 后在根目录
运行 `npm run deploy:web` 发布：`web` 分支以一个快进提交变为 GitHub 上的 `main`。
