# VisualTeX landing page

The web deployment uses two routes:

- `/` renders the VisualTeX product landing page.
- `/editor` renders the existing browser formula editor.

Cloudflare's SPA fallback in `wrangler.jsonc` keeps direct visits to `/editor` working.

## Download behavior

The header download action scrolls to the platform chooser. The chooser recommends macOS or Windows when detected and links directly to the official VisualTeX v1.2.7 installers hosted on the VisualTeX R2 download domain. Linux installers are no longer shown on the landing page. A separate link opens the full GitHub Releases page for checksums, older versions, and installation notes.

Each platform offers two direct downloads: the full installer with local OCR and the lightweight `no-ocr` installer with API OCR. Both editions retain formula editing and Office integration. The four filenames are `VisualTeX_1.2.7_aarch64.dmg`, `VisualTeX_1.2.7_aarch64-no-ocr.dmg`, `VisualTeX_1.2.7_x64-setup.exe`, and `VisualTeX_1.2.7_x64-no-ocr-setup.exe`.

## Page behaviour

- The formula field behind the page is typeset at build time and packed around the foreground.
- The hero lettering is written in on load; the hero has the two main actions.
- Buttons, the hero arrow, the equation-number loops and ink blots are hand-drawn shapes from
  `scripts/landing/shapes.mjs`. Everything animated has a static `prefers-reduced-motion` form.

Copy lives in `scripts/landing/copy.mjs`; run `npm run build:landing-art` after changing it.

## Local checks

```bash
npm ci
npm run test:landing
npm run build:web
npm run dev
```

Then verify both `http://localhost:5173/` and `http://localhost:5173/editor`.
