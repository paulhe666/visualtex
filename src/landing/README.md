# VisualTeX landing page

The web deployment uses two routes:

- `/` renders the VisualTeX product landing page.
- `/editor` renders the existing browser formula editor.

Cloudflare's SPA fallback in `wrangler.jsonc` keeps direct visits to `/editor` working.

## Download behavior

The header download action scrolls to the platform chooser. The chooser recommends macOS or Windows when detected and links directly to the official VisualTeX v1.2.1 installers hosted on the VisualTeX R2 download domain. Linux installers are no longer shown on the landing page. A separate link opens the full GitHub Releases page for checksums, older versions, and installation notes.

## Local checks

```bash
npm ci
npm run test:landing
npm run build:web
npm run dev
```

Then verify both `http://localhost:5173/` and `http://localhost:5173/editor`.
