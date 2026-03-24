# hypsis.ai — GitHub Pages

This repo powers **[hypsis.ai](https://hypsis.ai)** via GitHub Pages.

## Structure

```
index.html              ← Landing page (root)
deck/index.html         ← Investor deck (accessible at /deck)
assets/                 ← Shared assets (hero fresco, etc.)
favicon.ico             ← Favicon (ICO, multi-size)
favicon-32.png          ← Favicon (PNG 32×32)
apple-touch-icon.png    ← iOS home screen icon (180×180)
CNAME                   ← Custom domain config (hypsis.ai)
_config.yml             ← Jekyll config (minimal)
```

## URLs

| Page          | URL                          |
|---------------|------------------------------|
| Landing page  | https://hypsis.ai            |
| Investor deck | https://hypsis.ai/deck       |

## How to deploy

GitHub Pages deploys automatically from the `main` branch on push. No build step needed — everything is static HTML.

### Adding a new page

1. Create a folder with an `index.html` inside it (e.g. `deck/index.html` → serves at `/deck`)
2. If the page has its own assets, put them in a subfolder (e.g. `deck/assets/`)
3. Commit and push to `main`:

```bash
git add deck/
git commit -m "Add investor deck page"
git push origin main
```

4. Wait ~60 seconds for GitHub Pages to deploy
5. Verify at `https://hypsis.ai/<folder-name>`

### Updating an existing page

Edit the HTML file, commit, and push. That's it.

```bash
git add index.html
git commit -m "Update landing page copy"
git push origin main
```

### Shared assets

Put images and other shared files in `assets/`. Reference them with relative paths:

- From root `index.html`: `src="assets/hero-fresco.png"`
- From `deck/index.html`: `src="../assets/hero-fresco.png"` (or use absolute path `/assets/hero-fresco.png`)

### Custom domain

The `CNAME` file contains `hypsis.ai`. Do not delete it — GitHub Pages uses it to serve the site on the custom domain. DNS is configured separately in the domain registrar.

### Cache busting

Browsers cache aggressively. After deploying changes, hard-refresh (`Ctrl+Shift+R` / `Cmd+Shift+R`) to see updates. For shared links, append a query param if needed: `https://hypsis.ai/deck?v=2`.
