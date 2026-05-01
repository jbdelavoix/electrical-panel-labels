# Electrical Panel Labels

Web app to create, customize, and print electrical panel labels directly in the browser.

Live site: [jbdelavoix.github.io/electrical-panel-labels](https://jbdelavoix.github.io/electrical-panel-labels/)

## Features

- Build labels by rows and modules.
- Configure `Label height` and `Module width` in mm.
- Set a global modules-per-row value, with optional per-row override.
- Edit each module (icon, label text, size, color).
- Icon picker with multilingual labels and search.
- Import/export JSON configurations.
- Compressed permalink support (`gz-b64`) with live URL sync.
- Automatic print plan (paper format, orientation, splitting, page optimization).
- Dark/light UI with print always forced to light mode.
- UI localization: French, English, German, Spanish.

## Quick Start

1. Set your project name.
2. Choose your tape/label height (commonly 30 mm or 50 mm).
3. Keep module width at 18 mm (1 module = 18 mm standard), unless needed.
4. Set modules per row (often 13, but 9/18 are also common).
5. Fill rows, choose icons, and adjust module spans.
6. Open print preview and print at 100% scale.

## Run Locally

This project is a static app (`index.html`, no build step).

```bash
python3 -m http.server 8000
```

Then open [http://localhost:8000](http://localhost:8000).

## Project Structure

- `index.html`: layout, print CSS, and UI containers.
- `app.js`: state, rendering, i18n wiring, import/export, permalink, print logic.
- `i18n/*.json`: UI translations (`fr`, `en`, `de`, `es`) and icon labels.
- `config/icons.json`: icon catalog (keys, icon names).
- `config/paper-formats.json`: paper format definitions (dimensions/standard).

## Data and Default Behavior

- On first load without `data` in URL, the app bootstraps from a built-in default permalink payload.
- JSON export includes only relevant values (minimal payload, avoids unnecessary defaults).
- Module width default is `18 mm`, label height default is `30 mm`, global modules per row default is `13`.

## Print Notes

- Paper formats: A4, A3, Letter, Tabloid.
- Orientation is chosen automatically to minimize page count and avoid unnecessary splitting.
- If a row must be split, chunks are balanced when possible.
- Row and part labels are printed for split rows.

## Tech Stack

- HTML/CSS
- Vanilla JavaScript
- [Tailwind CSS (CDN)](https://tailwindcss.com/)
- [Lucide Icons (CDN)](https://lucide.dev/)

## Deployment

The site is published with GitHub Pages from this repository.

## Contributing

Contributions are welcome.

1. Fork the repository.
2. Create a feature branch from `master`.
3. Make focused changes (UI, translations, icons, print logic, docs, etc.).
4. Test locally with a static server and verify print preview.
5. Open a pull request with a clear summary and test notes.

For translation updates:

- UI strings live in `i18n/*.json`
- Icon metadata is in `config/icons.json`
- Paper formats are in `config/paper-formats.json`
