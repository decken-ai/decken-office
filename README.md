# `decken-office` — Official Microsoft PowerPoint & WPS Office Add-ins

Turn **Microsoft PowerPoint** (`Office.js 1.5`) and **WPS Office Presentation (`wpp` / WPS演示)** into a **better Gamma**—generating cohesive, expressive native slides from an outline (with optional reference image or corporate `.pptx` DNA) or upgrading rough slides & screenshots (`⌘V`) directly inside your active presentation.

## Supported Scenarios (`POST /v1/decks`)

1. **Generate from Outline (`from_outline`)**: Type or paste a slide outline -> insert native `<p:sp>`, `<a:tbl>`, and `<p:grpSp>` slides directly into PowerPoint or WPS.
2. **Outline + Brand Style DNA (`from_outline_with_reference`)**: Attach a reference slide image or corporate `.pptx` master to inherit palette and typography without rigid placeholder box-stuffing.
3. **Improve Rough `.pptx` or Reconstruct Screenshots (`improve_material`)**: Press `⌘V` / `Ctrl+V` on any Gamma/PDF/Figma screenshot or drop a rough `.pptx` draft to compile/upgrade into native editable OpenXML.

## Install

- **Microsoft PowerPoint**: Sideload [`powerpoint/manifest.xml`](./powerpoint/manifest.xml) (or [`https://app.decken.ai/addin/manifest.xml`](https://app.decken.ai/addin/manifest.xml)).
- **WPS Office Presentation**:
  ```bash
  npm run install:wps
  ```

## License

MIT
