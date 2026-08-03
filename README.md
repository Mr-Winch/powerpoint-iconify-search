# Iconify Search for PowerPoint

Search the full Iconify catalog, select one or more icons, and insert them into the current PowerPoint slide as editable SVG or transparent PNG.

## Highlights

- Native-looking PowerPoint task pane with automatic light/dark Office theme support
- Full Iconify collection catalog with collection, category, palette, style, license, and grid filters
- Multi-select with individual deselection and a clear-selection command
- Favorites organized into named accordion sections, with persistent storage and per-icon or per-section deletion
- SVG/PNG toggle, insertion size, PNG resolution, and a PowerPoint-style palette using the active slide's theme colors
- Persistent preferences, filters, zoom, and collection selection
- Consecutive icons are arranged in a compact grid on the slide
- Iconify API failover endpoints and sanitized SVG markup

## Install on Windows

The release installer is designed for a clean Windows computer. It does **not** require Git, Node.js, npm, or Python.

1. Install Microsoft PowerPoint or Microsoft 365.
2. Download and extract the complete release ZIP.
3. Double-click **Install Iconify Search for PowerPoint.cmd**.
4. When PowerPoint opens, choose **Home > Iconify > Search Icons**.
5. If the command has not appeared yet, choose **Home > Add-ins > Iconify Search** once.

The installer works per Windows user, detects PowerPoint and WebView2, registers the manifest, and opens PowerPoint. It does not use PowerShell, bypass execution policy, download programs, or run a silent installer. If WebView2 is missing, it stops and shows Microsoft's official download address. Internet access is required for Iconify searches.

To remove the add-in, double-click **Uninstall Iconify Search for PowerPoint.cmd**.

## Development

Requirements: Node.js 20.19 or newer and a current PowerPoint desktop installation.

```powershell
npm install
npm test
npm run validate
npm start
```

The development manifest uses `https://localhost:3000/`. The production manifest points to GitHub Pages.

## Architecture

This is an Office JavaScript task-pane add-in using an add-in-only XML manifest. The frontend is dependency-free JavaScript and CSS; development dependencies are only used for HTTPS localhost hosting, manifest validation, and sideloading.

Icon data is provided by [Iconify](https://iconify.design/). Each icon set retains its original license and attribution requirements.

## License

The extension source is released under the MIT License. Icons obtained through Iconify remain subject to the license of their individual collection.
