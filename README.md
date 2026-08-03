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

The installer is designed for a clean Windows computer and does not require Git, Node.js, npm, Python, PowerShell, or administrator access.

1. Extract the complete release ZIP.
2. Close every PowerPoint window.
3. Double-click **Install Iconify Search for PowerPoint.cmd**.
4. Close the installer when it finishes, then open PowerPoint yourself.
5. Choose **Home > Iconify > Search Icons**.

The installer intentionally does **not** launch PowerPoint; this avoids the Behavior Shield alert caused when `cmd.exe` registered the add-in and immediately opened Office. A script-free **Upload My Add-in** alternative is documented in `INSTALLATION.txt`.

To remove the add-in, close PowerPoint and double-click **Uninstall Iconify Search for PowerPoint.cmd**.

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
