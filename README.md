# n8n-nodes-xlsx-to-pdf

An n8n community node for converting spreadsheet binary attachments to PDF with LibreOffice.

This package does not bundle LibreOffice, Gotenberg, or any native conversion service. Install LibreOffice in the n8n runtime image/container and make `soffice` or `libreoffice` available on `PATH`, or set the node's LibreOffice Executable parameter to the exact binary path.

## Node

### Spreadsheet to PDF

Reads a spreadsheet from an n8n binary property, writes it to an isolated temporary directory, runs LibreOffice in headless mode, and returns the generated PDF as binary data.

Supported input extensions:

- `.xlsx`
- `.xls`
- `.xlsm`
- `.csv`
- `.ods`

Parameters:

- `Input Binary Property`: Binary property containing the spreadsheet. Defaults to `data`.
- `Output Binary Property`: Binary property where the PDF is written. Defaults to `data`.
- `Output File Name`: Optional PDF filename. Defaults to the input filename with `.pdf`.
- `Timeout (Seconds)`: Per-item LibreOffice timeout. Defaults to `60`.
- `LibreOffice Executable`: Optional executable path or command. Defaults to auto-detecting `soffice`, then `libreoffice`.

The node preserves `pairedItem`, supports multiple input items, respects `continueOnFail`, and includes conversion metadata under `json.xlsxToPdf`.

Landscape/page-orientation control is intentionally not exposed yet. LibreOffice conversion orientation depends on spreadsheet page styles and reliable per-conversion overrides need a document styling pass before export.

## Runtime Dependency

LibreOffice must be installed separately in the n8n environment.

Example Alpine-based n8n image:

```dockerfile
FROM n8nio/n8n:latest

USER root
RUN apk add --no-cache libreoffice
USER node
```

Use the equivalent package manager for Debian/Ubuntu-based images, for example `apt-get install libreoffice`.

## Development

```bash
npm install
npm run lint
npm run build
```

The package registers one compiled node entrypoint:

```json
"dist/nodes/SpreadsheetToPdf/SpreadsheetToPdf.node.js"
```
