# Oqerun Office Suite

**The open-source office suite for Oqerun OS.**

Oqerun Office Suite is a standalone application by **VRuntime Labs** that brings three workspaces together:

- **Oqerun Writer** — documents and rich text
- **Oqerun Sheets** — spreadsheets and formulas
- **Oqerun Slides** — presentations

The project is designed as one unified application while keeping each editor modular internally.

## Architecture

Oqerun Office separates the application shell, document model and storage provider:

```text
Oqerun Office
      |
      +-----------------------+
      |                       |
   Writer                  Sheets / Slides
      |                       |
      +----------+------------+
                 |
          Document Store
                 |
             VFS Adapter
                 |
       +---------+----------+
       |                    |
 Local Oqerun VFS       Cloud providers
                        (future)
```

Editors save structured Oqerun documents through the VFS boundary. The storage provider can therefore change without rewriting the editors.

## Current capabilities

### Writer
- Rich-text editing with font family and size controls
- Bold, italic, underline, strike-through, text color and highlighting
- Lists, indentation, alignment and line spacing
- Links, tables, images and equations
- Find and replace
- Page size, margins, headers and footers
- Open Recent, Open From, Save As and Download
- `.docx` import and standards-based `.docx` export
- Local draft persistence
- Structured .oqdoc document model
- VFS-backed local save

### Sheets
- Editable spreadsheet grid with multiple sheets
- Formula bar with preserved formulas
- `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `IF` and arithmetic expressions
- Cell references, ranges and circular-reference detection
- Copy/paste, row/column insertion, sorting and filtering
- Cell formatting and quick charts
- CSV, XLS, XLSX and ODS import/export through SheetJS
- Local structured workbook save

### Slides
- Multiple slides with layouts and themes
- Add, duplicate, delete and reorder slides
- Text boxes, shapes and images
- Select/edit/resize-position controls for slide elements
- Speaker notes and slide transitions
- Full presentation mode with keyboard navigation and fullscreen
- `.pptx` export
- Local structured presentation save

## Planned file support

### Writer
DOC, DOCX, ODT, RTF and additional compatible formats where practical.

### Sheets
XLS, XLSX, ODS, CSV and additional compatible formats where practical.

### Slides
PPT, PPTX, ODP and additional compatible formats where practical.

The native `.oqdoc`, `.oqsheet` and `.oqslide` representations remain the internal application documents. The current format engine adds browser-side DOCX import/export, spreadsheet standards import/export (XLS/XLSX/ODS/CSV), and PPTX export while keeping the native models as the editing source of truth.

## Oqerun OS integration

Integration is intentionally kept outside the Oqerun OS repository during initial development.

The future integration surface includes:

- Oqerun VFS
- application/file associations
- .oqpkg packaging
- Oqerun Store distribution
- opening supported office files directly from Oqerun OS

For example:

```text
.docx  -> Oqerun Office -> Writer
.xlsx  -> Oqerun Office -> Sheets
.pptx  -> Oqerun Office -> Slides
```

## Development

Requirements:

- Node.js 20+
- npm

Install dependencies and start the development server:

```bash
npm install
npm run dev
```

Build the application:

```bash
npm run build
```

## Project status

The project is now an active office-suite foundation rather than a placeholder. The unified Office shell, Office-style Backstage file workspace, VFS boundary, recent-file tracking, Writer, Sheets and Slides editors are established.

The current engineering stage focuses on richer editing and standards-based browser-side file exchange while preserving the internal Oqerun document models. Oqerun OS integration remains intentionally separate.

## License

MIT © 2026 VRuntime Labs
