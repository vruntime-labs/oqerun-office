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
- Rich-text editing
- Headings, bold, italic and underline
- Ordered and unordered lists
- Text alignment
- Links
- Table insertion
- Image insertion
- Local draft persistence
- Structured .oqdoc document model
- VFS-backed local save

### Sheets
- Editable spreadsheet grid
- Multiple sheets
- Basic numeric values
- Basic arithmetic formulas
- SUM(A1:A5)
- Local structured workbook save

### Slides
- Multiple slides
- Add, duplicate and delete slides
- Editable title/body content
- Slide thumbnails
- Presentation preview window
- Local structured presentation save

## Planned file support

### Writer
DOC, DOCX, ODT, RTF and additional compatible formats where practical.

### Sheets
XLS, XLSX, ODS, CSV and additional compatible formats where practical.

### Slides
PPT, PPTX, ODP and additional compatible formats where practical.

The current .oqdoc, .oqsheet and .oqslide representations are internal application documents; standards-based import/export is a later format-engine stage.

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

The project has moved beyond the placeholder stage. The unified shell, VFS boundary, document store, Writer foundation, functional Sheets workspace and functional Slides workspace are now established.

The next major engineering stage is standards-based import/export, beginning with DOCX/XLSX/PPTX while preserving the internal Oqerun document model.

## License

MIT © 2026 VRuntime Labs
