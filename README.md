# Oqerun Office Suite

**The open-source office suite for Oqerun OS.**

Oqerun Office Suite is a standalone application by **VRuntime Labs** that brings three workspaces together:

- **Oqerun Writer** — documents and rich text
- **Oqerun Sheets** — spreadsheets and formulas
- **Oqerun Slides** — presentations

The project is designed as one unified application while keeping each editor modular internally.

## Architecture

Oqerun Office is designed around a storage abstraction so editors do not need to know where a file is stored.

```text
Oqerun Office
      |
      v
   VFS Adapter
      |
  +---+------------------+
  |                      |
Local / Oqerun OS    Cloud providers
                      (Google Drive, future)
```

This boundary is intended to let the same Office code run independently on desktop platforms and integrate with Oqerun OS later.

## Planned file support

### Writer
DOC, DOCX, ODT, RTF and additional compatible formats where practical.

### Sheets
XLS, XLSX, ODS, CSV and additional compatible formats where practical.

### Slides
PPT, PPTX, ODP and additional compatible formats where practical.

## Oqerun OS integration

Integration is intentionally kept outside the Oqerun OS repository during initial development.

The future integration surface includes:

- Oqerun VFS
- application/file associations
- `.oqpkg` packaging
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

The repository is currently at the **foundation stage**. The application shell and VFS boundary are established; the Writer, Sheets and Slides editing engines will be built on top of this foundation.

## License

MIT © 2026 VRuntime Labs
