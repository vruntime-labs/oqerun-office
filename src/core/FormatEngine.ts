import * as XLSX from "xlsx";
import * as mammoth from "mammoth";
import {
  AlignmentType,
  Document,
  Footer,
  Header,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from "docx";
import pptxgen from "pptxgenjs";
import { parse as parsePptx } from "pptxtojson";
import type {
  PresentationContent,
  PresentationElement,
  SpreadsheetCell,
  SpreadsheetContent,
  WriterContent,
} from "./DocumentModel";

export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 100);
}

export function downloadText(text: string, fileName: string, mimeType = "text/plain"): void {
  downloadBlob(new Blob([text], { type: mimeType }), fileName);
}

function safeFileName(name: string): string {
  return name.trim().replace(/[^a-z0-9._-]+/gi, "-").replace(/^-+|-+$/g, "") || "untitled";
}

function htmlToDocxChildren(html: string): Array<Paragraph | Table> {
  const document = new DOMParser().parseFromString(html, "text/html");
  const children: Array<Paragraph | Table> = [];

  document.body.childNodes.forEach((node) => {
    if (!(node instanceof HTMLElement)) return;
    const tag = node.tagName.toLowerCase();

    if (/^h[1-6]$/.test(tag)) {
      const level = Number(tag.slice(1));
      const heading = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6][level - 1];
      children.push(new Paragraph({ heading, children: [new TextRun(node.textContent ?? "")] }));
      return;
    }

    if (tag === "ul" || tag === "ol") {
      Array.from(node.querySelectorAll(":scope > li")).forEach((item) => {
        children.push(new Paragraph({
          bullet: tag === "ul" ? { level: 0 } : undefined,
          numbering: tag === "ol" ? { reference: "oqerun-numbering", level: 0 } : undefined,
          children: [new TextRun(item.textContent ?? "")],
        }));
      });
      return;
    }

    if (tag === "table") {
      const rows = Array.from(node.querySelectorAll("tr")).map((row) =>
        new TableRow({
          children: Array.from(row.querySelectorAll("th,td")).map((cell) =>
            new TableCell({ children: [new Paragraph(cell.textContent ?? "")] }),
          ),
        }),
      );
      if (rows.length) children.push(new Table({ rows }));
      return;
    }

    const text = node.textContent ?? "";
    children.push(new Paragraph({
      alignment: node.style.textAlign === "center" ? AlignmentType.CENTER : node.style.textAlign === "right" ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [new TextRun(text)],
    }));
  });

  return children.length ? children : [new Paragraph("")];
}

export async function exportWriterDocx(content: WriterContent, fileName: string): Promise<void> {
  const settings = content.settings;
  const children = htmlToDocxChildren(content.html);
  const header = settings?.header?.trim()
    ? new Header({ children: [new Paragraph(settings.header.trim())] })
    : undefined;
  const footer = settings?.footer?.trim()
    ? new Footer({ children: [new Paragraph(settings.footer.trim())] })
    : undefined;

  const pageSize = settings?.pageSize === "Letter"
    ? { width: 12240, height: 15840 }
    : { width: 11906, height: 16838 };
  const margin = Math.max(360, Math.min(1800, (settings?.margin ?? 72) * 20));

  const doc = new Document({
    creator: "VRuntime Labs",
    numbering: {
      config: [{
        reference: "oqerun-numbering",
        levels: [{ level: 0, format: "decimal", text: "%1.", alignment: "left" }],
      }],
    },
    title: fileName,
    sections: [{
      headers: header ? { default: header } : undefined,
      footers: footer ? { default: footer } : undefined,
      properties: {
        page: {
          size: pageSize,
          margin: { top: margin, right: margin, bottom: margin, left: margin },
        },
      },
      children,
    }],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, safeFileName(fileName) + ".docx");
}

export async function importWriterDocx(file: File): Promise<{ html: string; plainText: string }> {
  const result = await mammoth.convertToHtml({ arrayBuffer: await file.arrayBuffer() });
  const html = result.value || "<p></p>";
  const parsed = new DOMParser().parseFromString(html, "text/html");
  return { html, plainText: parsed.body.innerText };
}

function sheetToArray(sheet: Record<string, unknown>): unknown[][] {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true }) as unknown[][];
}

function workbookFromContent(content: SpreadsheetContent): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new();
  for (const sheet of content.sheets) {
    const rows: unknown[][] = [];
    for (let row = 0; row < 200; row += 1) {
      const values: unknown[] = [];
      let hasValue = false;
      for (let col = 0; col < 50; col += 1) {
        const ref = XLSX.utils.encode_cell({ r: row, c: col });
        const raw = sheet.cells[ref];
        if (raw !== undefined && raw !== null && raw !== "") hasValue = true;
        if (typeof raw === "string" && raw.startsWith("=")) values.push({ t: "n", f: raw.slice(1) });
        else values.push(raw ?? "");
      }
      if (hasValue) rows.push(values);
    }
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name || "Sheet");
  }
  return workbook;
}

export function exportSheets(content: SpreadsheetContent, fileName: string, bookType: "xlsx" | "xls" | "ods" | "csv"): void {
  const workbook = workbookFromContent(content);
  const extension = bookType;
  if (bookType === "csv") {
    const first = workbook.Sheets[workbook.SheetNames[0]];
    const csv = XLSX.utils.sheet_to_csv(first);
    downloadText(csv, safeFileName(fileName) + ".csv", "text/csv");
    return;
  }
  XLSX.writeFile(workbook, safeFileName(fileName) + "." + extension, { bookType, compression: true });
}

function cellFromSheet(sheet: Record<string, unknown>, address: string): SpreadsheetCell {
  const cell = sheet[address] as { v?: unknown; f?: string } | undefined;
  if (!cell) return null;
  if (cell.f) return "=" + cell.f;
  if (typeof cell.v === "number" || typeof cell.v === "boolean" || typeof cell.v === "string") return cell.v;
  return cell.v == null ? null : String(cell.v);
}

export async function importSheets(file: File): Promise<SpreadsheetContent> {
  const workbook = XLSX.read(await file.arrayBuffer(), { cellFormula: true });
  const sheets = workbook.SheetNames.map((name) => {
    const source = workbook.Sheets[name] as Record<string, unknown>;
    const cells: Record<string, SpreadsheetCell> = {};
    const range = source["!ref"] as string | undefined;
    if (range) {
      const decoded = XLSX.utils.decode_range(range);
      for (let row = decoded.s.r; row <= decoded.e.r; row += 1) {
        for (let col = decoded.s.c; col <= decoded.e.c; col += 1) {
          const address = XLSX.utils.encode_cell({ r: row, c: col });
          const value = cellFromSheet(source, address);
          if (value !== null && value !== "") cells[address] = value;
        }
      }
    }
    return { name, cells };
  });
  return { activeSheet: 0, sheets: sheets.length ? sheets : [{ name: "Sheet 1", cells: {} }] };
}

function addSlideElement(slide: pptxgen.Slide, element: PresentationElement): void {
  const x = ((element.x ?? 10) / 100) * 13.333;
  const y = ((element.y ?? 10) / 100) * 7.5;
  const w = ((element.width ?? 25) / 100) * 13.333;
  const h = ((element.height ?? 15) / 100) * 7.5;

  if (element.type === "image") {
    slide.addImage({ data: element.value, x, y, w, h });
    return;
  }

  if (element.type === "shape") {
    slide.addText(element.value, {
      shape: pptxgen.ShapeType.rect,
      x, y, w, h,
      align: "center",
      valign: "middle",
      color: "FFFFFF",
      fill: { color: "5D78D6", transparency: 8 },
      line: { color: "5D78D6", width: 1 },
      margin: 0.08,
    });
    return;
  }

  slide.addText(element.value, {
    x, y, w, h,
    fontSize: 18,
    color: "182033",
    breakLine: false,
    margin: 0.06,
  });
}

export async function exportSlidesPptx(content: PresentationContent, fileName: string): Promise<void> {
  const presentation = new pptxgen();
  presentation.layout = "LAYOUT_WIDE";
  presentation.author = "VRuntime Labs";
  presentation.subject = "Oqerun Office presentation";
  presentation.title = fileName;

  for (const source of content.slides) {
    const slide = presentation.addSlide();
    const background = source.background.replace("#", "").toUpperCase();
    slide.background = { color: background };
    const foreground = source.theme === "light" || source.theme === "paper" ? "182033" : "F5F7FF";

    if (source.layout !== "blank") {
      slide.addText(source.title, {
        x: 0.9, y: 0.65, w: 11.5, h: 0.8,
        fontSize: 28, bold: true, color: foreground,
        margin: 0,
      });
    }
    if (source.layout === "title-body") {
      slide.addText(source.body, {
        x: 0.9, y: 1.65, w: 11.5, h: 4.5,
        fontSize: 20, color: foreground,
        valign: "top", breakLine: false,
        margin: 0,
      });
    }
    for (const element of source.elements) addSlideElement(slide, element);
    if (source.notes?.trim()) slide.addNotes(source.notes.trim());
  }

  await presentation.writeFile({ fileName: safeFileName(fileName) + ".pptx" });
}

export function importSlidesJson(content: PresentationContent): PresentationContent {
  return {
    activeSlide: Math.max(0, Math.min(content.activeSlide, content.slides.length - 1)),
    slides: content.slides.length ? content.slides : [{
      id: crypto.randomUUID(),
      title: "Title slide",
      body: "Add content here.",
      notes: "",
      theme: "light",
      background: "#ffffff",
      layout: "title-body",
      elements: [],
    }],
  };
}

export async function importSlidesPptx(file: File): Promise<PresentationContent> {
  const parsed = await parsePptx(await file.arrayBuffer(), {
    imageMode: "base64",
    videoMode: "none",
    audioMode: "none",
  });
  const width = Number(parsed.size?.width) || 960;
  const height = Number(parsed.size?.height) || 540;
  const htmlToText = (value: string): string => {
    const node = new DOMParser().parseFromString(value || "", "text/html");
    return node.body.innerText.trim();
  };

  const slides: PresentationContent["slides"] = (parsed.slides ?? []).map((source: any, index: number): PresentationContent["slides"][number] => {
    const elements: PresentationElement[] = (source.elements ?? []).flatMap((item: any) => {
      const type = item.type === "image" ? "image" : item.type === "shape" ? "shape" : item.type === "text" ? "text" : null;
      if (!type) return [];
      const content = htmlToText(item.content ?? "");
      const base64 = typeof item.base64 === "string" ? item.base64 : "";
      const value = type === "image" && base64
        ? (base64.startsWith("data:") ? base64 : `data:image/png;base64,${base64}`)
        : content;
      if (!value && type !== "image") return [];
      const fill = item.fill?.type === "color" ? item.fill.value : undefined;
      return [{
        id: crypto.randomUUID(),
        type,
        value,
        x: Number(item.left ?? 0) / width * 100,
        y: Number(item.top ?? 0) / height * 100,
        width: Number(item.width ?? 200) / width * 100,
        height: Number(item.height ?? 80) / height * 100,
        rotation: Number(item.rotate ?? 0),
        fontSize: Number(item.fontSize ?? 18),
        color: item.color ?? "#182033",
        background: fill,
      } satisfies PresentationElement];
    });

    const titleElement = elements.find((element) => element.type === "text" && element.y !== undefined && element.y < 30);
    const title = titleElement?.value || `Imported slide ${index + 1}`;
    const body = elements.filter((element) => element !== titleElement && element.type === "text").map((element) => element.value).join("\n");
    const background = source.fill?.type === "color" ? String(source.fill.value) : "#ffffff";
    const theme: PresentationSlide["theme"] = background.toLowerCase() === "#ffffff" || background.toLowerCase() === "#fff" ? "light" : "dark";

    return {
      id: crypto.randomUUID(),
      title,
      body,
      notes: source.note ?? "",
      theme,
      background,
      layout: (body ? "title-body" : title ? "title" : "blank") as PresentationSlide["layout"],
      transition: "none",
      elements,
    };
  });

  return {
    activeSlide: 0,
    slides: slides.length ? slides : [{
      id: crypto.randomUUID(),
      title: "Imported presentation",
      body: "",
      notes: "",
      theme: "light",
      background: "#ffffff",
      layout: "blank",
      transition: "none",
      elements: [],
    }],
  };
}
