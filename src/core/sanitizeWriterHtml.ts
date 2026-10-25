const ALLOWED_TAGS = new Set([
  "P", "BR", "H1", "H2", "H3", "H4", "H5", "H6",
  "STRONG", "B", "EM", "I", "U", "S", "SUB", "SUP",
  "UL", "OL", "LI", "BLOCKQUOTE", "PRE", "CODE",
  "TABLE", "THEAD", "TBODY", "TFOOT", "TR", "TH", "TD",
  "A", "IMG", "HR", "SPAN", "DIV"
]);

const DROP_WITH_CONTENT = new Set([
  "SCRIPT", "STYLE", "IFRAME", "OBJECT", "EMBED", "SVG", "MATH",
  "FORM", "INPUT", "BUTTON", "VIDEO", "AUDIO", "SOURCE",
  "LINK", "META", "BASE", "TEMPLATE"
]);

const ALLOWED_STYLE_PROPERTIES = new Set([
  "color", "background-color", "font-size", "font-family",
  "font-weight", "font-style", "text-decoration", "text-align",
  "line-height", "vertical-align", "margin-left", "text-indent"
]);

function safeStyle(value: string): string {
  if (/url\s*\(|expression\s*\(|javascript:|@import|\\/i.test(value)) return "";
  return value.split(";").map((declaration) => {
    const colon = declaration.indexOf(":");
    if (colon < 1) return "";
    const property = declaration.slice(0, colon).trim().toLowerCase();
    const setting = declaration.slice(colon + 1).trim();
    if (!ALLOWED_STYLE_PROPERTIES.has(property) || !setting) return "";
    if (!/^[#(),.%\w\s"'/-]+$/.test(setting) || /<|>|javascript/i.test(setting)) return "";
    return property + ": " + setting;
  }).filter(Boolean).join("; ");
}

function safeLink(value: string): boolean {
  const url = value.trim();
  return url.startsWith("#") || /^(https?:|mailto:)/i.test(url);
}

function safeImage(value: string): boolean {
  return value.length <= 3_000_000 &&
    /^data:image\/(png|jpeg|gif|webp);base64,[a-z0-9+/=]+$/i.test(value);
}

/**
 * Reduce imported and persisted Writer HTML to a small formatting allowlist.
 * External images, scripts, event handlers, embedded content, and unsafe URLs
 * are removed before the markup enters the editable document.
 */
export function sanitizeWriterHtml(input: string): string {
  if (typeof input !== "string") return "<p></p>";
  if (input.length > 5_000_000) throw new Error("Document HTML exceeds the 5 MB safety limit.");

  const parsed = new DOMParser().parseFromString(input, "text/html");
  const elements = Array.from(parsed.body.querySelectorAll("*"));
  if (elements.length > 20_000) throw new Error("Document contains too many HTML elements.");

  for (const element of elements.reverse()) {
    if (!element.isConnected) continue;
    if (DROP_WITH_CONTENT.has(element.tagName)) {
      element.remove();
      continue;
    }
    if (!ALLOWED_TAGS.has(element.tagName)) {
      element.replaceWith(...Array.from(element.childNodes));
      continue;
    }

    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase();
      const value = attribute.value;
      if (name.startsWith("on") || name === "srcdoc" || name === "formaction" || name === "contenteditable" || name === "id" || name === "name") {
        element.removeAttribute(attribute.name);
      } else if (name === "style") {
        const cleaned = safeStyle(value);
        if (cleaned) element.setAttribute("style", cleaned);
        else element.removeAttribute("style");
      } else if (name === "class") {
        if (element.tagName === "SPAN" && value.split(/\s+/).includes("writer-equation")) {
          element.setAttribute("class", "writer-equation");
        } else {
          element.removeAttribute("class");
        }
      } else if (name === "href" && element.tagName === "A") {
        if (!safeLink(value)) element.removeAttribute("href");
      } else if (name === "src" && element.tagName === "IMG") {
        if (!safeImage(value)) element.remove();
      } else if (name === "alt" && element.tagName === "IMG") {
        element.setAttribute("alt", value.slice(0, 300));
      } else if (["width", "height", "colspan", "rowspan"].includes(name)) {
        if (!/^\d{1,4}$/.test(value) || Number(value) > 2000) element.removeAttribute(attribute.name);
      } else {
        element.removeAttribute(attribute.name);
      }
    }

    if (element.tagName === "A" && element.hasAttribute("href")) {
      element.setAttribute("rel", "noopener noreferrer");
      element.removeAttribute("target");
    }
  }

  return parsed.body.innerHTML || "<p></p>";
}
