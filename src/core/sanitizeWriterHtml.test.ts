// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { sanitizeWriterHtml } from "./sanitizeWriterHtml";

describe("sanitizeWriterHtml security boundary", () => {
  it("removes scripts, embedded documents, and their executable content", () => {
    const result = sanitizeWriterHtml(
      '<p>Safe</p><script>window.pwned = true</script><iframe src="https://evil.example">bad</iframe><object data="x">object text</object>',
    );

    expect(result).toContain("<p>Safe</p>");
    expect(result).not.toMatch(/<script|<iframe|<object/i);
    expect(result).not.toContain("window.pwned");
    expect(result).not.toContain("object text");
  });

  it("removes event handlers and unsafe URL schemes", () => {
    const result = sanitizeWriterHtml(
      '<p onclick="alert(1)">Text</p><a href="javascript:alert(1)" onmouseover="alert(2)">link</a>',
    );

    expect(result).not.toMatch(/onclick|onmouseover|javascript:/i);
    expect(result).toContain("<p>Text</p>");
    expect(result).toContain("<a>link</a>");
  });

  it("keeps safe links but prevents opening them in a new tab", () => {
    const result = sanitizeWriterHtml('<a href="https://example.com" target="_blank">source</a>');
    const parsed = new DOMParser().parseFromString(result, "text/html");
    const link = parsed.querySelector("a");

    expect(link?.getAttribute("href")).toBe("https://example.com");
    expect(link?.getAttribute("rel")).toBe("noopener noreferrer");
    expect(link?.hasAttribute("target")).toBe(false);
  });

  it("keeps safe formatting and strips unsafe CSS", () => {
    const result = sanitizeWriterHtml(
      '<p style="color: red; font-weight: bold" onclick="alert(1)">Text</p><div style="background-image: url(javascript:alert(1))">Other</div>',
    );
    const parsed = new DOMParser().parseFromString(result, "text/html");
    const paragraph = parsed.querySelector("p");
    const div = parsed.querySelector("div");

    expect(paragraph?.getAttribute("style")).toContain("color: red");
    expect(paragraph?.getAttribute("style")).toContain("font-weight: bold");
    expect(paragraph?.hasAttribute("onclick")).toBe(false);
    expect(div?.hasAttribute("style")).toBe(false);
  });

  it("removes SVG, MathML, forms, and remote images", () => {
    const result = sanitizeWriterHtml(
      '<svg><script>alert(1)</script></svg><math><mi>x</mi></math><form><input value="secret"></form><img src="https://evil.example/pixel.png">',
    );

    expect(result).not.toMatch(/<svg|<math|<form|<input|evil\.example/i);
  });

  it("rejects HTML above the configured size limit", () => {
    expect(() => sanitizeWriterHtml("x".repeat(5_000_001))).toThrow(/5 MB/i);
  });
});
