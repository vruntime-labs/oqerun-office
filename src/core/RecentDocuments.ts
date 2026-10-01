export type RecentModule = "writer" | "sheets" | "slides";

export interface RecentDocument {
  path: string;
  name: string;
  module: RecentModule;
  modifiedAt: string;
}

const KEY = "oqerun-office-recent";
const LIMIT = 12;

export function recordRecent(document: RecentDocument): void {
  const current = getRecent().filter((item) => item.path !== document.path);
  current.unshift(document);
  localStorage.setItem(KEY, JSON.stringify(current.slice(0, LIMIT)));
}

export function getRecent(): RecentDocument[] {
  try {
    const value = localStorage.getItem(KEY);
    if (!value) return [];
    const parsed = JSON.parse(value) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((item): item is RecentDocument =>
      typeof item === "object" &&
      item !== null &&
      typeof (item as RecentDocument).path === "string" &&
      typeof (item as RecentDocument).name === "string" &&
      ["writer", "sheets", "slides"].includes((item as RecentDocument).module) &&
      typeof (item as RecentDocument).modifiedAt === "string",
    );
  } catch {
    return [];
  }
}

export function clearRecent(): void {
  localStorage.removeItem(KEY);
}
