export type RecentModule = "writer" | "sheets" | "slides";

export interface RecentDocument {
  path: string;
  name: string;
  module: RecentModule;
  modifiedAt: string;
  pinned?: boolean;
}

const KEY = "oqerun-office-recent";
const LIMIT = 20;

export function recordRecent(document: RecentDocument): void {
  const current = getRecent().filter((item) => item.path !== document.path);
  const existing = getRecent().find((item) => item.path === document.path);
  current.unshift({ ...document, pinned: existing?.pinned ?? document.pinned ?? false });
  const pinned = current.filter((item) => item.pinned);
  const normal = current.filter((item) => !item.pinned);
  localStorage.setItem(KEY, JSON.stringify([...pinned, ...normal].slice(0, LIMIT)));
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

export function togglePinned(path: string): void {
  const items = getRecent();
  const target = items.find((item) => item.path === path);
  if (!target) return;
  target.pinned = !target.pinned;
  const pinned = items.filter((item) => item.pinned);
  const normal = items.filter((item) => !item.pinned);
  localStorage.setItem(KEY, JSON.stringify([...pinned, ...normal].slice(0, LIMIT)));
}

export function removeRecent(path: string): void {
  localStorage.setItem(KEY, JSON.stringify(getRecent().filter((item) => item.path !== path)));
}

export function clearRecent(): void {
  localStorage.removeItem(KEY);
}
