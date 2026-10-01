import type { OfficeDocument } from "./DocumentModel";
import { LocalStorageVFS, type VFSAdapter } from "../storage/VFSAdapter";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

export class DocumentStore {
  constructor(private readonly vfs: VFSAdapter = new LocalStorageVFS()) {}

  async save<T>(document: OfficeDocument<T>): Promise<string> {
    const path = document.metadata.path ?? `/home/user/Documents/${this.safeName(document.metadata.name)}.oqdoc`;
    const next = {
      ...document,
      metadata: { ...document.metadata, path, modifiedAt: new Date().toISOString() },
    };
    await this.vfs.writeFile(path, encoder.encode(JSON.stringify(next)));
    return path;
  }

  async load<T>(path: string): Promise<OfficeDocument<T>> {
    const bytes = await this.vfs.readFile(path);
    return JSON.parse(decoder.decode(bytes)) as OfficeDocument<T>;
  }

  async list(path = "/home/user/Documents") {
    return this.vfs.list(path);
  }

  private safeName(name: string): string {
    return name.trim().replace(/[^a-z0-9._-]+/gi, "-").replace(/^-+|-+$/g, "") || "untitled";
  }
}
