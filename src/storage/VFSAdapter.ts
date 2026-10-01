export interface VFSFile {
  path: string;
  name: string;
  mimeType?: string;
  size?: number;
  modifiedAt?: string;
}

export interface VFSAdapter {
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, data: Uint8Array): Promise<void>;
  deleteFile(path: string): Promise<void>;
  list(path: string): Promise<VFSFile[]>;
}

export class UnconfiguredVFS implements VFSAdapter {
  async readFile(path: string): Promise<Uint8Array> {
    throw new Error(`No VFS provider configured for read: ${path}`);
  }

  async writeFile(path: string, _data: Uint8Array): Promise<void> {
    throw new Error(`No VFS provider configured for write: ${path}`);
  }

  async deleteFile(path: string): Promise<void> {
    throw new Error(`No VFS provider configured for delete: ${path}`);
  }

  async list(path: string): Promise<VFSFile[]> {
    throw new Error(`No VFS provider configured for list: ${path}`);
  }
}

export class LocalStorageVFS implements VFSAdapter {
  private readonly prefix = "oqerun-office:vfs:";

  async readFile(path: string): Promise<Uint8Array> {
    const value = localStorage.getItem(this.key(path));
    if (value === null) {
      throw new Error(`File not found: ${path}`);
    }
    return this.base64ToBytes(value);
  }

  async writeFile(path: string, data: Uint8Array): Promise<void> {
    localStorage.setItem(this.key(path), this.bytesToBase64(data));
  }

  async deleteFile(path: string): Promise<void> {
    localStorage.removeItem(this.key(path));
  }

  async list(path: string): Promise<VFSFile[]> {
    const root = this.normalize(path);
    const files: VFSFile[] = [];

    for (let index = 0; index < localStorage.length; index += 1) {
      const key = localStorage.key(index);
      if (!key || !key.startsWith(this.prefix)) continue;

      const filePath = key.slice(this.prefix.length);
      if (root !== "/" && !filePath.startsWith(`${root}/`)) continue;
      if (root === "/" || filePath.startsWith(`${root}/`)) {
        const name = filePath.split("/").pop() ?? filePath;
        const encoded = localStorage.getItem(key) ?? "";
        files.push({
          path: filePath,
          name,
          mimeType: this.mimeFor(name),
          size: this.base64ToBytes(encoded).byteLength,
        });
      }
    }

    return files.sort((a, b) => a.name.localeCompare(b.name));
  }

  private key(path: string): string {
    return this.prefix + this.normalize(path);
  }

  private normalize(path: string): string {
    const value = path.trim().replace(/\\+/g, "/");
    const withRoot = value.startsWith("/") ? value : `/${value}`;
    return withRoot.replace(/\/+/g, "/").replace(/\/$/, "") || "/";
  }

  private bytesToBase64(data: Uint8Array): string {
    let binary = "";
    const chunkSize = 0x8000;
    for (let index = 0; index < data.length; index += chunkSize) {
      binary += String.fromCharCode(...data.subarray(index, Math.min(index + chunkSize, data.length)));
    }
    return btoa(binary);
  }

  private base64ToBytes(value: string): Uint8Array {
    const binary = atob(value);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  }

  private mimeFor(name: string): string {
    const extension = name.toLowerCase().slice(name.lastIndexOf("."));
    const map: Record<string, string> = {
      ".oqdoc": "application/x-oqerun-document",
      ".json": "application/json",
      ".txt": "text/plain",
    };
    return map[extension] ?? "application/octet-stream";
  }
}
