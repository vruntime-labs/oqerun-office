export interface VFSFile { path: string; name: string; mimeType?: string; size?: number; modifiedAt?: string; }
export interface VFSAdapter { readFile(path: string): Promise<Uint8Array>; writeFile(path: string, data: Uint8Array): Promise<void>; deleteFile(path: string): Promise<void>; list(path: string): Promise<VFSFile[]>; }
export class UnconfiguredVFS implements VFSAdapter {
  async readFile(path: string): Promise<Uint8Array> { throw new Error(\`No VFS provider configured for read: \${path}\`); }
  async writeFile(path: string, _data: Uint8Array): Promise<void> { throw new Error(\`No VFS provider configured for write: \${path}\`); }
  async deleteFile(path: string): Promise<void> { throw new Error(\`No VFS provider configured for delete: \${path}\`); }
  async list(path: string): Promise<VFSFile[]> { throw new Error(\`No VFS provider configured for list: \${path}\`); }
}