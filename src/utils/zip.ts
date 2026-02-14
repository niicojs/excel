import { unzip, unzipSync, zip, zipSync, strFromU8, strToU8 } from 'fflate';

export type ZipFiles = Map<string, Uint8Array>;

export interface ZipStore {
  get(path: string): Uint8Array | undefined;
  set(path: string, content: Uint8Array): void;
  has(path: string): boolean;
  delete(path: string): void;
  getText(path: string): string | undefined;
  setText(path: string, content: string): void;
  toFiles(): Promise<ZipFiles>;
}

class EagerZipStore implements ZipStore {
  private _files: ZipFiles;

  constructor(files: ZipFiles) {
    this._files = files;
  }

  get(path: string): Uint8Array | undefined {
    return this._files.get(path);
  }

  set(path: string, content: Uint8Array): void {
    this._files.set(path, content);
  }

  has(path: string): boolean {
    return this._files.has(path);
  }

  delete(path: string): void {
    this._files.delete(path);
  }

  getText(path: string): string | undefined {
    const data = this._files.get(path);
    if (!data) return undefined;
    return strFromU8(data);
  }

  setText(path: string, content: string): void {
    this._files.set(path, strToU8(content));
  }

  toFiles(): Promise<ZipFiles> {
    return Promise.resolve(this._files);
  }
}

class LazyZipStore implements ZipStore {
  private _data: Uint8Array;
  private _files: ZipFiles = new Map();
  private _deleted: Set<string> = new Set();
  private _entryNames: Set<string> | null = null;

  constructor(data: Uint8Array) {
    this._data = data;
  }

  get(path: string): Uint8Array | undefined {
    if (this._deleted.has(path)) return undefined;
    const cached = this._files.get(path);
    if (cached) return cached;

    this._ensureIndex();
    if (this._entryNames && !this._entryNames.has(path)) return undefined;

    const result = unzipSync(this._data, {
      filter: (file) => file.name === path,
    });
    const data = result[path];
    if (data) {
      this._files.set(path, data);
    }
    return data;
  }

  set(path: string, content: Uint8Array): void {
    this._files.set(path, content);
    this._deleted.delete(path);
    if (this._entryNames) {
      this._entryNames.add(path);
    }
  }

  has(path: string): boolean {
    if (this._deleted.has(path)) return false;
    if (this._files.has(path)) return true;
    this._ensureIndex();
    return this._entryNames?.has(path) ?? false;
  }

  delete(path: string): void {
    this._files.delete(path);
    this._deleted.add(path);
    if (this._entryNames) {
      this._entryNames.delete(path);
    }
  }

  getText(path: string): string | undefined {
    const data = this.get(path);
    if (!data) return undefined;
    return strFromU8(data);
  }

  setText(path: string, content: string): void {
    this.set(path, strToU8(content));
  }

  async toFiles(): Promise<ZipFiles> {
    const unzipped = unzipSync(this._data);
    const files = new Map<string, Uint8Array>(Object.entries(unzipped));
    for (const path of this._deleted) {
      files.delete(path);
    }
    for (const [path, content] of this._files) {
      files.set(path, content);
    }
    return files;
  }

  private _ensureIndex(): void {
    if (this._entryNames) return;
    const names = new Set<string>();
    unzipSync(this._data, {
      filter: (file) => {
        names.add(file.name);
        return false;
      },
    });
    this._entryNames = names;
  }
}

export const createZipStore = (): ZipStore => {
  return new EagerZipStore(new Map());
};

/**
 * Reads a ZIP file and returns a map of path -> content
 * @param data - ZIP file as Uint8Array
 * @returns Promise resolving to a map of file paths to contents
 */
export const readZip = (data: Uint8Array, options?: { lazy?: boolean }): Promise<ZipStore> => {
  const lazy = options?.lazy ?? false;
  if (lazy) {
    return Promise.resolve(new LazyZipStore(data));
  }

  const isBun = typeof (globalThis as { Bun?: unknown }).Bun !== 'undefined';
  if (isBun) {
    try {
      const result = unzipSync(data);
      const files = new Map<string, Uint8Array>();
      for (const [path, content] of Object.entries(result)) {
        files.set(path, content);
      }
      return Promise.resolve(new EagerZipStore(files));
    } catch (error) {
      return Promise.reject(error);
    }
  }

  return new Promise((resolve, reject) => {
    unzip(data, (err, result) => {
      if (err) {
        reject(err);
        return;
      }
      const files = new Map<string, Uint8Array>();
      for (const [path, content] of Object.entries(result)) {
        files.set(path, content);
      }
      resolve(new EagerZipStore(files));
    });
  });
};

/**
 * Creates a ZIP file from a map of path -> content
 * @param files - Map of file paths to contents
 * @returns Promise resolving to ZIP file as Uint8Array
 */
export const writeZip = async (files: ZipStore): Promise<Uint8Array> => {
  const resolved = await files.toFiles();
  const zipData: Record<string, Uint8Array> = {};
  for (const [path, content] of resolved) {
    zipData[path] = content;
  }

  const isBun = typeof (globalThis as { Bun?: unknown }).Bun !== 'undefined';
  if (isBun) {
    try {
      return Promise.resolve(zipSync(zipData));
    } catch (error) {
      return Promise.reject(error);
    }
  }

  return new Promise((resolve, reject) => {
    zip(zipData, (err, result) => {
      if (err) {
        reject(err);
        return;
      }
      resolve(result);
    });
  });
};

/**
 * Reads a file from the ZIP as a UTF-8 string
 */
export const readZipText = (files: ZipStore, path: string): string | undefined => {
  return files.getText(path);
};

/**
 * Writes a UTF-8 string to the ZIP files map
 */
export const writeZipText = (files: ZipStore, path: string, content: string): void => {
  files.setText(path, content);
};
