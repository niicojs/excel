import { unzip, zip, zipSync, strFromU8, strToU8 } from 'fflate';

export type ZipFiles = Map<string, Uint8Array>;

/**
 * Reads a ZIP file and returns a map of path -> content
 * @param data - ZIP file as Uint8Array
 * @returns Promise resolving to a map of file paths to contents
 */
export const readZip = (data: Uint8Array): Promise<ZipFiles> => {
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
      resolve(files);
    });
  });
};

/**
 * Creates a ZIP file from a map of path -> content
 * @param files - Map of file paths to contents
 * @returns Promise resolving to ZIP file as Uint8Array
 */
export const writeZip = (files: ZipFiles): Promise<Uint8Array> => {
  const zipData: Record<string, Uint8Array> = {};
  for (const [path, content] of files) {
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
export const readZipText = (files: ZipFiles, path: string): string | undefined => {
  const data = files.get(path);
  if (!data) return undefined;
  return strFromU8(data);
};

/**
 * Writes a UTF-8 string to the ZIP files map
 */
export const writeZipText = (files: ZipFiles, path: string, content: string): void => {
  files.set(path, strToU8(content));
};
