import * as fflate from "fflate";
import { readFile, mkdir, writeFileSync, readdirSync, statSync } from "node:fs";
import { dirname, join, normalize } from "node:path";

function createDirectory(dir: string): Promise<void> {
  return new Promise((resolve, reject) => {
    mkdir(dir, { recursive: true }, (err) => {
      if (err) {
        reject(err);
      } else {
        resolve();
      }
    });
  });
}

export function unzip(xlsxFilePath: string, outputDir: string): Promise<void> {
  return new Promise((resolve, reject) => {
    readFile(xlsxFilePath, (err, data) => {
      if (err) {
        console.error(err);
        return;
      }

      fflate.unzip(new Uint8Array(data), async (unzipErr, unzipped) => {
        if (unzipErr) {
          reject(unzipErr);
          return;
        }

        for (const [filename, fileContent] of Object.entries(unzipped)) {
          try {
            const filepath = join(outputDir, filename);
            const fileDir = dirname(filepath);
            await createDirectory(fileDir);

            // When unzipping a file created with xlsx-kaku, some filenames with size 0 are included.
            if (fileContent.length > 0) {
              writeFileSync(filepath, Buffer.from(fileContent));
            }
          } catch (dirErr) {
            console.error(dirErr);
          }
        }
        resolve();
      });
    });
  });
}

export function listFiles(dir: string) {
  let fileList: string[] = [];

  const files = readdirSync(dir);
  files.forEach((file) => {
    const filePath = join(dir, file);
    const stat = statSync(filePath);

    if (stat.isDirectory()) {
      // ディレクトリの場合、再帰的に中のファイルをリストに追加
      fileList = fileList.concat(listFiles(filePath));
    } else {
      // ファイルの場合、ファイルパスをリストに追加
      fileList.push(filePath);
    }
  });

  return fileList;
}

export function removeBasePath(fullPath: string, basePath: string): string {
  const normalizedFullPath = normalize(fullPath);
  const normalizedBasePath = normalize(basePath);

  if (normalizedFullPath.startsWith(normalizedBasePath)) {
    return normalizedFullPath.substring(normalizedBasePath.length);
  }

  return fullPath;
}
