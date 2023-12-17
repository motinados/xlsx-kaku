import * as fflate from "fflate";
import { readFile, mkdir, writeFile } from "node:fs";
import { dirname, join } from "node:path";

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
            writeFile(filepath, Buffer.from(fileContent), (writeErr) => {
              if (writeErr) {
                console.error(writeErr);
              }
            });
          } catch (dirErr) {
            console.error(dirErr);
          }
        }
        resolve();
      });
    });
  });
}
