import * as fflate from "fflate";
import { readFile, writeFile, mkdir } from "node:fs";
import { basename, dirname, extname, join } from "node:path";

const EXPECTED_FILE_DIR = "tests/expected";

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

export function unzip(xlsxFilePath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const extension = extname(xlsxFilePath);
    const xlsxNameWithoutExtension = basename(xlsxFilePath, extension);

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
            const filepath = join(
              EXPECTED_FILE_DIR,
              xlsxNameWithoutExtension,
              filename
            );
            const dir = dirname(filepath);
            await createDirectory(dir);
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
