import { createWriteStream, statSync, readdirSync } from 'fs';
import { join, relative } from 'path';
import { ZipArchive } from 'archiver';

export async function zipDirectory(sourceDir: string, outputPath: string): Promise<void> {
  await new Promise<void>((resolveDone, reject) => {
    const output = createWriteStream(outputPath);
    const archive = new ZipArchive({ zlib: { level: 9 } });
    output.on('close', () => resolveDone());
    archive.on('error', reject);
    archive.pipe(output);

    walk(sourceDir, (filePath) => {
      const rel = relative(sourceDir, filePath);
      archive.file(filePath, { name: rel });
    });
    archive.finalize();
  });
}

function walk(dir: string, onFile: (path: string) => void): void {
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) walk(full, onFile);
    else if (st.isFile()) onFile(full);
  }
}
