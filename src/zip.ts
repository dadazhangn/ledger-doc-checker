import yauzl from 'yauzl';

const MAX_ENTRIES = 4000;
const MAX_EXPANDED = 80 * 1024 * 1024;
const MAX_XML = 8 * 1024 * 1024;

/** Validate actual expansion limits before passing an Office ZIP to another parser. */
export async function inspectZip(data: Buffer, wanted: (name: string) => boolean): Promise<Map<string, Buffer>> {
  if (data.length > 12 * 1024 * 1024) throw new Error('单文件超过 12 MB 上限。');
  const zip = await new Promise<yauzl.ZipFile>((resolve, reject) => {
    yauzl.fromBuffer(data, { lazyEntries: true, validateEntrySizes: true, strictFileNames: true }, (error, value) => {
      if (error || !value) reject(error ?? new Error('不是有效的 Office 文件。'));
      else resolve(value);
    });
  });
  return await new Promise((resolve, reject) => {
    let count = 0;
    let declared = 0;
    let expanded = 0;
    let ended = false;
    const names = new Set<string>();
    const selected = new Map<string, Buffer>();
    const fail = (error: unknown) => {
      if (ended) return;
      ended = true;
      zip.close();
      reject(error instanceof Error ? error : new Error(String(error)));
    };
    zip.on('error', fail);
    zip.on('end', () => { if (!ended) { ended = true; resolve(selected); } });
    zip.on('entry', (entry: yauzl.Entry) => {
      count++;
      declared += entry.uncompressedSize;
      if (count > MAX_ENTRIES || declared > MAX_EXPANDED) return fail(new Error('Office 压缩包超出条目数或解压大小限制。'));
      if (names.has(entry.fileName)) return fail(new Error('Office 压缩包存在重复条目。'));
      names.add(entry.fileName);
      if (entry.generalPurposeBitFlag & 1) return fail(new Error('暂不支持加密 Office 文件。'));
      if (entry.fileName.endsWith('/')) { zip.readEntry(); return; }
      const keep = wanted(entry.fileName);
      if (keep && entry.uncompressedSize > MAX_XML) return fail(new Error('文档 XML 超过 8 MB 限制。'));
      zip.openReadStream(entry, (error, stream) => {
        if (error || !stream) return fail(error ?? new Error('无法读取 Office 压缩条目。'));
        const chunks: Buffer[] = [];
        let size = 0;
        stream.on('error', fail);
        stream.on('data', (chunk: Buffer) => {
          expanded += chunk.length;
          size += chunk.length;
          if (expanded > MAX_EXPANDED || (keep && size > MAX_XML)) {
            stream.destroy();
            fail(new Error('Office 文件实际解压大小超过限制。'));
            return;
          }
          if (keep) chunks.push(chunk);
        });
        stream.on('end', () => {
          if (ended) return;
          if (keep) selected.set(entry.fileName, Buffer.concat(chunks));
          zip.readEntry();
        });
      });
    });
    zip.readEntry();
  });
}
