import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

async function inspect(input, output) {
  try {
    const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
    const result = await workbook.inspect({
      kind: 'region',
      sheetId: 'מידע כללי',
      range: 'D1:F2',
      maxChars: 1600,
    });
    const preview = await workbook.render({
      sheetName: 'מידע כללי',
      range: 'D1:F2',
      scale: 1,
      format: 'png',
    });
    await fs.writeFile(output, new Uint8Array(await preview.arrayBuffer()));
    return { input: path.basename(input), success: true, inspection: result.ndjson };
  } catch (error) {
    return { input: path.basename(input), success: false, error: String(error).slice(0, 1200) };
  }
}

const [input, output] = process.argv.slice(2);
console.log(JSON.stringify(await inspect(input, output)));
