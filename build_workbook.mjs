import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

function readArguments(values) {
  const result = {};
  for (let index = 0; index < values.length; index += 2) {
    assert(values[index].startsWith('--') && values[index + 1]);
    result[values[index].slice(2)] = values[index + 1];
  }
  for (const key of ['data', 'template', 'output', 'tool']) assert(result[key]);
  return result;
}

function fittedHeight(texts, widths, fontSize = 11) {
  let maximum = 1;
  for (let index = 0; index < texts.length; index += 1) {
    const capacity = Math.max(8, Math.floor(widths[index] * 7 / (fontSize * 0.70)));
    let lines = 0;
    for (const paragraph of String(texts[index] ?? '').split('\n')) {
      let used = 0;
      lines += 1;
      for (const word of paragraph.split(/\s+/)) {
        if (!word) continue;
        if (used && used + word.length + 1 > capacity) { lines += 1; used = 0; }
        used += word.length + (used ? 1 : 0);
        while (used > capacity) { lines += 1; used -= capacity; }
      }
    }
    maximum = Math.max(maximum, lines);
  }
  const points = Math.max(32, maximum * 14.5 + 14);
  assert(points <= 400, `Text needs ${points} pt; widen the description column`);
  return points;
}

function formatSheet(sheet, rows, widths, hebrewColumns, codeColumns) {
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const finalColumn = letters[widths.length - 1];
  const used = sheet.getRange(`A1:${finalColumn}${rows.length}`);
  used.format.font = { name: 'Arial', size: 11, color: '#111111' };
  used.format.wrapText = true;
  used.format.verticalAlignment = 'top';
  for (let index = 0; index < widths.length; index += 1) {
    sheet.getRange(`${letters[index]}1:${letters[index]}${rows.length}`).format.columnWidth = widths[index];
  }
  for (const column of hebrewColumns) {
    sheet.getRange(`${column}2:${column}${rows.length}`).format.horizontalAlignment = 'right';
  }
  for (const column of codeColumns) {
    sheet.getRange(`${column}2:${column}${rows.length}`).format.horizontalAlignment = 'left';
  }
  const header = sheet.getRange(`A1:${finalColumn}1`);
  header.format.fill = '#F2F2F2';
  header.format.font = { name: 'Arial', size: 11, bold: true, color: '#111111' };
  header.format.horizontalAlignment = 'center';
  header.format.verticalAlignment = 'center';
  header.format.rowHeight = 32;
  header.format.borders = { preset: 'doubleBottom', style: 'thin', color: '#B7B7B7' };
  for (let index = 1; index < rows.length; index += 1) {
    sheet.getRange(`A${index + 1}:${finalColumn}${index + 1}`).format.rowHeight = fittedHeight(rows[index], widths);
  }
  sheet.freezePanes.freezeRows(1);
}

async function savePreview(workbook, sheetName, range, filename) {
  const image = await workbook.render({ sheetName, range, scale: 1.5, format: 'png' });
  await fs.writeFile(filename, new Uint8Array(await image.arrayBuffer()));
}

async function main() {
  const args = readArguments(process.argv.slice(2));
  const data = JSON.parse(await fs.readFile(args.data, 'utf8'));
  const tool = data.tools.find(item => item.name === args.tool);
  assert(tool, 'Tool not found in the current source contract');
  assert.equal(data.agent_name, 'סוכן עדכון מדיניות אלוקציה');
  const outputFile = path.join(args.output, `${tool.name}.authored.xlsx`);
  try { await fs.access(outputFile); throw new Error('Output already exists'); }
  catch (error) { if (error.code !== 'ENOENT') throw error; }
  await fs.mkdir(args.output, { recursive: true });
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(args.template));
  const general = workbook.worksheets.getItemAt(0);
  const parameters = workbook.worksheets.getItemAt(1);
  if (tool.name === 'GetAllocationPolicy') {
    await savePreview(workbook, general.name, 'A1:D2', path.join(args.output, 'reference_general.png'));
  }
  for (const sheet of [general, parameters]) {
    for (const table of [...sheet.tables.items]) table.delete();
    sheet.getUsedRange().clear({ applyTo: 'all' });
  }
  const generalRows = [
    ['שם הסוכן', 'מזהה גרסה', 'שם הכלי', 'תיאור מלא מותאם לסוכן (כלי / פלט)'],
    [data.agent_name, data.agent_version, tool.name, tool.tool_description],
    [data.agent_name, data.agent_version, tool.name, `פלט ${tool.output_type}\n\n${tool.output_description}`],
  ];
  general.getRange('A1:D3').values = generalRows;
  assert.deepEqual(general.getRange('A1:D3').values, generalRows);
  const parameterRows = [
    ['שם הסוכן', 'מזהה גרסה', 'שם הכלי', 'סוג פרמטר (Input / Output)',
      'שם פרמטר / נתיב שדה', 'סוג', 'תיאור מקורי (מהקוד)', 'תיאור מלא מותאם לסוכן'],
  ];
  for (const field of tool.inputs) {
    parameterRows.push([data.agent_name, data.agent_version, tool.name,
      'Input', field.name, field.type, field.original, field.description]);
  }
  for (const field of tool.outputs) {
    parameterRows.push([data.agent_name, data.agent_version, tool.name,
      'Output', field.path, field.type, field.original, field.description]);
  }
  parameters.getRange(`A1:H${parameterRows.length}`).values = parameterRows;
  assert.deepEqual(parameters.getRange(`A1:H${parameterRows.length}`).values, parameterRows);
  const informationTable = general.tables.add('A1:D3', true, `${tool.name}Info`);
  informationTable.style = 'TableStyleMedium2';
  const fieldTable = parameters.tables.add(`A1:H${parameterRows.length}`, true, `${tool.name}Parameters`);
  fieldTable.style = 'TableStyleMedium2';
  const generalWidths = [24, 22, 29, 200];
  const parameterWidths = [24, 18, 27, 16, 49, 14, 65, 140];
  formatSheet(general, generalRows, generalWidths, ['A', 'B', 'D'], ['C']);
  formatSheet(parameters, parameterRows, parameterWidths, ['A', 'B', 'G', 'H'], ['C', 'D', 'E', 'F']);
  general.getRange('B2:B3').format.fill = '#FFF2CC';
  workbook.recalculate();
  assert.deepEqual(general.getRange('A1:D3').values, generalRows);
  assert.deepEqual(parameters.getRange(`A1:H${parameterRows.length}`).values, parameterRows);
  const formulaErrors = await workbook.inspect({kind:'match',
    searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
    options:{useRegex:true,maxResults:30},summary:'Final error scan'});
  assert(!formulaErrors.ndjson.includes('"kind":"match"'), formulaErrors.ndjson);
  await savePreview(workbook, general.name, 'A1:D3', path.join(args.output, `${tool.name}_general.png`));
  if (tool.outputs.length) {
    await savePreview(workbook, parameters.name, 'C1:H8', path.join(args.output, `${tool.name}_parameters_1.png`));
    await savePreview(workbook, parameters.name, `C9:H${parameterRows.length}`, path.join(args.output, `${tool.name}_parameters_2.png`));
  } else {
    await savePreview(workbook, parameters.name, `C1:H${parameterRows.length}`, path.join(args.output, `${tool.name}_parameters.png`));
  }
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputFile);
  const report = {tool:tool.name,agent:data.agent_name,version:data.agent_version,
    input_count:tool.inputs.length,output_field_count:tool.outputs.length,
    output_type:tool.output_type,parameter_rows:parameterRows.length - 1,
    all_originals_preserved:true,all_full_descriptions_verified:true,error_scan:formulaErrors.ndjson};
  await fs.writeFile(path.join(args.output, `${tool.name}_authoring_check.json`), JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report));
}

await main();
