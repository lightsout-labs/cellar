import JSZip from 'jszip';
import { builtinFormat, GENERAL } from '../format';
import {
  calculable,
  parse,
  print,
  shift,
  SHEET_COLUMNS,
  SHEET_ROWS,
  type Ast,
} from '../formula';
import type { SheetSnapshot, Snapshot, Style } from '../workbook';

/** What a file holds that Gridworks leaves out when it opens it (R31). */
export type DroppedFeature =
  | 'chart'
  | 'pivotTable'
  | 'picture'
  /** A formula Gridworks cannot calculate; its last result is kept. */
  | 'formula'
  /** A cell beyond the rows and columns of a sheet. */
  | 'cell';

export interface Dropped {
  readonly feature: DroppedFeature;
  readonly count: number;
}

/** A workbook read from a file, and what was left out of it. */
export interface Opened {
  readonly snapshot: Snapshot;
  /** Each feature left out, with how many there were, in a fixed order. */
  readonly dropped: Dropped[];
}

type Constant = SheetSnapshot['columns'][number][number];

/** The parts of a file that hold content Gridworks does not support. */
const UNSUPPORTED: [DroppedFeature, RegExp][] = [
  ['chart', /^xl\/charts\/chart\d+\.xml$/],
  ['pivotTable', /^xl\/pivotTables\/pivotTable\d+\.xml$/],
  ['picture', /^xl\/media\//],
];

const RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/**
 * Reads an `.xlsx` file as a workbook named `name`: its sheets, values,
 * formulas and number formats. Rejects a file that is not an `.xlsx` file.
 *
 * Our own reader, not ExcelJS, so that a 10 MB file opens within 5 seconds
 * (R47; ADR 0006): it scans the parts it needs and skips everything else.
 */
export async function readXlsx(
  bytes: Uint8Array,
  name: string,
): Promise<Opened> {
  const zip = await JSZip.loadAsync(bytes);
  const read = async (path: string): Promise<string> => {
    const file = zip.file(path);
    if (!file) throw new Error(`Not an .xlsx file: ${path} is missing`);
    return file.async('string');
  };
  const optional = (path: string | undefined) =>
    path && zip.file(path) ? read(path) : Promise.resolve('');

  const [workbookPath] = relationships(await optional('_rels/.rels'), '')
    .filter(({ type }) => type === 'officeDocument')
    .map(({ path }) => path);
  if (!workbookPath) throw new Error('Not an .xlsx file: no workbook');
  const workbookXml = await read(workbookPath);
  // Another kind of document, such as a `.docx` file, has another root; a
  // file whose elements have a namespace prefix is not read either.
  if (!/<workbook\b[^>]*\bxmlns="[^"]*spreadsheetml/.test(workbookXml)) {
    throw new Error('Not an .xlsx file: no spreadsheet workbook');
  }
  const parts = relationships(
    await optional(relationshipsPath(workbookPath)),
    directory(workbookPath),
  );
  const partOf = (type: string) => parts.find((part) => part.type === type);
  const [strings, styleXml] = await Promise.all([
    optional(partOf('sharedStrings')?.path).then(sharedStrings),
    optional(partOf('styles')?.path),
  ]);
  const { styles, styleOf } = cellStyles(styleXml);
  const counts = { formula: 0, cell: 0 };

  const sheets: SheetSnapshot[] = [];
  for (const tag of workbookXml.match(/<sheet\b[^>]*>/g) ?? []) {
    const { name: sheetName = '', 'r:id': id } = attributes(tag);
    const part = parts.find((each) => each.id === id);
    // A chart sheet is a chart, which the chart parts count.
    if (part?.type !== 'worksheet') continue;
    sheets.push(
      readSheet(
        `sheet${sheets.length + 1}`,
        sheetName,
        await read(part.path),
        strings,
        styleOf,
        counts,
      ),
    );
  }
  if (sheets.length === 0) {
    sheets.push({ id: 'sheet1', name: 'Sheet1', columns: [], formulas: [] });
  }

  const found: [DroppedFeature, number][] = [
    ...UNSUPPORTED.map(([feature, path]): [DroppedFeature, number] => [
      feature,
      zip.file(path).length,
    ]),
    ['formula', counts.formula],
    ['cell', counts.cell],
  ];
  return {
    snapshot: { version: 1, name, sheets, styles },
    dropped: found
      .filter(([, count]) => count > 0)
      .map(([feature, count]) => ({ feature, count })),
  };
}

// One element in the cells of a worksheet: a row, or a cell with its contents.
const ROW_OR_CELL = /<row\b([^>]*)>|<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
const REFERENCE = /\br="([A-Z]+)(\d+)"/;
const ROW_NUMBER = /\br="(\d+)"/;
const TYPE = /\bt="(\w+)"/;
const STYLE = /\bs="(\d+)"/;
const VALUE = /<v>([\s\S]*?)<\/v>/;
const FORMULA = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/;
const INLINE = /<is>([\s\S]*?)<\/is>/;
const SHARED_INDEX = /\bsi="(\d+)"/;

function readSheet(
  id: string,
  name: string,
  xml: string,
  strings: string[],
  styleOf: (index: number) => number,
  counts: { formula: number; cell: number },
): SheetSnapshot {
  const columns: Constant[][] = [];
  const styleIds: number[][] = [];
  const formulas: [number, number, string][] = [];
  // The first formula of each group shared down or across, where it is.
  const shared = new Map<
    string,
    { ast: Ast | null; row: number; col: number }
  >();
  if (!xml.includes('<worksheet')) {
    throw new Error(`Not an .xlsx file: ${name} is not a worksheet`);
  }
  const start = xml.indexOf('<sheetData');
  const before = xml.slice(0, Math.max(start, 0));
  const color =
    /<tabColor\b[^>]*\brgb="(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{6})"/.exec(
      before,
    )?.[1];

  let row = -1;
  let col = -1;
  ROW_OR_CELL.lastIndex = start < 0 ? xml.length : start;
  for (let match; (match = ROW_OR_CELL.exec(xml));) {
    const [, rowAttributes, cellAttributes = '', contents = ''] = match;
    if (rowAttributes !== undefined) {
      const number = ROW_NUMBER.exec(rowAttributes)?.[1];
      row = number ? Number(number) - 1 : row + 1;
      col = -1;
      continue;
    }
    const reference = REFERENCE.exec(cellAttributes);
    if (reference) {
      col = columnIndex(reference[1]!);
      row = Number(reference[2]) - 1;
    } else {
      col++;
    }
    if (row >= SHEET_ROWS || col >= SHEET_COLUMNS) {
      counts.cell++;
      continue;
    }

    let value = cellValue(TYPE.exec(cellAttributes)?.[1], contents, strings);
    const formula = contents.includes('<f') && FORMULA.exec(contents);
    if (formula) {
      const [, formulaAttributes = '', text] = formula;
      const group = SHARED_INDEX.exec(formulaAttributes)?.[1];
      let ast = text === undefined ? null : parse(decode(text));
      if (group !== undefined) {
        const first = shared.get(group);
        if (text !== undefined || !first) {
          shared.set(group, { ast, row, col });
        } else {
          ast =
            first.ast &&
            shift(first.ast, {
              type: 'offset',
              rows: row - first.row,
              cols: col - first.col,
            });
        }
      }
      if (ast && calculable(ast)) {
        formulas.push([row, col, `=${print(ast)}`]);
        value = null;
      } else {
        counts.formula++;
      }
    }
    if (value !== null) (columns[col] ??= [])[row] = value;
    const style = STYLE.exec(cellAttributes);
    if (style) {
      const styleId = styleOf(Number(style[1]));
      if (styleId !== 0) (styleIds[col] ??= [])[row] = styleId;
    }
  }

  return {
    id,
    name,
    ...(color && { color: `#${color.toLowerCase()}` }),
    columns: Array.from(columns, (column) =>
      Array.from(column ?? [], (value) => value ?? null),
    ),
    formulas,
    styleIds: Array.from(styleIds, (column) =>
      Array.from(column ?? [], (styleId) => styleId ?? 0),
    ),
  };
}

/** The value a cell holds, or the last result of its formula. */
function cellValue(
  type: string | undefined,
  contents: string,
  strings: string[],
): Constant {
  if (type === 'inlineStr') {
    const inline = INLINE.exec(contents)?.[1];
    return inline === undefined ? null : richText(inline);
  }
  const text = VALUE.exec(contents)?.[1];
  if (text === undefined) return null;
  switch (type) {
    case 's':
      return strings[Number(text)] ?? null;
    case 'b':
      return text === '1' || text === 'true';
    case 'str':
    case 'e':
      // An error is kept as its text, as Gridworks has no error constants.
      return decode(text);
    default: {
      const number = Number(text);
      return Number.isNaN(number) ? decode(text) : number;
    }
  }
}

/** The shared strings of a workbook, by index. */
function sharedStrings(xml: string): string[] {
  return Array.from(xml.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g), (match) =>
    richText(match[1] ?? ''),
  );
}

/** The text of a string that may be in runs, without its phonetic guide. */
function richText(xml: string): string {
  const text = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  return Array.from(text.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g), (match) =>
    decode(match[1]!),
  ).join('');
}

/**
 * The styles of the cells of a file, as Gridworks has them: the table of
 * styles, the default one first, and the index in it of each cell format of
 * the file.
 */
function cellStyles(xml: string): {
  styles: Style[];
  styleOf: (index: number) => number;
} {
  const custom = new Map<number, string>();
  for (const tag of xml.match(/<numFmt\b[^>]*>/g) ?? []) {
    const { numFmtId, formatCode } = attributes(tag);
    if (numFmtId && formatCode) custom.set(Number(numFmtId), formatCode);
  }
  const styles: Style[] = [{}];
  const index = new Map<string, number>();
  const cellFormats = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1];
  const ids = (cellFormats?.match(/<xf\b[^>]*>/g) ?? []).map((tag) => {
    const id = Number(attributes(tag).numFmtId ?? 0);
    const code = custom.get(id) ?? builtinFormat(id) ?? GENERAL;
    if (code === GENERAL) return 0;
    let styleId = index.get(code);
    if (styleId === undefined) {
      styleId = styles.push({ numberFormat: code }) - 1;
      index.set(code, styleId);
    }
    return styleId;
  });
  return { styles, styleOf: (format) => ids[format] ?? 0 };
}

/** The relationships of a part, with their targets as paths in the file. */
function relationships(
  xml: string,
  base: string,
): { id: string; type: string; path: string }[] {
  return (xml.match(/<Relationship\b[^>]*>/g) ?? []).map((tag) => {
    const { Id = '', Type = '', Target = '' } = attributes(tag);
    return {
      id: Id,
      type: Type.startsWith(RELATIONSHIP)
        ? Type.slice(RELATIONSHIP.length)
        : Type,
      path: resolve(base, Target),
    };
  });
}

/** The path of a target relative to a directory, or absolute in the file. */
function resolve(base: string, target: string): string {
  const parts: string[] = [];
  const full = target.startsWith('/') ? target.slice(1) : base + target;
  for (const part of full.split('/')) {
    if (part === '..') parts.pop();
    else if (part !== '.') parts.push(part);
  }
  return parts.join('/');
}

const directory = (path: string) => path.slice(0, path.lastIndexOf('/') + 1);

const relationshipsPath = (path: string) =>
  `${directory(path)}_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;

function attributes(tag: string): Record<string, string> {
  const found: Record<string, string> = {};
  for (const [, name, double, single] of tag.matchAll(
    /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g,
  )) {
    found[name!] = decode(double ?? single ?? '');
  }
  return found;
}

const ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/** Text with XML's entities and Excel's `_xHHHH_` escapes replaced. */
function decode(text: string): string {
  if (!text.includes('&') && !text.includes('_x')) return text;
  return text
    .replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (entity, name: string) =>
      name.startsWith('#')
        ? String.fromCodePoint(
            name[1] === 'x' || name[1] === 'X'
              ? parseInt(name.slice(2), 16)
              : Number(name.slice(1)),
          )
        : (ENTITIES[name] ?? entity),
    )
    .replace(/_x([0-9A-Fa-f]{4})_/g, (_, code: string) =>
      String.fromCharCode(parseInt(code, 16)),
    );
}

function columnIndex(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + letters.charCodeAt(i) - 64;
  }
  return col - 1;
}

/**
 * Writes a workbook as an `.xlsx` file. Formulas are written without their
 * results, and the file asks to be calculated when it is opened.
 */
export async function writeXlsx(snapshot: Snapshot): Promise<Uint8Array> {
  // Loaded only when a file is saved, so that it does not slow the start.
  const { default: ExcelJS } = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  workbook.calcProperties.fullCalcOnLoad = true;
  const styles = snapshot.styles ?? [{}];
  for (const sheet of snapshot.sheets) {
    const worksheet = workbook.addWorksheet(
      sheet.name,
      sheet.color === undefined
        ? {}
        : {
            properties: {
              tabColor: { argb: `FF${sheet.color.slice(1).toUpperCase()}` },
            },
          },
    );
    const cell = (row: number, col: number) =>
      worksheet.getCell(row + 1, col + 1);
    sheet.columns.forEach((column, col) => {
      column.forEach((value, row) => {
        if (value !== null) cell(row, col).value = value;
      });
    });
    for (const [row, col, text] of sheet.formulas) {
      const ast = parse(text.slice(1));
      const formula = ast ? print(ast, { file: true }) : text.slice(1);
      cell(row, col).value = { formula, date1904: false };
    }
    sheet.styleIds?.forEach((ids, col) => {
      ids.forEach((id, row) => {
        const code = styles[id]?.numberFormat;
        if (code !== undefined) cell(row, col).numFmt = code;
      });
    });
  }
  return new Uint8Array(await workbook.xlsx.writeBuffer());
}
