import JSZip from 'jszip';

/** A worksheet as Excel writes it: the XML inside `<sheetData>`. */
export interface SheetPart {
  readonly name: string;
  readonly rows: string;
  /** The XML before `<sheetData>`, such as `<sheetPr>`. */
  readonly before?: string;
}

/**
 * Builds an `.xlsx` file laid out as Excel saves one: a workbook part with
 * relationships, shared strings, styles and one part for each worksheet.
 * `parts` adds other files, such as charts.
 */
export async function xlsxFile({
  sheets,
  strings = [],
  numberFormats = [],
  cellFormats = [0],
  workbookProperties = '',
  parts = {},
}: {
  sheets: SheetPart[];
  /** Shared strings, as the XML inside each `<si>`. */
  strings?: string[];
  /** Custom number formats, as `[id, code]`. */
  numberFormats?: [number, string][];
  /** The number format identifier of each cell format, by index. */
  cellFormats?: number[];
  /** Attributes of `<workbookPr>`. */
  workbookProperties?: string;
  parts?: Record<string, string>;
}): Promise<Uint8Array> {
  const zip = new JSZip();
  const sheetEntries = sheets.map((_, index) => index + 1);
  zip.file(
    '[Content_Types].xml',
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheetEntries
        .map(
          (n) =>
            `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join('')}</Types>`,
    ),
  );
  zip.file(
    '_rels/.rels',
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
  );
  zip.file(
    'xl/workbook.xml',
    xml(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr ${workbookProperties}/><sheets>${sheets
        .map(
          (sheet, index) =>
            `<sheet name="${escape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 3}"/>`,
        )
        .join('')}</sheets></workbook>`,
    ),
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>${sheetEntries
        .map(
          (n) =>
            `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`,
        )
        .join('')}</Relationships>`,
    ),
  );
  zip.file(
    'xl/sharedStrings.xml',
    xml(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">${strings
        .map((string) => `<si>${string}</si>`)
        .join('')}</sst>`,
    ),
  );
  zip.file(
    'xl/styles.xml',
    xml(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${numberFormats.length}">${numberFormats
        .map(
          ([id, code]) =>
            `<numFmt numFmtId="${id}" formatCode="${escape(code)}"/>`,
        )
        .join(
          '',
        )}</numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><cellXfs count="${cellFormats.length}">${cellFormats
        .map(
          (id) =>
            `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"${id ? ' applyNumberFormat="1"' : ''}/>`,
        )
        .join('')}</cellXfs></styleSheet>`,
    ),
  );
  sheets.forEach((sheet, index) => {
    zip.file(
      `xl/worksheets/sheet${index + 1}.xml`,
      xml(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${sheet.before ?? ''}<sheetData>${sheet.rows}</sheetData></worksheet>`,
      ),
    );
  });
  for (const [path, contents] of Object.entries(parts)) {
    zip.file(path, contents);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const xml = (body: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;

const escape = (text: string) =>
  text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');

/**
 * The parts that a chart, a pivot table and a picture add to a file saved
 * by Excel.
 */
export const UNSUPPORTED_PARTS: Record<string, string> = {
  'xl/drawings/drawing1.xml': xml(
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
  ),
  'xl/charts/chart1.xml': xml(
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>',
  ),
  'xl/pivotTables/pivotTable1.xml': xml(
    '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1"/>',
  ),
  'xl/pivotCache/pivotCacheDefinition1.xml': xml(
    '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
  ),
  'xl/media/image1.png': 'not really a picture',
};
