import JSZip from "jszip";

const OFFICE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const SPREADSHEET_NAMESPACE =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export async function createRichXlsx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    relationships([
      [
        "rId1",
        `${OFFICE_RELATIONSHIP}/officeDocument`,
        "xl/workbook.xml",
      ],
    ]),
  );
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="${SPREADSHEET_NAMESPACE}"
 xmlns:r="${OFFICE_RELATIONSHIP}">
  <workbookPr date1904="0"/>
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden data" sheetId="2" state="hidden" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="ReportArea">'Summary'!$A$1:$E$2</definedName>
    <definedName name="LocalTail" localSheetId="0">Summary!$Z$100000</definedName>
    <definedName name="InvalidScope" localSheetId="99">Summary!$A$1</definedName>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Summary!$A$1:$Z$100000</definedName>
    <definedName name="ExternalRange">[remote.xlsx]Sheet1!$A$1</definedName>
  </definedNames>
  <calcPr calcId="0" calcMode="manual" fullCalcOnLoad="0" forceFullCalc="0"/>
</workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    relationships([
      [
        "rId1",
        `${OFFICE_RELATIONSHIP}/worksheet`,
        "worksheets/sheet1.xml",
      ],
      [
        "rId2",
        `${OFFICE_RELATIONSHIP}/worksheet`,
        "worksheets/sheet2.xml",
      ],
      [
        "rId3",
        `${OFFICE_RELATIONSHIP}/styles`,
        "styles.xml",
      ],
      [
        "rId4",
        `${OFFICE_RELATIONSHIP}/sharedStrings`,
        "sharedStrings.xml",
      ],
      [
        "rIdExternal",
        `${OFFICE_RELATIONSHIP}/externalLink`,
        "https://example.invalid/external.xlsx",
        "External",
      ],
    ]),
  );
  zip.file(
    "xl/styles.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Aptos"/></font>
    <font><b/><i/><color rgb="FF336699"/><sz val="14"/><name val="Aptos Display"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FF000000"/></left><right style="thin"/><top style="thin"/><bottom style="thin"/></border>
  </borders>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="4" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="10" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="1"><alignment horizontal="center" wrapText="1"/></xf>
  </cellXfs>
</styleSheet>`,
  );
  zip.file(
    "xl/sharedStrings.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><r><t>Merged </t></r><r><t>title</t></r></si>
  <si><t>Hidden value</t></si>
</sst>`,
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="${OFFICE_RELATIONSHIP}">
  <dimension ref="A1:Z100000"/>
  <sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultColWidth="9" defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="18" customWidth="1"/><col min="26" max="26" width="14" customWidth="1"/></cols>
  <sheetData>
    <row r="1" ht="30" customHeight="1"><c r="A1" t="s" s="4"><v>0</v></c></row>
    <row r="2">
      <c r="A2" s="1"><v>45292</v></c>
      <c r="B2" s="2"><v>12345.678</v></c>
      <c r="C2" s="3"><v>0.125</v></c>
      <c r="D2"><f>WEBSERVICE(&quot;https://example.invalid/value&quot;)</f><v>42</v></c>
      <c r="E2"><f>SUM(B2,1)</f><v>12346.678</v></c>
      <c r="F2"><f t="shared" si="1"/><v>99</v></c>
    </row>
    <row r="100000"><c r="Z100000" t="inlineStr"><is><t>Sparse tail</t></is></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
  <hyperlinks>
    <hyperlink ref="C2" r:id="rIdHyperlink" tooltip="External documentation"/>
    <hyperlink ref="E2" location="'Hidden data'!A1"/>
  </hyperlinks>
  <drawing r:id="rIdDrawing"/>
</worksheet>`,
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    relationships([
      [
        "rIdHyperlink",
        `${OFFICE_RELATIONSHIP}/hyperlink`,
        "https://example.invalid/docs",
        "External",
      ],
      [
        "rIdDrawing",
        `${OFFICE_RELATIONSHIP}/drawing`,
        "../drawings/drawing1.xml",
      ],
    ]),
  );
  zip.file(
    "xl/worksheets/sheet2.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData>
</worksheet>`,
  );
  zip.file(
    "xl/drawings/drawing1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="${OFFICE_RELATIONSHIP}">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="952500" cy="952500"/>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="Sample image" descr="One pixel fixture"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`,
  );
  zip.file(
    "xl/drawings/_rels/drawing1.xml.rels",
    relationships([
      [
        "rId1",
        `${OFFICE_RELATIONSHIP}/image`,
        "../media/image1.png",
      ],
    ]),
  );
  zip.file("xl/media/image1.png", onePixelPng());
  zip.file(
    "xl/connections.xml",
    `<?xml version="1.0"?><connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" name="Ignored"/></connections>`,
  );
  zip.file(
    "xl/externalLinks/externalLink1.xml",
    `<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`,
  );

  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
  });
}

export async function addOoxmlEntry(
  buffer: ArrayBuffer,
  path: string,
  content: string | Uint8Array,
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  zip.file(path, content);
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
  });
}

function relationships(
  entries: Array<
    [
      id: string,
      type: string,
      target: string,
      targetMode?: "External",
    ]
  >,
): string {
  const body = entries
    .map(([id, type, target, targetMode]) => {
      const mode = targetMode ? ` TargetMode="${targetMode}"` : "";
      return `<Relationship Id="${id}" Type="${type}" Target="${target}"${mode}/>`;
    })
    .join("");
  return `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PACKAGE_RELATIONSHIP}">${body}</Relationships>`;
}

function onePixelPng(): Uint8Array {
  return Uint8Array.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41,
    0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0xf0,
    0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99,
    0x3d, 0x1d, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
    0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
  ]);
}
