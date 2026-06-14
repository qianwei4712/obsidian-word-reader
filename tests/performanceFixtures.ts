import JSZip from "jszip";

const OFFICE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export async function createLargePptx(
  slideCount: number,
): Promise<ArrayBuffer> {
  const zip = new JSZip();
  const slideIds = Array.from(
    { length: slideCount },
    (_, index) =>
      `<p:sldId id="${256 + index}" rel:id="rId${index + 1}"/>`,
  ).join("");
  const slideRelationships = Array.from(
    { length: slideCount },
    (_, index) =>
      `<Relationship Id="rId${index + 1}" ` +
      `Type="${OFFICE_RELATIONSHIP}/slide" ` +
      `Target="slides/slide${index + 1}.xml"/>`,
  ).join("");

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`,
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:rel="${OFFICE_RELATIONSHIP}"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst>${slideIds}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${slideRelationships}
</Relationships>`,
  );

  for (let index = 0; index < slideCount; index += 1) {
    zip.file(
      `ppt/slides/slide${index + 1}.xml`,
      `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Performance slide ${index + 1}</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`,
    );
  }

  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
  });
}

export async function createLargeDocx(
  pageCount: number,
): Promise<ArrayBuffer> {
  const zip = new JSZip();
  const pages = Array.from({ length: pageCount }, (_, index) => {
    const pageBreak =
      index === pageCount - 1
        ? ""
        : '<w:r><w:br w:type="page"/></w:r>';
    return `<w:p><w:r><w:t>Performance page ${index + 1}</w:t></w:r>${pageBreak}</w:p>`;
  }).join("");

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${OFFICE_RELATIONSHIP}/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${pages}<w:sectPr/></w:body>
</w:document>`,
  );

  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
  });
}
