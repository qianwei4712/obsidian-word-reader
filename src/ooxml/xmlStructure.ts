export function validateXmlStructure(
  xml: string,
  path: string,
  expectedRootLocalName: string,
): void {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) {
    throw new Error(
      `Invalid XML in ${path}: document type declarations are not allowed`,
    );
  }

  const stack: string[] = [];
  let rootName: string | null = null;
  let offset = 0;
  while (offset < xml.length) {
    const openOffset = xml.indexOf("<", offset);
    if (openOffset < 0) {
      break;
    }
    if (stack.length === 0 && xml.slice(offset, openOffset).trim().length > 0) {
      throw new Error(`Invalid XML in ${path}: text outside root element`);
    }
    if (xml.startsWith("<!--", openOffset)) {
      offset = findMarkupEnd(xml, openOffset + 4, "-->", path);
      continue;
    }
    if (xml.startsWith("<![CDATA[", openOffset)) {
      offset = findMarkupEnd(xml, openOffset + 9, "]]>", path);
      continue;
    }
    if (xml.startsWith("<?", openOffset)) {
      offset = findMarkupEnd(xml, openOffset + 2, "?>", path);
      continue;
    }
    if (xml.startsWith("<!", openOffset)) {
      throw new Error(`Invalid XML in ${path}: unsupported declaration`);
    }

    const closeOffset = findTagEnd(xml, openOffset + 1, path);
    const source = xml.slice(openOffset + 1, closeOffset).trim();
    if (source.startsWith("/")) {
      const name = readTagName(source.slice(1), path);
      if (stack.pop() !== name) {
        throw new Error(`Invalid XML in ${path}: mismatched closing tag`);
      }
    } else {
      const selfClosing = source.endsWith("/");
      const name = readTagName(source, path);
      if (stack.length === 0) {
        if (rootName !== null) {
          throw new Error(`Invalid XML in ${path}: multiple root elements`);
        }
        rootName = name;
      }
      if (!selfClosing) {
        stack.push(name);
      }
    }
    offset = closeOffset + 1;
  }

  if (stack.length === 0 && xml.slice(offset).trim().length > 0) {
    throw new Error(`Invalid XML in ${path}: text outside root element`);
  }
  if (stack.length > 0) {
    throw new Error(`Invalid XML in ${path}: unclosed element`);
  }
  if (rootName?.split(":").at(-1) !== expectedRootLocalName) {
    throw new Error(
      `Invalid XML in ${path}: expected ${expectedRootLocalName} root element`,
    );
  }
}

function findMarkupEnd(
  xml: string,
  offset: number,
  marker: string,
  path: string,
): number {
  const endOffset = xml.indexOf(marker, offset);
  if (endOffset < 0) {
    throw new Error(`Invalid XML in ${path}: unclosed markup`);
  }
  return endOffset + marker.length;
}

function findTagEnd(xml: string, offset: number, path: string): number {
  let quote: string | null = null;
  for (let index = offset; index < xml.length; index += 1) {
    const character = xml[index];
    if (quote) {
      if (character === quote) {
        quote = null;
      }
      continue;
    }
    if (character === "\"" || character === "'") {
      quote = character;
    } else if (character === ">") {
      return index;
    }
  }
  throw new Error(`Invalid XML in ${path}: unclosed tag`);
}

function readTagName(source: string, path: string): string {
  const match = /^([A-Za-z_][\w.:-]*)/.exec(source.trim());
  if (!match) {
    throw new Error(`Invalid XML in ${path}: malformed tag name`);
  }
  return match[1];
}
