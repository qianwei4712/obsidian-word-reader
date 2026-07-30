export function validateXmlStructure(
  xml: string,
  path: string,
  expectedRootLocalName: string,
): void {
  const validator = new StreamingXmlStructureValidator(
    path,
    expectedRootLocalName,
  );
  validator.push(xml);
  validator.finish();
}

export class StreamingXmlStructureValidator {
  private readonly stack: string[] = [];
  private buffer = "";
  private rootName: string | null = null;
  private finished = false;

  constructor(
    private readonly path: string,
    private readonly expectedRootLocalName: string,
  ) {}

  push(chunk: string): void {
    if (this.finished) {
      throw new Error(
        `Invalid XML in ${this.path}: data after validation finished`,
      );
    }
    this.buffer += chunk;
    this.process(false);
  }

  finish(): void {
    if (this.finished) {
      return;
    }
    this.process(true);
    this.finished = true;
    if (this.stack.length > 0) {
      throw new Error(`Invalid XML in ${this.path}: unclosed element`);
    }
    if (
      this.rootName?.split(":").at(-1) !== this.expectedRootLocalName
    ) {
      throw new Error(
        `Invalid XML in ${this.path}: expected ${this.expectedRootLocalName} root element`,
      );
    }
  }

  private process(final: boolean): void {
    let offset = 0;
    while (offset < this.buffer.length) {
      const openOffset = this.buffer.indexOf("<", offset);
      if (openOffset < 0) {
        const trailingText = this.buffer.slice(offset);
        if (this.stack.length === 0 && trailingText.trim().length > 0) {
          throw new Error(
            `Invalid XML in ${this.path}: text outside root element`,
          );
        }
        offset = this.buffer.length;
        break;
      }
      if (
        this.stack.length === 0 &&
        this.buffer.slice(offset, openOffset).trim().length > 0
      ) {
        throw new Error(
          `Invalid XML in ${this.path}: text outside root element`,
        );
      }
      const remaining = this.buffer.slice(openOffset);
      if (
        !final &&
        ["<!--", "<![CDATA[", "<?"].some((marker) =>
          marker.startsWith(remaining),
        )
      ) {
        break;
      }
      if (this.buffer.startsWith("<!--", openOffset)) {
        const endOffset = this.buffer.indexOf("-->", openOffset + 4);
        if (endOffset < 0) {
          if (!final) {
            break;
          }
          throw new Error(`Invalid XML in ${this.path}: unclosed markup`);
        }
        offset = endOffset + 3;
        continue;
      }
      if (this.buffer.startsWith("<![CDATA[", openOffset)) {
        const endOffset = this.buffer.indexOf("]]>", openOffset + 9);
        if (endOffset < 0) {
          if (!final) {
            break;
          }
          throw new Error(`Invalid XML in ${this.path}: unclosed markup`);
        }
        offset = endOffset + 3;
        continue;
      }
      if (this.buffer.startsWith("<?", openOffset)) {
        const endOffset = this.buffer.indexOf("?>", openOffset + 2);
        if (endOffset < 0) {
          if (!final) {
            break;
          }
          throw new Error(`Invalid XML in ${this.path}: unclosed markup`);
        }
        offset = endOffset + 2;
        continue;
      }
      if (this.buffer.startsWith("<!", openOffset)) {
        throw new Error(
          `Invalid XML in ${this.path}: unsupported declaration`,
        );
      }

      const closeOffset = findTagEnd(
        this.buffer,
        openOffset + 1,
      );
      if (closeOffset < 0) {
        if (!final) {
          break;
        }
        throw new Error(`Invalid XML in ${this.path}: unclosed tag`);
      }
      const source = this.buffer.slice(openOffset + 1, closeOffset).trim();
      if (source.startsWith("/")) {
        const name = readTagName(source.slice(1), this.path);
        if (this.stack.pop() !== name) {
          throw new Error(
            `Invalid XML in ${this.path}: mismatched closing tag`,
          );
        }
      } else {
        const selfClosing = source.endsWith("/");
        const name = readTagName(source, this.path);
        if (this.stack.length === 0) {
          if (this.rootName !== null) {
            throw new Error(
              `Invalid XML in ${this.path}: multiple root elements`,
            );
          }
          this.rootName = name;
        }
        if (!selfClosing) {
          this.stack.push(name);
        }
      }
      offset = closeOffset + 1;
    }
    this.buffer = this.buffer.slice(offset);
    if (final && this.buffer.length > 0) {
      if (this.buffer.includes("<")) {
        throw new Error(`Invalid XML in ${this.path}: unclosed tag`);
      }
      if (this.stack.length === 0 && this.buffer.trim().length > 0) {
        throw new Error(
          `Invalid XML in ${this.path}: text outside root element`,
        );
      }
      this.buffer = "";
    }
  }
}

function findTagEnd(xml: string, offset: number): number {
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
  return -1;
}

function readTagName(source: string, path: string): string {
  const match = /^([A-Za-z_][\w.:-]*)/.exec(source.trim());
  if (!match) {
    throw new Error(`Invalid XML in ${path}: malformed tag name`);
  }
  return match[1];
}
