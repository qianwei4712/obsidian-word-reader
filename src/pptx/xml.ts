import { DOMParser } from "@xmldom/xmldom";

export function parseXml(xml: string, path: string): Document {
  const errors: string[] = [];
  const parser = new DOMParser({
    errorHandler: {
      warning: () => undefined,
      error: (message) => {
        errors.push(String(message));
      },
      fatalError: (message) => {
        errors.push(String(message));
      },
    },
  });
  const document = parser.parseFromString(xml, "application/xml");
  if (!document.documentElement || errors.length > 0) {
    throw new Error(`Invalid XML in ${path}: ${errors[0] ?? "missing root element"}`);
  }
  return document;
}

export function childElements(element: Element): Element[] {
  const result: Element[] = [];
  for (let index = 0; index < element.childNodes.length; index += 1) {
    const node = element.childNodes.item(index);
    if (node?.nodeType === 1) {
      result.push(node as Element);
    }
  }
  return result;
}

export function childrenNamed(element: Element, name: string): Element[] {
  return childElements(element).filter((child) => localName(child) === name);
}

export function firstChildNamed(
  element: Element | null | undefined,
  name: string,
): Element | null {
  if (!element) {
    return null;
  }
  return childElements(element).find((child) => localName(child) === name) ?? null;
}

export function descendantsNamed(element: Element, name: string): Element[] {
  const result: Element[] = [];
  const visit = (current: Element): void => {
    for (const child of childElements(current)) {
      if (localName(child) === name) {
        result.push(child);
      }
      visit(child);
    }
  };
  visit(element);
  return result;
}

export function firstDescendantNamed(
  element: Element | null | undefined,
  name: string,
): Element | null {
  if (!element) {
    return null;
  }
  const queue = childElements(element);
  while (queue.length > 0) {
    const current = queue.shift();
    if (!current) {
      break;
    }
    if (localName(current) === name) {
      return current;
    }
    queue.push(...childElements(current));
  }
  return null;
}

export function localName(element: Element): string {
  return element.localName || element.nodeName.split(":").at(-1) || element.nodeName;
}

export function attribute(element: Element | null | undefined, name: string): string | null {
  if (!element) {
    return null;
  }
  for (let index = 0; index < element.attributes.length; index += 1) {
    const item = element.attributes.item(index);
    if (
      item &&
      (item.localName === name ||
        item.name === name ||
        item.name.endsWith(`:${name}`))
    ) {
      return item.value;
    }
  }
  return null;
}

export function namespacedAttribute(
  element: Element | null | undefined,
  name: string,
  namespace: string,
): string | null {
  if (!element) {
    return null;
  }
  for (let index = 0; index < element.attributes.length; index += 1) {
    const item = element.attributes.item(index);
    if (item?.localName === name && item.namespaceURI === namespace) {
      return item.value;
    }
  }
  return null;
}

export function numberAttribute(
  element: Element | null | undefined,
  name: string,
  fallback = 0,
): number {
  const value = Number(attribute(element, name));
  return Number.isFinite(value) ? value : fallback;
}

export function textContent(element: Element | null | undefined): string {
  return element?.textContent ?? "";
}
