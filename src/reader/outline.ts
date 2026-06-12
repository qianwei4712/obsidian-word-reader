export interface ReaderHeading<T> {
  element: T;
  level: number;
  text: string;
}

export interface ReaderOutlineItem<T> extends ReaderHeading<T> {
  id: string;
  ancestorIds: string[];
  hasChildren: boolean;
}

export function buildReaderOutline<T>(
  headings: ReaderHeading<T>[],
): ReaderOutlineItem<T>[] {
  const parentStack: ReaderOutlineItem<T>[] = [];
  const items: ReaderOutlineItem<T>[] = [];

  for (const [index, heading] of headings.entries()) {
    while (
      parentStack.length > 0 &&
      parentStack[parentStack.length - 1].level >= heading.level
    ) {
      parentStack.pop();
    }

    const item: ReaderOutlineItem<T> = {
      ...heading,
      id: createOutlineId(index, heading.level, heading.text),
      ancestorIds: parentStack.map((parent) => parent.id),
      hasChildren: false,
    };
    const parent = parentStack[parentStack.length - 1];
    if (parent) {
      parent.hasChildren = true;
    }
    items.push(item);
    parentStack.push(item);
  }

  return items;
}

export function getVisibleOutlineId<T>(
  item: ReaderOutlineItem<T>,
  collapsedIds: ReadonlySet<string>,
): string {
  for (const ancestorId of item.ancestorIds) {
    if (collapsedIds.has(ancestorId)) {
      return ancestorId;
    }
  }
  return item.id;
}

function createOutlineId(index: number, level: number, text: string): string {
  return `${index}:${level}:${hashText(text)}`;
}

function hashText(text: string): string {
  let hash = 0x811c9dc5;
  for (const character of text) {
    hash ^= character.charCodeAt(0);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}
