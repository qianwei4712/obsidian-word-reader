import assert from "node:assert/strict";
import test from "node:test";

import { TextSearchIndex } from "../src/reader/textSearchIndex";

void test("TextSearchIndex builds once and returns immutable match offsets", async () => {
  const nodes = [
    { nodeValue: "Revenue revenue" },
    { nodeValue: "No match" },
  ] as unknown as Text[];
  const document = {
    defaultView: null,
    createTreeWalker() {
      let index = -1;
      const walker = {
        currentNode: nodes[0] as Node,
        nextNode() {
          index += 1;
          if (index >= nodes.length) {
            return false;
          }
          walker.currentNode = nodes[index];
          return true;
        },
      };
      return walker;
    },
  };
  for (const node of nodes) {
    Object.assign(node, {
      parentElement: {
        closest: () => null,
      },
    });
  }
  const root = { ownerDocument: document } as unknown as HTMLElement;
  const index = TextSearchIndex.build(root);
  const matches = await index.find("revenue");

  assert.equal(index.nodeCount, 2);
  assert.deepEqual(
    matches.map((match) => [match.start, match.end]),
    [[0, 7], [8, 15]],
  );
  assert.equal(nodes[0].nodeValue, "Revenue revenue");
});
