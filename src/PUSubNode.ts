import type { PHNodeHeapReader } from "./PHNodeHeapReader.js";
import { getHeapFrom } from "./PHUtil.js";
import type { PLSubNode } from "./PLSubNode.js";
import { getPropertyContext } from "./PropertyContextUtil.js";
import type { PUPropertyContext } from "./PUPropertyContext.js";
import type { PUTableContext } from "./PUTableContext.js";
import type { RawProperty } from "./RawProperty.js";
import { getTableContext } from "./TableContextUtil.js";

/**
 * Represents a sub-node within a property structure.
 */
export interface PUSubNode {
  /**
   * Gets a sub-node by its ID.
   * 
   * e.g.
   * - `0x692` for recipients
   * - `0x671` for attachments
   * - `ltpRowId` for the content of an attachment
   *
   * @param childNodeId The ID of the child node.
   */
  getSubNodeOf(childNodeId: number): Promise<PUSubNode | undefined>;

  /**
   * Extracts this sub-node as a table context.
   */
  extractAsTableContext(): Promise<PUTableContext>;

  /**
   * Extracts this sub-node as a property context.
   */
  extractAsPropertyContext(): Promise<PUPropertyContext>;
}

const passThruResolver = {
  resolveValueOf: async (key: number, type: number, value: ArrayBuffer, heap: PHNodeHeapReader): Promise<any> => {
    return value;
  },
};

function mixIntoOne(array: ArrayBuffer[]): ArrayBuffer {
  if (array.length === 0) {
    return new ArrayBuffer(0);
  }
  else if (array.length === 1) {
    return array[0];
  }
  else {
    const numBytes = array.reduce((prev, it) => prev + it.byteLength, 0);
    const one = new ArrayBuffer(numBytes);
    const dest = new Uint8Array(one);
    array.reduce(
      (nextPos, source) => {
        dest.set(new Uint8Array(source), nextPos);
        return nextPos + source.byteLength;
      },
      0
    );
    return one;
  }
}

type ResolveHeap = (hnid: number) => Promise<ArrayBuffer | undefined>;

function createResolveHeap(heap: PHNodeHeapReader): ResolveHeap {
  return async (hnid: number) => {
    return mixIntoOne(
      await heap.getHeapBuffers(hnid)
    );
  };
}

/**
 * Creates a new PUSubNode instance.
 *
 * @internal
 */
export function createPUSubNodeFrom(subNode: PLSubNode): PUSubNode {
  return {
    getSubNodeOf: async (childNodeId) => {
      const childNode = await subNode.getChildBy(childNodeId);
      return childNode ? createPUSubNodeFrom(childNode) : undefined;
    },
    extractAsTableContext: async () => {
      const heap = await getHeapFrom(subNode);
      const tc = await getTableContext(
        heap,
        passThruResolver
      );

      const rows = await tc.rows();

      return {
        resolveHeap: createResolveHeap(heap.getReader()),
        numRows: rows.length,
        getRow: async (index: number): Promise<RawProperty[]> => {
          return await rows[index].listRaw();
        },
      };
    },
    extractAsPropertyContext: async () => {
      const heap = await getHeapFrom(subNode);

      const pc = await getPropertyContext(
        heap,
        passThruResolver
      );

      return {
        resolveHeap: createResolveHeap(heap.getReader()),
        properties: await pc.listRaw(),
      };
    },
  };
}
