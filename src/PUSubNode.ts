import { PHNodeHeapReader } from "./PHNodeHeapReader";
import { getHeapFrom } from "./PHUtil";
import { PLSubNode } from "./PLSubNode";
import { getPropertyContext } from "./PropertyContextUtil";
import { PUTableContext } from "./PUTableContext";
import { RawProperty } from "./RawProperty";
import { getTableContext } from "./TableContextUtil";

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
  extractAsPropertyContext(): Promise<RawProperty[]>;
}

const passThruResolver = {
  resolveValueOf: async (key: number, type: number, value: ArrayBuffer, heap: PHNodeHeapReader): Promise<any> => {
    return value;
  },
};

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

      return await pc.listRaw();
    },
  };
}
