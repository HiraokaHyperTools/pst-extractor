import type { PLNode } from "./PLNode.js";
import { createPUSubNodeFrom, type PUSubNode } from "./PUSubNode.js";

/**
 * Represents a node in the PST file structure.
 */
export interface PUNode {
  /**
   * The unique identifier for the node.
   */
  nodeId: number;

  /**
   * Gets the parent node of this node.
   */
  getParent(): Promise<PUNode | undefined>;

  /**
   * Gets the child nodes of this node.
   */
  getChildren(): Promise<PUNode[]>;

  /**
   * Gets the sub-node of this node.
   */
  getSubNode(): Promise<PUSubNode>;

  /**
   * Gets the sibling node of this node.
   * 
   * @param nidType The node ID type to match.
   */
  getSiblingNode(nidType: number): Promise<PUNode | undefined>;
}

function createPUNodeFromOr(node: PLNode | undefined): PUNode | undefined {
  return node ? createPUNodeFrom(node) : undefined;
}

/**
 * Creates a new PUNode instance.
 *
 * @internal
 */
export function createPUNodeFrom(node: PLNode): PUNode {
  return {
    nodeId: node.nodeId,
    getParent: async () => createPUNodeFromOr(node.getParent()),
    getChildren: async () => node.getChildren().map(createPUNodeFrom),
    getSubNode: async () => createPUSubNodeFrom(node.getSubNode()),

    getSiblingNode: async (nidType: number) => createPUNodeFromOr(node.getSiblingNode(nidType)),
  };
}
