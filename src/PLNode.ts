/**
 * PST lower level node
 */
import type { PLSubNode } from "./PLSubNode.js";

export interface PLNode {
  nodeId: number;
  getParent(): PLNode | undefined;
  getChildren(): PLNode[];
  getSubNode(): PLSubNode;

  getSiblingNode(nidType: number): PLNode | undefined;
}
