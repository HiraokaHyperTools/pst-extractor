/**
 * PST lower level store
 */

import type { PLNode } from "./PLNode.js";

export interface PLStore {
  getOneNodeBy(nodeId: number): PLNode | undefined;
  getOneNodeByOrError(nodeId: number): PLNode;
  close(): void;
  storeSupportMask: number | undefined;
}
