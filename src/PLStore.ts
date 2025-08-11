/**
 * PST lower level store
 */

import { PLNode } from "./PLNode";

export interface PLStore {
  getOneNodeBy(nodeId: number): PLNode | undefined;
  getOneNodeByOrError(nodeId: number): PLNode;
  close(): void;
  storeSupportMask: number | undefined;
}
