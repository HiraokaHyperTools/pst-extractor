/**
 * PST lower level store
 */

import { PLNode } from "./PLNode";

export interface PLStore {
  getOneNodeBy(nodeId: number): PLNode
}
