import type { PHNodeHeapReader } from "./PHNodeHeapReader.js";

export interface PHNodeHeap {
  /**
   * bType
   * 
   * Set to `0` if heap structure doesn't exist, in case of loading absent subNode
   */
  bClientSig: number;

  /**
   * hnid
   * 
   * Set to `0` if heap structure doesn't exist, in case of loading absent subNode
   */
  userRootHnid: number;

  getReader(): PHNodeHeapReader;
}
