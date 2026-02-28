import type { PUNode } from "./PUNode.js";

export interface IPSTMessageStore {
  /**
   * Requests access to the user node of the internal PST structure.
   */
  requestAccessToUserNode(): Promise<PUNode | undefined>;
}