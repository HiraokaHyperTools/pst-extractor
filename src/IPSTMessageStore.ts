import type { IPSTObject } from "./IPSTObject.js";
import type { PUNode } from "./PUNode.js";

export interface IPSTMessageStore extends IPSTObject {
  /**
   * Requests access to the user node of the internal PST structure.
   */
  requestAccessToUserNode(): Promise<PUNode | undefined>;
}