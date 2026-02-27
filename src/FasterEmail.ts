import { PSTMessage } from "./PSTMessage.class.js";

export interface FasterEmail {
  displayName: string;
  messageClass: string;
  primaryNodeId: number;
  getMessage: () => Promise<PSTMessage>;
}
