import { PSTMessage } from "./PSTMessage.class";

export interface FasterEmail {
  displayName: string;
  messageClass: string;
  primaryNodeId: number;
  getMessage: () => Promise<PSTMessage>;
}
