import { RawProperty } from "./RawProperty";

export interface PUPropertyContext {
  properties: RawProperty[];
  resolveHeap: (heap: number) => Promise<ArrayBuffer | undefined>;
}
