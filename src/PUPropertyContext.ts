import type { RawProperty } from "./RawProperty.js";

export interface PUPropertyContext {
  properties: RawProperty[];
  resolveHeap: (heap: number) => Promise<ArrayBuffer | undefined>;
}
