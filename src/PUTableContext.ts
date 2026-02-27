import type { RawProperty } from "./RawProperty.js";

export interface PUTableContext {
  numRows: number;
  getRow(index: number): Promise<RawProperty[]>;
  resolveHeap: (heap: number) => Promise<ArrayBuffer | undefined>;
}
