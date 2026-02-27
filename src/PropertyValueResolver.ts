import type { PHNodeHeapReader } from "./PHNodeHeapReader.js";

export interface PropertyValueResolver {
  resolveValueOf(key: number, type: number, value: ArrayBuffer, heap: PHNodeHeapReader): Promise<any>
}
