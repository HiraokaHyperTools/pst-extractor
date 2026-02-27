import type { Property } from "./Property.js";
import type { RawProperty } from "./RawProperty.js";

export interface PropertyContext {
  listRaw(): Promise<RawProperty[]>;
  list(): Promise<Property[]>;
}
