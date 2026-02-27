import type { Property } from "./Property.js";
import type { RawProperty } from "./RawProperty.js";

export interface TableRow {
  listRaw(): Promise<RawProperty[]>;
  list(): Promise<Property[]>;
}
