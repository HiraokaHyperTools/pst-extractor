import { RawProperty } from "./RawProperty";

export interface PUTableContext {
  numRows: number;
  getRow(index: number): Promise<RawProperty[]>;
}
