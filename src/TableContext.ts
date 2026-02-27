import type { TableRow } from "./TableRow.js";

export interface TableContext {
  rows(): Promise<TableRow[]>
}
