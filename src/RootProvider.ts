import type { PLNode } from "./PLNode.js";
import type { PLSubNode } from "./PLSubNode.js";
import type { PropertyValueResolver } from "./PropertyValueResolver.js";
import { PSTFolder } from "./PSTFolder.class.js";
import { PSTMessage } from "./PSTMessage.class.js";

export interface RootProvider {
  resolver: PropertyValueResolver;
  getNameToIdMapItem(key: number, idx: number): number;
  getStringToIdMapItem(key: string): number;
  getItemOf(node: PLNode, subNode: PLSubNode): Promise<PSTMessage>;
  getFolderOf(node: PLNode): Promise<PSTFolder>;
}
