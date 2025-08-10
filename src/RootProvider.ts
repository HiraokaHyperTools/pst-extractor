import { PropertyFinder } from "./PAUtil";
import { PLNode } from "./PLNode";
import { PLSubNode } from "./PLSubNode";
import { PropertyValueResolver } from "./PropertyValueResolver"
import { PSTFolder } from "./PSTFolder.class";
import { PSTMessage } from "./PSTMessage.class";

export interface RootProvider {
  resolver: PropertyValueResolver;
  getNameToIdMapItem(key: number, idx: number): number;
  getItemOf(node: PLNode, subNode: PLSubNode): Promise<PSTMessage>;
  getFolderOf(node: PLNode): Promise<PSTFolder>;
}
