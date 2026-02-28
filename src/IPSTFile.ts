import type { IPSTFolder } from "./IPSTFolder.js";
import type { IPSTMessageStore } from "./IPSTMessageStore.js";
import type { PUNode } from "./PUNode.js";

export interface IPSTFile {
  /**
   * Close the file.
   */
  close(): Promise<void>;

  /**
   * Get name to ID map item.
   */
  getNameToIdMapItem(key: number, idx: number): number;

  /**
   * Get public string to id map item.
   */
  getPublicStringToIdMapItem(key: string): number;

  /**
   * Get property name from id.
   */
  getPropertyName(
    propertyId: number,
    bNamed: boolean
  ): string | undefined;

  /**
   * Get name to id map key.
   */
  getNameToIdMapKey(propId: number): Long | undefined;

  /**
   * Get the message store of the PST file.  Note that this doesn't really
   * have much information, better to look under the root folder.
   */
  getMessageStore(): Promise<IPSTMessageStore>;

  /**
   * Get the root folder for the PST file
   */
  getRootFolder(): Promise<IPSTFolder>;

  /**
   * Get the `Top of Outlook data file` of the PST file
   */
  getTopOfOutlookDataFile(): Promise<IPSTFolder>;

  /**
   * Requests access to the user node of the internal PST structure.
   * 
   * - `NID_MESSAGE_STORE` 33
   * - `NID_NAME_TO_ID_MAP` 97
   * - `ROOT_FOLDER_DESCRIPTOR_IDENTIFIER` 290 (unsafe)
   */
  requestAccessToUserNode(nodeId: number): Promise<PUNode | undefined>;

  /**
   * Get the store support mask should be set to `PidTagStoreSupportMask`.
   */
  getStoreSupportMask(): number | undefined;

  /**
   * JSON stringify the object properties.
   */
  toJSON(): any;
}