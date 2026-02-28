import type { FasterEmail } from "./FasterEmail.js";
import type { IPSTMessage } from "./IPSTMessage.js";
import type { IPSTObject } from "./IPSTObject.js";
import type { GetFasterEmailListOptions } from "./PSTFolder.class.js";
import type { PUNode } from "./PUNode.js";

export interface IPSTFolder extends IPSTObject {
  /**
   * Get folders in one fell swoop, since there's not usually thousands of them.
   */
  getSubFolders(): Promise<IPSTFolder[]>;

  getSubFolder(index: number): Promise<IPSTFolder>;

  /**
   * The number of child folders in this folder
   * @readonly
   */
  getSubFolderCount(): Promise<number>;

  /**
   * Number of emails in this folder
   * @readonly
   */
  getEmailCount(): Promise<number>;

  getEmail(index: number): Promise<IPSTMessage>;

  getEmails(): Promise<IPSTMessage[]>;

  getFasterEmailList(options?: GetFasterEmailListOptions): Promise<FasterEmail[]>;

  /**
   * Contains a constant that indicates the folder type.
   * https://msdn.microsoft.com/en-us/library/office/cc815373.aspx
   * @readonly
   */
  get folderType(): number;

  /**
   * Contains the number of messages in a folder, as computed by the message store.
   * For a number calculated by the library use getEmailCount
   * @readonly
   */
  get contentCount(): number;

  /**
   * Contains the number of unread messages in a folder, as computed by the message store.
   * https://msdn.microsoft.com/en-us/library/office/cc841964.aspx
   * @readonly
   */
  get unreadCount(): number;

  /**
   * Contains TRUE if a folder contains subfolders.
   * once again, read from the PST, use getSubFolderCount if you want to know
   * @readonly
   */
  get hasSubfolders(): boolean;

  /**
   * Contains a text string describing the type of a folder. Although this property is
   * generally ignored, versions of Microsoft® Exchange Server prior to Exchange Server
   * 2003 Mailbox Manager expect this property to be present.
   * https://msdn.microsoft.com/en-us/library/office/cc839839.aspx
   * @readonly
   */
  get containerClass(): string;

  /**
   * Contains a bitmask of flags describing capabilities of an address book container.
   * https://msdn.microsoft.com/en-us/library/office/cc839610.aspx
   * @readonly
   */
  get containerFlags(): number;

  /**
   * Requests access to the user node of the internal PST structure.
   */
  requestAccessToUserNode(): Promise<PUNode | undefined>;

  /**
   * JSON stringify the object properties.
   */
  toJSON(): any;
}