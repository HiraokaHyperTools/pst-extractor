import Long from 'long';
import { PSTFolder } from './PSTFolder.class.js';
import { PSTMessageStore } from './PSTMessageStore.class.js';
import { PSTUtil } from './PSTUtil.class.js';
import { NodeMap } from './NodeMap.class.js';
import type { PSTOpts } from './PSTOpts.js';
import { createPropertyFinder } from './PAUtil.js';
import type { PLStore } from './PLStore.js';
import { getHeapFrom } from './PHUtil.js';
import { getPropertyContext } from './PropertyContextUtil.js';
import { PropertyValueResolverV1 } from './PropertyValueResolverV1.js';
import type { PLNode } from './PLNode.js';
import { PSTMessage } from './PSTMessage.class.js';
import type { PropertyValueResolver } from './PropertyValueResolver.js';
import type { PLSubNode } from './PLSubNode.js';
import type { RootProvider } from './RootProvider.js';
import { createPUNodeFrom, type PUNode } from './PUNode.js';
// eslint-disable-next-line @typescript-eslint/no-var-requires

export class PSTFile {
  /**
   * @internal
   */
  public static ENCRYPTION_TYPE_NONE = 0

  /**
   * @internal
   */
  public static ENCRYPTION_TYPE_COMPRESSIBLE = 1

  public static MESSAGE_STORE_DESCRIPTOR_IDENTIFIER = 33

  public static ROOT_FOLDER_DESCRIPTOR_IDENTIFIER = 290

  /**
   * @internal
   */
  public static PST_TYPE_ANSI = 14

  /**
   * @internal
   */
  public static PST_TYPE_ANSI_2 = 15

  /**
   * @internal
   */
  public static PST_TYPE_UNICODE = 23

  /**
   * @internal
   */
  public static PST_TYPE_2013_UNICODE = 36

  /**
   * @internal
   */
  public static PS_PUBLIC_STRINGS = 0

  /**
   * @internal
   */
  public static PS_INTERNET_HEADERS = 3

  /**
   * @internal
   */
  public static PSETID_Messaging = 7

  /**
   * @internal
   */
  public static PSETID_Note = 8

  /**
   * @internal
   */
  public static PSETID_PostRss = 9

  /**
   * @internal
   */
  public static PSETID_Task = 10

  /**
   * @internal
   */
  public static PSETID_UnifiedMessaging = 11

  /**
   * @internal
   */
  public static PS_MAPI = 12

  /**
   * @internal
   */
  public static PSETID_AirSync = 13

  /**
   * @internal
   */
  public static PSETID_Sharing = 14

  private _store: PLStore;
  private _resolver: PropertyValueResolver;

  // node tree maps
  private _nodeMap: NodeMap = new NodeMap();

  /**
   * Creates an instance of PSTFile.  File is opened in constructor.
   * @internal
   * @param {string} fileName
   * @memberof PSTFile
   */
  public constructor(
    store: PLStore,
    nodeMap: NodeMap,
    opts?: PSTOpts
  ) {
    const convertAnsiString = (opts && opts.convertAnsiString)
      || PSTUtil.createConvertAnsiString(
        (opts && opts.ansiEncoding) || "latin1"
      );
    this._nodeMap = nodeMap;
    this._store = store;
    this._resolver = new PropertyValueResolverV1(
      convertAnsiString,
      opts?.provideTypeConverterOf,
      opts?.provideFallbackTypeConverterOf
    );
  }

  /**
   * Close the file.
   * @memberof PSTFile
   */
  public async close(): Promise<void> {
    await this._store.close();
  }

  /**
   * Get name to ID map item.
   * @param {number} key
   * @param {number} idx
   * @returns {number}
   * @memberof PSTFile
   */
  public getNameToIdMapItem(key: number, idx: number): number {
    return this._nodeMap.getId(key, idx)
  }

  /**
   * Get public string to id map item.
   * @static
   * @param {string} key
   * @returns {number}
   * @memberof PSTFile
   */
  public getPublicStringToIdMapItem(key: string): number {
    return this._nodeMap.getId(key)
  }

  /**
   * Get property name from id.
   * @param {number} propertyId
   * @param {boolean} bNamed
   * @returns {string}
   * @memberof PSTFile
   */
  public getPropertyName(
    propertyId: number,
    bNamed: boolean
  ): string | undefined {
    return PSTUtil.propertyName.get(propertyId)
  }

  /**
   * Get name to id map key.
   * @param {number} propId
   * @returns {long}
   * @memberof PSTFile
   */
  public getNameToIdMapKey(propId: number): Long | undefined {
    return this._nodeMap.getNumericName(propId)
  }

  /**
   * Get the message store of the PST file.  Note that this doesn't really
   * have much information, better to look under the root folder.
   * @returns {PSTMessageStore}
   * @memberof PSTFile
   */
  public async getMessageStore(): Promise<PSTMessageStore> {
    const node = this._store.getOneNodeBy(
      PSTFile.MESSAGE_STORE_DESCRIPTOR_IDENTIFIER
    );
    if (node === undefined) {
      throw new Error("MESSAGE_STORE_DESCRIPTOR not found");
    }
    const heap = await getHeapFrom(node.getSubNode());
    const pc = await getPropertyContext(
      heap,
      this._resolver
    );
    const propertyFinder = createPropertyFinder(await pc.list());

    return new PSTMessageStore(
      this.getRootProvider(), node, node.getSubNode(), propertyFinder
    );
  }

  private async getFolderOf(node: PLNode): Promise<PSTFolder> {
    const heap = await getHeapFrom(node.getSubNode());
    const pc = await getPropertyContext(
      heap,
      this._resolver
    );
    const propertyFinder = createPropertyFinder(await pc.list());

    const output: PSTFolder = new PSTFolder(
      this.getRootProvider(),
      node,
      node.getSubNode(),
      propertyFinder
    );
    return output
  }

  private async getItemOf(
    node: PLNode,
    subNode: PLSubNode
  ): Promise<PSTMessage> {
    return await PSTUtil.createAppropriatePSTMessageObject(
      this.getRootProvider(),
      node,
      subNode,
      this._resolver
    );
  }

  /**
   * Get the root folder for the PST file
   * @returns {PSTFolder}
   * @memberof PSTFile
   */
  public async getRootFolder(): Promise<PSTFolder> {
    const node = this._store.getOneNodeBy(
      PSTFile.ROOT_FOLDER_DESCRIPTOR_IDENTIFIER
    );
    if (node === undefined) {
      throw new Error("ROOT_FOLDER_DESCRIPTOR not found");
    }
    return await this.getFolderOf(node);
  }

  /**
   * Get the `Top of Outlook data file` of the PST file
   * @returns {PSTFolder}
   * @memberof PSTFile
   */
  public async getTopOfOutlookDataFile(): Promise<PSTFolder> {
    const node = this._store.getOneNodeBy(32802);
    if (node === undefined) {
      throw new Error("Top of Outlook data file not found");
    }
    return await this.getFolderOf(node);
  }

  private getRootProvider(): RootProvider {
    return {
      resolver: this._resolver,
      getNameToIdMapItem: this.getNameToIdMapItem.bind(this),
      getItemOf: this.getItemOf.bind(this),
      getFolderOf: this.getFolderOf.bind(this),
      getStringToIdMapItem: this.getPublicStringToIdMapItem.bind(this),
    } as RootProvider
  }

  /**
   * Requests access to the user node of the internal PST structure.
   * 
   * - `NID_MESSAGE_STORE` 33
   * - `NID_NAME_TO_ID_MAP` 97
   * - `ROOT_FOLDER_DESCRIPTOR_IDENTIFIER` 290 (unsafe)
   */
  public async requestAccessToUserNode(nodeId: number): Promise<PUNode | undefined> {
    const node = this._store.getOneNodeBy(nodeId);
    return node ? createPUNodeFrom(node) : undefined;
  }

  /**
   * Get the store support mask should be set to `PidTagStoreSupportMask`.
   */
  public getStoreSupportMask(): number | undefined {
    return this._store.storeSupportMask;
  }


  /**
   * JSON stringify the object properties.
   * @returns {string}
   * @memberof PSTFile
   */
  public toJSON(): any {
    return this
  }
}
