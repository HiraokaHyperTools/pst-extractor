/* eslint-disable @typescript-eslint/no-explicit-any */
import { OutlookProperties } from './OutlookProperties.js';
import { PSTFile } from './PSTFile.class.js';
import { PSTObject } from './PSTObject.class.js';
import { PSTUtil } from './PSTUtil.class.js';
import type { PLNode } from './PLNode.js';
import { createPropertyFinder, type PropertyFinder } from './PAUtil.js';
import { PSTMessage } from './PSTMessage.class.js';
import { getTableContext } from './TableContextUtil.js';
import { getHeapFrom } from './PHUtil.js';
import type { PLSubNode } from './PLSubNode.js';
import { CollectionAsyncProvider } from './CollectionAsyncProvider.js';
import { SingleAsyncProvider } from './SingleAsyncProvider.js';
import type { RootProvider } from './RootProvider.js';
import type { FasterEmail } from './FasterEmail.js';
import { getPropertyContext } from './PropertyContextUtil.js';
import type { PropertyValueResolver } from './PropertyValueResolver.js';
import { createPUNodeFrom, type PUNode } from './PUNode.js';

export interface GetFasterEmailListOptions {
  progress?: (current: number, count: number) => void;
}

/**
 * Represents a folder in the PST File.  Allows you to access child folders or items.
 * Items are accessed through a sort of cursor arrangement.  This allows for
 * incremental reading of a folder which may have _lots_ of emails.
 * @export
 * @class PSTFolder
 * @extends {PSTObject}
 */
export class PSTFolder extends PSTObject {
  private _subFoldersProvider: SingleAsyncProvider<CollectionAsyncProvider<PSTFolder>>;
  private _emailsProvider: SingleAsyncProvider<CollectionAsyncProvider<PSTMessage>>;

  /**
   * Creates an instance of PSTFolder.
   * Represents a folder in the PST File.  Allows you to access child folders or items.
   * Items are accessed through a sort of cursor arrangement.  This allows for
   * incremental reading of a folder which may have _lots_ of emails.
   * @internal
   * @param {PSTFile} rootProvider
   * @param {DescriptorIndexNode} descriptorIndexNode
   * @param {Map<number, PSTDescriptorItem>} [localDescriptorItems]
   * @memberof PSTFolder
   */
  constructor(
    rootProvider: RootProvider,
    node: PLNode,
    subNode: PLSubNode,
    propertyFinder: PropertyFinder
  ) {
    super(rootProvider, node, subNode, propertyFinder)

    this._subFoldersProvider = new SingleAsyncProvider();
    this._emailsProvider = new SingleAsyncProvider();
  }

  private async getEmailsProvider(): Promise<CollectionAsyncProvider<PSTMessage>> {
    return this._emailsProvider.getOrCreate(
      async () => {
        const targets: {
          node: PLNode,
          propertyFinder: PropertyFinder | undefined
        }[] = [];

        if (this.getNodeType() === PSTUtil.NID_TYPE_SEARCH_FOLDER) {
          // some folder types don't have children:
        }
        else {
          // trying to read emailsTable PSTTable7C
          const contentsTableNode = this._node.getSiblingNode(PSTUtil.NID_TYPE_CONTENTS_TABLE);

          let doFallback = true;

          if (contentsTableNode !== undefined) {
            try {
              const contentsTableNodeReader = contentsTableNode.getSubNode();
              const heap = await getHeapFrom(
                contentsTableNodeReader
              );

              const tc = await getTableContext(
                heap,
                this._rootProvider.resolver
              );

              const rows = await tc.rows();

              const orderOfNodes: {
                nodeId: number,
                propertyFinder: PropertyFinder,
              }[] = [];

              for (let row of rows) {
                const props = createPropertyFinder(await row.list());
                const prop = props.findByKey(0x67f2);
                if (prop !== undefined && typeof prop.value === 'number') {
                  orderOfNodes.push({ nodeId: prop.value, propertyFinder: props });
                }
              }

              const childNodeIdMap = new Map(
                this._node.getChildren()
                  .map(node => [node.nodeId, node])
              );

              for (let { nodeId, propertyFinder } of orderOfNodes) {
                const found = childNodeIdMap.get(nodeId);
                if (found !== undefined) {
                  targets.push({
                    node: found,
                    propertyFinder: propertyFinder,
                  });
                }
              }
              doFallback = false;
            }
            catch (ex) {
              // console.error(ex);

              // There are some unknown cases that TC of email list is broken.
              // Especially on ost file.
              // Thus fallback is still required.

              // .../Folder#0 // Error: getTableContext.list(rowIndex=0) resolving property key=0x001a type=0x001f of subNode of nodeId=32974,nidType=14 failure --> Error: heap=0x11604460 of subNode of nodeId=32974,nidType=14 not found

              // In this case, Outlook 2003 will try to recover the broken TC of that folder with a kind of fallback mode.
            }
          }

          if (doFallback) {
            //console.log("fallback");
            // fallback to children as listed in the descriptor b-tree
            targets.length = 0;
            for (let node of this._node.getChildren()) {
              if (this.getNodeType(node.nodeId) === PSTUtil.NID_TYPE_NORMAL_MESSAGE) {
                targets.push({ node, propertyFinder: undefined });
              }
            }
          }
        }

        return new CollectionAsyncProvider(
          targets.length,
          async (index) => {
            if (!(index in targets)) {
              throw new RangeError(`email index ${index} out of range. maximum index is ${targets.length - 1}.`);
            }
            return await this._rootProvider.getItemOf(
              targets[index].node,
              targets[index].node.getSubNode()
            );

            // Some important properties are not provided thru table context, like:
            //
            // - msg.senderName
            // - msg.sentRepresentingEmailAddress
            // - appt.localeId
            // - contact.initials
            // - activity.bodyPrefix
            // 
            // Thus we need to load full set of properties
            // from property context of corresponding sub node.
          }
        );
      }
    );
  }

  private async getSubFoldersProvider(): Promise<CollectionAsyncProvider<PSTFolder>> {
    return this._subFoldersProvider.getOrCreate(
      async () => {
        try {
          const targets: PLNode[] = [];

          for (let node of this._node.getChildren()) {
            const nodeType = this.getNodeType(node.nodeId);
            if (false
              || nodeType === PSTUtil.NID_TYPE_NORMAL_FOLDER
            ) {
              targets.push(node);
            }
          }

          return new CollectionAsyncProvider(
            targets.length,
            async (index) => {
              if (!(index in targets)) {
                throw new RangeError(`folder index ${index} out of range. maximum index is ${targets.length - 1}.`);
              }
              return await this._rootProvider.getFolderOf(targets[index]);
            }
          );
        } catch (err) {
          console.error(
            "PSTFolder::getSubFolders Can't get child folders for folder " +
            this.displayName +
            '\n' +
            err
          )
          throw err
        }
      }
    );
  }

  /**
   * Get folders in one fell swoop, since there's not usually thousands of them.
   * @returns {PSTFolder[]}
   * @memberof PSTFolder
   */
  public async getSubFolders(): Promise<PSTFolder[]> {
    return await (await this.getSubFoldersProvider()).all();
  }

  public async getSubFolder(index: number): Promise<PSTFolder> {
    return await (await this.getSubFoldersProvider()).get(index);
  }

  /**
   * The number of child folders in this folder
   * @readonly
   * @type {number}
   * @memberof PSTFolder
   */
  public async getSubFolderCount(): Promise<number> {
    return (await this.getSubFoldersProvider()).count;
  }

  /**
   * Number of emails in this folder
   * @readonly
   * @type {number}
   * @memberof PSTFolder
   */
  public async getEmailCount(): Promise<number> {
    return (await this.getEmailsProvider()).count;
  }

  public async getEmail(index: number): Promise<PSTMessage> {
    return (await (await this.getEmailsProvider()).get(index));
  }

  public async getEmails(): Promise<PSTMessage[]> {
    return (await (await this.getEmailsProvider()).all());
  }

  public async getFasterEmailList(options?: GetFasterEmailListOptions): Promise<FasterEmail[]> {
    const list: FasterEmail[] = [];

    if (this.getNodeType() === PSTUtil.NID_TYPE_SEARCH_FOLDER) {
      // ignore
    }
    else {
      const rootProvider = this._rootProvider;
      const innerResolver: PropertyValueResolver = {
        async resolveValueOf(key, type, value, heap) {
          if (false
            || key === OutlookProperties.PR_DISPLAY_NAME
            || key === OutlookProperties.PR_SUBJECT
            || key === OutlookProperties.PR_MESSAGE_CLASS
          ) {
            return rootProvider.resolver.resolveValueOf(key, type, value, heap);
          }
          return undefined;
        },
      };

      const nodes = this._node.getChildren();
      const progress = options?.progress || ((_, __) => { });
      const count = nodes.length;
      let index = -1;

      for (let node of nodes) {
        progress(++index, count);

        if (this.getNodeType(node.nodeId) === PSTUtil.NID_TYPE_NORMAL_MESSAGE) {
          const subNode = node.getSubNode();

          const heap = await getHeapFrom(subNode);
          const pc = await getPropertyContext(
            heap,
            innerResolver
          );

          const propList = await pc.list();

          function getValueOfAny(keys: number[]): string {
            return `${propList.filter(it => keys.indexOf(it.key) !== -1)[0]?.value}`;
          }

          list.push(
            {
              displayName: getValueOfAny([OutlookProperties.PR_SUBJECT, OutlookProperties.PR_DISPLAY_NAME]),
              messageClass: getValueOfAny([OutlookProperties.PR_MESSAGE_CLASS]),
              primaryNodeId: node.nodeId,
              async getMessage() {
                return await rootProvider.getItemOf(
                  node,
                  node.getSubNode()
                );
              }
            }
          );
        }
      }
    }

    return list;
  }

  /**
   * Contains a constant that indicates the folder type.
   * https://msdn.microsoft.com/en-us/library/office/cc815373.aspx
   * @readonly
   * @type {number}
   * @memberof PSTFolder
   */
  public get folderType(): number {
    return this.getIntItem(OutlookProperties.PR_FOLDER_TYPE)
  }

  /**
   * Contains the number of messages in a folder, as computed by the message store.
   * For a number calculated by the library use getEmailCount
   * @readonly
   * @type {number}
   * @memberof PSTFolder
   */
  public get contentCount(): number {
    return this.getIntItem(OutlookProperties.PR_CONTENT_COUNT)
  }

  /**
   * Contains the number of unread messages in a folder, as computed by the message store.
   * https://msdn.microsoft.com/en-us/library/office/cc841964.aspx
   * @readonly
   * @type {number}
   * @memberof PSTFolder
   */
  public get unreadCount(): number {
    return this.getIntItem(OutlookProperties.PR_CONTENT_UNREAD)
  }

  /**
   * Contains TRUE if a folder contains subfolders.
   * once again, read from the PST, use getSubFolderCount if you want to know
   * @readonly
   * @type {boolean}
   * @memberof PSTFolder
   */
  public get hasSubfolders(): boolean {
    return false
      || this.getBooleanItem(OutlookProperties.PR_SUBFOLDERS)
      || this.getIntItem(OutlookProperties.PR_SUBFOLDERS) != 0
  }

  /**
   * Contains a text string describing the type of a folder. Although this property is
   * generally ignored, versions of Microsoft® Exchange Server prior to Exchange Server
   * 2003 Mailbox Manager expect this property to be present.
   * https://msdn.microsoft.com/en-us/library/office/cc839839.aspx
   * @readonly
   * @type {string}
   * @memberof PSTFolder
   */
  public get containerClass(): string {
    return this.getStringItem(OutlookProperties.PR_CONTAINER_CLASS)
  }

  /**
   * Contains a bitmask of flags describing capabilities of an address book container.
   * https://msdn.microsoft.com/en-us/library/office/cc839610.aspx
   * @readonly
   * @type {number}
   * @memberof PSTFolder
   */
  public get containerFlags(): number {
    return this.getIntItem(OutlookProperties.PR_CONTAINER_FLAGS)
  }

  /**
   * Requests access to the user node of the internal PST structure.
   */
  public async requestAccessToUserNode(): Promise<PUNode | undefined> {
    return createPUNodeFrom(this._node);
  }

  /**
   * JSON stringify the object properties.
   * @returns {string}
   * @memberof PSTFolder
   */
  public toJSON(): any {
    const clone = Object.assign(
      {
        folderType: this.folderType,
        contentCount: this.contentCount,
        unreadCount: this.unreadCount,
        hasSubfolders: this.hasSubfolders,
        containerClass: this.containerClass,
        containerFlags: this.containerFlags,
        primaryNodeId: this.primaryNodeId,
      },
      this
    )
    return clone
  }
}
