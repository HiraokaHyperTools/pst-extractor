import type { IPSTMessageStore } from './IPSTMessageStore.js';
import type { PropertyFinder } from './PAUtil.js';
import type { PLNode } from './PLNode.js';
import type { PLSubNode } from './PLSubNode.js';
import { PSTFile } from './PSTFile.class.js';
import { PSTObject } from './PSTObject.class.js';
import { createPUNodeFrom, type PUNode } from './PUNode.js';
import type { RootProvider } from './RootProvider.js';

export class PSTMessageStore extends PSTObject implements IPSTMessageStore {
  /**
   * Creates an instance of PSTMessageStore.
   * Not much use other than to get the "name" of the PST file.
   * @internal
   * @param {PSTFile} rootProvider
   * @param {DescriptorIndexNode} descriptorIndexNode
   */
  constructor(
    rootProvider: RootProvider,
    node: PLNode,
    subNode: PLSubNode,
    propertyFinder: PropertyFinder
  ) {
    super(rootProvider, node, subNode, propertyFinder)
  }

  /**
   * Requests access to the user node of the internal PST structure.
   */
  public async requestAccessToUserNode(): Promise<PUNode | undefined> {
    return createPUNodeFrom(this._node);
  }
}
